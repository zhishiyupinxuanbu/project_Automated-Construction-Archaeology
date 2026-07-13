#!/usr/bin/env python3
import json
import os
import re
import subprocess
import tempfile
import time
from pathlib import Path

import requests


REFERENCE_LIBRARY_ROOT = Path(
    os.environ.get("HERITAGE_REFERENCE_LIBRARY", "/Users/drevan01/Desktop/文物影响评估与保护方案资料库")
)
ROOT = REFERENCE_LIBRARY_ROOT / "01_法规政策与标准/环境与施工标准资料库"
OUT = ROOT / "reference_report_scan"
OUT.mkdir(parents=True, exist_ok=True)
REPORT_ROOT = REFERENCE_LIBRARY_ROOT / "05_报告与方案样本/文物影响评估报告"

FILES = [
    REPORT_ROOT / "高速公路保护方案.pdf",
    REPORT_ROOT / "高速公路影响评估报告.pdf",
    REPORT_ROOT / "京新高速公路(G7)临河至白疙瘩段穿越居延遗址文物影响评估报告.pdf",
    REPORT_ROOT / "内蒙古呼伦贝尔市新建阿荣旗至莫旗铁路文物影响评估报告.pdf",
    REPORT_ROOT / "新建赤峰至京沈高铁喀左站铁路工程穿越燕北长城文物影响评估报告.pdf",
]

OCR_URL = "http://127.0.0.1:8000/ocr"
STD_TITLE = re.compile(r"《[^》]*(?:标准|规范|规程|导则|准则|指南|技术标准|环境质量标准|排放标准|卫生标准|质量标准|安全技术条件)[^》]*》")
STD_CODE = re.compile(
    r"(?<![A-Za-z])(?:GB/T|GBZ/T|GBZ|GB|HJ|JTG/T|JTG|JT/T|JTJ|TB/T|TB|CJJ/T|CJJ|DL/T|SL|NY/T|DZ/T|MT/T|DB\\d{0,2}/T|DB\\d{0,2})\\s*[A-Z]*\\s*[0-9][0-9A-Za-z./]*(?:\\s*[—－-]\\s*\\d{2,4})?",
    re.I,
)
KEYWORDS = ["技术规范", "标准", "规范", "导则", "噪声", "振动", "水土", "污水", "废水", "环境", "公路", "铁路"]


def _result_paths(pdf: Path) -> tuple[Path, Path, Path]:
    safe_stem = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", pdf.stem)
    return (
        OUT / f"{safe_stem}.pages.jsonl",
        OUT / f"{safe_stem}.txt",
        OUT / f"{safe_stem}.summary.json",
    )


def _existing_pages(jsonl_path: Path) -> set[int]:
    pages = set()
    if not jsonl_path.exists():
        return pages
    for line in jsonl_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        try:
            payload = json.loads(line)
        except json.JSONDecodeError:
            continue
        if "page" in payload:
            pages.add(int(payload["page"]))
    return pages


def _page_count(pdf: Path) -> int:
    output = subprocess.check_output(["pdfinfo", str(pdf)], text=True, stderr=subprocess.STDOUT)
    for line in output.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":", 1)[1].strip())
    raise RuntimeError(f"Cannot determine page count for {pdf}")


def _render_page(pdf: Path, page_number: int, temp_dir: Path) -> Path:
    prefix = temp_dir / f"page-{page_number:04d}"
    cmd = [
        "pdftoppm",
        "-png",
        "-r",
        "140",
        "-f",
        str(page_number),
        "-l",
        str(page_number),
        "-singlefile",
        str(pdf),
        str(prefix),
    ]
    subprocess.run(cmd, check=True)
    image_path = prefix.with_suffix(".png")
    if not image_path.exists():
        raise FileNotFoundError(image_path)
    return image_path


def ocr_pdf(pdf: Path, max_pages: int | None = None) -> dict:
    jsonl_path, text_path, summary_path = _result_paths(pdf)
    completed_pages = _existing_pages(jsonl_path)
    text_parts: list[str] = []
    page_records: list[dict] = []
    total_pages = _page_count(pdf)
    if max_pages is not None:
        total_pages = min(total_pages, max_pages)
    with tempfile.TemporaryDirectory(prefix="report_ocr_") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        for index in range(1, total_pages + 1):
            if index in completed_pages:
                print(f"{pdf.name}: page {index}/{total_pages} already done", flush=True)
                continue
            started = time.time()
            page = _render_page(pdf, index, temp_dir)
            with page.open("rb") as image:
                response = requests.post(OCR_URL, files={"file": (page.name, image, "image/png")}, timeout=600)
            response.raise_for_status()
            data = response.json()
            page_text = data.get("text", "")
            record = {"page": index, "text_chars": len(page_text), "text": page_text}
            with jsonl_path.open("a", encoding="utf-8") as output:
                output.write(json.dumps(record, ensure_ascii=False) + "\n")
            try:
                page.unlink()
            except FileNotFoundError:
                pass
            print(f"{pdf.name}: page {index}/{total_pages} chars={len(page_text)} seconds={time.time() - started:.1f}", flush=True)

    for line in jsonl_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        try:
            record = json.loads(line)
        except json.JSONDecodeError:
            continue
        page_records.append(record)
    page_records.sort(key=lambda item: item["page"])
    for record in page_records:
        text_parts.append(f"===== PAGE {record['page']} =====\n{record.get('text', '')}")
    text = "\n".join(text_parts)
    text_path.write_text(text, encoding="utf-8")
    titles = sorted(set(STD_TITLE.findall(text)))
    codes = sorted(set(re.sub(r"\s+", " ", match.group(0)).strip() for match in STD_CODE.finditer(text)))
    context = []
    for line in text.splitlines():
        clean = re.sub(r"\s+", " ", line.strip())
        if clean and (STD_TITLE.search(clean) or STD_CODE.search(clean) or any(keyword in clean for keyword in KEYWORDS)):
            context.append(clean)
    result = {
        "file": str(pdf),
        "text_chars": len(text),
        "titles": titles,
        "codes": codes,
        "context": context,
        "pages": page_records,
    }
    summary_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> None:
    print("OCR reference report standards job started", flush=True)
    results = []
    for pdf in FILES:
        print(f"START {pdf}", flush=True)
        results.append(ocr_pdf(pdf))
        print(f"DONE {pdf}", flush=True)
    result_path = OUT / "高速扫描报告_标准检索结果.json"
    result_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

    summary_lines = []
    for result in results:
        summary_lines.append(f"# {Path(result['file']).name}")
        summary_lines.append("")
        summary_lines.append("## 标准名称")
        summary_lines.extend(f"- {item}" for item in result["titles"] or ["未识别到"])
        summary_lines.append("")
        summary_lines.append("## 标准编号")
        summary_lines.extend(f"- {item}" for item in result["codes"] or ["未识别到"])
        summary_lines.append("")
        summary_lines.append("## 相关上下文")
        summary_lines.extend(f"- {item}" for item in result["context"][:120])
        summary_lines.append("")
    (OUT / "高速扫描报告_标准检索摘要.md").write_text("\n".join(summary_lines), encoding="utf-8")
    print(f"Saved: {result_path}", flush=True)


if __name__ == "__main__":
    main()
