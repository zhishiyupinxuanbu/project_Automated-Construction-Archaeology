#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build Obsidian-friendly Markdown index notes for the document library.

The script is intentionally read-mostly: it does not move source files and does
not OCR PDFs. It reuses cached text when available so the vault becomes easier
to search without turning every generation into a slow parsing job.
"""

from __future__ import annotations

import datetime as dt
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

import gongwen_agent as ga


ROOT = Path(__file__).resolve().parents[1]
INDEX_DIR = ROOT / "0.资料索引"
CARD_DIR = INDEX_DIR / "文件卡片"
TYPE_DIR = INDEX_DIR / "按文种"
TEXT_DIR = ROOT / ga.WORK_DIR_NAME / ga.TEXT_DIR_NAME
MAX_EXCERPT_CHARS = 900
TOP_LEVEL_PAGES = [
    "00-资料库总览.md",
    "01-按文种索引.md",
    "02-模板索引.md",
    "03-待补充文本清单.md",
]

EXTRA_BUSINESS_ALIASES = {
    "申请开展考古勘探工作请示": ["申请开展", "开展XX项目考古勘探工作", "开展考古勘探工作", "申请勘探"],
    "文物调查报告": ["文物调查的报告"],
}

VALID_BUSINESS_TYPES = set(ga.BUSINESS_ALIASES) | set(EXTRA_BUSINESS_ALIASES) | {"政策法规与规范", "其他"}


def md_escape(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ").strip()


def safe_note_name(name: str, fallback: str = "未命名") -> str:
    name = re.sub(r"[\\/:*?\"<>|#^[\\]]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return (name or fallback)[:80]


def tag_value(text: str) -> str:
    text = re.sub(r"\s+", "", text.strip())
    return re.sub(r"[#\\[\\](),，。；;:：/\\\\]+", "-", text) or "未识别"


def wiki_link(rel_path: str, label: str = "原文件") -> str:
    return f"[[{rel_path}|{label}]]"


def read_index_records() -> Dict[str, Dict[str, object]]:
    index_file = ROOT / ga.WORK_DIR_NAME / ga.INDEX_NAME
    if not index_file.exists():
        return {}
    try:
        data = json.loads(index_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    return {item.get("path", ""): item for item in data.get("records", []) if item.get("path")}


def infer_business(path: Path) -> str:
    parts = path.relative_to(ROOT).parts
    rel_text = " ".join(parts) + " " + path.stem
    if any("政策法规" in part for part in parts):
        return "政策法规与规范"
    for business_type, aliases in EXTRA_BUSINESS_ALIASES.items():
        if business_type in rel_text or any(alias in rel_text for alias in aliases):
            return business_type
    business_type = ga.infer_business_from_path(path)
    if business_type in {"公文参考", "公文模板", "政策法规与规范"}:
        return "政策法规与规范" if "政策法规" in rel_text else "其他"
    return business_type or "其他"


def cached_text_for(path: Path) -> Tuple[str, str]:
    text_file = TEXT_DIR / f"{ga.stable_id(path)}.txt"
    if text_file.exists():
        text = text_file.read_text(encoding="utf-8", errors="ignore")
        status = "已缓存文本" if len(text) >= ga.MIN_USEFUL_TEXT else "缓存文本较少"
        return text, status
    if path.suffix.lower() == ".docx":
        text, parse_status = ga.extract_docx_text(path)
        if text:
            return text, f"Word快速提取:{parse_status}"
        return "", f"Word提取失败:{parse_status}"
    return "", "未缓存文本，建议后续入库或OCR"


def first_lines(text: str, max_chars: int = MAX_EXCERPT_CHARS) -> str:
    text = ga.clean_text(text)
    if not text:
        return "暂无可用文本摘要。"
    return text[:max_chars] + ("…" if len(text) > max_chars else "")


def build_record(path: Path, indexed: Dict[str, Dict[str, object]]) -> Dict[str, object]:
    rel = str(path.relative_to(ROOT))
    text, text_status = cached_text_for(path)
    index_item = indexed.get(rel, {})
    title = ga.title_from_name(path)
    indexed_business_type = str(index_item.get("business_type") or "")
    business_type = indexed_business_type if indexed_business_type in VALID_BUSINESS_TYPES else infer_business(path)
    region = str(index_item.get("region") or ga.infer_region(rel))
    role = str(index_item.get("source_role") or ga.infer_source_role(path))
    is_template = bool(index_item.get("is_template", ga.is_template_file(path)))
    return {
        "id": ga.stable_id(path),
        "path": rel,
        "title": title,
        "business_type": business_type,
        "region": region,
        "role": role,
        "is_template": is_template,
        "extension": path.suffix.lower().lstrip("."),
        "project_name_guess": ga.guess_project_name(title),
        "issuing_org_guess": ga.guess_issuing_org(title),
        "date_guess": str(index_item.get("date_guess") or ga.guess_date(text, path)),
        "text_status": text_status,
        "text_chars": len(text),
        "excerpt": first_lines(text),
    }


def write_card(record: Dict[str, object]) -> Path:
    note_name = safe_note_name(f"{record['id']}-{record['title']}")
    card_path = CARD_DIR / f"{note_name}.md"
    tags = [
        f"#文种/{tag_value(str(record['business_type']))}",
        f"#角色/{tag_value(str(record['role']))}",
    ]
    if record["region"]:
        tags.append(f"#地区/{tag_value(str(record['region']))}")
    if record["is_template"]:
        tags.append("#资料类型/模板")

    frontmatter = [
        "---",
        f"source_path: \"{record['path']}\"",
        f"business_type: \"{record['business_type']}\"",
        f"region: \"{record['region']}\"",
        f"role: \"{record['role']}\"",
        f"is_template: {str(bool(record['is_template'])).lower()}",
        f"extension: \"{record['extension']}\"",
        f"text_status: \"{record['text_status']}\"",
        f"text_chars: {record['text_chars']}",
        f"date_guess: \"{record['date_guess']}\"",
        "---",
        "",
    ]
    content = "\n".join(
        frontmatter
        + [
            f"# {record['title']}",
            "",
            f"- 原文件：{wiki_link(str(record['path']), Path(str(record['path'])).name)}",
            f"- 文种：{record['business_type']}",
            f"- 地区：{record['region'] or '未识别'}",
            f"- 角色：{record['role']}",
            f"- 模板：{'是' if record['is_template'] else '否'}",
            f"- 文本状态：{record['text_status']}（{record['text_chars']} 字）",
            f"- 推测项目：{record['project_name_guess'] or '未识别'}",
            f"- 推测发文单位：{record['issuing_org_guess'] or '未识别'}",
            f"- 标签：{' '.join(tags)}",
            "",
            "## 摘要",
            "",
            str(record["excerpt"]),
            "",
            "## 生成提示",
            "",
            "- 生成公文时优先核对原文件中的项目名称、发文主体、受文单位、面积口径、附件名称和日期。",
            "- 如文本状态提示未缓存或文本较少，涉及事实提取时应打开原文件复核。",
            "",
        ]
    )
    card_path.write_text(content, encoding="utf-8")
    return card_path


def clean_generated_pages() -> None:
    for filename in TOP_LEVEL_PAGES:
        path = INDEX_DIR / filename
        if path.exists():
            path.unlink()
    for folder in (CARD_DIR, TYPE_DIR):
        if not folder.exists():
            continue
        for path in folder.glob("*.md"):
            path.unlink()


def write_type_page(business_type: str, records: List[Dict[str, object]], card_paths: Dict[str, Path]) -> Path:
    page = TYPE_DIR / f"{safe_note_name(business_type)}.md"
    rows = []
    for record in sorted(records, key=lambda item: (not item["is_template"], str(item["region"]), str(item["title"]))):
        card_rel = str(card_paths[str(record["id"])].relative_to(ROOT))
        rows.append(
            "| "
            + " | ".join(
                [
                    wiki_link(card_rel, md_escape(str(record["title"]))),
                    md_escape(str(record["region"]) or "未识别"),
                    md_escape(str(record["role"])),
                    "是" if record["is_template"] else "否",
                    wiki_link(str(record["path"]), "原文件"),
                    md_escape(str(record["text_status"])),
                ]
            )
            + " |"
        )
    content = "\n".join(
        [
            f"# {business_type}",
            "",
            f"- 文件数：{len(records)}",
            f"- 模板数：{sum(1 for item in records if item['is_template'])}",
            f"- 返回：[[0.资料索引/00-资料库总览|资料库总览]]",
            "",
            "| 文件卡片 | 地区 | 角色 | 模板 | 原文件 | 文本状态 |",
            "| --- | --- | --- | --- | --- | --- |",
            *rows,
            "",
        ]
    )
    page.write_text(content, encoding="utf-8")
    return page


def write_overview(records: List[Dict[str, object]], type_pages: Dict[str, Path]) -> None:
    by_type = Counter(str(record["business_type"]) for record in records)
    by_region = Counter(str(record["region"]) or "未识别" for record in records)
    by_status = Counter(str(record["text_status"]).split(":", 1)[0] for record in records)
    templates = [record for record in records if record["is_template"]]
    needs_text = [record for record in records if "未缓存" in str(record["text_status"]) or int(record["text_chars"]) < ga.MIN_USEFUL_TEXT]

    type_rows = []
    for business_type, count in sorted(by_type.items()):
        page_rel = str(type_pages[business_type].relative_to(ROOT))
        type_templates = sum(1 for record in records if record["business_type"] == business_type and record["is_template"])
        type_rows.append(f"| {wiki_link(page_rel, business_type)} | {count} | {type_templates} |")

    region_rows = [f"| {md_escape(region)} | {count} |" for region, count in sorted(by_region.items())]
    status_rows = [f"| {md_escape(status)} | {count} |" for status, count in sorted(by_status.items())]

    overview = "\n".join(
        [
            "# 资料库总览",
            "",
            f"- 生成时间：{dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"- 原始 Word/PDF 文件数：{len(records)}",
            f"- 模板文件数：{len(templates)}",
            f"- 需要后续补充文本或 OCR 的文件数：{len(needs_text)}",
            "",
            "## 快速入口",
            "",
            "- [[0.资料索引/01-按文种索引|按文种索引]]",
            "- [[0.资料索引/02-模板索引|模板索引]]",
            "- [[0.资料索引/03-待补充文本清单|待补充文本清单]]",
            "",
            "## 文种统计",
            "",
            "| 文种 | 文件数 | 模板数 |",
            "| --- | ---: | ---: |",
            *type_rows,
            "",
            "## 地区统计",
            "",
            "| 地区 | 文件数 |",
            "| --- | ---: |",
            *region_rows,
            "",
            "## 文本状态",
            "",
            "| 状态 | 文件数 |",
            "| --- | ---: |",
            *status_rows,
            "",
        ]
    )
    (INDEX_DIR / "00-资料库总览.md").write_text(overview, encoding="utf-8")

    by_type_index = "\n".join(
        [
            "# 按文种索引",
            "",
            *[f"- {wiki_link(str(type_pages[business_type].relative_to(ROOT)), business_type)}（{count}）" for business_type, count in sorted(by_type.items())],
            "",
        ]
    )
    (INDEX_DIR / "01-按文种索引.md").write_text(by_type_index, encoding="utf-8")

    template_rows = []
    for record in sorted(templates, key=lambda item: (str(item["business_type"]), str(item["title"]))):
        template_rows.append(
            f"| {wiki_link(str(record['path']), md_escape(str(record['title'])))} | {md_escape(str(record['business_type']))} | {md_escape(str(record['region']) or '未识别')} | {md_escape(str(record['text_status']))} |"
        )
    (INDEX_DIR / "02-模板索引.md").write_text(
        "\n".join(
            [
                "# 模板索引",
                "",
                "| 模板文件 | 文种 | 地区 | 文本状态 |",
                "| --- | --- | --- | --- |",
                *template_rows,
                "",
            ]
        ),
        encoding="utf-8",
    )

    needs_rows = []
    for record in sorted(needs_text, key=lambda item: (str(item["business_type"]), str(item["path"]))):
        needs_rows.append(
            f"| {wiki_link(str(record['path']), md_escape(str(record['title'])))} | {md_escape(str(record['business_type']))} | {md_escape(str(record['region']) or '未识别')} | {md_escape(str(record['text_status']))} |"
        )
    (INDEX_DIR / "03-待补充文本清单.md").write_text(
        "\n".join(
            [
                "# 待补充文本清单",
                "",
                "这些文件目前没有足够的可检索文本。后续可通过轻量 OCR 或人工摘录补强，避免生成公文时临时读取 PDF。",
                "",
                "| 原文件 | 文种 | 地区 | 文本状态 |",
                "| --- | --- | --- | --- |",
                *needs_rows,
                "",
            ]
        ),
        encoding="utf-8",
    )


def main() -> int:
    INDEX_DIR.mkdir(parents=True, exist_ok=True)
    CARD_DIR.mkdir(parents=True, exist_ok=True)
    TYPE_DIR.mkdir(parents=True, exist_ok=True)
    clean_generated_pages()

    indexed = read_index_records()
    source_files = list(ga.iter_source_files(ROOT))
    records = [build_record(path, indexed) for path in source_files]

    card_paths = {}
    for record in records:
        card_paths[str(record["id"])] = write_card(record)

    by_type: Dict[str, List[Dict[str, object]]] = defaultdict(list)
    for record in records:
        by_type[str(record["business_type"])].append(record)

    type_pages = {}
    for business_type, items in by_type.items():
        type_pages[business_type] = write_type_page(business_type, items, card_paths)

    write_overview(records, type_pages)
    print(f"已生成 Obsidian 索引：{INDEX_DIR.relative_to(ROOT)}")
    print(f"文件卡片：{len(records)}")
    print(f"文种页：{len(type_pages)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
