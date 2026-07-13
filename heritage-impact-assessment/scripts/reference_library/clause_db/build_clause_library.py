#!/usr/bin/env python3
"""Build the external clause library for heritage impact assessment."""

from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REFERENCE_LIBRARY_ROOT = Path(
    os.environ.get("HERITAGE_REFERENCE_LIBRARY", "/Users/drevan01/Desktop/文物影响评估与保护方案资料库")
)
STANDARD_LIBRARY_ROOT = REFERENCE_LIBRARY_ROOT / "01_法规政策与标准"
LIB_ROOT = STANDARD_LIBRARY_ROOT / "法规条文库"
STANDARDS_ROOT = STANDARD_LIBRARY_ROOT / "环境与施工标准资料库"
STANDARDS_INDEX = STANDARDS_ROOT / "standards_index.json"

YSD_ASSET_CANDIDATES = [
    Path(__file__).resolve().parents[3] / "assets/17-元上都遗址保护管理规划关键条文.md",
    Path("/Users/drevan01/.codex/skills/heritage-impact-assessment/assets/17-元上都遗址保护管理规划关键条文.md"),
]


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")


def split_list(value: str) -> list[str]:
    value = value.strip()
    if not value:
        return []
    parts = re.split(r"[、,，;；]\s*", value)
    return [p.strip() for p in parts if p.strip()]


def parse_backtick_tags(value: str) -> list[str]:
    tags = re.findall(r"`([^`]+)`", value)
    if tags:
        return [t.strip() for t in tags if t.strip()]
    return split_list(value)


def find_ysd_asset() -> Path | None:
    for path in YSD_ASSET_CANDIDATES:
        if path.exists():
            return path
    return None


def parse_ysd_asset(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    text = path.read_text(encoding="utf-8")
    source = {
        "source_id": "PLAN_YSD_PROTECTION_MANAGEMENT_202509",
        "source_title": "元上都遗址保护管理规划修编·规划文本",
        "source_type": "protection_plan",
        "source_code": "",
        "source_version": "2025-09",
        "source_status": "current",
        "local_path": str(path),
        "source_note": "由既有 skill 资产导入；条文来源为第33条遗产区划管理规定。",
        "content_sha256": sha256_text(text),
        "generated_at": now_iso(),
    }

    heading_pattern = re.compile(r"^(#{2,4})\s+(.+?)\s*$", re.M)
    headings = list(heading_pattern.finditer(text))
    matches = [m for m in headings if m.group(1) == "####" and re.match(r"YSD-[A-Z0-9-]+$", m.group(2).strip())]
    clauses: list[dict[str, Any]] = []

    for i, match in enumerate(matches):
        clause_id = match.group(2).strip()
        start = match.end()
        end = len(text)
        for heading in headings:
            if heading.start() > match.start():
                end = heading.start()
                break
        block = text[start:end].strip()
        lines = block.splitlines()

        meta: dict[str, str] = {}
        body_lines: list[str] = []
        in_body = False
        for raw in lines:
            line = raw.strip()
            if not line:
                if body_lines:
                    body_lines.append("")
                continue
            m = re.match(r"^-\s*([a-zA-Z_]+):\s*(.*)$", line)
            if m and not in_body:
                meta[m.group(1)] = m.group(2).strip()
            else:
                in_body = True
                body_lines.append(line)

        tags = parse_backtick_tags(meta.get("tags", ""))
        clause_text = "\n".join(body_lines).strip()
        if not clause_text:
            continue

        clauses.append(
            {
                "clause_id": clause_id,
                "source_id": source["source_id"],
                "source_title": source["source_title"],
                "source_type": source["source_type"],
                "source_code": source["source_code"],
                "source_version": source["source_version"],
                "source_status": source["source_status"],
                "granularity": "clause",
                "clause_location": meta.get("location", ""),
                "clause_text": clause_text,
                "control_object": "",
                "control_value": "",
                "control_unit": "",
                "applies_to": meta.get("applies_to", ""),
                "zones": [t for t in tags if t.startswith("zone:")],
                "topics": [t for t in tags if t.startswith("topic:")],
                "trigger_keywords": split_list(meta.get("trigger_keywords", "")),
                "project_types": [],
                "heritage_types": ["world_heritage"],
                "impact_factors": [],
                "local_path": str(path),
                "source_url": "",
                "notes": "从元上都遗址保护管理规划关键条文资产导入。",
            }
        )

    return [source], clauses


def impact_control_object(entry: dict[str, Any]) -> str:
    tags = entry.get("factor_tags") or []
    title = entry.get("title", "")
    if "工程设计" in tags or "公路工程" in tags or "农村公路" in tags:
        return "公路工程设计、技术等级、设计速度、交通量和横断面"
    if "噪声" in tags or "声环境" in title:
        return "噪声、声环境"
    if "振动" in tags:
        return "振动"
    if "废气" in tags or "大气" in title:
        return "废气、大气环境"
    if "废水" in tags or "水" in title:
        return "废水、水环境"
    if "危险货物运输" in tags or "危险货物" in title:
        return "危险货物运输车辆"
    return "环境与施工影响因子"


def standard_clause_id(entry: dict[str, Any]) -> str:
    raw = entry.get("id") or entry.get("code") or entry.get("title")
    return "STD_" + re.sub(r"[^A-Za-z0-9]+", "_", raw).strip("_").upper() + "_GENERAL"


def parse_standards_index(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    if not path.exists():
        return [], []
    data = read_json(path)
    entries = data.get("entries", []) if isinstance(data, dict) else data
    sources: list[dict[str, Any]] = []
    clauses: list[dict[str, Any]] = []

    for entry in entries:
        source_id = "STANDARD_" + re.sub(r"[^A-Za-z0-9]+", "_", entry.get("id", entry.get("code", ""))).strip("_").upper()
        title = entry.get("title", "")
        code = entry.get("code", "")
        source_title = f"{title}（{code}）" if code else title
        status = "current" if entry.get("current", True) else "replaced"
        local_pdf = str((STANDARDS_ROOT / entry["pdf"]).resolve()) if entry.get("pdf") else ""
        local_text = str((STANDARDS_ROOT / entry["text"]).resolve()) if entry.get("text") else ""

        sources.append(
            {
                "source_id": source_id,
                "source_title": source_title,
                "source_type": "standard",
                "source_code": code,
                "source_version": entry.get("effective_date", ""),
                "source_status": status,
                "local_path": local_pdf,
                "text_path": local_text,
                "source_url": entry.get("source_url", ""),
                "source_page": entry.get("source_page", ""),
                "source_tier": entry.get("source_tier", ""),
                "replaces": entry.get("replaces", []),
                "ocr_required": entry.get("ocr_required", False),
                "text_chars": entry.get("text_chars", 0),
                "generated_at": now_iso(),
            }
        )

        factors = entry.get("factor_tags", [])
        clauses.append(
            {
                "clause_id": standard_clause_id(entry),
                "source_id": source_id,
                "source_title": source_title,
                "source_type": "standard",
                "source_code": code,
                "source_version": entry.get("effective_date", ""),
                "source_status": status,
                "granularity": "source_level",
                "clause_location": "全文/适用范围（待按具体条款、表格或限值细分）",
                "clause_text": f"项目涉及{ '、'.join(factors) if factors else '相关影响因子' }时，应调取并对照《{title}》（{code}）现行文本中的适用范围、控制要求、限值、监测、评价方法或工程技术指标。",
                "control_object": impact_control_object(entry),
                "control_value": "",
                "control_unit": "",
                "applies_to": "涉及该环境或施工影响因子的建设项目",
                "zones": [],
                "topics": entry.get("topics", []),
                "trigger_keywords": list(dict.fromkeys([*factors, *(entry.get("trigger_keywords") or []), title, code])),
                "project_types": entry.get("project_types", []),
                "heritage_types": entry.get("heritage_types", []),
                "impact_factors": factors,
                "local_path": local_pdf,
                "text_path": local_text,
                "source_url": entry.get("source_url", ""),
                "notes": "标准已纳入本地资料库；本条为 source-level 候选，后续可继续拆分为具体条款和限值表。",
            }
        )

    return sources, clauses


def build_tag_index(clauses: list[dict[str, Any]]) -> dict[str, list[str]]:
    index: dict[str, list[str]] = {}
    tag_fields = ["zones", "topics", "impact_factors", "project_types", "heritage_types", "trigger_keywords"]
    for clause in clauses:
        for field in tag_fields:
            for tag in clause.get(field, []) or []:
                index.setdefault(tag, []).append(clause["clause_id"])
    return {k: sorted(set(v)) for k, v in sorted(index.items())}


def main() -> None:
    sources: list[dict[str, Any]] = []
    clauses: list[dict[str, Any]] = []

    ysd_asset = find_ysd_asset()
    if ysd_asset:
        ysd_sources, ysd_clauses = parse_ysd_asset(ysd_asset)
        sources.extend(ysd_sources)
        clauses.extend(ysd_clauses)

    std_sources, std_clauses = parse_standards_index(STANDARDS_INDEX)
    sources.extend(std_sources)
    clauses.extend(std_clauses)

    write_jsonl(LIB_ROOT / "sources.jsonl", sources)
    write_jsonl(LIB_ROOT / "clause_library.jsonl", clauses)
    (LIB_ROOT / "tag_index.json").write_text(
        json.dumps(build_tag_index(clauses), ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )

    manifest = {
        "library_name": "法规条文库",
        "generated_at": now_iso(),
        "source_count": len(sources),
        "clause_count": len(clauses),
        "clause_granularity": {
            "clause": sum(1 for c in clauses if c.get("granularity") == "clause"),
            "source_level": sum(1 for c in clauses if c.get("granularity") == "source_level"),
        },
    }
    (LIB_ROOT / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    print(json.dumps(manifest, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
