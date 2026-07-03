#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Preflight gate for region + document-type template routing.

This script intentionally reads the current project or packaged skill indexes
before any memory-derived rule can influence drafting. It fails when a planned
generic route would skip an indexed local template.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path


DOC_TYPE_ALIASES = {
    "文物保护许可申请": "文物保护许可申请",
    "文物核查请示": "文物保护许可申请",
    "核查请示": "文物保护许可申请",
    "文物手续": "文物保护许可申请",
    "文物保护许可": "文物保护许可申请",
    "许可申请": "文物保护许可申请",
    "文物调查事宜请示": "文物保护许可申请",
    "办理文物调查事宜": "文物保护许可申请",
    "申请办理文物调查": "文物保护许可申请",
    "查询用地是否涉及文物": "文物保护许可申请",
    "用地是否涉及文物": "文物保护许可申请",
    "临时用地涉文查询": "文物保护许可申请",
    "勘探验收请示": "勘探验收请示",
    "考古勘探验收": "勘探验收请示",
    "勘探计划备案请示": "勘探计划备案请示",
    "勘探报告备案请示": "勘探报告备案请示",
    "申请开展考古勘探工作请示": "申请开展考古勘探工作请示",
    "开展勘探请示": "申请开展考古勘探工作请示",
    "考古勘探工作支持请示": "考古勘探工作支持请示",
    "协助勘探请示": "考古勘探工作支持请示",
    "开工请示": "开工请示",
    "文物调查报告": "文物调查报告",
}

GENERIC_REGION_MARKERS = {"", "未识别", "通用", "无"}
REGION_ALIASES = {
    "锡盟": "锡林郭勒盟",
    "锡林郭勒": "锡林郭勒盟",
    "锡林浩特": "锡林浩特市",
    "锡市": "锡林浩特市",
    "胜利矿区": "锡林浩特市",
    "胜利煤田": "锡林浩特市",
    "胜利矿田": "锡林浩特市",
    "西乌": "西乌珠穆沁旗",
    "西乌旗": "西乌珠穆沁旗",
}


@dataclass(frozen=True)
class TemplateRow:
    name: str
    source: str
    doc_type: str
    region: str
    status: str


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def normalize_doc_type(value: str) -> str:
    raw = normalize_text(value)
    return DOC_TYPE_ALIASES.get(raw, raw)


def normalize_region(value: str) -> str:
    raw = normalize_text(value).replace("地区", "")
    return REGION_ALIASES.get(raw, raw)


def find_root(start: Path) -> tuple[Path, str]:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists() and (candidate / "0.资料索引/02-模板索引.md").exists():
            return candidate, "workspace"
        if (candidate / "SKILL.md").exists() and (
            candidate / "references/knowledge-index/02-模板索引.md"
        ).exists():
            return candidate, "skill"
    raise SystemExit("ERROR: run inside 公文撰写资料库 or the gongwen-drafting skill.")


def template_index_path(root: Path, mode: str) -> Path:
    if mode == "workspace":
        return root / "0.资料索引/02-模板索引.md"
    return root / "references/knowledge-index/02-模板索引.md"


def extract_link(cell: str) -> tuple[str, str]:
    match = re.search(r"\[\[([^|\]]+)(?:\|([^\]]+))?\]\]", cell)
    if not match:
        text = cell.strip()
        return text, text
    source = match.group(1).strip()
    name = (match.group(2) or source).strip()
    return name, source


def split_markdown_row(line: str) -> list[str]:
    cells: list[str] = []
    current: list[str] = []
    in_wikilink = False
    chars = line.strip()
    index = 0
    if chars.startswith("|"):
        chars = chars[1:]
    if chars.endswith("|"):
        chars = chars[:-1]
    while index < len(chars):
        if chars.startswith("[[", index):
            in_wikilink = True
            current.append("[[")
            index += 2
            continue
        if chars.startswith("]]", index):
            in_wikilink = False
            current.append("]]")
            index += 2
            continue
        char = chars[index]
        if char == "|" and not in_wikilink:
            cells.append("".join(current).strip())
            current = []
        else:
            current.append(char)
        index += 1
    cells.append("".join(current).strip())
    return cells


def parse_template_index(path: Path) -> list[TemplateRow]:
    rows: list[TemplateRow] = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if not line.startswith("|") or "---" in line:
            continue
        cells = split_markdown_row(line)
        if len(cells) < 4 or cells[1] == "文种":
            continue
        name, source = extract_link(cells[0])
        rows.append(
            TemplateRow(
                name=name,
                source=source,
                doc_type=normalize_doc_type(cells[1]),
                region=normalize_region(cells[2]),
                status=cells[3].strip(),
            )
        )
    return rows


def is_region_match(row_region: str, requested_region: str) -> bool:
    if row_region in GENERIC_REGION_MARKERS:
        return False
    if row_region == requested_region:
        return True
    return row_region in requested_region or requested_region in row_region


def resolve_route(rows: list[TemplateRow], region: str, doc_type: str) -> dict:
    requested_region = normalize_region(region)
    requested_doc_type = normalize_doc_type(doc_type)
    doc_rows = [row for row in rows if row.doc_type == requested_doc_type]
    local_rows = [row for row in doc_rows if is_region_match(row.region, requested_region)]
    generic_rows = [row for row in doc_rows if row.region in GENERIC_REGION_MARKERS]

    if local_rows:
        route = "local"
        chosen = local_rows[0]
    elif generic_rows:
        route = "generic"
        chosen = generic_rows[0]
    else:
        route = "unknown"
        chosen = None

    return {
        "route": route,
        "region": requested_region,
        "doc_type": requested_doc_type,
        "template": None
        if chosen is None
        else {
            "name": chosen.name,
            "source": chosen.source,
            "region": chosen.region,
            "status": chosen.status,
        },
        "local_candidates": [
            {
                "name": row.name,
                "source": row.source,
                "region": row.region,
                "status": row.status,
            }
            for row in local_rows
        ],
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="检查地区+文种是否必须走地方模板")
    parser.add_argument("--region", required=True, help="地区，如：库伦旗、乌兰察布、伊金霍洛旗")
    parser.add_argument("--doc-type", required=True, help="文种，如：文物核查请示、文物保护许可申请")
    parser.add_argument(
        "--planned-route",
        choices=("local", "generic", "unknown"),
        help="当前准备采用的路线；若与索引冲突则报错",
    )
    parser.add_argument(
        "--planned-template",
        help="当前准备采用的模板名或路径；若地方模板存在但计划模板不匹配则报错",
    )
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--root", type=Path, help="资料库根目录或 gongwen-drafting skill 根目录")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root, mode = find_root(args.root or Path.cwd())
    rows = parse_template_index(template_index_path(root, mode))
    result = resolve_route(rows, args.region, args.doc_type)

    errors: list[str] = []
    if args.planned_route and args.planned_route != result["route"]:
        if result["route"] == "local" and args.planned_route == "generic":
            errors.append(
                f"地区模板门禁失败：{result['region']} + {result['doc_type']} 已有地方模板，不能走通用模板。"
            )
        else:
            errors.append(f"计划路线 {args.planned_route} 与索引路线 {result['route']} 不一致。")

    if result["route"] == "local" and args.planned_template:
        planned = normalize_text(args.planned_template)
        candidates = result["local_candidates"]
        if not any(planned in normalize_text(item["name"]) or planned in normalize_text(item["source"]) for item in candidates):
            errors.append(
                "地区模板门禁失败：计划模板未命中地方模板候选，候选为 "
                + "；".join(item["name"] for item in candidates)
            )

    payload = {"mode": mode, "root": str(root), **result, "errors": errors}
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(f"mode: {mode}")
        print(f"region: {result['region']}")
        print(f"doc_type: {result['doc_type']}")
        print(f"route: {result['route']}")
        if result["template"]:
            print(f"template: {result['template']['name']} ({result['template']['source']})")
        if result["local_candidates"]:
            print("local_candidates:")
            for item in result["local_candidates"]:
                print(f"  - {item['name']} | {item['source']} | {item['status']}")
        if errors:
            print("errors:", file=sys.stderr)
            for error in errors:
                print(f"  - {error}", file=sys.stderr)

    return 2 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
