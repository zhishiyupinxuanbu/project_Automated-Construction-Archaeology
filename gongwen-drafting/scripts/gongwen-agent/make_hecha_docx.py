#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成固定版式的文物保护许可/核查请示 Word 文档。"""

from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path

from gongwen_format import HechaRequestData, create_hecha_docx, write_audit_report


ROOT = Path(__file__).resolve().parents[1]


def default_output_dir() -> Path:
    desktop = Path.home() / "Desktop"
    if desktop.exists():
        return desktop
    return ROOT / ".gongwen_agent" / "outputs"


def today_chinese() -> str:
    today = dt.date.today()
    return f"{today.year}年{today.month}月{today.day}日"


def load_data(path: Path) -> HechaRequestData:
    raw = json.loads(path.read_text(encoding="utf-8"))
    if not raw.get("date_text"):
        raw["date_text"] = today_chinese()
    return HechaRequestData(
        issuing_org=raw["issuing_org"],
        recipient_org=raw["recipient_org"],
        project_name=raw["project_name"],
        location=raw.get("location", ""),
        scale=raw.get("scale", ""),
        project_area=raw.get("project_area", ""),
        approval_file=raw.get("approval_file", "审批机关立项选址核准文件"),
        basis_sentence=raw.get(
            "basis_sentence",
            "按照各级文物管理要求及《内蒙古自治区文物局关于做好基本建设用地考古工作的通知》（内文物发〔2025〕6号）文件要求",
        ),
        date_text=raw["date_text"],
        doc_number=raw.get("doc_number", ""),
        signer=raw.get("signer", ""),
        add_red_header=raw.get("add_red_header", True),
        body_paragraphs=raw.get("body_paragraphs", []),
        attachments=raw.get("attachments", []),
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成固定版式核查请示 DOCX")
    parser.add_argument("input_json", help="项目要素 JSON")
    parser.add_argument("--out", help="输出 DOCX 路径")
    parser.add_argument("--audit", action="store_true", help="生成后同步输出校验报告")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input_json)
    data = load_data(input_path)
    if args.out:
        output_path = Path(args.out)
    else:
        output_path = default_output_dir() / f"{data.issuing_org}关于办理{data.project_name}用地范围内文物保护许可的请示.docx"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    create_hecha_docx(data, output_path)
    print(output_path)
    if args.audit:
        report = output_path.with_suffix(".版式校验.md")
        write_audit_report(output_path, report)
        print(report)


if __name__ == "__main__":
    main()
