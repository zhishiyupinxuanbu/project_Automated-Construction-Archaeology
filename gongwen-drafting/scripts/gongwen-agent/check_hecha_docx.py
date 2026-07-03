#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验核查请示 Word 文档是否符合项目固定版式。"""

from __future__ import annotations

import argparse
from pathlib import Path

from gongwen_format import audit_hecha_docx, write_audit_report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="校验固定版式核查请示 DOCX")
    parser.add_argument("docx", help="待校验 DOCX")
    parser.add_argument("--report", help="可选：输出 Markdown 校验报告")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    docx_path = Path(args.docx)
    errors = audit_hecha_docx(docx_path)
    if args.report:
        write_audit_report(docx_path, Path(args.report))
    if errors:
        print("校验未通过：")
        for item in errors:
            print(f"- {item}")
        raise SystemExit(1)
    print("校验通过")


if __name__ == "__main__":
    main()
