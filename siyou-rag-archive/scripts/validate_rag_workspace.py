#!/usr/bin/env python3
"""Validate the standard 四有档案 RAG workspace structure."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


REQUIRED_DIRS = [
    "00_manifest",
    "01_pages",
    "02_ocr",
    "03_text_clean",
    "04_extract",
    "05_classify",
    "06_vector_db",
    "07_review_html",
    "08_exports",
    "09_logs",
]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workspace", help="Path to 00_RAG工作区")
    args = parser.parse_args()

    workspace = Path(args.workspace).expanduser().resolve()
    errors: list[str] = []

    if not workspace.exists():
        errors.append(f"Workspace does not exist: {workspace}")
    elif not workspace.is_dir():
        errors.append(f"Workspace is not a directory: {workspace}")

    for dirname in REQUIRED_DIRS:
        path = workspace / dirname
        if not path.is_dir():
            errors.append(f"Missing required directory: {dirname}")

    config_path = workspace / "rag_workspace.json"
    if not config_path.is_file():
        errors.append("Missing rag_workspace.json")
    else:
        try:
            config = json.loads(config_path.read_text(encoding="utf-8"))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Invalid rag_workspace.json: {exc}")
        else:
            if config.get("schema_version") != "siyou-rag-archive-v1.0":
                errors.append("Unexpected schema_version in rag_workspace.json")

    if errors:
        for error in errors:
            print(f"[ERROR] {error}")
        return 1

    print(f"[OK] Valid RAG workspace: {workspace}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
