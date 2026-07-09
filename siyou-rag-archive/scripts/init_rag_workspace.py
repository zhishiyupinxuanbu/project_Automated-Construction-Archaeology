#!/usr/bin/env python3
"""Create the standard 四有档案 RAG workspace under a project folder."""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path


DIRS = [
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
    parser.add_argument("project_root", help="Project folder that contains original materials")
    parser.add_argument("--name", default="00_RAG工作区", help="Workspace directory name")
    args = parser.parse_args()

    project_root = Path(args.project_root).expanduser().resolve()
    workspace = project_root / args.name
    workspace.mkdir(parents=True, exist_ok=True)
    for dirname in DIRS:
        (workspace / dirname).mkdir(exist_ok=True)

    config = {
        "schema_version": "siyou-rag-archive-v1.0",
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "project_root": str(project_root),
        "workspace": str(workspace),
        "ocr_default": "/Users/drevan01/Desktop/OCR",
        "required_dirs": DIRS,
    }
    (workspace / "rag_workspace.json").write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    log = workspace / "09_logs" / f"init-{datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
    log.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    print(workspace)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
