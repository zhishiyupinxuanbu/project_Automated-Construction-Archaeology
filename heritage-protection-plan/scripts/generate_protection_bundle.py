#!/usr/bin/env python3
"""生成文物保护方案固定格式工作包。"""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成文物保护方案固定格式工作包")
    parser.add_argument("--project-name", "--项目名称", dest="project_name", required=True, help="项目名称")
    parser.add_argument("--output-dir", "--输出目录", dest="output_dir", required=True, help="输出目录")
    parser.add_argument("--overwrite", "--覆盖", dest="overwrite", action="store_true", help="允许覆盖已存在的同名文件")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    skill_dir = Path(__file__).resolve().parents[1]
    assets_dir = skill_dir / "assets"
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    today = dt.date.today().isoformat()
    templates = sorted(assets_dir.glob("*模板.md"))
    if not templates:
        raise SystemExit(f"未找到模板文件：{assets_dir}")

    for template_path in templates:
        target_path = output_dir / template_path.name.replace("模板", "")
        if target_path.exists() and not args.overwrite:
            raise SystemExit(f"文件已存在，未覆盖：{target_path}")
        content = template_path.read_text(encoding="utf-8")
        content = content.replace("{{项目名称}}", args.project_name)
        content = content.replace("{{生成日期}}", today)
        target_path.write_text(content, encoding="utf-8")

    print(f"已生成文物保护方案工作包：{output_dir}")
    print(f"项目名称：{args.project_name}")
    print(f"文件数量：{len(templates)}")


if __name__ == "__main__":
    main()
