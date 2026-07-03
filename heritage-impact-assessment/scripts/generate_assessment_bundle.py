#!/usr/bin/env python3
"""Generate a v0.1 heritage impact assessment working bundle."""

from __future__ import annotations

import argparse
from pathlib import Path


TEMPLATES = [
    "01-资料缺项清单模板.md",
    "02-项目信息提取表模板.md",
    "03-影响识别矩阵模板.md",
    "04-保护措施建议模板.md",
    "05-报告草稿大纲模板.md",
    "06-专家审核清单模板.md",
    "07-文物影响评估报告正文模板.md",
    "08-章节证据索引模板.md",
    "09-docx成稿自检清单模板.md",
    "10-项目资料可用信息清单模板.md",
    "11-信息入文章节映射表模板.md",
    "12-遗产价值分级量表模板.md",
    "13-第五章影响评估表模板.md",
    "14-附表综合评估大表模板.md",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成文物影响评估 v0.1 工作包")
    parser.add_argument("--项目名称", dest="project_name", required=True)
    parser.add_argument("--输出目录", dest="output_dir", required=True)
    parser.add_argument("--覆盖", dest="overwrite", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    skill_dir = Path(__file__).resolve().parents[1]
    assets_dir = skill_dir / "assets"
    output_dir = Path(args.output_dir).expanduser().resolve()

    if output_dir.exists() and any(output_dir.iterdir()) and not args.overwrite:
        raise SystemExit(f"输出目录非空，如需覆盖请添加 --覆盖：{output_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)

    for template_name in TEMPLATES:
        source = assets_dir / template_name
        target_name = template_name.replace("模板", "")
        target = output_dir / target_name
        text = source.read_text(encoding="utf-8").replace("{{项目名称}}", args.project_name)
        target.write_text(text, encoding="utf-8")

    index = output_dir / "00-工作包说明.md"
    index.write_text(
        "\n".join(
            [
                f"# {args.project_name} 文物影响评估 v0.1 工作包",
                "",
                "本工作包由 heritage-impact-assessment v0.1 生成，用于资料检查、影响识别、措施建议和报告草稿协作。",
                "",
                "## 文件清单",
                "",
                *[f"- {name.replace('模板', '')}" for name in TEMPLATES],
                "",
                "## 使用边界",
                "",
                "- 项目事实必须来自资料或人工确认。",
                "- 影响等级、审批结论和关键技术判断需专业负责人审核。",
                "- 正文中的【固定表格入口】必须替换为已填写的实际表格后，再生成 Word 成稿。",
            ]
        ),
        encoding="utf-8",
    )

    print(f"已生成文物影响评估工作包：{output_dir}")


if __name__ == "__main__":
    main()
