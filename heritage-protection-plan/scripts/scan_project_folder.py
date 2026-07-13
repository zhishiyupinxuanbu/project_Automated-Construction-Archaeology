#!/usr/bin/env python3
"""扫描文物保护方案项目资料夹，输出轻量资料盘点。"""

from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path


CATEGORIES: list[tuple[str, tuple[str, ...]]] = [
    ("文物保护方案", ("文物保护方案", "保护实施方案", "保护措施")),
    ("文物影响评估", ("文物影响评估", "影响评估报告", "文评")),
    ("专家意见", ("专家意见", "评审", "修改内容", "审查意见")),
    ("文物调查勘探", ("调查报告", "勘探报告", "考古", "验收意见")),
    ("工程资料", ("建设方案", "工程方案", "平面图", "宗地图", "坐标", "红线", "kml", "dwg")),
    ("审批批复", ("批复", "复函", "请示", "函", "许可")),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="扫描项目资料夹并按保护方案资料类型归类")
    parser.add_argument("folder", help="项目资料夹路径")
    parser.add_argument("--output", "-o", help="输出 Markdown 文件路径；默认打印到终端")
    parser.add_argument("--limit", type=int, default=300, help="明细最多列出多少条，默认 300")
    parser.add_argument("--all", action="store_true", help="列出全部文件明细")
    return parser.parse_args()


def classify(path: Path) -> str:
    name = path.name.lower()
    for category, keywords in CATEGORIES:
        if any(keyword.lower() in name for keyword in keywords):
            return category
    return "其他资料"


def main() -> None:
    args = parse_args()
    root = Path(args.folder).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        raise SystemExit(f"资料夹不存在或不是目录：{root}")

    rows = []
    for path in sorted(p for p in root.rglob("*") if p.is_file() and p.name != ".DS_Store"):
        rel = path.relative_to(root)
        rows.append((classify(path), rel.as_posix(), path.suffix.lower() or "无扩展名"))

    counts = Counter(category for category, _, _ in rows)
    limit = len(rows) if args.all else max(args.limit, 0)
    visible_rows = rows[:limit]

    lines = [
        f"# 项目资料盘点：{root.name}",
        "",
        f"- 资料夹：`{root}`",
        f"- 文件数量：{len(rows)}",
        f"- 明细列出：{len(visible_rows)}",
        "",
        "## 分类统计",
        "",
        "| 类别 | 数量 |",
        "|---|---:|",
    ]
    lines.extend(f"| {category} | {counts[category]} |" for category in sorted(counts))
    lines.extend([
        "",
        "## 文件明细",
        "",
        "| 类别 | 文件 | 格式 |",
        "|---|---|---|",
    ])
    lines.extend(f"| {category} | {rel} | {suffix} |" for category, rel, suffix in visible_rows)
    if len(visible_rows) < len(rows):
        lines.append(f"| 省略 | 还有 {len(rows) - len(visible_rows)} 个文件未列出，可用 `--all` 输出全量明细。 |  |")
    lines.append("")
    lines.append("## 待人工判断")
    lines.append("")
    lines.append("- `其他资料` 中是否有保护方案成稿所需的关键附件。")
    lines.append("- 同名或多版本资料需要确认使用版本。")

    output = "\n".join(lines)
    if args.output:
        target = Path(args.output).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(output + "\n", encoding="utf-8")
        print(f"已写入资料盘点：{target}")
    else:
        print(output)


if __name__ == "__main__":
    main()
