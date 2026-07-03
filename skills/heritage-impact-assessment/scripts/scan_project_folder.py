#!/usr/bin/env python3
"""Scan a project folder and create draft evidence tables for impact assessment."""

from __future__ import annotations

import argparse
from pathlib import Path


IGNORE_NAMES = {".DS_Store", "Thumbs.db"}
IGNORE_SUFFIXES = {".log"}


INFO_RULES = [
    ("项目基本信息", ["合同", "项目资料", "项目说明", "建设", "改扩建", "规划", "平面图"], "一、总则；三、建设项目规划概况", "项目名称、建设单位、位置、建设内容、规模"),
    ("文物基本信息", ["文物", "长城", "遗址", "调查的报告"], "二、建设项目涉及文物概况", "文物名称、级别、类型、保存现状、价值"),
    ("保护区划信息", ["保护范围", "建设控制地带", "建控", "遗产区", "缓冲区", "本体范围", "生态景观规划保护区"], "二、建设项目涉及文物概况；四、空间关系", "保护范围、建设控制地带、遗产区、缓冲区、本体范围等材料原文术语及内容"),
    ("空间关系信息", ["红线", "坐标", "kml", "ovkml", "范围", "正射", "位置", "距离", "重叠", "叠合", "管控区"], "四、项目用地范围与文物空间分布关系", "红线、项目坐标、文保点坐标、最近距离、重叠面积、叠合关系、所在管控区"),
    ("工程设计信息", ["规划", "总平面", "平面图", "设计", "Model"], "三、建设项目规划概况；五、设计合规性评估", "总平面、建筑布局、高度体量、功能"),
    ("建设期信息", ["施工", "开工", "探孔", "普探", "勘探单元", "标准孔", "剖线"], "五、建设期影响评估；六、施工期减缓措施", "施工扰动、地下文物风险、勘探过程"),
    ("运营期信息", ["运营", "旅游", "道路", "配套", "设施"], "五、运营期影响评估；六、运营期减缓措施", "运营功能、人流车流、环境影响"),
    ("调查勘探信息", ["考古", "勘探", "调查", "验收", "探孔", "标准孔", "剖线"], "四、调查勘探工作；五、地下文物影响", "调查范围、勘探结果、验收意见"),
    ("审查评审信息", ["函件", "备案", "告知", "评审", "意见", "复核", "请示"], "一、必要性；七、结论建议", "文物部门意见、评审意见、审查要求"),
    ("图件照片信息", [".jpg", ".jpeg", ".png", ".pdf", "影像", "现状照", "踏查照", "位置图", "示意图"], "二、文物现状；四、空间关系；正文图件", "图件、照片、影像资料"),
]


CHAPTER_RULES = [
    ("一、总则", "评估项目基础信息；文物影响评估必要性", ["项目基本信息", "审查评审信息"], "用事实交代项目和开展文评的必要性"),
    ("二、建设项目涉及文物概况", "文物所在区域概况；【文物名称】概述六个固定小节；6.价值评估固定段和表1", ["文物基本信息", "保护区划信息", "图件照片信息"], "正式名称、级别、保护区划必须有依据；项目价值评级写在固定表格之后"),
    ("三、建设项目规划概况", "项目概况；建设及运营方案", ["项目基本信息", "工程设计信息", "运营期信息"], "先写工程事实，为影响分析铺垫"),
    ("四、项目用地范围与文物空间分布关系", "相对位置关系；选址分析；调查勘探工作", ["空间关系信息", "调查勘探信息", "图件照片信息"], "距离、重叠、管控区关系必须可核验"),
    ("五、建设项目可能对文物造成的影响分析与评估", "设计合规性；建设期；运营期；综合影响", ["工程设计信息", "建设期信息", "运营期信息", "调查勘探信息"], "影响路径对应真实工程事实"),
    ("六、减缓措施建议", "小节标题按第五章影响问题确定", ["工程设计信息", "建设期信息", "运营期信息"], "措施必须回应第五章影响；可采用阶段式或问题式组织"),
    ("七、文物影响评估结论及建议", "评估结论；其他建议", ["审查评审信息", "空间关系信息"], "不写成审批意见"),
    ("八、支撑法律法规及文件", "法律法规及文件", [], "法规需另行核验"),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="扫描项目资料文件夹并生成文评证据组织底稿")
    parser.add_argument("--项目资料目录", dest="project_dir", required=True)
    parser.add_argument("--输出目录", dest="output_dir", required=True)
    parser.add_argument("--项目名称", dest="project_name", default="未命名项目")
    parser.add_argument("--覆盖", dest="overwrite", action="store_true")
    return parser.parse_args()


def classify(path: Path) -> list[str]:
    haystack = " ".join([path.name, *path.parts]).lower()
    categories: list[str] = []
    for category, keywords, *_ in INFO_RULES:
        if any(keyword.lower() in haystack for keyword in keywords):
            categories.append(category)
    if not categories:
        suffix = path.suffix.lower()
        if suffix in {".docx", ".doc", ".pdf"}:
            categories.append("项目基本信息")
        elif suffix in {".xlsx", ".xls", ".csv", ".kml", ".ovkml"}:
            categories.append("空间关系信息")
        elif suffix in {".jpg", ".jpeg", ".png", ".tif", ".tiff"}:
            categories.append("图件照片信息")
        else:
            categories.append("其他")
    return list(dict.fromkeys(categories))


def collect_files(project_dir: Path) -> list[tuple[Path, list[str]]]:
    rows = []
    for path in sorted(project_dir.rglob("*")):
        if not path.is_file():
            continue
        if path.name in IGNORE_NAMES or path.suffix.lower() in IGNORE_SUFFIXES:
            continue
        rows.append((path, classify(path.relative_to(project_dir))))
    return rows


def md_escape(value: object) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def write_file_inventory(output_dir: Path, project_dir: Path, rows: list[tuple[Path, list[str]]]) -> None:
    lines = [
        "# 项目资料目录",
        "",
        f"项目资料目录：`{project_dir}`",
        "",
        "| 序号 | 相对路径 | 文件类型 | 初判信息类别 | 可能用途 |",
        "| --- | --- | --- | --- | --- |",
    ]
    for idx, (path, categories) in enumerate(rows, start=1):
        rel = path.relative_to(project_dir)
        use = "；".join(rule[2] for rule in INFO_RULES if rule[0] in categories) or "待人工判断"
        lines.append(f"| {idx} | `{md_escape(rel)}` | {md_escape(path.suffix or '无扩展名')} | {md_escape('；'.join(categories))} | {md_escape(use)} |")
    (output_dir / "00-项目资料目录.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_available_info(output_dir: Path, project_dir: Path, rows: list[tuple[Path, list[str]]]) -> None:
    by_category: dict[str, list[Path]] = {}
    for path, categories in rows:
        for category in categories:
            by_category.setdefault(category, []).append(path.relative_to(project_dir))

    rule_map = {rule[0]: rule for rule in INFO_RULES}
    lines = [
        "# 项目资料可用信息清单",
        "",
        "| 信息类别 | 可用信息 | 来源文件/路径 | 可用于报告位置 | 可信度 | 风险/待确认 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for category, files in by_category.items():
        if category == "其他":
            continue
        _, _, chapter, facts = rule_map.get(category, (category, [], "待人工判断", "待人工判断"))
        source = "<br>".join(f"`{md_escape(p)}`" for p in files[:12])
        if len(files) > 12:
            source += f"<br>...另 {len(files) - 12} 项"
        risk = "高风险，需人工核验" if category in {"保护区划信息", "空间关系信息", "调查勘探信息", "审查评审信息"} else "待核验"
        lines.append(f"| {category} | {facts} | {source} | {chapter} | 待核验 | {risk} |")
    lines.extend(["", "## 不足以直接写入正文的信息", "", "| 信息 | 问题 | 需要补充或确认 |", "| --- | --- | --- |", "|  |  |  |"])
    (output_dir / "10-项目资料可用信息清单.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_chapter_mapping(output_dir: Path, project_dir: Path, rows: list[tuple[Path, list[str]]]) -> None:
    category_files: dict[str, list[Path]] = {}
    for path, categories in rows:
        for category in categories:
            category_files.setdefault(category, []).append(path.relative_to(project_dir))

    lines = [
        "# 信息入文章节映射表",
        "",
        "| 可用信息/事实 | 来源文件/路径 | 写入章节 | 写入小节 | 写法要求 | 缺项/风险 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for chapter, section, categories, writing_rule in CHAPTER_RULES:
        if not categories:
            lines.append(f"| 法规、政策、规范依据 | 【待补充法规依据】 | {chapter} | {section} | {writing_rule} | 待核验 |")
            continue
        for category in categories:
            files = category_files.get(category, [])
            source = "<br>".join(f"`{md_escape(p)}`" for p in files[:8]) if files else "【待补充】"
            if len(files) > 8:
                source += f"<br>...另 {len(files) - 8} 项"
            risk = "高风险" if category in {"保护区划信息", "空间关系信息", "调查勘探信息", "审查评审信息"} else "待核验"
            lines.append(f"| {category} | {source} | {chapter} | {section} | {writing_rule} | {risk} |")
    (output_dir / "11-信息入文章节映射表.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    args = parse_args()
    project_dir = Path(args.project_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    if not project_dir.exists() or not project_dir.is_dir():
        raise SystemExit(f"项目资料目录不存在或不是文件夹：{project_dir}")
    if output_dir.exists() and any(output_dir.iterdir()) and not args.overwrite:
        raise SystemExit(f"输出目录非空，如需覆盖请添加 --覆盖：{output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = collect_files(project_dir)
    write_file_inventory(output_dir, project_dir, rows)
    write_available_info(output_dir, project_dir, rows)
    write_chapter_mapping(output_dir, project_dir, rows)

    print(f"已扫描 {len(rows)} 个文件，输出目录：{output_dir}")


if __name__ == "__main__":
    main()
