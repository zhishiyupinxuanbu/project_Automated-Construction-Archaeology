#!/usr/bin/env python3
"""Upgrade all legacy company templates into smart baseline templates.

Legacy templates are treated as source material only. Generated smart baseline
templates are written to 智能生成报告技能资料/知识库/3.模板与表单/模板/报告、计划 with the naming convention:

智能报告生成基准模板_公司全称_有遗迹/无遗迹.docx
"""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import build_smart_template


def find_project_root() -> Path:
    here = Path(__file__).resolve()
    candidates = [here.parent, *here.parents, Path.cwd(), *Path.cwd().resolve().parents]
    for candidate in candidates:
        if (
            (candidate / "智能生成报告技能资料" / "知识库" / "3.模板与表单").exists()
            and ((candidate / "过程资料").exists() or (candidate / "基础信息").exists())
        ):
            return candidate
    for candidate in candidates:
        if candidate.name == "智能生成报告技能资料":
            return candidate.parent
    return here.parents[1]


ROOT = find_project_root()
LEGACY_DIR = ROOT / "基础信息" / "旧模板"
OUT_DIR = ROOT / "智能生成报告技能资料" / "知识库" / "3.模板与表单" / "模板" / "报告、计划"
REPORT = ROOT / "过程资料" / "模板升级报告.md"
PROTECTED_BASELINES = {
    "智能报告生成基准模板_北京卓凡文博技术有限公司_无遗迹.docx",
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)

COMPANY_NAMES = [
    "北京卓凡文博技术有限公司",
    "内蒙古煊迹考古勘探有限公司",
    "内蒙古峰驰考古勘探有限公司",
    "三门峡市文物考古勘探有限公司",
    "河南燧火文物保护有限公司",
]


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def element_text(elem: ET.Element) -> str:
    out: list[str] = []
    for node in elem.iter():
        if node.tag == qn("w:t") and node.text:
            out.append(node.text)
        elif node.tag == qn("w:tab"):
            out.append(" ")
    return "".join(out).strip().replace("[]", "")


def text_nodes(elem: ET.Element) -> list[ET.Element]:
    return [node for node in elem.iter(qn("w:t"))]


def set_text(elem: ET.Element, value: str) -> None:
    nodes = text_nodes(elem)
    if not nodes:
        return
    nodes[0].text = value
    nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in nodes[1:]:
        node.text = ""


def body(root: ET.Element) -> ET.Element | None:
    return root.find("w:body", NS)


def replace_contains(root: ET.Element, old: str, new: str) -> int:
    count = 0
    for elem in root.findall(".//w:p", NS) + root.findall(".//w:tc", NS):
        text = element_text(elem)
        if old in text:
            set_text(elem, text.replace(old, new))
            count += 1
    return count


def replace_exact(root: ET.Element, old: str, new: str) -> int:
    count = 0
    for elem in root.findall(".//w:p", NS) + root.findall(".//w:tc", NS):
        if element_text(elem) == old:
            set_text(elem, new)
            count += 1
    return count


def replace_prefix_paragraph(root: ET.Element, prefix: str, value: str) -> int:
    count = 0
    for para in root.findall(".//w:p", NS):
        text = element_text(para)
        if text.startswith(prefix):
            set_text(para, prefix + value)
            count += 1
    return count


def template_info(path: Path) -> tuple[str, str]:
    name = path.name
    company = next((item for item in COMPANY_NAMES if item in name), "")
    if not company:
        raise ValueError(f"无法识别模板公司：{path}")
    relic_state = "有遗迹" if "有遗迹" in name else "无遗迹"
    return company, relic_state


def output_path_for(source: Path) -> Path:
    company, relic_state = template_info(source)
    return OUT_DIR / f"智能报告生成基准模板_{company}_{relic_state}.docx"


def apply_generic_smart_rules(root: ET.Element, relic_state: str) -> dict[str, int]:
    stats: dict[str, int] = {}

    prefix_values = {
        "项目名称：": "{{项目名称}}",
        "建设单位：": "{{建设单位}}",
        "项目位置：": "{{项目位置}}",
        "项目面积：": "{{项目面积}}",
        "调查面积：": "{{调查面积}}",
        "勘探面积：": "{{勘探面积}}",
        "勘探时间：": "{{勘探时间}}",
        "遗迹：": "{{遗迹结论}}",
        "遗迹现象：": "{{遗迹结论}}",
    }
    stats["prefix_fields"] = sum(replace_prefix_paragraph(root, key, value) for key, value in prefix_values.items())

    generic_replacements = {
        "标准孔20个": "标准孔{{标准孔数量}}个",
        "2条地层剖线": "{{剖线数量}}条地层剖线",
        "400米×400米": "{{勘探单元规格}}",
        "A399": "A{{普探列最大编号}}",
        "B399": "B{{普探行最大编号}}",
        "A799": "A{{重点勘探列最大编号}}",
        "B799": "B{{重点勘探行最大编号}}",
    }
    stats["generic_replacements"] = sum(replace_contains(root, old, new) for old, new in generic_replacements.items())

    # Reuse the original Beijing conversion rules where they match. They are
    # intentionally match-based, so nonmatching company-specific fixed content is preserved.
    try:
        stats["legacy_transform"] = sum(build_smart_template.transform_document(root).values())
    except Exception:
        stats["legacy_transform"] = 0

    conclusion = (
        "{{IF_无遗迹}}经过对勘探区域的系统钻探及土样分析，在本项目已完成勘探的{{勘探面积}}范围内，未发现古代文化遗存。"
        "{{ENDIF}}{{IF_有遗迹}}{{勘探成果综合结论}}{{ENDIF}}"
    )
    stats["conclusion_placeholders"] = sum(
        replace_contains(root, text, conclusion)
        for text in [
            "经过对勘探区域的系统钻探及土样分析，在本项目已完成勘探的{{勘探面积}}范围内，未发现古代文化遗存。",
            "综合考古调查与勘探结果，在本项目已完成勘探区域内，未发现古代文化遗存。",
        ]
    )

    if relic_state == "有遗迹":
        stats["relic_state_marker"] = replace_contains(root, "未发现文化遗存", "{{遗迹结论}}")
    return stats


def upgrade_template(
    source: Path,
    overwrite: bool = False,
    overwrite_protected: bool = False,
) -> tuple[Path, dict[str, int]]:
    target = output_path_for(source)
    if target.exists() and target.name in PROTECTED_BASELINES and not overwrite_protected:
        return target, {"skipped_protected": 1}
    if target.exists() and not overwrite:
        return target, {"skipped_existing": 1}
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    root = ET.fromstring(entries["word/document.xml"])
    _, relic_state = template_info(source)
    stats = apply_generic_smart_rules(root, relic_state)
    entries["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    return target, stats


def main() -> int:
    parser = argparse.ArgumentParser(description="批量升级旧模板为智能基准模板")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的智能基准模板")
    parser.add_argument(
        "--overwrite-protected",
        action="store_true",
        help="允许覆盖受保护的已定稿基准模板；默认保护北京卓凡无遗迹模板",
    )
    args = parser.parse_args()

    lines = [
        "# 模板升级报告",
        "",
        "| 公司 | 遗迹状态 | 来源旧模板 | 输出智能基准模板 | 状态 | 规则命中 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for source in sorted(LEGACY_DIR.glob("*.docx")):
        company, relic_state = template_info(source)
        target, stats = upgrade_template(
            source,
            overwrite=args.overwrite,
            overwrite_protected=args.overwrite_protected,
        )
        if stats.get("skipped_protected"):
            status = "已保护，跳过"
        elif stats.get("skipped_existing"):
            status = "已存在，跳过"
        else:
            status = "已生成"
        hit_count = sum(value for key, value in stats.items() if key != "skipped_existing")
        lines.append(f"| {company} | {relic_state} | {source.name} | {target.name} | {status} | {hit_count} |")
        print(target)
    REPORT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(REPORT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
