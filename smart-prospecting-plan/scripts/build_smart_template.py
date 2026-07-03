#!/usr/bin/env python3
"""Build a smart-generation DOCX template from the read-only base template."""

from __future__ import annotations

import copy
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "r": R_NS}
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)

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
SOURCE = ROOT / "基础信息" / "北京卓凡文博技术有限公司—勘探报告模板（无遗迹）.docx"
OUTPUT = (
    ROOT
    / "智能生成报告技能资料"
    / "知识库"
    / "3.模板与表单"
    / "模板"
    / "报告、计划"
    / "北京卓凡文博技术有限公司-基准模板"
    / "智能报告生成基准模板_北京卓凡文博技术有限公司_无遗迹.docx"
)
PLACEHOLDER_DOC = ROOT / "过程资料" / "智能生成模板占位符清单.md"


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def text_nodes(elem: ET.Element) -> list[ET.Element]:
    return [node for node in elem.iter(qn("w:t"))]


def element_text(elem: ET.Element) -> str:
    out = []
    for node in elem.iter():
        if node.tag == qn("w:t") and node.text:
            out.append(node.text)
        elif node.tag == qn("w:tab"):
            out.append(" ")
    return "".join(out).strip()


def set_text(elem: ET.Element, value: str) -> None:
    nodes = text_nodes(elem)
    if not nodes:
        return
    nodes[0].text = value
    nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in nodes[1:]:
        node.text = ""


def body_blocks(root: ET.Element) -> list[ET.Element]:
    body = root.find("w:body", NS)
    return list(body) if body is not None else []


def clone_para_after(body: ET.Element, ref: ET.Element, value: str) -> ET.Element:
    new_para = copy.deepcopy(ref)
    set_text(new_para, value)
    blocks = list(body)
    body.insert(blocks.index(ref) + 1, new_para)
    return new_para


def add_marker_after_caption(root: ET.Element, caption_contains: str, marker: str) -> bool:
    body = root.find("w:body", NS)
    if body is None:
        return False
    for block in list(body):
        if block.tag == qn("w:p") and caption_contains in element_text(block):
            clone_para_after(body, block, marker)
            return True
    return False


def replace_prefix_paragraph(root: ET.Element, prefix: str, value: str) -> int:
    count = 0
    for p in root.findall(".//w:p", NS):
        text = element_text(p).replace("[]", "")
        if text.startswith(prefix):
            set_text(p, prefix + value)
            count += 1
    return count


def replace_all_text(root: ET.Element, replacements: dict[str, str]) -> int:
    changed = 0
    for elem in root.findall(".//w:p", NS) + root.findall(".//w:tc", NS):
        text = element_text(elem).replace("[]", "")
        if not text:
            continue
        new_text = text
        for old, new in sorted(replacements.items(), key=lambda item: len(item[0]), reverse=True):
            new_text = new_text.replace(old, new)
        if new_text != text:
            set_text(elem, new_text)
            changed += 1
    return changed


def replace_exact_or_contains(root: ET.Element, replacements: dict[str, str]) -> int:
    changed = 0
    for p in root.findall(".//w:p", NS):
        text = element_text(p).replace("[]", "")
        if text in replacements:
            set_text(p, replacements[text])
            changed += 1
            continue
        for old, new in replacements.items():
            if old in text:
                set_text(p, text.replace(old, new))
                changed += 1
                break
    return changed


def simplify_table(tbl: ET.Element, rows: list[list[str]]) -> None:
    existing = tbl.findall("w:tr", NS)
    if not existing:
        return
    template = existing[-1]
    for tr in existing:
        tbl.remove(tr)
    for row in rows:
        new_tr = copy.deepcopy(template)
        cells = new_tr.findall("w:tc", NS)
        while len(cells) < len(row):
            cells.append(copy.deepcopy(cells[-1]))
            new_tr.append(cells[-1])
        for idx, value in enumerate(row):
            set_text(cells[idx], value)
        tbl.append(new_tr)


def table_after_caption(root: ET.Element, caption_contains: str) -> ET.Element | None:
    seen = False
    for block in body_blocks(root):
        if block.tag == qn("w:p") and caption_contains in element_text(block):
            seen = True
        elif seen and block.tag == qn("w:tbl"):
            return block
    return None


def transform_document(root: ET.Element) -> dict[str, int]:
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
    }
    stats["prefix_fields"] = sum(replace_prefix_paragraph(root, k, v) for k, v in prefix_values.items())

    replacements = {
        "零碳产业园宝丰西侧至经十三路东侧拟出让地块场平项目": "{{项目名称}}",
        "鄂尔多斯蒙苏经济开发区管理委员会": "{{建设单位}}",
        "内蒙古自治区鄂尔多斯市伊金霍洛旗": "{{项目位置}}",
        "832450平方米": "{{项目面积}}",
        "530595平方米": "{{勘探面积}}",
        "2026年1月19日至2026年2月2日": "{{勘探时间}}",
        "2026年2月": "{{报告年月}}",
        "109°31'27″": "{{经度}}",
        "39°34'38″": "{{纬度}}",
        "建筑占压、人工垫土及季节性积水": "{{不可勘探原因}}",
        "建筑、人工堆积及积水": "{{现场限制因素}}",
        "标准孔20个": "标准孔{{标准孔数量}}个",
        "2条地层剖线": "{{剖线数量}}条地层剖线",
        "部分区域因存在{{不可勘探原因}}，暂不具备勘探条件": "{{IF_存在不可勘探区域}}部分区域因存在{{不可勘探原因}}，暂不具备勘探条件{{ENDIF}}",
    }
    stats["global_replacements"] = replace_all_text(root, replacements)

    fixes = {
        "为本次工作的核心目的是": "本次工作的核心目的是",
        "（3《内蒙古自治区文物保护条例》": "（3）《内蒙古自治区文物保护条例》",
        "或（根据某某文件，……）": "{{IF_有文物审查意见}}根据{{文物审查意见文件名}}，{{文物审查意见结论}}。{{ENDIF}}",
        "5、遗迹编号": "5、遗迹编号",
        "遗迹编号根据2009年颁布施行的《田野考古工作规程》附录一《野外作业技术要点》第三点《考古发掘》中，对考古遗迹单位的编号符号的规定，遗迹单位符号采用其汉语拼音的第一个字的大写字母表示，如：T-探方（沟）；H-灰坑；F-房屋（址）；M-墓葬；G-沟；J-井；L-路；Y-窑；Z-灶；Q-墙。": "{{IF_无遗迹}}本次勘探未发现遗迹单位，未进行遗迹编号。{{ENDIF}}{{IF_有遗迹}}遗迹单位编号依据《田野考古工作规程》执行，按遗迹类型拼音首字母编号，如H-灰坑、M-墓葬、F-房址、G-沟等。{{ENDIF}}",
        "勘探过程中，利用RTK对发现的遗迹进行精确定位、测绘及矢量化处理，利用文字、线图、照片等技术手段对各类信息进行记录。": "{{IF_无遗迹}}勘探过程中，利用RTK对探孔、剖线、标准孔及项目范围进行定位和测绘，并通过文字、图件、照片等方式记录工作信息。{{ENDIF}}{{IF_有遗迹}}勘探过程中，利用RTK对发现的遗迹进行精确定位、测绘及矢量化处理，利用文字、线图、照片等技术手段对各类信息进行记录。{{ENDIF}}",
        "经过对勘探区域的系统钻探及土样分析，在本项目已完成勘探的{{勘探面积}}范围内，未发现古代文化遗存。": "经过对勘探区域的系统钻探及土样分析，在本项目已完成勘探的{{勘探面积}}范围内，未发现古代文化遗存。",
    }
    stats["fixes"] = replace_exact_or_contains(root, fixes)

    caption_replacements = {
        "图1 项目地块在内蒙古自治区位置示意图": "图1 {{图:项目地块在内蒙古自治区位置示意图}}",
        "图2 项目地块在鄂尔多斯市位置示意图": "图2 {{图:项目地块在鄂尔多斯市位置示意图}}",
        "图3 项目地块在伊金霍洛旗位置示意图": "图3 {{图:项目地块在旗县区位置示意图}}",
        "图4  项目地块卫星图": "图4 {{图:项目地块卫星图}}",
        "图 5  项目红线四至坐标图": "图5 {{图:项目红线四至坐标图}}",
        "图6  勘探流程图": "图6 {{图:勘探流程图}}",
        "图7 项目勘探区域": "图7 {{图:项目勘探区域图}}",
        "图45  勘探单元布置照": "图45 {{图:勘探单元布置照}}",
        "图46  划分勘探单元图": "图46 {{图:划分勘探单元图}}",
        "图47 勘探布孔示意图": "图47 {{图:勘探布孔示意图}}",
        "图82 剖面位置图": "图82 {{图:剖面位置图}}",
        "图83 A-A′  地层堆积剖线图": "图83 {{图:A-A′地层堆积剖线图}}",
        "图84 B-B′  地层堆积剖线图": "图84 {{图:B-B′地层堆积剖线图}}",
        "图85 标准探孔位置示意图": "图85 {{图:标准探孔位置示意图}}",
        "图86 资料整理工作照": "图86 {{图:资料整理工作照}}",
    }
    stats["caption_markers"] = replace_exact_or_contains(root, caption_replacements)

    markers = {
        "表 1  红线四至坐标": "{{#红线坐标}}本表由 Excel「红线坐标」生成。{{/红线坐标}}",
        "表 4  勘探单元坐标": "{{#勘探单元}}本表由 Excel「勘探单元」生成。{{/勘探单元}}",
        "附表一  勘探单元坐标表": "{{#附表_勘探单元}}本附表由 Excel「勘探单元」生成。{{/附表_勘探单元}}",
        "附表二  标准孔坐标表": "{{#标准孔坐标}}本附表由 Excel「标准孔与剖线」生成。{{/标准孔坐标}}",
        "附表三  考古勘探剖面探孔记录表（A-A’）": "{{#剖线_AA}}本附表由 Excel「剖线记录」生成。{{/剖线_AA}}",
        "附表四  考古勘探剖面探孔记录表（B-B’）": "{{#剖线_BB}}本附表由 Excel「剖线记录」生成。{{/剖线_BB}}",
    }
    stats["section_markers"] = sum(1 for caption, marker in markers.items() if add_marker_after_caption(root, caption, marker))

    table = table_after_caption(root, "表 1  红线四至坐标")
    if table is not None:
        simplify_table(
            table,
            [
                ["四至范围坐标", "", ""],
                ["{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
                ["{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
                ["{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
                ["{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
            ],
        )
        stats["redline_table"] = 1

    table = table_after_caption(root, "表 4  勘探单元坐标")
    if table is not None:
        simplify_table(
            table,
            [
                ["勘探单元坐标", "", "", ""],
                ["{{单元编号}}", "{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
                ["", "{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
                ["", "{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
                ["", "{{角点}}", "{{X坐标}}", "{{Y坐标}}"],
            ],
        )
        stats["unit_table"] = 1

    return stats


def build() -> None:
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    PLACEHOLDER_DOC.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(SOURCE, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    root = ET.fromstring(entries["word/document.xml"])
    stats = transform_document(root)
    entries["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(OUTPUT, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    PLACEHOLDER_DOC.write_text(
        """# 智能生成模板占位符清单

## 基础字段

- `{{项目名称}}`
- `{{建设单位}}`
- `{{项目位置}}`
- `{{项目面积}}`
- `{{调查面积}}`
- `{{勘探面积}}`
- `{{勘探时间}}`
- `{{报告年月}}`
- `{{遗迹结论}}`
- `{{经度}}`
- `{{纬度}}`
- `{{不可勘探原因}}`
- `{{现场限制因素}}`
- `{{标准孔数量}}`
- `{{剖线数量}}`

## 条件段

- `{{IF_无遗迹}}...{{ENDIF}}`
- `{{IF_有遗迹}}...{{ENDIF}}`
- `{{IF_存在不可勘探区域}}...{{ENDIF}}`
- `{{IF_有文物审查意见}}...{{ENDIF}}`

## 重复表格

- `{{#红线坐标}}...{{/红线坐标}}`
- `{{#勘探单元}}...{{/勘探单元}}`
- `{{#标准孔坐标}}...{{/标准孔坐标}}`
- `{{#剖线_AA}}...{{/剖线_AA}}`
- `{{#剖线_BB}}...{{/剖线_BB}}`

## 图片位

图题中的 `{{图:...}}` 是稳定图片插入位，后续生成器可按名称匹配图片，不再依赖模板图片序号。
""",
        encoding="utf-8",
    )

    print(OUTPUT)
    print(PLACEHOLDER_DOC)
    print(stats)


if __name__ == "__main__":
    build()
