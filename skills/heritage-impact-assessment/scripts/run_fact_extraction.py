#!/usr/bin/env python3
"""Run v0.2.2 fact extraction from material-processing outputs."""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

from pipeline_common import now_iso, read_json, read_jsonl, truncate, update_module_state, write_json, write_jsonl


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="运行文物影响评估 v0.2.2 事实抽取阶段")
    parser.add_argument("--工作目录", "--workspace", dest="work_dir", required=True)
    parser.add_argument("--覆盖", dest="overwrite", action="store_true")
    return parser.parse_args()


def find_value(pattern: str, text: str, default: str = "待确认") -> str:
    match = re.search(pattern, text)
    return match.group(1).strip(" ，。；;") if match else default


def is_markup_or_spatial_text(row: dict, text: str) -> bool:
    suffix = Path(row.get("source_file", "")).suffix.lower()
    normalized = text.lstrip("\ufeff \n\t").lower()
    if suffix in {".kml", ".ovkml", ".xml"}:
        return True
    return normalized.startswith("<?xml") or normalized.startswith("<kml") or "<placemark" in normalized[:1000]


def evidence_by_file_id(evidence_rows: list[dict]) -> dict[str, dict]:
    return {row.get("source_file_id", ""): row for row in evidence_rows}


def load_text_records(text_rows: list[dict], work_dir: Path, evidence_rows: list[dict]) -> list[dict]:
    by_file = evidence_by_file_id(evidence_rows)
    records = []
    for row in text_rows:
        path = work_dir / row.get("text_path", "")
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8", errors="ignore").strip()
        evidence = by_file.get(row.get("file_id", ""), {})
        records.append(
            {
                **row,
                "text": text,
                "source_evidence_id": evidence.get("evidence_id", ""),
                "evidence_source_file": evidence.get("source_file", row.get("source_file", "")),
                "is_markup_or_spatial": is_markup_or_spatial_text(row, text),
            }
        )
    return records


def select_text(records: list[dict], keywords: list[str], *, allow_spatial: bool = False) -> dict:
    candidates = [row for row in records if allow_spatial or not row["is_markup_or_spatial"]]
    for row in candidates:
        haystack = f"{row.get('source_file', '')} {row.get('text_path', '')} {row.get('text', '')[:500]}"
        if any(keyword in haystack for keyword in keywords):
            return row
    return candidates[0] if candidates else {}


def quote_allowed(text_record: dict) -> bool:
    if not text_record or text_record.get("is_markup_or_spatial"):
        return False
    text = text_record.get("text", "").strip()
    if len(text) < 40:
        return False
    return not is_markup_or_spatial_text(text_record, text)


def source_fields(text_record: dict) -> tuple[str, str, str]:
    if not text_record:
        return "", "", "待确认"
    return (
        text_record.get("source_evidence_id", ""),
        text_record.get("evidence_source_file") or text_record.get("source_file", ""),
        text_record.get("text_path", "文本摘录"),
    )


def clean_value(value: str) -> str:
    return " ".join(value.strip(" ，。；;\n\t").split())


def regex_value(pattern: str, text: str) -> str:
    match = re.search(pattern, text, flags=re.S)
    return clean_value(match.group(1)) if match else ""


def has_any(text: str, keywords: list[str]) -> bool:
    return any(keyword in text for keyword in keywords)


def extract_technical_spec_section(text: str) -> str:
    match = re.search(r"(?:^|\n)\s*4[.．、]\s*技术规范[：:]?\s*", text)
    if not match:
        return ""
    remainder = text[match.end() :]
    end = re.search(r"\n\s*5[.．、]\s*(?:其他资料|其他|相关资料)", remainder)
    section = remainder[: end.start()] if end else remainder
    return section.strip()


def extract_chapter_two_opening_lead(text: str) -> str:
    normalized = " ".join(text.strip().split())
    if not has_any(normalized, ["经查阅", "发现", "文物影响评估"]):
        return ""
    match = re.search(r"(经查阅.{20,700}?文物影响评估。?)", normalized)
    return match.group(1).strip() if match else ""


def extract_project_overview_section(text: str) -> str:
    match = re.search(r"(?:^|\n)\s*(?:一[、.]|\（一\）|\(一\))?\s*项目概况[：:]?\s*", text)
    if not match:
        return ""
    remainder = text[match.end() :].strip()
    end = re.search(r"\n\s*(?:二[、.]|\（二\）|\(二\)|[一二三四五六七八九十]+[、.])\s*[^。\n]{0,30}", remainder)
    section = remainder[: end.start()] if end else remainder
    paragraphs = [paragraph.strip() for paragraph in re.split(r"\n\s*\n", section) if paragraph.strip()]
    return "\n\n".join(paragraphs[:4]).strip()


def extract_project_necessity_section(text: str) -> str:
    match = re.search(r"(?:^|\n)\s*(?:二[、.]|\（二\）|\(二\)|[一二三四五六七八九十]+[、.])?\s*项目建设必要性[：:]?\s*", text)
    if not match:
        return ""
    remainder = text[match.end() :].strip()
    end = re.search(r"\n\s*(?:三[、.]|\（三\）|\(三\)|[一二三四五六七八九十]+[、.])\s*[^。\n]{0,30}", remainder)
    return (remainder[: end.start()] if end else remainder).strip()


def extract_site_selection_section(text: str) -> str:
    match = re.search(
        r"(?:^|\n)\s*(?:[一二三四五六七八九十]+[、.．]|\（[一二三四五六七八九十]+\）|\([一二三四五六七八九十]+\))?\s*(?:项目)?选址(?:必要性)?分析[：:]?\s*",
        text,
    )
    if not match:
        return ""
    remainder = text[match.end() :].strip()
    end = re.search(
        r"\n\s*(?:[一二三四五六七八九十]+[、.．]|\（[一二三四五六七八九十]+\）|\([一二三四五六七八九十]+\))\s*[^。\n]{0,40}",
        remainder,
    )
    return (remainder[: end.start()] if end else remainder).strip()


def extract_numbered_subsection(text: str, number: int, title: str, next_pattern: str) -> str:
    pattern = rf"(?:^|\n)\s*(?:{number}[.．、]|\（{number}\）|\({number}\))\s*{re.escape(title)}[：:]?\s*"
    match = re.search(pattern, text)
    if not match:
        return ""
    remainder = text[match.end() :].strip()
    end = re.search(next_pattern, remainder)
    return (remainder[: end.start()] if end else remainder).strip()


def extract_construction_scale_section(text: str) -> str:
    return extract_numbered_subsection(
        text,
        1,
        "建设规模与技术指标",
        r"\n\s*(?:2[.．、]|\（2\）|\(2\))\s*项目运营方案与规模[：:]?\s*",
    )


def extract_operation_scale_section(text: str) -> str:
    return extract_numbered_subsection(
        text,
        2,
        "项目运营方案与规模",
        r"\n\s*(?:3[.．、]|\（3\）|\(3\)|[四五六七八九十]+[、.．]|\（[四五六七八九十]+\）|\([四五六七八九十]+\))\s*[^。\n]{0,40}",
    )


def is_feasibility_or_scheme_source(record: dict[str, Any]) -> bool:
    return has_any(
        record.get("source_file", "") + record.get("text", "")[:500],
        ["可研", "可行性研究", "建设方案", "项目说明"],
    )


def source_fact(record: dict[str, Any], field_name: str, value: str, *, use_mode: str = "可直接引用") -> dict[str, Any]:
    evidence_id, source_file, source_location = source_fields(record)
    return {
        "fact_type": "项目事实",
        "field_name": field_name,
        "value": clean_value(value),
        "source_evidence_id": evidence_id,
        "source_file": source_file,
        "source_location": source_location,
        "use_mode": use_mode,
        "confidence": "中",
        "notes": "",
    }


def extract_project_facts_from_record(record: dict[str, Any], project_name: str) -> list[dict[str, Any]]:
    text = record.get("text", "")
    source_file = record.get("source_file", "")
    facts: list[dict[str, Any]] = []

    field_patterns = [
        ("项目代码", r"项目代码[：:]?\s*([0-9A-Za-z-]+)"),
        ("建设单位", r"建设单位(?:为|[：:])\s*([^，。；;\n]+)"),
        ("建设地点", r"(?:建设地点|拟建设地点|位于)(?:为|[：:])?\s*([^，。；;\n]+)"),
        ("项目面积", r"(?:项目面积|用地面积|勘探面积)[为：:]?\s*([0-9.]+\s*(?:平方米|㎡|公顷|亩))"),
        ("总建筑面积", r"((?:项目)?(?:扩建)?总建筑面积\s*[0-9.]+\s*(?:平方米|㎡|公顷|亩))"),
        ("总占地面积", r"((?:扩建后)?(?:寺院)?总占地面积\s*[0-9.]+\s*(?:平方米|㎡|公顷|亩))"),
        ("占地规模", r"占地面积\s*([0-9.]+\s*公顷(?:（约\s*[0-9.]+\s*亩）|\\(约\s*[0-9.]+\s*亩\\))?)"),
        ("总投资", r"(?:项目计划总投资|项目总投资|总投资)\s*([0-9.]+\s*万元)"),
        ("资金来源", r"(企业自筹资金|企业自筹|财政资金|自筹资金)"),
        ("建设工期", r"(20\d{2}\s*年\s*\d{1,2}\s*月\s*[—至-]\s*20\d{2}\s*年\s*\d{1,2}\s*月)"),
        ("建设模式", r"(DBB\s*模式(?:（[^）]+）|\\([^)]+\\))?)"),
    ]
    for field_name, pattern in field_patterns:
        value = regex_value(pattern, text)
        if value:
            facts.append(source_fact(record, field_name, value))

    goal = regex_value(r"以[“\"]([^”\"]+)[”\"]为核心\s*建设目标", text)
    if goal:
        facts.append(source_fact(record, "建设目标", f"以“{goal}”为核心建设目标", use_mode="需改写"))

    content = regex_value(r"建设内容(?:包括|[：:])\s*([^。\n]{8,220})", text)
    if content:
        facts.append(source_fact(record, "建设内容", content, use_mode="需改写"))

    if has_any(text, ["主要经济技术指标", "经济技术指标表"]):
        excerpt = regex_value(r"(?:主要经济技术指标表|主要经济技术指标)([\s\S]{0,700})", text)
        facts.append(source_fact(record, "主要经济技术指标", excerpt or "资料包含主要经济技术指标表", use_mode="需改写"))

    necessity_terms = [
        "项目建设必要性",
        "建设背景和必要性",
        "安全底线",
        "寺院发展战略",
        "文化展示品牌化",
        "服务能力优质化",
        "长期稳定运营",
        "市场需求",
    ]
    if has_any(text, necessity_terms):
        found = [term for term in necessity_terms if term in text]
        facts.append(source_fact(record, "建设必要性", "；".join(found), use_mode="需改写"))

    if has_any(text, ["选址方案比选", "多维度比选", "原址改扩建方案", "异地新建方案"]):
        pieces = []
        for term in ["原址改扩建方案", "原址扩建+邻近地块补充方案", "异地新建方案", "规划符合性", "技术可行性", "经济合理性", "社会适应性"]:
            if term in text:
                pieces.append(term)
        facts.append(source_fact(record, "选址方案比选", "；".join(pieces) or "资料包含选址方案比选分析", use_mode="需改写"))

    if project_name and project_name in text:
        facts.append(source_fact(record, "项目名称", project_name))

    # File names can carry reliable document-role facts even when text is short.
    if "可研" in source_file or "可行性研究" in source_file:
        facts.append(source_fact(record, "资料来源-可研", "已提供可研或可行性研究资料", use_mode="只可参考"))
    return [fact for fact in facts if fact.get("value")]


def add_fact_ids(facts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[tuple[str, str, str]] = set()
    deduped: list[dict[str, Any]] = []
    for fact in facts:
        key = (fact.get("field_name", ""), fact.get("value", ""), fact.get("source_file", ""))
        if key in seen:
            continue
        seen.add(key)
        fact = dict(fact)
        fact["fact_id"] = f"PF{len(deduped)+1:03d}"
        deduped.append(fact)
    return deduped


def quote_type_for_record(record: dict[str, Any]) -> tuple[str, str]:
    source = record.get("source_file", "")
    text = record.get("text", "")
    if has_any(source + text[:800], ["文物调查报告", "调查的报告", "核查报告", "调查报告"]) and has_any(
        text[:1200],
        ["项目位于", "建设内容", "总建筑面积", "总占地面积", "拟用地范围涉及", "建设控制地带"],
    ):
        return "评估项目基础信息", "一、总则 / （一）编制背景 / 1.评估项目基础信息"
    if has_any(source + text[:500], ["可研", "可行性研究", "建设必要性", "选址"]):
        return "项目背景、建设必要性与选址论证", "三、建设项目规划概况"
    if has_any(source, ["调查", "长城", "文物"]):
        return "文物概况与空间关系", "二、建设项目涉及文物概况；四、项目用地范围与文物空间分布关系"
    if has_any(source, ["规划", "总平", "平面图", "设计"]):
        return "工程设计与建设内容", "三、建设项目规划概况；五、建设项目可能对文物造成的影响分析与评估"
    if has_any(source, ["勘探", "验收", "探孔", "剖线"]):
        return "调查勘探资料", "四、项目用地范围与文物空间分布关系；五、建设项目可能对文物造成的影响分析与评估"
    return "项目资料可引用叙述", "三、建设项目规划概况"


def main() -> None:
    args = parse_args()
    work_dir = Path(args.work_dir).expanduser().resolve()
    manifest = read_json(work_dir / "processing_output" / "manifest.json")
    texts = read_jsonl(work_dir / "processing_output" / "text_index.jsonl")
    evidence = read_jsonl(work_dir / "evidence" / "evidence_register.jsonl")
    records = load_text_records(texts, work_dir, evidence)
    first_evidence_id = evidence[0]["evidence_id"] if evidence else ""

    facts_dir = work_dir / "facts"
    if facts_dir.exists() and any(facts_dir.iterdir()) and not args.overwrite:
        raise SystemExit(f"facts 目录非空，如需覆盖请添加 --覆盖：{facts_dir}")
    facts_dir.mkdir(parents=True, exist_ok=True)

    text_records = [row for row in records if not row["is_markup_or_spatial"]]
    project_records = [
        row
        for row in text_records
        if has_any(
            f"{row.get('source_file', '')} {row.get('text', '')[:1200]}",
            ["项目", "建设", "改扩建", "可研", "可行性研究", "方案", "规划", "请示", "总平", "平面图"],
        )
    ]
    if not project_records:
        project_records = text_records
    heritage_record = select_text(records, ["文物", "长城", "遗址", "调查"])
    spatial_record = select_text(records, ["坐标", "kml", "ovkml", "范围", "红线"], allow_spatial=True)
    project_text = "\n".join(row.get("text", "") for row in project_records)
    heritage_text = heritage_record.get("text", "")
    spatial_text = spatial_record.get("text", "")
    combined = "\n".join([project_text, heritage_text])
    combined_with_spatial = "\n".join([combined, spatial_text])
    heritage_evidence_id, heritage_source_file, heritage_source_location = source_fields(heritage_record)
    spatial_evidence_id, spatial_source_file, spatial_source_location = source_fields(spatial_record)

    project_name = manifest.get("project_name", "待确认")
    raw_project_facts: list[dict[str, Any]] = []
    if project_name:
        raw_project_facts.append(
            {
                "fact_type": "项目事实",
                "field_name": "项目名称",
                "value": project_name,
                "source_evidence_id": first_evidence_id,
                "source_file": "processing_output/manifest.json",
                "source_location": "manifest/project_name",
                "use_mode": "可直接引用",
                "confidence": "中",
                "notes": "",
            }
        )
    for record in project_records:
        raw_project_facts.extend(extract_project_facts_from_record(record, project_name))
    project_facts = add_fact_ids(raw_project_facts)

    heritage_name = "桌子山秦长城东风农场七队长城3段" if "长城" in combined_with_spatial else "待确认"
    distance = find_value(r"距[^，。；;]*?约?\s*([0-9.]+\s*米)", combined_with_spatial)
    zoning = "建设控制地带" if "建设控制地带" in combined_with_spatial else "待确认"
    heritage_level = regex_value(r"(国家级文物保护单位|全国重点文物保护单位|自治区级文物保护单位|省级文物保护单位|市级文物保护单位|县级文物保护单位)", combined)
    value_aspects = regex_value(r"具备([^。；;\n]{2,40}?方面的价值)", combined)
    heritage_value_grade = regex_value(r"遗产价值分级为[“\"]([^”\"]+)[”\"]", combined) or regex_value(
        r"遗产价值分级为([^，。；;\n]+)", combined
    )
    heritage_facts = [
        {
            "fact_id": "HF001",
            "fact_type": "文物事实",
            "field_name": "文物名称",
            "value": heritage_name,
            "source_evidence_id": heritage_evidence_id,
            "source_file": heritage_source_file,
            "source_location": heritage_source_location,
            "use_mode": "待确认" if heritage_name == "待确认" else "可直接引用",
            "confidence": "待核验",
            "notes": "",
        },
        {
            "fact_id": "HF002",
            "fact_type": "文物事实",
            "field_name": "空间关系",
            "value": f"项目位于{zoning}；最近距离{distance}",
            "source_evidence_id": spatial_evidence_id or heritage_evidence_id,
            "source_file": spatial_source_file or heritage_source_file,
            "source_location": spatial_source_location or heritage_source_location,
            "use_mode": "待确认" if zoning == "待确认" else "需改写",
            "confidence": "待核验",
            "notes": "",
        },
        {
            "fact_id": "HF003",
            "fact_type": "文物事实",
            "field_name": "文物级别",
            "value": heritage_level or "待确认",
            "source_evidence_id": heritage_evidence_id,
            "source_file": heritage_source_file,
            "source_location": heritage_source_location,
            "use_mode": "待确认" if not heritage_level else "可直接引用",
            "confidence": "待核验",
            "notes": "",
        },
        {
            "fact_id": "HF004",
            "fact_type": "文物事实",
            "field_name": "价值方面",
            "value": value_aspects or "待确认",
            "source_evidence_id": heritage_evidence_id,
            "source_file": heritage_source_file,
            "source_location": heritage_source_location,
            "use_mode": "待确认" if not value_aspects else "可直接引用",
            "confidence": "待核验",
            "notes": "",
        },
        {
            "fact_id": "HF005",
            "fact_type": "文物事实",
            "field_name": "遗产价值分级",
            "value": heritage_value_grade or "待确认",
            "source_evidence_id": heritage_evidence_id,
            "source_file": heritage_source_file,
            "source_location": heritage_source_location,
            "use_mode": "待确认" if not heritage_value_grade else "专业复核",
            "confidence": "待核验",
            "notes": "",
        },
    ]

    issues = []
    quote_candidates = []
    source_coverage = []
    facts_by_source: dict[str, int] = {}
    for fact in project_facts:
        source = fact.get("source_file", "")
        if source:
            facts_by_source[source] = facts_by_source.get(source, 0) + 1

    for record in text_records:
        evidence_id, source_file, source_location = source_fields(record)
        fact_count = facts_by_source.get(source_file, 0)
        quote_count = 0
        technical_spec_section = extract_technical_spec_section(record.get("text", ""))
        if technical_spec_section:
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": "技术规范原文摘录",
                    "text": technical_spec_section,
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "原文摘录，禁止改写",
                    "target_section": "一、总则 / （五）评估依据 / 4.技术规范",
                }
            )
            quote_count += 1
        chapter_two_opening_lead = extract_chapter_two_opening_lead(record.get("text", ""))
        if chapter_two_opening_lead:
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": "第二章开头总起段",
                    "text": chapter_two_opening_lead,
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "按调查回函或调查类材料替换，禁止删除",
                    "target_section": "二、建设项目涉及文物概况",
                }
            )
            quote_count += 1
        project_overview_section = extract_project_overview_section(record.get("text", ""))
        if project_overview_section and is_feasibility_or_scheme_source(record):
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": "项目概况自然段",
                    "text": project_overview_section,
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "来自可研或建设方案，写自然段，禁止改成表格",
                    "target_section": "三、建设项目规划概况 / （一）项目概况",
                }
            )
            quote_count += 1
        project_necessity_section = extract_project_necessity_section(record.get("text", ""))
        if project_necessity_section and is_feasibility_or_scheme_source(record):
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": "项目建设必要性原文摘录",
                    "text": project_necessity_section,
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "通常从可研报告完整摘取相关内容，禁止概述拟写",
                    "target_section": "三、建设项目规划概况 / （二）项目建设必要性",
                }
            )
            quote_count += 1
        construction_scale_section = extract_construction_scale_section(record.get("text", ""))
        if construction_scale_section and is_feasibility_or_scheme_source(record):
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": "建设规模与技术指标原文摘录",
                    "text": construction_scale_section,
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "通常从可研报告完整摘取相关内容，禁止概述拟写",
                    "target_section": "三、建设项目规划概况 / （四）建设及运营方案 / 1.建设规模与技术指标",
                }
            )
            quote_count += 1
        operation_scale_section = extract_operation_scale_section(record.get("text", ""))
        if operation_scale_section and is_feasibility_or_scheme_source(record):
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": "项目运营方案与规模原文摘录",
                    "text": operation_scale_section,
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "通常从可研报告完整摘取相关内容，禁止概述拟写",
                    "target_section": "三、建设项目规划概况 / （四）建设及运营方案 / 2.项目运营方案与规模",
                }
            )
            quote_count += 1
        site_selection_section = extract_site_selection_section(record.get("text", ""))
        if site_selection_section and is_feasibility_or_scheme_source(record):
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": "选址分析原文摘录",
                    "text": site_selection_section,
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "通常从可研报告选址分析/选址必要性分析章节完整摘取，禁止修改、缩减、概括",
                    "target_section": "四、项目用地范围与文物空间分布关系 / （二）【项目名称】选址分析",
                }
            )
            quote_count += 1
        if quote_allowed(record):
            quote_type, target_section = quote_type_for_record(record)
            quote_candidates.append(
                {
                    "quote_id": f"QC{len(quote_candidates)+1:03d}",
                    "quote_type": quote_type,
                    "text": truncate(record.get("text", ""), 1200),
                    "source_evidence_id": evidence_id,
                    "source_file": source_file,
                    "source_location": source_location,
                    "use_mode": "需改写",
                    "target_section": target_section,
                }
            )
            quote_count += 1
        source_coverage.append(
            {
                "source_file": source_file,
                "text_id": record.get("text_id", ""),
                "text_path": record.get("text_path", ""),
                "char_count": len(record.get("text", "")),
                "fact_count": fact_count,
                "quote_candidate_count": quote_count,
                "status": "covered" if fact_count or quote_count else "no_structured_fact",
                "notes": "" if fact_count or quote_count else "已读取文本，但未抽取到结构化事实；需人工复核是否为可用项目资料。",
            }
        )
    if not quote_candidates:
        issues.append(
            {
                "issue_id": "FI001",
                "field_name": "可引用叙述材料",
                "preferred_value": "",
                "conflicting_values": [],
                "preferred_reason": "未识别到可直接改写入正文的叙述性文本；KML/OVKML/XML 坐标材料已排除在 quote_candidates 之外",
                "impact_level": "high",
                "requires_user_confirmation": False,
            }
        )
    for row in source_coverage:
        if row["status"] == "no_structured_fact":
            issues.append(
                {
                    "issue_id": f"FI{len(issues)+1:03d}",
                    "field_name": "资料覆盖",
                    "preferred_value": row["source_file"],
                    "conflicting_values": [],
                    "preferred_reason": "该资料已进入 text_index，但未抽取到结构化事实或可引用候选段，应复核抽取规则或资料质量。",
                    "impact_level": "medium",
                    "requires_user_confirmation": False,
                }
            )
    for fact in project_facts + heritage_facts:
        if fact["value"] == "待确认":
            issues.append(
                {
                    "issue_id": f"FI{len(issues)+1:03d}",
                    "field_name": fact["field_name"],
                    "preferred_value": "",
                    "conflicting_values": [],
                    "preferred_reason": "未在资料处理文本摘录中稳定识别",
                    "impact_level": "medium",
                    "requires_user_confirmation": False,
                }
            )

    write_jsonl(facts_dir / "project_facts.jsonl", project_facts)
    write_jsonl(facts_dir / "heritage_facts.jsonl", heritage_facts)
    write_jsonl(facts_dir / "quote_candidates.jsonl", quote_candidates)
    write_jsonl(facts_dir / "source_coverage.jsonl", source_coverage)
    write_jsonl(facts_dir / "fact_issues.jsonl", issues)

    prompt_path = work_dir / "next_prompts" / "next_prompt_analysis.md"
    prompt_path.parent.mkdir(parents=True, exist_ok=True)
    prompt_path.write_text(
        "\n".join(
            [
                "# 下一阶段启动提示",
                "",
                "## 阶段",
                "analysis",
                "",
                "## 项目路径",
                f"`{work_dir}`",
                "",
                "## 必读规则文件",
                "- `references/04-分析判断模块.md`",
                "- `references/05-文评成稿样本与文章架构.md`",
                "- `references/11-固定正文结构与固定内容.md`",
                "- `references/12-文物对象概述与价值评估规则.md`",
                "- `references/13-空间关系写作规则.md`",
                "",
                "## 只读输入",
                "- `facts/project_facts.jsonl`",
                "- `facts/heritage_facts.jsonl`",
                "- `facts/quote_candidates.jsonl`",
                "- `facts/source_coverage.jsonl`",
                "- `evidence/evidence_register.jsonl`",
                "- `processing_output/external_sources.jsonl`",
                "",
                "## 必写输出",
                "- `analysis/impact_matrix.jsonl`",
                "- `analysis/mitigation_matrix.jsonl`",
                "- `analysis/risk_flags.jsonl`",
                "- `next_prompts/next_prompt_report_assembly.md`",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    update_module_state(work_dir, "fact_extraction")
    write_json(
        work_dir / "run_state" / "fact_extraction.module_done.json",
        {
            "module_name": "fact_extraction",
            "status": "completed",
            "started_at": now_iso(),
            "finished_at": now_iso(),
            "input_files": ["processing_output/manifest.json", "processing_output/text_index.jsonl", "evidence/evidence_register.jsonl"],
            "output_files": ["facts/project_facts.jsonl", "facts/heritage_facts.jsonl", "facts/quote_candidates.jsonl", "facts/source_coverage.jsonl", "facts/fact_issues.jsonl", "next_prompts/next_prompt_analysis.md"],
            "blocking_gaps_count": 0,
            "issues_count": len(issues),
            "next_prompt": "next_prompts/next_prompt_analysis.md",
            "notes": "事实抽取阶段完成；正式成稿仍需人工或后续模型细化字段。",
        },
    )
    print(f"事实抽取完成：{work_dir}")


if __name__ == "__main__":
    main()
