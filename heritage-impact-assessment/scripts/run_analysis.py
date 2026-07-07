#!/usr/bin/env python3
"""Run v0.2.2 analysis from extracted facts."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

from pipeline_common import now_iso, read_jsonl, update_module_state, write_json, write_jsonl


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="运行文物影响评估 v0.2.2 分析判断阶段")
    parser.add_argument("--工作目录", dest="work_dir", required=True)
    parser.add_argument("--覆盖", dest="overwrite", action="store_true")
    return parser.parse_args()


def value(rows: list[dict], field_name: str, default: str = "待确认") -> str:
    for row in rows:
        if row.get("field_name") == field_name:
            return row.get("value") or default
    return default


def load_chapter5_pairs(work_dir: Path) -> list[dict]:
    human_input = work_dir / "human_input"
    json_path = human_input / "chapter5_fact_rule_pairs.json"
    csv_path = human_input / "chapter5_fact_rule_pairs.csv"
    if json_path.exists():
        data = json.loads(json_path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return list(data.get("pairs") or [])
        if isinstance(data, list):
            return data
    if csv_path.exists():
        with csv_path.open(encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))
    return []


def main() -> None:
    args = parse_args()
    work_dir = Path(args.work_dir).expanduser().resolve()
    analysis_dir = work_dir / "analysis"
    if analysis_dir.exists() and any(analysis_dir.iterdir()) and not args.overwrite:
        raise SystemExit(f"analysis 目录非空，如需覆盖请添加 --覆盖：{analysis_dir}")
    analysis_dir.mkdir(parents=True, exist_ok=True)

    project = read_jsonl(work_dir / "facts" / "project_facts.jsonl")
    heritage = read_jsonl(work_dir / "facts" / "heritage_facts.jsonl")
    chapter5_pairs = load_chapter5_pairs(work_dir)
    heritage_name = value(heritage, "文物名称")
    relation = value(heritage, "空间关系")

    impact_matrix = [
        {
            "impact_id": "IM-001",
            "heritage_object": heritage_name,
            "project_action": "建设项目设计与新建/改扩建活动",
            "spatial_relation": relation,
            "phase": "设计阶段",
            "impact_pathway": "工程体量、高度、色彩、外立面和夜景照明可能影响文物相关历史风貌。",
            "evidence_ids": ["E0001"],
            "preliminary_judgment": "可能存在风貌影响，需结合最终设计文件专业复核。",
            "professional_review_required": True,
            "judgment_limit": "不得据此替代最终审批或专业负责人结论。",
            "mitigation_direction": "控制高度、体量、色彩、照明和景观设施。",
            "confirmation_points": "最终总平、立面、高度和照明方案。",
        },
        {
            "impact_id": "IM-002",
            "heritage_object": heritage_name,
            "project_action": "施工活动",
            "spatial_relation": relation,
            "phase": "施工期",
            "impact_pathway": "机械、扬尘、排水、临时堆载和人员活动可能造成施工扰动。",
            "evidence_ids": ["E0001"],
            "preliminary_judgment": "直接扰动风险需结合保护范围和施工组织设计判断。",
            "professional_review_required": True,
            "judgment_limit": "缺少施工组织设计时不得写最终等级。",
            "mitigation_direction": "设置边界、围挡、巡查、扬尘排水控制和交底机制。",
            "confirmation_points": "施工便道、堆料区、排水方向和禁入边界。",
        },
        {
            "impact_id": "IM-003",
            "heritage_object": heritage_name,
            "project_action": "运营活动",
            "spatial_relation": relation,
            "phase": "运营期",
            "impact_pathway": "人流车流、照明、垃圾排水和活动外溢可能累积影响周边环境秩序。",
            "evidence_ids": ["E0001"],
            "preliminary_judgment": "运营期影响以管理风险为主，需纳入日常管控。",
            "professional_review_required": True,
            "judgment_limit": "需结合运营方案和主管部门意见确认。",
            "mitigation_direction": "限定活动范围、规范导视照明、建立巡查和告知机制。",
            "confirmation_points": "运营规模、开放时段、照明和排水方案。",
        },
    ]
    mitigation_matrix = [
        {
            "mitigation_id": "MT-001",
            "impact_id": "IM-001",
            "heritage_object": heritage_name,
            "phase": "设计阶段",
            "risk_source": "建控地带内工程设计",
            "control_measure": "最终总平、建筑高度、体量、外立面、屋面形式、色彩和夜景照明应提交专业复核。",
            "responsible_party": "建设单位/设计单位",
            "monitoring_or_acceptance": "形成报审图纸和风貌控制说明。",
            "evidence_ids": ["E0001"],
            "professional_confirmation_points": "建筑高度和风貌控制指标。",
        },
        {
            "mitigation_id": "MT-002",
            "impact_id": "IM-002",
            "heritage_object": heritage_name,
            "phase": "施工期",
            "risk_source": "施工扰动",
            "control_measure": "施工前明确文物本体、保护范围、建设控制地带和施工边界，严禁越界施工、堆料和弃土。",
            "responsible_party": "建设单位/施工单位/监理单位",
            "monitoring_or_acceptance": "交底记录、围挡照片、巡查台账。",
            "evidence_ids": ["E0001"],
            "professional_confirmation_points": "施工组织设计和现场边界。",
        },
        {
            "mitigation_id": "MT-003",
            "impact_id": "IM-003",
            "heritage_object": heritage_name,
            "phase": "运营期",
            "risk_source": "运营活动外溢",
            "control_measure": "限定人员活动范围，规范导视、照明、垃圾收集和排水，建立日常巡查机制。",
            "responsible_party": "建设单位/运营管理单位",
            "monitoring_or_acceptance": "运营巡查记录和问题整改台账。",
            "evidence_ids": ["E0001"],
            "professional_confirmation_points": "运营管理制度。",
        },
    ]
    write_jsonl(analysis_dir / "impact_matrix.jsonl", impact_matrix)
    write_jsonl(analysis_dir / "mitigation_matrix.jsonl", mitigation_matrix)
    write_jsonl(analysis_dir / "chapter5_fact_rule_pairs.jsonl", chapter5_pairs)
    write_jsonl(
        analysis_dir / "chapter5_pairing_issues.jsonl",
        []
        if chapter5_pairs
        else [
            {
                "issue_id": "C5P001",
                "issue": "未找到第五章事实条文人工配对结果；正式生成第五章判断前应先完成 human_input/chapter5_fact_rule_pairs.json。",
                "blocking_for_formal_report": True,
            }
        ],
    )
    write_jsonl(analysis_dir / "risk_flags.jsonl", [{"risk_id": "R001", "risk": "所有影响判断均为初步判断，需专业负责人确认。"}])
    write_jsonl(analysis_dir / "analysis_notes.jsonl", [{"note_id": "AN001", "note": "分析矩阵由事实抽取成果生成，需结合正式模板扩写。"}])
    write_jsonl(analysis_dir / "expert_review_points.jsonl", [{"point_id": "ER001", "point": "重点复核空间关系、建控要求、设计高度体量和施工组织。"}])

    prompt = work_dir / "next_prompts" / "next_prompt_report_assembly.md"
    prompt.write_text(
        "\n".join(
            [
                "# 下一阶段启动提示",
                "",
                "## 阶段",
                "report_assembly",
                "",
                "## 项目路径",
                f"`{work_dir}`",
                "",
                "## 必读规则文件",
                "- `references/05-报告拼装与版式模块.md`",
                "- `references/05-文评成稿样本与文章架构.md`",
                "- `references/10-docx成稿规则.md`",
                "- `references/11-固定正文结构与固定内容.md`",
                "- `references/14-Word字体与版式迁移要求.md`",
                "",
                "## 必写输出",
                "- `report_clean.md`",
                "- `report_with_evidence.md`",
                "- `report_clean.docx`",
                "- `report_with_evidence.docx`",
                "- `report_evidence_map.jsonl`",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    update_module_state(work_dir, "analysis")
    write_json(
        work_dir / "run_state" / "analysis.module_done.json",
        {
            "module_name": "analysis",
            "status": "completed",
            "started_at": now_iso(),
            "finished_at": now_iso(),
            "input_files": ["facts/project_facts.jsonl", "facts/heritage_facts.jsonl", "human_input/chapter5_fact_rule_pairs.json"],
            "output_files": ["analysis/impact_matrix.jsonl", "analysis/mitigation_matrix.jsonl", "analysis/chapter5_fact_rule_pairs.jsonl", "analysis/chapter5_pairing_issues.jsonl", "next_prompts/next_prompt_report_assembly.md"],
            "blocking_gaps_count": 0,
            "issues_count": 0,
            "next_prompt": "next_prompts/next_prompt_report_assembly.md",
            "notes": "分析判断阶段完成。",
        },
    )
    print(f"分析判断完成：{work_dir}")


if __name__ == "__main__":
    main()
