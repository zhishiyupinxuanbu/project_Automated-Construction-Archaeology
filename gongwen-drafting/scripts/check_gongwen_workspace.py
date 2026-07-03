#!/usr/bin/env python3
"""Lightweight health check for the gongwen drafting workspace."""

from __future__ import annotations

import re
import sys
from pathlib import Path


WORKSPACE_MARKERS = ["AGENTS.md", "agent/gongwen_agent.py", "0.资料索引/00-AI工作台.md"]
SKILL_MARKERS = ["SKILL.md", "references/routing.md", "assets/templates"]
WORKSPACE_REQUIRED_PATHS = [
    "AGENTS.md",
    "00_开始使用_请先读.md",
    "agent_handoff/AI_AGENT_START_HERE.md",
    "agent_handoff/PROJECT_MEMORY_EXPORT.md",
    "agent_handoff/skills/gongwen-drafting/SKILL.md",
    "agent_handoff/skills/gongwen-hecha-fixed-format/SKILL.md",
    "agent/check_template_route.py",
    "agent/gongwen_agent.py",
    "agent/gongwen_format.py",
    "agent/make_hecha_docx.py",
    "agent/check_hecha_docx.py",
    "0.资料索引/02-模板索引.md",
    "0.资料索引/文种手册/00-文种手册总览.md",
    "0.资料索引/健康检查/2026-06-15-资料库规范化日志.md",
]
SKILL_REQUIRED_PATHS = [
    "SKILL.md",
    "references/routing.md",
    "references/drafting-rules.md",
    "references/word-output.md",
    "references/resource-map.md",
    "references/project-rules/AGENTS.md",
    "references/project-handoff/PROJECT_MEMORY_EXPORT.md",
    "references/knowledge-index/文种手册/00-文种手册总览.md",
    "references/policies-and-standards",
    "references/reference-documents",
    "references/fixed-format",
    "assets/fixed-format",
    "assets/fonts",
    "assets/templates",
    "scripts/gongwen-agent/check_template_route.py",
    "scripts/gongwen-agent/gongwen_agent.py",
    "scripts/gongwen-agent/gongwen_format.py",
    "scripts/gongwen-agent/make_hecha_docx.py",
    "scripts/gongwen-agent/check_hecha_docx.py",
    "agents/openai.yaml",
]
SKILL_FORBIDDEN_PATHS = [
    "assets/templates/模板——市级调查报告——长呼复线大东管道联络线项目文物调查报告20260402.docx",
    "assets/templates/模版——自治区——关于XXX项目文物调查的报告.docx",
    "assets/reference-documents/10.文物调查报告",
    "assets/reference-documents",
    "assets/policies-and-standards",
    "references/knowledge-index/文种手册/文物调查报告-作战手册.md",
    "references/knowledge-index/按文种/文物调查报告.md",
    "references/knowledge-index/03-待补充文本清单.md",
    "scripts/gongwen-agent/wecom_adapter",
    "scripts/gongwen-agent/迁移打包工具",
]
SKILL_STALE_TEXT_PATHS = [
    "references/knowledge-index/模板内化/manifest.json",
    "references/knowledge-index/模板内化/00-模板内化总览.md",
    "references/project-handoff/AI_AGENT_START_HERE.md",
    "references/project-handoff/PROJECT_MEMORY_EXPORT.md",
    "references/project-rules/AGENTS.md",
]
SKILL_STALE_PHRASES = [
    "文物调查报告-作战手册.md",
    "模板——市级调查报告——长呼复线大东管道联络线项目文物调查报告20260402.docx",
    "模版——自治区——关于XXX项目文物调查的报告.docx",
    "0.资料索引/模板内化/模版——自治区——关于XXX项目文物调查的报告.md",
]
FORBIDDEN_TERMS = ["调查面积", "勘探面积"]
KEY_PHRASES = [
    "请示类文件正文必须先写背景",
    "默认保存到电脑桌面",
    "check_template_route.py",
    "固定版式 Word 样张",
    "红头、标题、正文、附件、落款",
    "字体字号、段落缩进、附件悬挂缩进",
]
SKILL_LAYOUT_REQUIRED_PATHS = [
    "assets/fixed-format/固定版式——文物保护许可核查请示格式范本.docx",
    "references/fixed-format/固定版式——文物保护许可核查请示格式范本.版式校验.md",
    "scripts/gongwen-agent/固定版式说明.md",
]
SKILL_LAYOUT_RULE_FILES = [
    "SKILL.md",
    "references/word-output.md",
    "references/project-rules/AGENTS.md",
    "scripts/gongwen-agent/固定版式说明.md",
    "references/fixed-format/固定版式——文物保护许可核查请示格式范本.版式校验.md",
]
SKILL_STALE_LAYOUT_PHRASES = [
    "python3 agent/gongwen_agent.py",
    "python3 agent/make_hecha_docx.py",
    "python3 agent/check_hecha_docx.py",
    "python3 agent/check_template_route.py",
    "`agent/固定版式说明.md`",
    "agent_handoff/skills/gongwen-drafting/scripts/check_gongwen_workspace.py",
    "1.政策法规与规范/固定版式——文物保护许可核查请示格式范本.docx",
]


def find_root(start: Path) -> tuple[Path, str]:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if all((candidate / marker).exists() for marker in WORKSPACE_MARKERS):
            return candidate, "workspace"
        if all((candidate / marker).exists() for marker in SKILL_MARKERS):
            return candidate, "skill"
    raise SystemExit("ERROR: run this script inside 公文撰写资料库 or gongwen-drafting skill.")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def main() -> int:
    root, mode = find_root(Path.cwd())
    errors: list[str] = []
    warnings: list[str] = []

    required_paths = WORKSPACE_REQUIRED_PATHS if mode == "workspace" else SKILL_REQUIRED_PATHS
    for rel in required_paths:
        if not (root / rel).exists():
            errors.append(f"missing required path: {rel}")
    if mode == "skill":
        for rel in SKILL_FORBIDDEN_PATHS:
            if (root / rel).exists():
                errors.append(f"survey report resource should live in gongwen-survey-report: {rel}")
        for rel in SKILL_LAYOUT_REQUIRED_PATHS:
            if not (root / rel).exists():
                errors.append(f"missing fixed-format layout resource: {rel}")
        for rel in SKILL_STALE_TEXT_PATHS:
            path = root / rel
            if path.exists():
                content = read_text(path)
                for phrase in SKILL_STALE_PHRASES:
                    if phrase in content:
                        errors.append(f"{rel} contains stale survey-report reference: {phrase}")
        for rel in SKILL_LAYOUT_RULE_FILES:
            path = root / rel
            if path.exists():
                content = read_text(path)
                for phrase in KEY_PHRASES[3:]:
                    if phrase not in content:
                        errors.append(f"{rel} missing fixed-format rule phrase: {phrase}")
                for phrase in SKILL_STALE_LAYOUT_PHRASES:
                    if phrase in content:
                        errors.append(f"{rel} contains stale layout path or command: {phrase}")

    agents = root / ("AGENTS.md" if mode == "workspace" else "references/project-rules/AGENTS.md")
    if agents.exists():
        text = read_text(agents)
        for phrase in KEY_PHRASES:
            if phrase not in text:
                errors.append(f"AGENTS.md missing key phrase: {phrase}")
    else:
        text = ""

    skill_prefix = "agent_handoff/skills/gongwen-drafting/" if mode == "workspace" else ""
    for rel in [
        f"{skill_prefix}SKILL.md",
        f"{skill_prefix}references/drafting-rules.md",
    ]:
        path = root / rel
        if path.exists():
            content = read_text(path)
            for term in FORBIDDEN_TERMS:
                if term not in content:
                    warnings.append(f"{rel} does not mention forbidden term: {term}")

    placeholders = []
    for rel in [
        f"{skill_prefix}SKILL.md",
        f"{skill_prefix}references/routing.md",
        f"{skill_prefix}references/drafting-rules.md",
        f"{skill_prefix}references/word-output.md",
        f"{skill_prefix}references/resource-map.md",
    ]:
        path = root / rel
        if path.exists():
            for idx, line in enumerate(read_text(path).splitlines(), start=1):
                if re.search(r"\bTODO\b|\[TODO", line):
                    placeholders.append(f"{rel}:{idx}")
    if placeholders:
        errors.append("placeholder text remains: " + ", ".join(placeholders))

    print(f"{mode}: {root}")
    if warnings:
        print("warnings:")
        for warning in warnings:
            print(f"  - {warning}")
    if errors:
        print("errors:")
        for error in errors:
            print(f"  - {error}")
        return 1
    print(f"OK: gongwen drafting {mode} checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
