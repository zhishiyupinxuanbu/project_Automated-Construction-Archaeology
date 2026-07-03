#!/usr/bin/env python3
"""Validate the installed smart-prospecting-plan skill package."""

from __future__ import annotations

import sys
from pathlib import Path


SKILL_NAME = "smart-prospecting-plan"
EXPECTED_TOP = {"SKILL.md", "scripts", "references", "assets", "agents", "VERSION", "CHANGELOG.md", "DEVELOPMENT_LOG.md", "MANIFEST.json"}
FORBIDDEN_NAMES = {
    ".DS_Store",
    ".env.local",
    "__pycache__",
    "基础信息",
    "过程资料",
    "生成报告",
    "课题申报",
}


def fail(message: str, errors: list[str]) -> None:
    errors.append(message)


def main() -> int:
    skill_root = Path(__file__).resolve().parents[1]
    errors: list[str] = []

    top_names = {path.name for path in skill_root.iterdir()}
    extra = sorted(top_names - EXPECTED_TOP)
    missing = sorted(EXPECTED_TOP - top_names)
    if extra:
        fail(f"top-level unexpected entries: {', '.join(extra)}", errors)
    if missing:
        fail(f"top-level missing entries: {', '.join(missing)}", errors)

    skill_md = skill_root / "SKILL.md"
    if not skill_md.exists():
        fail("SKILL.md is missing", errors)
    else:
        text = skill_md.read_text(encoding="utf-8", errors="ignore")
        if f"name: {SKILL_NAME}" not in text:
            fail(f"SKILL.md name is not {SKILL_NAME}", errors)
        if "我要写计划" not in text or "生成计划" not in text:
            fail("SKILL.md description lacks expected Chinese plan triggers", errors)

    for path in skill_root.rglob("*"):
        if path.name in FORBIDDEN_NAMES or path.suffix == ".pyc" or path.name.startswith("~$"):
            fail(f"forbidden package entry: {path.relative_to(skill_root)}", errors)

    required_paths = [
        "scripts/run_smart_report_workflow.py",
        "scripts/create_manual_form_from_project.py",
        "scripts/fill_smart_template_from_form.py",
        "scripts/requirements.txt",
        "references/test-plan.md",
        "references/safety-checklist.md",
        "assets/templates/forms/人工填写表模板.xlsx",
        "assets/company-personnel-library",
        "assets/templates/plans",
    ]
    for rel in required_paths:
        if not (skill_root / rel).exists():
            fail(f"required path missing: {rel}", errors)

    traceability_paths = [
        "VERSION",
        "CHANGELOG.md",
        "DEVELOPMENT_LOG.md",
        "MANIFEST.json",
        "scripts/verify_manifest.py",
        "agents/openai.yaml",
    ]
    for rel in traceability_paths:
        if not (skill_root / rel).exists():
            fail(f"traceability path missing: {rel}", errors)

    version_path = skill_root / "VERSION"
    if version_path.exists() and not version_path.read_text(encoding="utf-8").strip():
        fail("VERSION is empty", errors)

    plan_templates = list((skill_root / "assets/templates/plans").glob("*.docx"))
    if not plan_templates:
        fail("no plan templates found", errors)
    report_template_dir = skill_root / "assets/templates/reports"
    if report_template_dir.exists() and list(report_template_dir.glob("*.docx")):
        fail("report templates must not be bundled in the plan skill", errors)

    if errors:
        print("FAIL")
        for item in errors:
            print(f"- {item}")
        return 1
    print("OK smart-prospecting-plan skill package")
    return 0


if __name__ == "__main__":
    sys.exit(main())
