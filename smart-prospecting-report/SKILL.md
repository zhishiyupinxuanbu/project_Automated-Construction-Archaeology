---
name: smart-prospecting-report
description: 中文显示名：智能勘探报告。生成考古调查勘探报告 DOCX 和检查成果。Use when the user says 我要写报告、生成报告、帮我写考古调查勘探报告、智能勘探报告、考古勘探报告、勘探报告、报告模板匹配、报告专项检查、问题摘要、回函识别，or asks to create a report from a project source folder, manual workbook, smart report workbook, Word report templates, drawings, photos, review replies, coordinates, and company/personnel assets.
---

# Smart Prospecting Report

This skill is a standalone package for generating archaeological survey/prospecting reports. It must not depend on any other skill or on the original `/Users/ivan/Desktop/智能勘探报告` workspace.

## Chinese Display Name

- 中文显示名：`智能勘探报告`。
- 底层 skill `name` 仍保留 `smart-prospecting-report`，以保持 Codex 标准命名、文件夹匹配和跨电脑迁移稳定。
- 中文调用主要依赖 `description` 与 `agents/openai.yaml`，用户可以直接说“智能勘探报告”“我要写报告”“生成报告”。

## Directory Contract

- Keep this skill in the standard structure: `SKILL.md`, `scripts/`, `references/`, `assets/`.
- Read project source materials from a user-provided project folder.
- Write runtime outputs in the current working directory:
  - `过程资料/项目名称/` for manual forms, temporary checks, backups, OCR, region overview artifacts, and issue summaries.
- Write final report DOCX and final checking outputs directly to the user's Desktop by default.
- Never write generated outputs into the source-material folder.
- Never require `基础信息/`, `过程资料/`, or `生成报告/` to be bundled inside the skill.

## Bundled Resources

- `scripts/run_smart_report_workflow.py`: main report workflow entrypoint.
- `scripts/create_manual_form_from_project.py`: prefill the user-facing manual workbook from a project folder.
- `scripts/fill_smart_template_from_form.py`: fill a report template from the smart workbook.
- `scripts/check_smart_report.py`: report QA helper.
- `scripts/region_overview_agent.py`: optional region-overview generation helper.
- `assets/templates/reports/`: report baseline templates only.
- `assets/templates/forms/`: user-facing manual workbook template.
- `assets/company-personnel-library/`: company and personnel assets used by report generation.
- `references/`: workflow, field, image, template, and QA rules.

## Required Gate Order

For a new report run, always keep these gates visible and stop for user confirmation where required:

1. Confirm the project source-material folder.
2. Run preflight on the source folder.
3. Generate and prefill the project-specific manual workbook.
4. Wait for the user to inspect and confirm the manual workbook.
5. Re-read the latest workbook from disk and build the smart report workbook.
6. Confirm the personnel set when multiple sets exist.
7. Match the report template from the confirmed workbook.
8. Tell the user the matched template path and wait for explicit confirmation.
9. Generate the report and final QA outputs.

Do not use compact legacy confirmation workbooks such as `*_人工确认表_精简版.xlsx`.

## Commands

Run commands from the project working directory, not from inside the skill folder.

```bash
python /path/to/smart-prospecting-report/scripts/run_smart_report_workflow.py doctor
python /path/to/smart-prospecting-report/scripts/run_smart_report_workflow.py preflight /path/to/project-source
python /path/to/smart-prospecting-report/scripts/run_smart_report_workflow.py prepare /path/to/project-source --source-confirmed
python /path/to/smart-prospecting-report/scripts/run_smart_report_workflow.py build-form /path/to/manual-form.xlsx
python /path/to/smart-prospecting-report/scripts/run_smart_report_workflow.py recommend /path/to/smart-form.xlsx --smart --personnel-set 人员信息1
python /path/to/smart-prospecting-report/scripts/run_smart_report_workflow.py generate --form /path/to/smart-form.xlsx --project-dir /path/to/project-source --personnel-set 人员信息1 --confirm-template --use-region-api
```

Use `generate` only. Do not use `generate-plan` from this skill.

## References

Read these files only when the task needs the detail:

- `references/新报告接收流程.md`: full intake and confirmation gate workflow.
- `references/智能勘探报告通用流程.md`: end-to-end report workflow.
- `references/智能模板字段计算规则.md`: field derivation rules.
- `references/智能报告替换变量与资料路径对照.md`: placeholder and source-path mapping.
- `references/有分区模板匹配检查.md`: partition template matching checks.
- `references/test-plan.md`: install, structure, doctor, smoke, and trigger tests.
- `references/safety-checklist.md`: package, runtime, API, and content safety checks.

## Local Versioning

- `VERSION`: current local package version, currently `v1.0.1`.
- `CHANGELOG.md`: release-style change history.
- `DEVELOPMENT_LOG.md`: development decisions, validation notes, and maintenance log.
- `MANIFEST.json`: file inventory with sizes and SHA-256 hashes for migration integrity.
- `scripts/verify_manifest.py`: verifies the manifest after copying or installing this skill.

When changing bundled scripts, references, templates, or assets, update the version files, regenerate `MANIFEST.json`, and rerun validation.

## Validation

Before delivery or after moving this skill to another computer, run:

```bash
python /path/to/smart-prospecting-report/scripts/self_test_skill.py
python /path/to/smart-prospecting-report/scripts/verify_manifest.py
python /path/to/smart-prospecting-report/scripts/run_smart_report_workflow.py doctor
```

If dependencies are missing, install packages listed in `scripts/requirements.txt`.
