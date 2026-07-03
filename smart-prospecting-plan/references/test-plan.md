# Test Plan

Use this test plan after installing or updating `smart-prospecting-plan`.

## 1. Structure Test

Run:

```bash
python ~/.codex/skills/smart-prospecting-plan/scripts/self_test_skill.py
```

Pass criteria:

- Top level contains only `SKILL.md`, `scripts/`, `references/`, and `assets/`.
- `SKILL.md` has `name: smart-prospecting-plan`.
- No `.env.local`, `.DS_Store`, `__pycache__`, `*.pyc`, project source folders, or generated output folders exist inside the skill.
- Plan templates exist under `assets/templates/plans/`.
- No report templates exist under `assets/templates/reports/`.

## 2. Doctor Test

Run from any project working directory:

```bash
python -B ~/.codex/skills/smart-prospecting-plan/scripts/run_smart_report_workflow.py doctor
```

Pass criteria:

- The final plan output directory is the user's Desktop.
- The plan template matrix is covered.
- The form template, company/personnel library, scripts, references, and assets are found.

## 3. Workflow Smoke Test

Use a small desensitized project folder and run:

```bash
python -B ~/.codex/skills/smart-prospecting-plan/scripts/run_smart_report_workflow.py preflight /path/to/project-source
python -B ~/.codex/skills/smart-prospecting-plan/scripts/run_smart_report_workflow.py prepare /path/to/project-source --source-confirmed
```

Pass criteria:

- Source files are not modified.
- Runtime files are created under the working directory's `过程资料/`.
- No runtime files are written into the skill folder.
- Formal generation is not attempted until the manual workbook, personnel set, and matched template are confirmed.

## 4. Trigger Test

Use Chinese prompts such as:

- 我要写计划
- 生成计划
- 帮我写考古调查勘探工作计划
- 勘探计划
- 计划模板匹配

Pass criteria: Codex selects `smart-prospecting-plan`, not the report skill.
