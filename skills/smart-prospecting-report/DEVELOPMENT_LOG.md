# Development Log

## 2026-07-03 - v1.0.1 Chinese display-name metadata

### Goal

Make the skill easier to recognize and invoke in Chinese while preserving Codex-compatible ASCII skill names.

### Decision

- Do not rename `name:` or the skill folder to Chinese. The skill-creator naming rule expects lowercase letters, digits, and hyphens.
- Add `agents/openai.yaml` with a Chinese `display_name`, short description, and default prompt.
- Add Chinese aliases to the `description` field because description is the actual trigger surface.

### Validation

Run `scripts/self_test_skill.py`, `scripts/verify_manifest.py --strict`, and `scripts/run_smart_report_workflow.py doctor` after regenerating `MANIFEST.json`.

## 2026-06-29 - v1.0 formal local delivery version

### Goal

Mark `smart-prospecting-report` as the current formal local delivery version for colleague installation and cross-computer migration. Preserve traceability without introducing Git as a required maintenance tool.

### What Changed

- Set `VERSION` to `v1.0`.
- Rebuilt `CHANGELOG.md` around the formal `v1.0` delivery version.
- Rebuilt `MANIFEST.json` so every tracked file has current size and SHA-256 hash values.
- Kept the standalone package contract: no dependency on the original `/Users/ivan/Desktop/智能勘探报告` workspace.
- Kept default final 报告 output on the user's Desktop.

### Safety Notes

- Do not store project source materials, process outputs, generated reports/plans, API keys, or local `.env` files in the skill package.
- Keep `assets/` limited to reusable templates, form templates, company/personnel assets, and approved static libraries.
- Keep runtime process files in the current working directory under `过程资料/项目名称/`.
- Keep final 报告 outputs on the user's Desktop by default.

### Validation To Run After Any Copy Or Update

```bash
python3 ~/.codex/skills/smart-prospecting-report/scripts/self_test_skill.py
python3 ~/.codex/skills/smart-prospecting-report/scripts/verify_manifest.py --strict
python3 ~/.codex/skills/smart-prospecting-report/scripts/run_smart_report_workflow.py doctor
```

### Next Maintenance Rule

Whenever scripts, references, templates, or bundled assets change, update `VERSION`, add a `CHANGELOG.md` entry, add a `DEVELOPMENT_LOG.md` entry, regenerate `MANIFEST.json`, and rerun self-test plus manifest verification.
