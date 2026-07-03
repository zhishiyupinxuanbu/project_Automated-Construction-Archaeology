# Safety Checklist

Use this checklist before publishing, copying, installing, or running `smart-prospecting-report`.

## Package Safety

- [ ] Top level contains only `SKILL.md`, `scripts/`, `references/`, and `assets/`.
- [ ] `SKILL.md` has `name: smart-prospecting-report`.
- [ ] The skill does not contain `基础信息/`, `过程资料/`, `生成报告/`, `课题申报/`, or real project source folders.
- [ ] The skill does not contain generated reports, checking reports, issue summaries, OCR outputs, or temporary diagnostics.
- [ ] The skill does not contain `.env.local`, API keys, passwords, tokens, cookies, or personal credentials.
- [ ] The skill does not contain `.DS_Store`, `__pycache__`, `*.pyc`, lock files, or office temp files such as `~$*.docx`.
- [ ] Report templates live only under `assets/templates/reports/`.
- [ ] Plan templates are not bundled in this report skill.
- [ ] Company/personnel assets are under `assets/company-personnel-library/`.

## Runtime Safety

- [ ] Confirm the project source folder before preflight or preparation.
- [ ] Run preflight before generating or filling workbooks.
- [ ] Do not write outputs into the project source folder.
- [ ] Do not edit source materials in place.
- [ ] Stop for manual workbook confirmation before building the smart workbook.
- [ ] Stop for personnel-set confirmation when multiple sets exist.
- [ ] Stop for matched-template confirmation before formal generation.
- [ ] Leave process artifacts under the working directory's `过程资料/`.
- [ ] Write final report outputs to the user's Desktop by default.

## API Safety

- [ ] Do not package `.env.local`.
- [ ] If region overview API is used, send only the county/banner name and generic writing rules.
- [ ] Never send project name, coordinates, exact location, construction content, workbook content, images, or other private project materials to external APIs.

## Content Safety

- [ ] Do not invent missing review-reply titles, document numbers, or conclusions.
- [ ] If a field, table, reply, or image cannot be matched reliably, mark/report the issue instead of hiding it.
- [ ] Do not silently switch to plan generation from this report skill.
