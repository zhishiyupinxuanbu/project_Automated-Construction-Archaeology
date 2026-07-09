---
name: administration
description: Build exe-ready text-volume archive materials from manually provided candidate files. Use when the user wants 行政卷/行政续补/法律文书卷/参考资料卷/论文卷 materials converted from PDFs or image-page folders into administration_data_chinese.xlsx plus images/1/<序号.标准题名>/page_001.png, with field cleanup, page ordering, and audit/review outputs. 行政卷 may fill 文号; legal/reference/paper volumes leave 文号 blank.
---

# Administration

## Core Rule

Treat the user's input paths as manually provided candidates for one text-style volume. Do not claim the volume set is complete, and do not restart broad raw-material classification unless the user asks for that. Your job is to normalize the given materials into an exe-readable package and surface uncertain fields for review.

This skill is used for the same exe package shape across:

- `administration`: 行政管理文件卷 / 行政续补, may fill `文号` when reliable.
- `legal`: 法律文书卷, leave `文号` blank.
- `reference`: 参考资料卷, leave `文号` blank.
- `paper`: 论文卷, leave `文号` blank.

## Output Contract

Always build this package shape:

- `administration_data_chinese.xlsx`
- `images/1/<序号.标准题名>/page_001.png`, `page_002.png`, ...
- `administration_review.xlsx`
- `audit.json`
- `生成说明.md`

The workbook must have one sheet named `Sheet1` with exactly these headers:

`项目、卷、文号、编制单位（承担单位）、题名、编制时间、批准单位、批准时间、张数、备注、扫描时间、归还时间`

Use `项目=1` by default. Excel `题名` must exactly match the folder name under `images/1/`. `张数` must equal the page image count.
Write `编制时间` as text in Chinese date format, such as `2020年6月18日`, not as an Excel date or `YYYY-MM-DD`. Leave the main Excel `备注` column blank; put issues and review notes only in `administration_review.xlsx`, `audit.json`, and `source_manifest.csv`.

Before writing the workbook, delete confirmed blank page images such as blank back-side scans. Excel `张数` must use the page count after blank-page deletion.

## Input Handling

Support these input shapes:

- One PDF file: one record.
- One folder containing page images such as `image00001.jpg`: one record, sorted by natural numeric order.
- One parent folder containing multiple PDFs and/or image folders: each PDF and each image folder becomes a separate record.
- A folder that has both root-level images and child folders: root-level images become one record; each child folder becomes a separate record; mark this in review.

Accept casual or messy names. Keep the original path in review outputs.

Sort output records by the current file's signing/issue date from earliest to latest, then assign the final `1.`, `2.`, `3.` sequence numbers. Do not keep input order unless dates are missing or tied.

## Field Handling

- Derive the standard title from the file/folder name first.
- Use page text/OCR from all retained pages to extract the current file date. Do not assume the date is on the first or last page; attachments can appear after the signature page.
- For forwarding documents, distinguish the outer/current forwarding document date from the forwarded original document date. Prefer dates before attachments or before the forwarded original text starts; downgrade dates found in attachment tables, forwarded originals, contact blocks, or funding/project lists.
- Preserve document-number bracket style exactly when detected; never normalize `【】`, `〔〕`, `[]`, `（）`, or spaces into another form.
- Do not fill `文号` or `编制单位（承担单位）` from cited policy basis text such as `按照/依据/根据《...》（国发〔2016〕17号）`. Those are references, not the current file's document number or issuer. If the input is only an attachment/body without a red-head or signing block, leave current-file `文号` and issuer blank in the main Excel.
- Leave uncertain main-table fields blank. Put the uncertainty reason in `administration_review.xlsx`; do not write `待核` or reasons into the main Excel `文号`, `编制单位（承担单位）`, or `备注` columns.
- Do not turn a document-number year into a full date unless the source has the actual date.

## Script

Use the bundled script for deterministic package generation:

```bash
python3 ~/.codex/skills/administration/scripts/build_administration_package.py \
  --output /path/to/output-package \
  --volume-type administration \
  /path/to/pdf-or-folder ...
```

Useful options:

- `--volume-type administration|legal|reference|paper`: choose field behavior. Non-administration volumes leave `文号` blank.
- `--project 1`: set the Excel `项目` value and image level, default `1`.
- `--start-seq 1`: first output sequence number, default `1`.
- `--dpi 180`: PDF render DPI, default `180`.
- OCR all retained pages for date candidates by default.
- `--no-ocr-date-pages`: skip all-page OCR and use embedded PDF text / filenames only.
- `--ocr-dir /Users/drevan01/Desktop/OCR`: local OCR root, default matches this workspace.
- The local Paddle OCR service must be running for scanned image/PDF date recognition. If OCR is unavailable, mark dates as 待核 rather than guessing.
- Blank pages are removed by default; use `--keep-blank-pages` only for debugging.

After running, inspect `生成说明.md`, `audit.json`, and `administration_review.xlsx` before calling the package final.

## Validation Checklist

- Every Excel row has a matching `images/1/<题名>/` folder.
- Every folder's page count equals Excel `张数`.
- Page filenames are continuous as `page_001.png`, `page_002.png`, ...
- Output order is by recognized current-file date. Review output must list date candidates and the selected date source, especially for forwarding documents.
- Review output lists original paths, inferred fields, and any issues such as nested folders, page gaps, deleted blank pages, OCR failure, blank dates, or document numbers marked `待核`.
- The final answer must state that v1.0 processes only the given candidate inputs and does not certify administrative completeness.
