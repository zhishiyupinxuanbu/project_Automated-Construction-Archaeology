---
name: 图纸
description: Use when extracting, checking, deduplicating, naming, sorting, and preparing image files and drawing_data_chinese.xlsx for cultural heritage drawing/photo volumes. Covers PDF/Word image extraction, avoiding black/unreadable images, naming from image titles, using an images folder, de-duplicating identical images, sorting by drawing category/time, and filling Excel paths.
---

# 图纸

Use this skill when the user asks to extract photos/drawings from Word/PDF source files and prepare the `images` folder plus `drawing_data_chinese.xlsx` for a drawing volume.

## Core Rules

- Final image folder must be named exactly `images`.
- Excel image paths must be relative paths: `images/<image filename>`.
- Final deliverables in the user's requested project folder should only be the `images` folder and the filled Excel workbook, unless the user asks for extra reports or review files.
- When extracting from Word documents, extract both the drawing/photo images and the related information needed for Excel.
- If the same folder contains a drawing catalog or drawing list, use it to match image titles/order and fill Excel.
- Check extracted images before delivery. Do not deliver black, blank, unreadable, or broken images.
- Check duplicates by image content, not just filename. If two images are exactly the same, keep one.
- If duplicates are removed, update Excel to match the kept images only.
- Do not invent `绘制时间`, `责任人`, or `绘制人`. Fill them only when clearly present in the source. Otherwise leave blank.

## Preferred Extraction

### Standalone Image Sources

If the user separately provides image files such as `.jpg`, `.jpeg`, `.png`, `.tif`, `.tiff`, `.webp`, or `.bmp`, include them as candidate drawing images.

1. Read the image directly and verify it opens normally.
2. Use the title visible on the image first for naming.
3. If the image has no visible title, use nearby folder/catalog context or the source filename.
4. Include the image in duplicate checking with all PDF/Word extracted images.
5. Save the final kept copy as JPG in `images`.

### PDF Sources

Prefer extracting embedded images from the drawing-sheet pages (`图纸册页`) inside PDFs with a PDF library. For each PDF:

1. Read all pages.
2. Identify pages or text sections marked as `图纸册页`, `图纸编号`, `图纸`, `名称`, or drawing-list entries.
3. Extract embedded images from those drawing-sheet pages first.
4. Do not treat ordinary registration-form photos as drawings unless the user explicitly asks for photos too.
5. Use the largest valid image on the matched drawing-sheet page unless source context indicates otherwise.
6. Save as JPG for compatibility.
7. Verify every output image opens and is not black/blank.

This is safer than binary scraping and avoids black-image problems.

### Word Sources

Word files may be `.docx`, XML-based `.doc`, or old binary `.doc`. For each Word source:

1. Extract drawing-sheet pages (`图纸册页`) and their candidate drawing images.
2. Also read surrounding document text, including nearby paragraphs, captions, tables, and page text around the drawing sheet.
3. Use the visible title on the image first. If the image has no title, use nearby text or captions to identify the image name.
4. Capture related Excel information only when clearly present: `绘制时间`, `责任人`, `绘制人`.
5. Keep source order so untitled images can still be sorted conservatively.
6. Save output as JPG for compatibility.
7. Verify every output image opens and is not black/blank.

When reading Word information fields, normalize field labels before matching:

- Treat labels with spaces as the same field, such as `绘 制 人` = `绘制人`.
- Check common variants such as `绘制人`, `绘 制 人`, `制图人`, `绘图人`, `绘制时间`, `绘 制 时 间`, `制图时间`, `绘图时间`, `责任人`.
- If the Word file uses mail-merge fields, also read the displayed merge-field values, such as `MERGEFIELD draftdrawer` for the drawing person and `MERGEFIELD drafttime` for the drawing time.
- If the same document contains several drawing/image records, match each drawing to the nearest related information row or table. Do not apply a value to unrelated images unless the connection is clear.

For `.docx`, inspect `word/media` and relationships/context when possible.

For XML-based `.doc`, images may be stored as base64 in `w:binData`; decode these embedded images instead of assuming there are no pictures.

Old binary `.doc` files can contain bad preview layers. Avoid blindly scraping image bytes from the file. If direct extraction produces black images:

1. Prefer rendering the page, exporting to PDF, or using another known-good generated source.
2. Match the rendered/extracted image by document/object title.
3. Re-export as JPG.
4. Run brightness/validity checks.

When Word extraction finds identical drawings/photos in different places or documents, keep only one final copy and make Excel point only to that kept copy.

## Image Validation

For every output image:

- Confirm file opens with an image library.
- Confirm dimensions are reasonable.
- Compute average brightness; very low brightness may indicate a black image.
- Visually inspect samples, especially any image flagged as dark.
- If an image is black but technically valid, do not accept it; find a better source or render/extract another way.

## Deduplication

Before finalizing:

1. Convert each image to a normalized RGB representation.
2. Hash pixel bytes plus image size.
3. Treat matching hashes as identical images.
4. Keep the first image in sorted/final order.
5. Remove later duplicates or do not copy them into final `images`.
6. Record duplicate pairs in a short report when useful.

Always run this duplicate audit after Word/PDF extraction and again after final renaming. Identical photos/drawings must not appear twice in `images` or twice in Excel.

## Naming Images

Name images from the title visible on the image whenever possible.

- If image title says `凤凰岭长城1段平面位置图`, filename should be `凤凰岭长城1段平面位置图.jpg`.
- Do not keep extraction suffixes like `_01`.
- Do not use generic names like `image1.jpg` unless the user explicitly wants that.
- If final ordered naming is requested, prefix the sorted number: `1.凤凰岭长城1段平面位置图.jpg`.
- When the image has no visible title, infer from nearby source text or source filename, then keep a conservative title.

For Word sources, prefer naming evidence in this order:

1. title visible inside the image
2. image caption or paragraph immediately before/after the image
3. table row or section title connected to the image
4. source filename plus conservative description

## Drawing Catalog Matching

If the source folder, final folder, or a user-provided "put together" folder contains a drawing catalog/list such as `目录`, `图纸目录`, `卷内目录`, or a Word/Excel table listing drawings:

1. Read the catalog before final naming and Excel filling.
2. Match catalog rows to images by visible image title, source filename, object/site name, and source order.
3. Use the catalog title and order when the catalog clearly corresponds to the image.
4. If the catalog has fields such as drawing name, sequence number, drawing time, responsible person, or drawing person, fill Excel from those fields only when the row clearly matches that image.
5. If the catalog conflicts with the title visible on the image, prefer the visible image title for the filename, and note the mismatch in the report.
6. If a catalog row has no matching image, do not create an Excel row without an image. If an image has no matching catalog row, still include it using the best available image/source title.

The final Excel row order must match the final image order, and catalog-derived information must point to the same image path in `images/<filename>`.

## Sorting Drawings

When the user asks to sort drawing-volume images, use this category order:

1. 总体图纸: 地形地貌图、地质图、行政区域图、文物分布图、保护范围控制地带图
2. 考古图纸: 考古发掘平面图、典型地层剖面图、重要遗迹分布图、横/纵/平剖面图、典型基础图等
3. 建筑图纸: 建筑群体总平面图、单体平面图、立面图、剖面图、结构图、节点大样图等

Within the same category, sort by drawing/source time if clearly present. If time is absent or unreliable, keep source order or filename sequence.

If all extracted images are `平面位置图` and no overall/archaeology keywords appear, classify them as 建筑图纸 unless the user says otherwise.

## Excel Filling

For `drawing_data_chinese.xlsx`, keep columns:

1. `标题`
2. `图片路径`
3. `绘制时间`
4. `责任人`
5. `绘制人`

Fill rows from final image order:

- `标题`: image filename without extension, e.g. `1.凤凰岭长城1段平面位置图`
- `图片路径`: `images/1.凤凰岭长城1段平面位置图.jpg`
- `绘制时间`: fill only if clearly present; otherwise blank
- `责任人`: fill only if clearly present; otherwise blank
- `绘制人`: fill only if clearly present; otherwise blank

For Word sources, fill these fields from visible captions, nearby text, or related tables only when the connection to that image is clear. If the document has general project metadata but it is not clearly tied to the drawing/photo, leave the image row blank for those fields.

When rewriting the workbook, remove old data rows first so stale rows do not remain.

## Final Checks

Before final response, report:

- source file count
- extracted image count
- kept unique image count
- duplicate count, if any
- broken/black image count
- Excel row count, if Excel was filled
- output folder path

Before delivery, clean the user's requested project folder so the final visible result contains only:

1. `images`
2. the filled Excel workbook, usually `drawing_data_chinese.xlsx`

Temporary extraction folders, contact sheets, check reports, scripts, intermediate images, and diagnostic files should be removed or kept outside the project folder unless the user explicitly asks to keep them.
