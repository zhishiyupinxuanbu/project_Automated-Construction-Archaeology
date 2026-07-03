---
name: rename-blue-point-images
description: Rename survey, archaeology, drafting, or standard-hole image files according to the blue point label and the unit cell containing the blue point, optionally using 标准孔坐标.xlsx as the authoritative unit mapping, and correct standard-hole photo orientation from visible TK labels. Use when Codex needs to inspect image files such as .jpg/.jpeg/.png, read labels like A031B017 next to a blue point marker, identify units like U01-U11 from a coordinate workbook or drafting grid lines, build names like U01-A031B017.jpg, rotate field photos such as TK01.jpg-TK10.jpg so the TK label is upright at the top, and safely batch rename or rotate files.
---

# Rename Blue Point Images

## Overview

Use this skill to turn image filenames such as `1.jpg`, `2.jpg`, or generic exports into names taken from the blue point label and the unit containing that point, for example `U01-A031B017.jpg`. Also use it to normalize standard-hole field-photo orientation so the white `TKxx` label is upright at the top of each photo.

The workflow is intentionally verification-first. Do not rely on OCR alone for final naming; use visual inspection or a contact sheet for the point label. If `标准孔坐标.xlsx` exists, use it as the authoritative source for unit prefixes; otherwise run image-based unit detection.

## Workflow

### Standard-Hole Photo Orientation

1. Preview proposed rotations:

```bash
python3 scripts/rename_blue_point_images.py orient-photos "/path/to/12.标准孔照"
```

2. Review the output. The script detects the bright white `TKxx` label component and chooses the rotation that places it horizontally near the top.

3. Apply only after the dry run looks right:

```bash
python3 scripts/rename_blue_point_images.py orient-photos "/path/to/12.标准孔照" --apply
```

4. Generate a contact sheet for QA:

```bash
python3 scripts/rename_blue_point_images.py contact-sheet "/path/to/12.标准孔照" --output "/tmp/oriented-standard-hole-photos.jpg"
```

If a photo has no detectable `TKxx` label or the label is obscured, inspect it manually and rotate it separately.

For known exceptions, pass manual overrides as counterclockwise degrees:

```bash
python3 scripts/rename_blue_point_images.py orient-photos "/path/to/11.标准孔照" --rotate TK12.jpg:90 --apply
```

### Blue-Point Drawing Renaming

1. Inventory the target folder and count image files:

```bash
find "/path/to/folder" -maxdepth 1 -type f | sort
```

2. Generate a contact sheet for visual reading:

```bash
python3 scripts/rename_blue_point_images.py contact-sheet "/path/to/folder" --output "/tmp/blue-point-contact.jpg"
```

Open or view the contact sheet. The intended label is the blue text adjacent to the saturated blue point marker, not magenta unit labels such as `U01`, `U02`, or red/purple boundary labels.

3. Create a CSV mapping with `source,target` columns. The `target` can be only the point label at first; `--auto-unit` will add the `Uxx-` prefix.

```csv
source,target
1.jpg,A031B017.jpg
2.jpg,A084B066.jpg
```

Use `assets/mapping_template.csv` as the starting shape if useful.

4. If a coordinate workbook exists, preview unit mapping from it:

```bash
python3 scripts/rename_blue_point_images.py rename "/path/to/folder" \
  --mapping mapping.csv \
  --auto-unit \
  --unit-workbook "/path/to/标准孔坐标.xlsx"
```

The workbook path is preferred for irregular unit layouts or merged cells.

5. If no workbook exists, preview automatic unit detection:

```bash
python3 scripts/rename_blue_point_images.py detect-units "/path/to/folder" --mapping mapping.csv
```

The output includes `source,unit,label,target,x,y`. Check any low-confidence or visually ambiguous boundary cases before applying.

6. Validate with a dry run that prefixes the detected unit:

```bash
python3 scripts/rename_blue_point_images.py rename "/path/to/folder" --mapping mapping.csv --auto-unit
```

7. Apply only after the dry run has no missing files, duplicate targets, invalid labels, target collisions, or suspect unit detections:

```bash
python3 scripts/rename_blue_point_images.py rename "/path/to/folder" --mapping mapping.csv --auto-unit --unit-workbook "/path/to/标准孔坐标.xlsx" --apply
```

8. Verify the final directory:

```bash
find "/path/to/folder" -maxdepth 1 -type f | sort
```

## Label Rules

- Default final target label pattern is `U\d{2}-A\d{3}B\d{3}`, such as `U01-A031B017`.
- If `标准孔坐标.xlsx` has `勘探单元` and `标准孔编号` columns, use that mapping first and carry merged/blank unit cells downward.
- The unit is detected from the saturated blue point center and the purple unit grid lines.
- If the point center is on or very near a unit boundary, inspect the image manually and override the mapping if needed.
- Preserve uppercase `A` and `B`.
- Preserve image extensions unless a project has a different naming convention.
- If a label is unclear, inspect the original image at high detail before adding it to the mapping.
- If two images appear to have the same blue point label, stop and ask the user how to handle the duplicate.

## Safety Rules

- Always dry-run before `--apply`.
- Never overwrite an existing target file.
- Never rename partially if validation fails.
- If working in a protected or source-material folder, mention the exact folder and the planned mapping before applying changes.
- For live project materials, keep a small audit note in the conversation with the mapping that was applied.
- For photo orientation, start with dry run and only write when the planned rotations are plausible.
- If the automatic label heuristic is wrong for a specific photo, use `--rotate filename:90/180/270` and verify with a contact sheet.

## Bundled Script

`scripts/rename_blue_point_images.py` provides:

- `contact-sheet`: creates a labeled contact sheet for visual review.
- `detect-units`: detects blue point centers, infers `U01`-`U06`, and prints proposed final targets.
- `rename --auto-unit`: validates a CSV mapping, prefixes units from `--unit-workbook` or detected image geometry, and performs an atomic two-stage rename when `--apply` is passed.
- `orient-photos`: detects `TKxx` label orientation, accepts manual `--rotate filename:degrees` overrides, and rotates standard-hole field photos in place when `--apply` is passed.

The script requires Python 3. `contact-sheet`, `detect-units`, and image-based `rename --auto-unit` require Pillow (`PIL`) and NumPy. `--unit-workbook` requires openpyxl. Plain `rename` without `--auto-unit` uses only the Python standard library.
