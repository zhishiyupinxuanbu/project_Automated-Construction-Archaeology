---
name: project-data-management
description: Use when creating, filling, searching, organizing, or auditing project data archives for Chinese cultural-heritage/project documents. Handles mounted NAS/SMB shares such as DREVAN-P3, project folder scaffolds, placeholder test data, red-head official documents, execution-material formatting, and recurring archive compliance checks against the user's project data template rules.
---

# 项目资料管理

Use this skill for local project data archive work, including NAS shares mounted under `/Volumes`, especially `DREVAN-P3` SMB shares.

## Core Rules

- Treat each mounted SMB share under `/Volumes` as an independent NAS share; there may be no writable `/Volumes/DREVAN-P3` server root.
- Real project folders are named like `YYYYMMDD-项目名称`.
- Build project folder structures from the real template pattern, not only from the written spec.
- User/company-generated red-head official documents must go under `4.成果资料`, usually in the closest workflow subfolder such as `2.三级联调`, `3.用地申请`, or `4.发掘申请`.
- Client-provided original materials are classified by material type, usually under `1.项目资料`, `2.商务资料`, or `3.执行资料`.
- Ask before deleting, overwriting, moving large batches, or renaming real project materials. Test folders under `/Volumes/测试2` may be edited when the user explicitly asks.

For full classification and audit rules, read `references/archive-rules.md`.

## Quick Commands

List mounted NAS shares:

```bash
python /Users/hero/.codex/skills/project-data-management/scripts/nas_archive.py list-mounts
```

Create a template-style project scaffold:

```bash
python /Users/hero/.codex/skills/project-data-management/scripts/nas_archive.py create-project --root /Volumes/测试2 --name xx项目
```

Audit one or more mounted shares:

```bash
python /Users/hero/.codex/skills/project-data-management/scripts/nas_archive.py audit --roots /Volumes/测试2
```

Audit all currently mounted NAS shares:

```bash
python /Users/hero/.codex/skills/project-data-management/scripts/nas_archive.py audit-mounted
```

Default audit reports are saved to:

`/Users/hero/Desktop/NAS归档检查报告/nas-audit-YYYYMMDD.md`

## Workflow

1. Confirm mounted shares with `list-mounts` or `mount | grep smb`.
2. For creating test/project structures, use `create-project`; add placeholder content only when requested.
3. For generated red-head documents, create/save them under `4.成果资料`, not under `1.项目资料`.
4. For audits, run `audit` or `audit-mounted`, then summarize:
   - missing required folders,
   - likely misplaced red-head documents,
   - missing key evidence files,
   - non-template execution folder names,
   - empty required folders in active projects.
5. Do not modify real project files during an audit unless the user asks for fixes.
