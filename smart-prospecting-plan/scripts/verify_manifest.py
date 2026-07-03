#!/usr/bin/env python3
"""Verify local version-control manifest for this skill package."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path


EXCLUDED_DIRS = {"__pycache__"}
EXCLUDED_NAMES = {".DS_Store", ".env.local"}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def iter_files(root: Path):
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        rel_parts = path.relative_to(root).parts
        if any(part in EXCLUDED_DIRS for part in rel_parts):
            continue
        if path.name in EXCLUDED_NAMES or path.suffix == ".pyc" or path.name.startswith("~$"):
            continue
        rel = path.relative_to(root).as_posix()
        if rel == "MANIFEST.json":
            continue
        yield rel, path


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify a smart prospecting skill manifest.")
    parser.add_argument("--strict", action="store_true", help="fail when extra files not listed in MANIFEST.json are present")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    manifest_path = root / "MANIFEST.json"
    version_path = root / "VERSION"
    errors: list[str] = []
    warnings: list[str] = []

    if not manifest_path.exists():
        print("FAIL")
        print("- MANIFEST.json is missing")
        return 1
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))

    if version_path.exists():
        version_text = version_path.read_text(encoding="utf-8").strip()
        if version_text != manifest.get("version"):
            errors.append(f"VERSION mismatch: VERSION={version_text}, MANIFEST={manifest.get('version')}")
    else:
        errors.append("VERSION is missing")

    expected = {item["path"]: item for item in manifest.get("files", [])}
    actual = {rel: path for rel, path in iter_files(root)}

    for rel, item in sorted(expected.items()):
        path = root / rel
        if not path.exists():
            errors.append(f"missing file: {rel}")
            continue
        size = path.stat().st_size
        if size != item.get("size_bytes"):
            errors.append(f"size mismatch: {rel}: {size} != {item.get('size_bytes')}")
        digest = sha256_file(path)
        if digest != item.get("sha256"):
            errors.append(f"sha256 mismatch: {rel}")

    extras = sorted(set(actual) - set(expected))
    if extras:
        message = "extra files not in manifest: " + ", ".join(extras[:20])
        if args.strict:
            errors.append(message)
        else:
            warnings.append(message)

    if errors:
        print("FAIL")
        for item in errors:
            print(f"- {item}")
        for item in warnings:
            print(f"warning: {item}")
        return 1

    print(f"OK {manifest.get('skill_name')} {manifest.get('version')}")
    print(f"tracked_files={len(expected)} generated_at={manifest.get('generated_at')}")
    for item in warnings:
        print(f"warning: {item}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
