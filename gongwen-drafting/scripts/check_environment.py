#!/usr/bin/env python3
"""Check runtime dependencies for the self-contained gongwen skill."""

from __future__ import annotations

import importlib.util
import platform
import shutil
import subprocess
import sys
from pathlib import Path


def status(label: str, ok: bool, detail: str, required: bool = True) -> bool:
    prefix = "[OK]" if ok else ("[FAIL]" if required else "[WARN]")
    print(f"{prefix} {label}: {detail}")
    return ok or not required


def command_version(command: str) -> str:
    path = shutil.which(command)
    if not path:
        return "not found"
    version_arg = "-v" if command in {"pdftotext", "pdftoppm"} else "--version"
    try:
        result = subprocess.run(
            [command, version_arg],
            check=False,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=5,
        )
        first = result.stdout.splitlines()[0] if result.stdout else path
        return f"{path} ({first})"
    except Exception:
        return path


def main() -> int:
    here = Path(__file__).resolve()
    skill_root = here.parents[1]
    results = [
        status("skill root", (skill_root / "SKILL.md").exists(), str(skill_root)),
        status("Python >= 3.9", sys.version_info >= (3, 9), sys.version.split()[0] + f" {sys.executable}"),
        status("python-docx", importlib.util.find_spec("docx") is not None, "installed"),
        status("PyYAML", importlib.util.find_spec("yaml") is not None, "installed"),
        status("pdftotext", bool(shutil.which("pdftotext")), command_version("pdftotext"), required=False),
        status("pdftoppm", bool(shutil.which("pdftoppm")), command_version("pdftoppm"), required=False),
        status("tesseract", bool(shutil.which("tesseract")), command_version("tesseract"), required=False),
        status("Paddle OCR local dir", Path("/Users/drevan01/Desktop/OCR").exists(), "/Users/drevan01/Desktop/OCR", required=False),
        status("template route gate", (skill_root / "scripts" / "gongwen-agent" / "check_template_route.py").exists(), "scripts/gongwen-agent/check_template_route.py"),
        status("fixed-format generator", (skill_root / "scripts" / "gongwen-agent" / "make_hecha_docx.py").exists(), "scripts/gongwen-agent/make_hecha_docx.py"),
        status("embedded templates", (skill_root / "assets" / "templates").exists(), "assets/templates"),
    ]
    print(f"[INFO] OS: {platform.platform()}")
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
