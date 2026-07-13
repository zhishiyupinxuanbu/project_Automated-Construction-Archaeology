#!/usr/bin/env python3
"""Start reference report OCR as a detached daemon."""

from __future__ import annotations

import os
import sys
from pathlib import Path


REFERENCE_LIBRARY_ROOT = Path(
    os.environ.get("HERITAGE_REFERENCE_LIBRARY", "/Users/drevan01/Desktop/文物影响评估与保护方案资料库")
)
ROOT = REFERENCE_LIBRARY_ROOT / "01_法规政策与标准/环境与施工标准资料库"
SCRIPT = ROOT / "scripts" / "ocr_reference_report_standards.py"
LOG = ROOT / "logs" / "reference_report_ocr.log"
PID = ROOT / "logs" / "reference_report_ocr.pid"


def main() -> None:
    LOG.parent.mkdir(parents=True, exist_ok=True)

    first_pid = os.fork()
    if first_pid > 0:
        print(first_pid)
        return

    os.setsid()

    second_pid = os.fork()
    if second_pid > 0:
        os._exit(0)

    os.chdir(str(ROOT))
    os.umask(0o022)

    with open(os.devnull, "rb", buffering=0) as stdin, open(LOG, "ab", buffering=0) as log:
        os.dup2(stdin.fileno(), 0)
        os.dup2(log.fileno(), 1)
        os.dup2(log.fileno(), 2)

    PID.write_text(str(os.getpid()), encoding="utf-8")
    os.execv(sys.executable, [sys.executable, str(SCRIPT)])


if __name__ == "__main__":
    main()
