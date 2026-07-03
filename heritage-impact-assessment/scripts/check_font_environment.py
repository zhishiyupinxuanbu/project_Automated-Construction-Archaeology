#!/usr/bin/env python3
"""Check whether the target Word fonts are likely available."""

from __future__ import annotations

import argparse
import platform
import subprocess
from pathlib import Path


REQUIRED_FONTS = {
    "宋体": ["宋体", "SimSun", "Songti", "STSong"],
    "Times New Roman": ["Times New Roman", "TimesNewRoman"],
}


FONT_DIRS = [
    Path.home() / "Library/Fonts",
    Path("/Library/Fonts"),
    Path("/System/Library/Fonts"),
    Path("/System/Library/AssetsV2"),
    Path("C:/Windows/Fonts"),
    Path("/usr/share/fonts"),
    Path("/usr/local/share/fonts"),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="检查文物影响评估 Word 成稿所需字体环境")
    parser.add_argument("--严格", dest="strict", action="store_true", help="缺少任一字体时返回非零退出码")
    return parser.parse_args()


def collect_font_text() -> str:
    chunks: list[str] = []

    for command in (["fc-list"], ["system_profiler", "SPFontsDataType"]):
        try:
            result = subprocess.run(command, capture_output=True, text=True, timeout=15, check=False)
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
        chunks.append(result.stdout)
        chunks.append(result.stderr)

    for font_dir in FONT_DIRS:
        if not font_dir.exists():
            continue
        try:
            for path in font_dir.rglob("*"):
                if path.suffix.lower() in {".ttf", ".ttc", ".otf", ".dfont"}:
                    chunks.append(path.name)
                    chunks.append(str(path))
        except (OSError, PermissionError):
            continue

    return "\n".join(chunks).casefold()


def main() -> None:
    args = parse_args()
    font_text = collect_font_text()
    missing: list[str] = []

    print("文物影响评估 Word 字体环境检查")
    print(f"系统：{platform.platform()}")
    print("")

    for display_name, aliases in REQUIRED_FONTS.items():
        found = any(alias.casefold() in font_text for alias in aliases)
        status = "已找到" if found else "未找到"
        print(f"- {display_name}: {status}")
        if not found:
            missing.append(display_name)

    print("")
    if missing:
        print("处理建议：")
        print("- 不要在 skill 包中直接打包或分发宋体、Times New Roman 等系统或商业字体文件。")
        print("- 请在目标机器安装系统字体、Office 字体或单位确认的正版字体包。")
        print("- 正式提交前，在 Word 中复核中文宋体、西文和数字 Times New Roman 是否生效。")
    else:
        print("检查结果：所需字体环境初步可用。正式提交前仍应在 Word 中复核版式。")

    if missing and args.strict:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
