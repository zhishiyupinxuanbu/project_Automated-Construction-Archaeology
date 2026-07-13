from __future__ import annotations

import unittest
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
REQUIRED = (
    "SKILL.md",
    "agents/openai.yaml",
    "assets",
    "references",
    "scripts",
    "tests",
    "CHANGELOG.md",
    ".gitignore",
)
FORBIDDEN_NAMES = {".DS_Store", ".pytest_cache", "source-originals"}
OLD_LIBRARY_ROOT = "/Users/drevan01/Desktop/影响评估与保护方案skill"


class SkillPackageLayoutTest(unittest.TestCase):
    def test_required_package_entries_exist(self) -> None:
        missing = [name for name in REQUIRED if not (SKILL_DIR / name).exists()]
        self.assertEqual(missing, [])

    def test_package_excludes_runtime_and_raw_reference_materials(self) -> None:
        forbidden = sorted(
            str(path.relative_to(SKILL_DIR))
            for path in SKILL_DIR.rglob("*")
            if path.name in FORBIDDEN_NAMES
        )
        self.assertEqual(forbidden, [])

    def test_text_files_do_not_reference_old_library_root(self) -> None:
        offenders = []
        for path in SKILL_DIR.rglob("*"):
            if path == Path(__file__):
                continue
            if path.is_file() and path.suffix.lower() in {".md", ".py", ".yaml", ".yml"}:
                if OLD_LIBRARY_ROOT in path.read_text(encoding="utf-8", errors="ignore"):
                    offenders.append(str(path.relative_to(SKILL_DIR)))
        self.assertEqual(sorted(offenders), [])

    def test_runtime_caches_are_ignored(self) -> None:
        ignore_text = (SKILL_DIR / ".gitignore").read_text(encoding="utf-8")
        self.assertIn("__pycache__/", ignore_text)
        self.assertIn(".pytest_cache/", ignore_text)


if __name__ == "__main__":
    unittest.main()
