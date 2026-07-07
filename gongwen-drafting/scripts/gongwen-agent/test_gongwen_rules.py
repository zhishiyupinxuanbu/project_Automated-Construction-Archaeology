import unittest
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[2]


class GongwenRuleTests(unittest.TestCase):
    def read_reference(self, name: str) -> str:
        return (SKILL_ROOT / "references" / name).read_text(encoding="utf-8")

    def test_bureau_to_bureau_documents_must_route_as函(self):
        routing = self.read_reference("routing.md")
        drafting = self.read_reference("drafting-rules.md")
        combined = f"{routing}\n{drafting}"
        self.assertIn("收发单位均为“局”", combined)
        self.assertIn("只能写“函”", combined)
        self.assertIn("不可出现下级对上级的“请示”类口径", combined)


if __name__ == "__main__":
    unittest.main()
