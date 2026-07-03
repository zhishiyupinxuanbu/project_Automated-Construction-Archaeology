#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
import check_template_route as route


class TemplateRouteTests(unittest.TestCase):
    def test_xilinhot_hecha_uses_local_template(self):
        root = Path(__file__).resolve().parents[2]
        rows = route.parse_template_index(root / "references/knowledge-index/02-模板索引.md")

        result = route.resolve_route(rows, "锡市", "文物核查请示")

        self.assertEqual(result["route"], "local")
        self.assertEqual(result["region"], "锡林浩特市")
        self.assertEqual(result["doc_type"], "文物保护许可申请")
        self.assertEqual(result["template"]["name"], "锡林浩特市临时用地是否涉及文物请示")

    def test_shengli_mining_area_routes_to_xilinhot_template(self):
        root = Path(__file__).resolve().parents[2]
        rows = route.parse_template_index(root / "references/knowledge-index/02-模板索引.md")

        result = route.resolve_route(rows, "胜利矿区", "文物核查请示")

        self.assertEqual(result["route"], "local")
        self.assertEqual(result["region"], "锡林浩特市")
        self.assertEqual(result["template"]["name"], "锡林浩特市临时用地是否涉及文物请示")


if __name__ == "__main__":
    unittest.main()
