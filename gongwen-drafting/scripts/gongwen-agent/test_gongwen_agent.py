#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tempfile
import unittest
from pathlib import Path
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
import gongwen_agent as ga


class GongwenAgentTests(unittest.TestCase):
    def test_infer_region(self):
        self.assertEqual(ga.infer_region("伊旗某项目"), "伊金霍洛旗")
        self.assertEqual(ga.infer_region("准格尔旗某项目"), "准格尔旗")

    def test_guess_project_name(self):
        title = "伊金霍洛旗文物局关于纳林希里煤矿及选煤项目用地范围有关文物事宜的请示"
        self.assertEqual(ga.guess_project_name(title), "纳林希里煤矿及选煤项目")

    def test_fact_checks_preserve_user_facts(self):
        project = ga.ProjectInput(
            business_type="文物保护许可申请",
            issuing_org="测试公司",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试公司",
            location="测试地点",
            land_area="1.23公顷",
        )
        body = ga.build_body(project, "文物保护许可申请", ["附件一"])
        checks = ga.make_checks(project, "\n".join([ga.make_title(project, project.business_type), *body]))
        self.assertTrue(all(item["status"] == "pass" for item in checks))

    def test_docx_generation(self):
        if ga.Document is None:
            self.skipTest("python-docx not installed")
        project = ga.ProjectInput(
            business_type="文物保护许可申请",
            issuing_org="测试公司",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试公司",
            location="测试地点",
            scale="测试建设内容",
            land_area="1.23公顷",
            approval_basis="测试核准文件",
        )
        body = ga.build_body(project, project.business_type, ["附件一"])
        checks = ga.make_checks(project, "\n".join(body))
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "draft.docx"
            ga.create_docx(project, project.business_type, body, checks, output)
            self.assertTrue(output.exists())
            self.assertGreater(output.stat().st_size, 0)
            with ZipFile(output) as docx:
                xml_text = docx.read("word/document.xml").decode("utf-8")
            self.assertNotIn("事实校验", xml_text)
            self.assertNotIn("参考来源", xml_text)
            self.assertNotIn("通过：", xml_text)

    def test_missing_scale_is_omitted_without_placeholder(self):
        project = ga.ProjectInput(
            business_type="勘探验收请示",
            issuing_org="测试单位",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试单位",
            location="测试地点",
            land_area="100平方米",
        )
        body = ga.build_body(project, project.business_type, [])
        draft = "\n".join(body)

        self.assertNotIn("待补充", draft)
        self.assertNotIn("建设内容为", draft)

    def test_survey_report_uses_report_not_request_language(self):
        project = ga.ProjectInput(
            business_type="文物调查报告",
            issuing_org="测试文物考古研究院",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试建设单位",
            location="测试地点",
            scale="测试建设内容",
            land_area="1.23公顷",
            approval_basis="《测试文物局关于测试项目联合开展考古调查的请示》（测试〔2026〕1号）",
            special_notes="项目用地范围内未发现登记在册不可移动文物及其他文物遗迹",
        )
        body = ga.build_body(project, project.business_type, ga.default_attachment_items(project, project.business_type))
        draft = "\n".join([ga.make_title(project, project.business_type), *body])

        self.assertIn("关于测试项目文物调查的报告", draft)
        self.assertIn("专此", body)
        self.assertIn("基本建设项目文物调查表", draft)
        self.assertNotIn("妥否，请批示", draft)
        self.assertNotIn("特此请示", draft)
        self.assertNotIn("现申请", draft)

    def test_docx_title_single_paragraph_no_manual_break(self):
        if ga.Document is None:
            self.skipTest("python-docx not installed")
        project = ga.ProjectInput(
            business_type="文物保护许可申请",
            issuing_org="测试公司",
            recipient_org="测试文物局",
            project_name="很长很长很长很长很长很长很长很长很长很长很长的测试项目",
            construction_unit="测试公司",
            location="测试地点",
            scale="测试建设内容",
            land_area="1.23公顷",
        )
        body = ga.build_body(project, project.business_type, ["附件一"])
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "draft.docx"
            ga.create_docx(project, project.business_type, body, [], output)
            with ZipFile(output) as docx:
                xml = docx.read("word/document.xml")
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(xml)
        first_para = root.find(".//w:p", ns)
        text = "".join(node.text or "" for node in first_para.findall(".//w:t", ns))
        self.assertEqual(text, ga.make_title(project, project.business_type))
        self.assertIsNone(first_para.find(".//w:br", ns))

    def test_docx_signature_centers_on_date_with_tab_stop(self):
        if ga.Document is None:
            self.skipTest("python-docx not installed")
        project = ga.ProjectInput(
            business_type="文物保护许可申请",
            issuing_org="测试公司",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试公司",
            location="测试地点",
            scale="测试建设内容",
            land_area="1.23公顷",
        )
        body = ga.build_body(project, project.business_type, ["附件一"])
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "draft.docx"
            ga.create_docx(project, project.business_type, body, [], output)
            with ZipFile(output) as docx:
                xml = docx.read("word/document.xml")
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(xml)
        signature_para = None
        for para in root.findall(".//w:p", ns):
            text = "".join(node.text or "" for node in para.findall(".//w:t", ns))
            if text == "测试公司":
                signature_para = para
                break
        self.assertIsNotNone(signature_para)
        indent = signature_para.find("./w:pPr/w:ind", ns)
        self.assertTrue(indent is None or indent.attrib.get(f"{{{ns['w']}}}left") in {None, "0"})
        tab = signature_para.find("./w:pPr/w:tabs/w:tab", ns)
        self.assertIsNotNone(tab)
        self.assertEqual(tab.attrib.get(f"{{{ns['w']}}}val"), "center")
        self.assertIsNotNone(signature_para.find(".//w:tab", ns))

    def test_docx_attachment_items_use_hanging_indent(self):
        if ga.Document is None:
            self.skipTest("python-docx not installed")
        project = ga.ProjectInput(
            business_type="文物保护许可申请",
            issuing_org="测试公司",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试公司",
            location="测试地点",
            scale="测试建设内容",
            land_area="1.23公顷",
            approval_basis="测试核准文件",
        )
        attachment = "测试项目用地经纬度坐标 Excel 表"
        body = ga.build_body(project, project.business_type, [attachment])
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "draft.docx"
            ga.create_docx(project, project.business_type, body, [], output)
            with ZipFile(output) as docx:
                xml = docx.read("word/document.xml")
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(xml)
        attachment_para = None
        for para in root.findall(".//w:p", ns):
            text = "".join(node.text or "" for node in para.findall(".//w:t", ns))
            if text == f"附件：1.{attachment}":
                attachment_para = para
                break
        self.assertIsNotNone(attachment_para)
        indent = attachment_para.find("./w:pPr/w:ind", ns)
        self.assertIsNotNone(indent)
        self.assertEqual(indent.attrib.get(f"{{{ns['w']}}}left"), "2240")
        self.assertEqual(indent.attrib.get(f"{{{ns['w']}}}hanging"), "1600")

    def test_default_attachment_items_replace_project_and_basis(self):
        project = ga.ProjectInput(
            business_type="文物保护许可申请",
            issuing_org="测试公司",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试公司",
            approval_basis="《关于核实项目选址是否涉及文物遗址的复函》（右中文物字〔2024〕83号）",
        )
        attachments = ga.default_attachment_items(project, project.business_type)
        self.assertEqual(attachments[0], "关于核实项目选址是否涉及文物遗址的复函")
        self.assertIn("测试项目用地经纬度坐标 Excel 表", attachments)
        self.assertNotIn("XX项目用地经纬度坐标 Excel 表", attachments)

    def test_yanshou_uses_new_template_and_generic_third_party(self):
        project = ga.ProjectInput(
            business_type="勘探验收请示",
            issuing_org="测试公司",
            recipient_org="测试文物局",
            project_name="测试项目",
            construction_unit="测试公司",
            location="测试地点",
            scale="测试建设内容",
            land_area="1.23公顷",
            special_notes="三门峡市文物考古勘探有限公司，发现灰坑遗迹3处",
        )
        title = ga.make_title(project, project.business_type)
        body = ga.build_body(project, project.business_type, [])
        draft = "\n".join([title, *body])
        self.assertEqual(title, "关于申请测试项目考古勘探验收的函")
        self.assertIn("第三方公司", draft)
        self.assertNotIn("三门峡市文物考古勘探有限公司", draft)
        self.assertIn("此函。", body)
        self.assertEqual(body[-1], "测试项目考古调查、勘探工作报告")
        self.assertNotIn("测试项目考古调查、勘探工作计划", body)


if __name__ == "__main__":
    unittest.main()
