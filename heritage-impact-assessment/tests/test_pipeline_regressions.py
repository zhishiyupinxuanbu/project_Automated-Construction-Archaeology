from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
JUEHAI_PROJECT_DIR = Path("/Users/drevan01/Desktop/影响评估与保护方案skill/觉海寺项目/1.项目资料")


def read_jsonl(path: Path) -> list[dict]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def run_script(script: str, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SKILL_DIR / "scripts" / script), *args],
        text=True,
        capture_output=True,
    )


class PipelineRegressionTest(unittest.TestCase):
    def test_material_processing_keeps_late_project_text_without_default_truncation(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            late_marker = "选址方案比选分析：原址改扩建方案为推荐方案。"
            (project_dir / "项目可研.txt").write_text(
                "前置文字" * 35000 + late_marker,
                encoding="utf-8",
            )

            result = run_script(
                "run_material_processing.py",
                "--项目资料目录",
                str(project_dir),
                "--输出目录",
                str(output_dir),
                "--项目名称",
                "测试项目",
                "--覆盖",
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            text_index = read_jsonl(output_dir / "processing_output" / "text_index.jsonl")
            extracted_text = "\n".join(
                (output_dir / row["text_path"]).read_text(encoding="utf-8", errors="ignore")
                for row in text_index
            )
            self.assertIn(late_marker, extracted_text)
            chunk_index = read_jsonl(output_dir / "processing_output" / "text_chunk_index.jsonl")
            self.assertGreater(len(chunk_index), 1)
            chunk_text = "\n".join(
                (output_dir / row["chunk_path"]).read_text(encoding="utf-8", errors="ignore")
                for row in chunk_index
            )
            self.assertIn(late_marker, chunk_text)

    def test_fact_extraction_uses_all_project_text_sources(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目可研.txt").write_text(
                "海南区觉海寺改扩建配套设施改造项目以“保障安全、传承文化、提升功能、服务发展”为核心建设目标。"
                "项目计划总投资361.90万元，企业自筹资金。建设工期2025年12月—2027年9月。"
                "项目建设必要性包括安全底线、寺院发展战略和长期稳定运营。"
                "选址方案比选分析设置原址改扩建方案、原址扩建+邻近地块补充方案、异地新建方案三个方案。",
                encoding="utf-8",
            )
            (project_dir / "开工请示.txt").write_text(
                "建设单位为海南区觉海寺，项目面积7386平方米，位于乌海市海南区拉僧庙镇赛汗乌素村。",
                encoding="utf-8",
            )
            (project_dir / "规划平面图说明.txt").write_text(
                "建设内容包括三圣殿、山门、钟楼、鼓楼、展览室、传统文化宣传室、流通处、客堂、小寮、护法殿、禅堂、库房、念经堂、灯房。",
                encoding="utf-8",
            )

            result = run_script(
                "run_material_processing.py",
                "--项目资料目录",
                str(project_dir),
                "--输出目录",
                str(output_dir),
                "--项目名称",
                "海南区觉海寺改扩建配套设施改造项目",
                "--覆盖",
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            result = run_script("run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖")
            self.assertEqual(result.returncode, 0, result.stderr)

            project_facts = read_jsonl(output_dir / "facts" / "project_facts.jsonl")
            by_field = {row["field_name"]: row["value"] for row in project_facts}
            for field in [
                "建设单位",
                "项目面积",
                "建设目标",
                "总投资",
                "建设工期",
                "建设必要性",
                "选址方案比选",
                "建设内容",
            ]:
                self.assertIn(field, by_field)
                self.assertNotEqual(by_field[field], "待确认")

            coverage = read_jsonl(output_dir / "facts" / "source_coverage.jsonl")
            covered_sources = {row["source_file"] for row in coverage if row["fact_count"] > 0}
            self.assertIn("项目可研.txt", covered_sources)
            self.assertIn("开工请示.txt", covered_sources)
            self.assertIn("规划平面图说明.txt", covered_sources)

    def test_material_processing_extracts_text_from_juehai_pdf(self) -> None:
        if not JUEHAI_PROJECT_DIR.exists():
            self.skipTest("觉海寺项目资料目录不存在")
        with tempfile.TemporaryDirectory() as tmp:
            output_dir = Path(tmp) / "work"
            result = run_script(
                "run_material_processing.py",
                "--项目资料目录",
                str(JUEHAI_PROJECT_DIR),
                "--输出目录",
                str(output_dir),
                "--项目名称",
                "海南区觉海寺改扩建配套设施改造项目",
                "--覆盖",
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            text_index = read_jsonl(output_dir / "processing_output" / "text_index.jsonl")
            pdf_rows = [row for row in text_index if row["source_file"].endswith(".pdf")]
            self.assertTrue(pdf_rows, "PDF 文件必须进入 text_index，不能只登记为图件或文件级证据")

            extracted_text = "\n".join(
                (output_dir / row["text_path"]).read_text(encoding="utf-8", errors="ignore")
                for row in pdf_rows
            )
            self.assertIn("东风农场七队长城", extracted_text)
            self.assertIn("建设控制地带", extracted_text)

    def test_fact_extraction_does_not_promote_kml_xml_to_quote_candidates(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目文物调查范围.kml").write_text(
                """<?xml version="1.0" encoding="UTF-8"?>
<kml><Document><name>测试项目用地范围</name><Placemark><name>J1</name>
<Point><coordinates>106.1,39.1,0</coordinates></Point></Placemark></Document></kml>""",
                encoding="utf-8",
            )

            result = run_script(
                "run_material_processing.py",
                "--项目资料目录",
                str(project_dir),
                "--输出目录",
                str(output_dir),
                "--项目名称",
                "测试项目",
                "--覆盖",
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            result = run_script("run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖")
            self.assertEqual(result.returncode, 0, result.stderr)

            quote_candidates = read_jsonl(output_dir / "facts" / "quote_candidates.jsonl")
            combined_quotes = "\n".join(row.get("text", "") for row in quote_candidates)
            self.assertNotIn("<?xml", combined_quotes)
            self.assertNotIn("<kml", combined_quotes)

            fact_issues = read_jsonl(output_dir / "facts" / "fact_issues.jsonl")
            issue_fields = {row["field_name"] for row in fact_issues}
            self.assertIn("可引用叙述材料", issue_fields)

    def test_survey_report_project_basics_feed_first_chapter_without_process_narration(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "文物调查报告.txt").write_text(
                "海南区觉海寺改扩建配套设施改造项目位于内蒙古自治区乌海市海南区拉僧庙镇赛汗乌素村黄河后组89号，"
                "建设改扩建配套设施。项目扩建总建筑面积7386平方米，扩建后寺院总占地面积16919.64平方米。"
                "建设内容包括三圣殿、山门、钟楼、鼓楼、展览室、传统文化宣传室、流通处、客堂、小寮、护法殿、禅堂、库房、念经堂、灯房等设施。"
                "经查阅内蒙古自治区现有不可移动文物名录、内蒙古自治区长城资源调查数据库及现场调查，"
                "发现海南区觉海寺改扩建配套设施改造项目拟用地范围涉及乌海市海南区第五批自治区级文物保护单位桌子山秦长城--东风农场七队长城3段建设控制地带。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "海南区觉海寺改扩建配套设施改造项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            project_facts = read_jsonl(output_dir / "facts" / "project_facts.jsonl")
            by_field = {row["field_name"]: row["value"] for row in project_facts}
            self.assertIn("扩建后寺院总占地面积16919.64平方米", by_field.get("总占地面积", ""))

            quote_candidates = read_jsonl(output_dir / "facts" / "quote_candidates.jsonl")
            survey_quotes = [row for row in quote_candidates if "文物调查报告" in row.get("source_file", "")]
            self.assertTrue(survey_quotes)
            self.assertIn("一、总则", survey_quotes[0]["target_section"])
            self.assertIn("评估项目基础信息", survey_quotes[0]["quote_type"])

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            self.assertIn("1.评估项目基础信息", report)
            self.assertIn("2.文物影响评估必要性", report)
            self.assertIn("扩建后寺院总占地面积16919.64平方米", report)
            self.assertIn(
                "现结合项目当前地理位置、施工方案，以及相关法律法规和行业标准，"
                "对桌子山秦长城东风农场七队长城3段可能受到的影响进行评估。"
                "具体来看，海南区觉海寺改扩建配套设施改造项目在设计、建设、运营过程中，"
                "施工方式、车辆、人员等因素可能对不可移动文物本体及周边环境造成影响，"
                "为统筹推进项目实施与文物保护，实现项目建设与文化遗产保护协同发展，"
                "需系统性评估海南区觉海寺改扩建配套设施改造项目在建设及运营期间对所涉及文物的影响。",
                report,
            )
            self.assertIn(
                "建设项目文物影响评估是对建设项目当前计划方案的实施对文物本体及其周边环境造成的影响进行分析、预测和评估，"
                "提出预防或减轻不良影响的对策和措施，并在规划和建设事中、事后进行跟踪监测。",
                report,
            )
            self.assertIn(
                "《中华人民共和国文物保护法》2024年11月8日修订版中第二十八条规定，"
                "在文物保护单位的保护范围内不得进行文物保护工程以外的其他建设工程或者爆破、钻探、挖掘等作业；",
                report,
            )
            self.assertIn(
                "第二十九条指出，在文物保护单位的建设控制地带内进行建设工程，不得破坏文物保护单位的历史风貌，"
                "工程设计方案应当根据文物保护单位的级别和建设工程对文物保护单位历史风貌的影响程度，经国家规定的文物行政部门同意后，依法取得建设工程规划许可。",
                report,
            )
            self.assertIn(
                "根据《内蒙古自治区人民政府办公厅关于加强工程建设文物保护前置审查工作的通知》（内政办发〔2024〕40号）、"
                "《内蒙古自治区文物局关于做好基本建设用地考古工作的通知》（内文物发〔2025〕6号）文件内容，"
                "对考古调查和勘探工作中发现文物遗存需原址保护的，或涉及各级文物保护单位保护范围和建设控制地带的建设用地，"
                "用地单位应提供《文物保护方案》和《文物影响评估报告》，由接收申请的文物行政部门组织包括自治区文物考古研究院专家在内的专家进行评审，并由参与评审的单位共同出具意见。",
                report,
            )
            self.assertIn("（三）评估内容", report)
            self.assertIn(
                "根据文物影响的程度，明确海南区觉海寺改扩建配套设施改造项目是否可行，提出调整建议和针对性的减缓措施。"
                "并根据海南区觉海寺改扩建配套设施改造项目的文物影响评估结论，提出其他相关建议。",
                report,
            )
            self.assertNotIn("后续方案需重点复核的影响问题", report)
            self.assertNotIn("最终影响等级、建设可接受性和审批结论", report)
            self.assertIn("（五）评估依据", report)
            self.assertIn("1.宪章公约", report)
            fixed_charter_lines = [
                "(1) 《国际古迹保护与修复宪章》（1964）（International Charter for the Conservation and Restoration of Monuments and Sites）；",
                "(2) 《考古遗产保护与管理宪章》（1990）（Charter for the Protection and Management of the Archaeological Heritage）；",
                "(3) 《中国文物古迹保护准则》（2000）（Principles for the Conservation of Heritage Sites in China）；",
                "(4) 《会安草案——亚洲最佳保护范例》（2005）（Hoi An Protocols for Best Conservation of Historic Towns and Urban Areas）；",
                "(5) 《西安宣言》（2005）（Xi'an Declaration on the Conservation of the Setting of Heritage Structures, Sites and Areas）；",
                "(6) 《文化遗产阐释与展示宪章》（2008）（Charter on the Interpretation and Presentation of Cultural Heritage Sites）。",
            ]
            for fixed_line in fixed_charter_lines:
                self.assertIn(fixed_line, report)
            self.assertNotIn("1.宪章公约：", report)
            self.assertNotIn("《会安草案一一亚洲最佳保护范例》", report)
            self.assertIn("2.法律法规", report)
            fixed_law_lines = [
                "(1) 《中华人民共和国文物保护法》2024年11月8日修订，2025年3月1日起施行；",
                "(2) 《中华人民共和国文物保护法实施条例》2017年10月7日修订；",
                "(3) 《内蒙古自治区文物保护条例》2005年12月1日修订；",
            ]
            for fixed_line in fixed_law_lines:
                self.assertIn(fixed_line, report)
            self.assertIn("(4) 《长城保护条例》2006年12月1日起施行。", report)
            self.assertNotIn("2.法律法规：", report)
            self.assertNotIn("《长城保护条例》。", report)
            self.assertIn("3.文件规定", report)
            fixed_policy_lines = [
                "(1)《国务院关于加强文化遗产保护的通知》（国发〔2005〕42号）；",
                "(2)《国务院关于进一步加强文物工作的指导意见》（国发〔2016〕17号）；",
                "(3)《关于加强基本建设工程中考古工作的指导意见》；",
                "(4)《中共中央办公厅、国务院办公厅关于加强文物保护利用改革的若干意见》（2018年10月8日）；",
                "(5)《内蒙古自治区人民政府关于进一步加强文物保护与利用工作的意见》（内政字〔2004〕260号）；",
                "(6)《内蒙古自治区人民政府办公厅关于加强工程建设文物保护前置审查工作的通知》（内政办发〔2024〕40号）；",
                "(7)《内蒙古自治区文物局关于做好基本建设用地考古工作的通知》（内文物发〔2025〕6号）。",
            ]
            for fixed_line in fixed_policy_lines:
                self.assertIn(fixed_line, report)
            self.assertNotIn("3.文件规定：", report)
            self.assertNotIn("(1) 《国务院关于加强文化遗产保护的通知》", report)
            self.assertNotIn("本报告围绕海南区觉海寺改扩建配套设施改造项目开展文物影响评估", report)
            self.assertNotIn("本报告为工作草稿", report)
            self.assertNotIn("重新开展资料整理、OCR、事实抽取", report)

    def test_technical_specs_are_quoted_as_complete_project_source_section(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            technical_section = (
                "4.技术规范\n"
                "(1)《宗教活动场所建设标准》；\n"
                "(2)《古建筑修缮工程施工规范》（JGJ 159-2008）；\n"
                "(3)《文物保护工程施工质量验收规范》（WW/T 0041-2014）；\n"
                "(4)《工业企业总平面图设计规范》GB50187—2012；\n"
                "(5)《建筑设计防火规范》（GB50016-2014）（2018版）；\n"
                "(6)《民用建筑设计统一标准》（GB50352-2019）；\n"
                "(7)《建筑内部装修设计防火规范》（GB50222-2017）；\n"
                "(8)《屋面工程技术规范》（GB50345-2012）；\n"
                "(9)《外墙外保温工程技术标准》（JGJ 144-2019）；\n"
                "(10)其他国家颁布的有关工程建设的政策及法规，现行设计规范与标准，以及地方性有关设计要求及规定。"
            )
            (project_dir / "项目可研.txt").write_text(
                "项目名称：测试技术规范项目。\n"
                "建设地点：测试地点。\n"
                f"{technical_section}\n"
                "5.其他资料\n"
                "本节为其他资料。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试技术规范项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            quote_candidates = read_jsonl(output_dir / "facts" / "quote_candidates.jsonl")
            tech_quotes = [row for row in quote_candidates if row.get("quote_type") == "技术规范原文摘录"]
            self.assertTrue(tech_quotes)
            self.assertIn("(10)其他国家颁布的有关工程建设的政策及法规", tech_quotes[0]["text"])

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            self.assertIn("4.技术规范", report)
            for line in technical_section.splitlines()[1:]:
                self.assertIn(line, report)
            self.assertNotIn("本项目资料明确引用《考古勘探工作规程（试行）》", report)
            self.assertNotIn("具体清单需由设计文件或施工方案进一步核验", report)

    def test_other_materials_keep_fixed_first_two_and_exclude_drawings_images_acceptance(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            material_names = [
                "内蒙古自治区长城保护规划（2020—2035年）.txt",
                "测试项目可行性研究报告.txt",
                "测试项目考古调查、勘探工作报告.txt",
                "乌海市海南区城乡规划（2021-2035）.txt",
                "测试项目考古勘探项目验收意见书.txt",
                "规划平面图.txt",
                "项目红线.kml",
                "正射影像说明.txt",
                "坐标表.txt",
                "现场照片.txt",
            ]
            for name in material_names:
                (project_dir / name).write_text(
                    "测试项目位于测试地点，建设内容包括配套设施。",
                    encoding="utf-8",
                )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试其他资料项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            self.assertIn("5.其他资料", report)
            other_materials_section = report.split("#### 5.其他资料", 1)[1].split("## 二、建设项目涉及文物概况", 1)[0]
            expected_lines = [
                "(1) 《世界文化遗产影响评估指南》（2011年）；",
                "(2) 《中国文物古迹保护准则》（2015年）；",
                "(3) 《内蒙古自治区长城保护规划（2020—2035年）》；",
                "(4) 《测试项目可行性研究报告》；",
                "(5) 《测试项目考古调查、勘探工作报告》；",
                "(6) 《乌海市海南区城乡规划（2021-2035）》。",
            ]
            for line in expected_lines:
                self.assertIn(line, other_materials_section)
            for excluded in ["验收意见书", "规划平面图", "项目红线", "正射影像", "坐标表", "现场照片", "KML/OVKML"]:
                self.assertNotIn(excluded, other_materials_section)

    def test_evaluation_method_is_fixed_with_project_and_heritage_replacements(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "文物调查报告.txt").write_text(
                "测试评估方法项目位于测试地点，拟用地范围涉及桌子山秦长城东风农场七队长城3段建设控制地带。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试评估方法项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            self.assertIn("（六）评估方法", report)
            method_section = report.split("### （六）评估方法", 1)[1].split("## 二、建设项目涉及文物概况", 1)[0]
            self.assertIn("#### 1.调查法", method_section)
            self.assertIn("#### 2.综合分析法", method_section)
            self.assertIn("#### 3.预测法", method_section)
            self.assertIn(
                "项目信息。收集测试评估方法项目的建设性质、具体内容、规模、范围、工程参数、相关审批文件、批复意见以及所在区域的城市或控制性规划。",
                method_section,
            )
            self.assertIn(
                "实地勘察。对测试评估方法项目区域及桌子山秦长城东风农场七队长城3段建设控制地带进行详细的实地踏勘。",
                method_section,
            )
            self.assertIn(
                "根据测试评估方法项目的特点确立文物影响评价因子，识别出关键的影响路径和最敏感的文物要素",
                method_section,
            )
            self.assertIn(
                "依据项目建设工程的《测试评估方法项目可行性研究报告》中的建设内容和施工方法",
                method_section,
            )
            self.assertNotIn("海南区觉海寺改扩建配套设施改造项目", method_section)

    def test_chapter_two_keeps_opening_lead_from_investigation_reply(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            opening_lead = (
                "经查阅内蒙古自治区现有不可移动文物名录、内蒙古自治区长城资源调查数据库及现场调查，"
                "发现测试调查回函项目拟用地范围涉及测试文物点建设控制地带。因此，本报告主要针对测试文物点进行文物影响评估。"
            )
            (project_dir / "调查回函.txt").write_text(opening_lead, encoding="utf-8")

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试调查回函项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            quotes = read_jsonl(output_dir / "facts" / "quote_candidates.jsonl")
            chapter_two_quotes = [row for row in quotes if row.get("quote_type") == "第二章开头总起段"]
            self.assertTrue(chapter_two_quotes)
            self.assertEqual(chapter_two_quotes[0]["text"], opening_lead)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_two = report.split("## 二、建设项目涉及文物概况", 1)[1].split("## 三、建设项目规划概况", 1)[0]
            self.assertIn(opening_lead, chapter_two)
            self.assertLess(chapter_two.index(opening_lead), chapter_two.index("项目涉及文物对象"))

    def test_chapter_two_region_overview_has_required_web_sourced_subsections(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "调查回函.txt").write_text(
                "经查阅现有不可移动文物名录及现场调查，发现测试区域项目拟用地范围涉及测试文物点建设控制地带。因此，本报告主要针对测试文物点进行文物影响评估。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试区域项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_two = report.split("## 二、建设项目涉及文物概况", 1)[1].split("## 三、建设项目规划概况", 1)[0]
            required_headings = [
                "### （一）文物所在区域概况",
                "#### 1.地理位置",
                "#### 2.自然概况",
                "#### 3.社会经济情况",
                "#### 4.历史沿革",
            ]
            positions = [chapter_two.index(heading) for heading in required_headings]
            self.assertEqual(positions, sorted(positions))
            self.assertIn("【待联网检索：文物所在区域地理位置资料", chapter_two)

    def test_chapter_two_heritage_overview_has_fixed_subsections_from_investigation_reply(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "文物调查回函.txt").write_text(
                "经查阅现有不可移动文物名录及现场调查，发现测试遗址项目拟用地范围涉及桌子山秦长城东风农场七队长城3段建设控制地带。"
                "因此，本报告主要针对桌子山秦长城东风农场七队长城3段进行文物影响评估。"
                "桌子山秦长城东风农场七队长城3段为长城遗址，属于自治区级文物保护单位，具备历史、军事、社会方面的价值，遗产价值分级为“中”。"
                "调查材料载明其保存现状、保护范围及建设控制地带。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试遗址项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_two = report.split("## 二、建设项目涉及文物概况", 1)[1].split("## 三、建设项目规划概况", 1)[0]
            required_headings = [
                "### （二）桌子山秦长城东风农场七队长城3段概述",
                "#### 1.遗址基本情况",
                "#### 2.遗址现状评估",
                "#### 3.调查、发掘、保护工程情况",
                "#### 4.保护范围及建设控制地带",
                "#### 5.价值陈述",
                "#### 6.价值评估",
            ]
            positions = [chapter_two.index(heading) for heading in required_headings]
            self.assertEqual(positions, sorted(positions))
            self.assertIn("【待依据文物调查回函或调查类材料补充：遗址基本情况", chapter_two)
            value_section = chapter_two.split("#### 6.价值评估", 1)[1]
            self.assertIn(
                "《世界文化遗产影响评估指南》中的价值评估准则指出，对文物开展文物影响评估时，"
                "应考量其具备的文物价值、当地价值或国家价值，以及国家研究规划所明确的优先顺序与建议。",
                value_section,
            )
            self.assertIn("表1 遗产价值分级量表", value_section)
            self.assertIn("| 分级 | 考古学 | 建成遗址或历史城市景观 | 历史景观 | 非物质文化遗产或相关内容 |", value_section)
            self.assertIn("具有某些隐含（如无法看到）或潜在历史重要性的建筑物", value_section)
            self.assertIn(
                "根据上述文物价值陈述，结合遗产价值分级量表，桌子山秦长城东风农场七队长城3段属于已公布为保护单位的建筑物"
                "（自治区级文物保护单位），具备历史、军事、社会方面的价值。因此评估桌子山秦长城东风农场七队长城3段的遗产价值分级为“中”。",
                value_section,
            )
            self.assertNotIn("根据现有资料，", value_section)
            self.assertNotIn("暂按", value_section)

    def test_chapter_three_project_overview_uses_feasibility_natural_paragraphs_not_table(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n"
                "海南区觉海寺改扩建配套设施改造项目以“保障安全、传承文化、提升功能、服务发展”为核心建设目标。"
                "通过对山门、念佛堂等核心配套设施的改扩建，提升宗教活动场所的安全保障水平。\n\n"
                "项目拟建于内蒙古自治区乌海市海南区拉僧庙镇赛汗乌素村黄河后组89号，建设改扩建配套设施项目。"
                "项目扩建总建筑面积7386平方米，建设内容包括三圣殿、山门、钟楼、鼓楼、展览室等设施。\n\n"
                "项目核心建设任务聚焦寺院主要基础设施的改扩建及配套设施完善，具体包括完善传统寺院礼仪和景观节点、"
                "实施宗教功能空间建设、建设配套设施并同步完善给排水、供电、消防等基础设施。\n"
                "二、项目建设必要性\n"
                "本节为必要性内容。",
                encoding="utf-8",
            )
            (project_dir / "文物调查回函.txt").write_text(
                "经查阅现有不可移动文物名录及现场调查，发现海南区觉海寺改扩建配套设施改造项目拟用地范围涉及桌子山秦长城东风农场七队长城3段建设控制地带。"
                "因此，本报告主要针对桌子山秦长城东风农场七队长城3段进行文物影响评估。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "海南区觉海寺改扩建配套设施改造项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_three = report.split("## 三、建设项目规划概况", 1)[1].split("## 四、项目用地范围与文物空间分布关系", 1)[0]
            self.assertIn("### （一）项目概况", chapter_three)
            self.assertIn("海南区觉海寺改扩建配套设施改造项目以“保障安全、传承文化、提升功能、服务发展”为核心建设目标。", chapter_three)
            self.assertIn("项目拟建于内蒙古自治区乌海市海南区拉僧庙镇赛汗乌素村黄河后组89号", chapter_three)
            self.assertIn("项目核心建设任务聚焦寺院主要基础设施的改扩建及配套设施完善", chapter_three)
            self.assertNotIn("表3 项目基本信息", chapter_three)
            self.assertNotIn("| 项目 | 内容 | 来源 |", chapter_three)
            self.assertNotIn("| --- | --- | --- |", chapter_three)

    def test_chapter_three_project_necessity_uses_complete_feasibility_section(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n"
                "测试项目以改善基础设施为核心建设目标。\n"
                "二、项目建设必要性\n"
                "1.符合安全底线的刚性要求\n"
                "觉海寺始建于清代乾隆年间，至今已有200余年历史，现存建筑普遍存在墙体开裂、屋顶渗漏等结构安全隐患，"
                "且消防通道狭窄、消防设施严重不足，直接违反《宗教活动场所消防安全管理规定》，已成为制约宗教活动正常开展的核心风险点。"
                "实施改扩建项目是彻底消除安全隐患、保障信众和游客人身安全、确保宗教活动合规有序开展的刚性要求。\n\n"
                "2.寺院发展战略的核心支撑\n"
                "觉海寺长远发展聚焦“宗教传承规范化、文化展示品牌化、服务能力优质化”三大目标，本次项目是战略落地的关键载体："
                "通过修缮历史建筑、完善安全设施筑牢宗教传承的合规根基；通过配套文化展示厅、历史陈列区将寺院从单一宗教场所升级为区域特色宗教文化地标；"
                "通过扩建接待设施解决当前服务能力不足、信众游客满意度偏低的问题，全面提升寺院运营品质。"
                "项目实施对觉海寺发展战略的实现具有重要价值，是寺院实现可持续发展的必须举措。\n"
                "三、支持性文件取得情况\n"
                "本节为支持性文件内容。",
                encoding="utf-8",
            )
            (project_dir / "文物调查回函.txt").write_text(
                "经查阅现有不可移动文物名录及现场调查，发现测试项目拟用地范围涉及桌子山秦长城东风农场七队长城3段建设控制地带。"
                "因此，本报告主要针对桌子山秦长城东风农场七队长城3段进行文物影响评估。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_three = report.split("## 三、建设项目规划概况", 1)[1].split("## 四、项目用地范围与文物空间分布关系", 1)[0]
            self.assertIn("### （二）项目建设必要性", chapter_three)
            self.assertIn("#### 1.符合安全底线的刚性要求", chapter_three)
            self.assertIn("觉海寺始建于清代乾隆年间，至今已有200余年历史", chapter_three)
            self.assertIn("#### 2.寺院发展战略的核心支撑", chapter_three)
            self.assertIn("觉海寺长远发展聚焦“宗教传承规范化、文化展示品牌化、服务能力优质化”三大目标", chapter_three)
            self.assertNotIn("现有资料未提供完整可研必要性章节", chapter_three)
            self.assertNotIn("本节依据开工请示", chapter_three)
            self.assertNotIn("需建设单位确认", chapter_three)

    def test_chapter_three_supporting_files_only_allows_approval_or_survey_reply_images(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            from PIL import Image

            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n测试项目位于测试地点，建设内容为基础设施改造。\n"
                "二、项目建设必要性\n1.测试必要性\n本项目建设具有明确必要性。",
                encoding="utf-8",
            )
            survey_pdf = project_dir / "文物调查回函.pdf"
            Image.new("RGB", (500, 700), "white").save(survey_pdf, "PDF")
            (project_dir / "项目红线与正射影像.txt").write_text("红线与正射影像资料不得放入支持性文件取得情况。", encoding="utf-8")
            (project_dir / "KML_OVKML坐标文件.txt").write_text("KML/OVKML 坐标文件不得放入支持性文件取得情况。", encoding="utf-8")
            (project_dir / "考古调查勘探工作报告.txt").write_text("考古调查勘探工作报告不得放入支持性文件取得情况。", encoding="utf-8")
            (project_dir / "考古勘探项目验收意见书.txt").write_text("验收意见书不得放入支持性文件取得情况。", encoding="utf-8")
            (project_dir / "既有文物影响评估报告.txt").write_text("既有文评不得放入支持性文件取得情况。", encoding="utf-8")
            (project_dir / "文物保护方案.txt").write_text("文物保护方案不得放入支持性文件取得情况。", encoding="utf-8")

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            supporting_files = read_jsonl(output_dir / "processing_output" / "supporting_files.jsonl")
            self.assertEqual([row["source_file"] for row in supporting_files], ["文物调查回函.pdf"])
            self.assertTrue(supporting_files[0]["image_paths"], "支持性 PDF 必须转为图片供正文插入")
            for image_path in supporting_files[0]["image_paths"]:
                self.assertTrue((output_dir / image_path).exists(), image_path)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_three = report.split("## 三、建设项目规划概况", 1)[1].split("## 四、项目用地范围与文物空间分布关系", 1)[0]
            self.assertIn("### （三）支持性文件取得情况", chapter_three)
            self.assertIn("文物调查回函.pdf", chapter_three)
            self.assertIn("![文物调查回函.pdf", chapter_three)
            for forbidden in ["项目红线", "正射影像", "KML", "OVKML", "坐标文件", "考古调查勘探工作报告", "验收意见书", "既有文物影响评估报告", "文物保护方案"]:
                self.assertNotIn(forbidden, chapter_three)

    def test_chapter_three_construction_and_operation_plan_uses_complete_feasibility_subsections(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n测试项目位于测试地点，建设内容为基础设施改造。\n"
                "二、项目建设必要性\n1.测试必要性\n本项目建设具有明确必要性。\n"
                "三、建设及运营方案\n"
                "1.建设规模与技术指标\n"
                "项目扩建总建筑面积7386平方米，扩建后寺院总占地面积16919.64平方米。"
                "建设内容包括山门、念佛堂、寮房、禅堂、鼓楼等基础设施。\n\n"
                "主要技术指标包括建筑面积、占地面积、给排水、电气、消防和道路硬化等内容，"
                "本段为可研原文中的技术指标说明，不得改写为摘要。\n"
                "2.项目运营方案与规模\n"
                "项目建成后主要服务于寺院宗教活动、传统文化展示、公共服务和场所安全管理。"
                "运营规模按日常宗教活动、节庆活动和游客接待分级组织，配套客堂、流通处和展览室承担接待服务功能。\n"
                "四、投资估算\n本章不属于建设及运营方案。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_three = report.split("## 三、建设项目规划概况", 1)[1].split("## 四、项目用地范围与文物空间分布关系", 1)[0]
            self.assertIn("### （四）建设及运营方案", chapter_three)
            self.assertIn("#### 1.建设规模与技术指标", chapter_three)
            self.assertIn("项目扩建总建筑面积7386平方米，扩建后寺院总占地面积16919.64平方米", chapter_three)
            self.assertIn("本段为可研原文中的技术指标说明，不得改写为摘要", chapter_three)
            self.assertIn("#### 2.项目运营方案与规模", chapter_three)
            self.assertIn("项目建成后主要服务于寺院宗教活动、传统文化展示、公共服务和场所安全管理", chapter_three)
            self.assertIn("运营规模按日常宗教活动、节庆活动和游客接待分级组织", chapter_three)
            self.assertNotIn("本章不属于建设及运营方案", chapter_three)
            self.assertNotIn("稍加提炼整理", chapter_three)

    def test_chapter_four_opening_lead_is_fixed_before_first_subsection(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n测试项目位于测试地点，建设内容为基础设施改造。\n"
                "二、项目建设必要性\n1.测试必要性\n本项目建设具有明确必要性。",
                encoding="utf-8",
            )
            (project_dir / "文物调查回函.txt").write_text(
                "经查阅现有不可移动文物名录及现场调查，发现测试项目拟用地范围涉及桌子山秦长城东风农场七队长城3段建设控制地带。"
                "距桌子山秦长城东风农场七队长城3段本体最近处约192米。因此，本报告主要针对桌子山秦长城东风农场七队长城3段进行文物影响评估。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "测试项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_four = report.split("## 四、项目用地范围与文物空间分布关系", 1)[1].split("## 五、建设项目可能对文物造成的影响分析与评估", 1)[0]
            fixed_lead = (
                "本着“既有利于文物保护、又有利于基本建设”的“两利”方针，"
                "在建设施工前对测试项目涉及范围进行了现场勘察，"
                "确定桌子山秦长城东风农场七队长城3段与本项目的位置关系。"
            )
            self.assertIn(fixed_lead, chapter_four)
            self.assertIn("### （一）项目用地与文物相对位置关系", chapter_four)
            self.assertLess(chapter_four.index(fixed_lead), chapter_four.index("### （一）项目用地与文物相对位置关系"))
            self.assertNotIn("根据现有空间资料，", chapter_four.split("### （一）项目用地与文物相对位置关系", 1)[0])

    def test_chapter_four_first_section_uses_corner_coordinate_table_and_fixed_zoning_text_without_auto_figures(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            from openpyxl import Workbook
            from PIL import Image

            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n海南区觉海寺改扩建配套设施改造项目位于海南区拉僧庙镇赛汗乌素村。\n"
                "二、项目建设必要性\n1.测试必要性\n本项目建设具有明确必要性。",
                encoding="utf-8",
            )
            (project_dir / "文物调查回函.txt").write_text(
                "根据内蒙古自治区现有不可移动文物名录、内蒙古自治区长城资源调查数据库及现场调查情况，"
                "海南区觉海寺改扩建配套设施改造项目用地范围涉及乌海市海南区第五批自治区级文物保护单位"
                "桌子山秦长城——东风农场七队长城3段建设控制地带，不涉及其保护范围。",
                encoding="utf-8",
            )
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["序号", "经度（E）", "纬度（N）"])
            sheet.append(["J1", "106°45'34.3880\"", "39°29'56.0821\""])
            sheet.append(["J2", "106°45'35.0000\"", "39°29'57.0000\""])
            workbook.save(project_dir / "项目拐点经纬度坐标表.xlsx")
            Image.new("RGB", (20, 20), "white").save(project_dir / "制图成果项目红线示意图.png")
            Image.new("RGB", (20, 20), "white").save(project_dir / "勘探报告项目勘探区图.png")

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "海南区觉海寺改扩建配套设施改造项目",
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            coordinates = read_jsonl(output_dir / "processing_output" / "project_corner_coordinates.jsonl")
            self.assertEqual([row["point_id"] for row in coordinates], ["J1", "J2"])

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_four = report.split("## 四、项目用地范围与文物空间分布关系", 1)[1].split("## 五、建设项目可能对文物造成的影响分析与评估", 1)[0]
            first_section = chapter_four.split("### （一）项目用地与文物相对位置关系", 1)[1]
            self.assertIn("海南区觉海寺改扩建配套设施改造项目位于海南区拉僧庙镇赛汗乌素村", first_section)
            self.assertIn("项目拐点坐标见表4。", first_section)
            self.assertIn("表4 海南区觉海寺改扩建配套设施改造项目拐点坐标表", first_section)
            self.assertIn("| 序号 | 经度（E） | 纬度（N） |", first_section)
            self.assertIn("| J1 | 106°45'34.3880\" | 39°29'56.0821\" |", first_section)
            self.assertIn(
                "根据内蒙古自治区现有不可移动文物名录、内蒙古自治区长城资源调查数据库及现场调查情况，"
                "海南区觉海寺改扩建配套设施改造项目用地范围涉及乌海市海南区第五批自治区级文物保护单位"
                "桌子山秦长城——东风农场七队长城3段建设控制地带，不涉及其保护范围。",
                first_section,
            )
            self.assertNotIn("制图成果", chapter_four)
            self.assertNotIn("勘探报告项目勘探区图", chapter_four)
            self.assertNotIn("图件索引已建立", chapter_four)

    def test_chapter_four_project_named_site_selection_analysis_uses_complete_feasibility_section(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            project_name = "海南区觉海寺改扩建配套设施改造项目"
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n海南区觉海寺改扩建配套设施改造项目位于海南区拉僧庙镇赛汗乌素村。\n"
                "二、项目建设必要性\n1.测试必要性\n本项目建设具有明确必要性。\n"
                "三、建设及运营方案\n1.建设规模与技术指标\n本节为建设规模。\n2.项目运营方案与规模\n本节为运营规模。\n"
                "四、选址分析\n"
                "本项目为海南区觉海寺改扩建配套设施改造项目，属于在现有宗教活动场所范围内的升级改造工程，"
                "不涉及新增场址或线路建设。为确保选址合理性与合规性，本次分析基于“原址改扩建”核心原则，"
                "通过对原址深化利用方案的多维度论证，并结合虚拟备选方案比选，明确场址土地权属、利用状况及相关合规性条件，"
                "为项目落地提供选址层面的坚实支撑。\n\n"
                "1.原址深化利用方案\n"
                "原址深化利用方案依托既有寺院用地和配套设施条件，能够保持宗教活动组织和场地使用连续性。\n"
                "2.备选方案比较\n"
                "经比较，异地新建方案不符合现阶段项目建设条件，邻近地块补充方案需另行开展用地论证。\n"
                "五、其他章节\n这部分不属于选址分析。",
                encoding="utf-8",
            )
            (project_dir / "文物调查回函.txt").write_text(
                "经查阅现有不可移动文物名录及现场调查，发现海南区觉海寺改扩建配套设施改造项目拟用地范围涉及桌子山秦长城东风农场七队长城3段建设控制地带。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    project_name,
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_four = report.split("## 四、项目用地范围与文物空间分布关系", 1)[1].split("## 五、建设项目可能对文物造成的影响分析与评估", 1)[0]
            self.assertIn(f"### （二）{project_name}选址分析", chapter_four)
            self.assertNotIn("### （二）项目选址分析", chapter_four)
            self.assertIn("本项目为海南区觉海寺改扩建配套设施改造项目，属于在现有宗教活动场所范围内的升级改造工程", chapter_four)
            self.assertIn("1.原址深化利用方案", chapter_four)
            self.assertIn("原址深化利用方案依托既有寺院用地和配套设施条件", chapter_four)
            self.assertIn("2.备选方案比较", chapter_four)
            self.assertIn("经比较，异地新建方案不符合现阶段项目建设条件", chapter_four)
            self.assertNotIn("这部分不属于选址分析", chapter_four)
            self.assertNotIn("现有资料未提供完整多方案比选文本", chapter_four)
            self.assertNotIn("初步分析", chapter_four)

    def test_chapter_four_exploration_work_uses_fixed_text_without_images_or_freewriting(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            project_name = "海南区觉海寺改扩建配套设施改造项目"
            company_name = "内蒙古文物考古工程有限公司"
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n海南区觉海寺改扩建配套设施改造项目位于海南区拉僧庙镇赛汗乌素村。\n"
                "二、项目建设必要性\n本项目建设具有明确必要性。\n"
                "四、选址分析\n本项目采用原址改扩建方案。",
                encoding="utf-8",
            )
            (project_dir / "文物调查回函.txt").write_text(
                "根据内蒙古自治区现有不可移动文物名录、内蒙古自治区长城资源调查数据库及现场调查情况，"
                "海南区觉海寺改扩建配套设施改造项目用地范围涉及桌子山秦长城东风农场七队长城3段建设控制地带，不涉及其保护范围。",
                encoding="utf-8",
            )
            (project_dir / f"{project_name}考古调查、勘探工作报告.txt").write_text(
                f"海南区觉海寺改扩建配套设施改造项目委托{company_name}开展考古调查、勘探工作。"
                "本项目考古调查、勘探面积为7386平方米，采用考古调查、普通勘探和重点勘探相结合的方式，"
                "主要以洛阳铲人工钻探为工具。勘探中采用2米x2米梅花孔布孔，必要时以1米x1米梅花孔加密。"
                "通过考古调查及勘探，认定在勘探区域内未发现文物及遗迹现象。",
                encoding="utf-8",
            )
            (project_dir / "图4勘探单元布置图.png").write_bytes(b"placeholder image")

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    project_name,
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_four = report.split("## 四、项目用地范围与文物空间分布关系", 1)[1].split("## 五、建设项目可能对文物造成的影响分析与评估", 1)[0]
            self.assertIn("### （三）勘探工作", chapter_four)
            self.assertNotIn("### （三）调查、勘探或核查工作", chapter_four)
            self.assertIn(
                f"因此，{project_name}委托{company_name}针对该项目区域展开了全面的考古调查勘探。",
                chapter_four,
            )
            self.assertIn(f"根据《{project_name}考古调查、勘探工作报告》", chapter_four)
            self.assertIn("该次考古调查、勘探实际完成面积为7386平方米", chapter_four)
            self.assertIn("认定在勘探区域内未发现文物及遗迹现象。", chapter_four)
            self.assertNotIn("洛阳铲", chapter_four)
            self.assertNotIn("梅花孔", chapter_four)
            self.assertNotIn("GPS测点", chapter_four)
            self.assertNotIn("图4勘探单元布置图", chapter_four)
            self.assertNotIn("![", chapter_four)

    def test_chapter_five_opening_lead_is_fixed_before_first_subsection(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "project"
            output_dir = root / "work"
            project_dir.mkdir()
            project_name = "海南区觉海寺改扩建配套设施改造项目"
            heritage_name = "东风农场七队长城3段"
            (project_dir / "项目可研.txt").write_text(
                "一、项目概况\n海南区觉海寺改扩建配套设施改造项目位于海南区拉僧庙镇赛汗乌素村。\n"
                "二、项目建设必要性\n本项目建设具有明确必要性。\n"
                "四、选址分析\n本项目采用原址改扩建方案。",
                encoding="utf-8",
            )
            (project_dir / "文物调查回函.txt").write_text(
                "经查阅现有不可移动文物名录及现场调查，发现海南区觉海寺改扩建配套设施改造项目拟用地范围涉及"
                f"{heritage_name}建设控制地带。",
                encoding="utf-8",
            )

            for command in [
                [
                    "run_material_processing.py",
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    project_name,
                    "--覆盖",
                ],
                ["run_fact_extraction.py", "--workspace", str(output_dir), "--覆盖"],
                ["run_analysis.py", "--工作目录", str(output_dir), "--覆盖"],
                ["run_report_assembly.py", "--workspace", str(output_dir), "--覆盖"],
            ]:
                result = run_script(command[0], *command[1:])
                self.assertEqual(result.returncode, 0, result.stderr)

            report = (output_dir / "report_clean.md").read_text(encoding="utf-8")
            chapter_five = report.split("## 五、建设项目可能对文物造成的影响分析与评估", 1)[1].split("## 六、减缓措施建议", 1)[0]
            expected_first = (
                "由于文化遗产的不可再生性，文物、遗产领域的法律法规及相关文件均反复指出，"
                "发展建设项目规划必须将文物保护纳入考量，通过履行规定的审批流程，确保项目建设尽可能规避对文物的破坏。"
                "建设单位应提供《文物保护方案》《文物影响评估报告》等文件，确保在施工前充分考虑到涉及的文物情况，"
                "避免建设活动导致相关文物与遗产的完整性的丧失，从而破坏遗产的真实性。"
            )
            expected_second = (
                f"因此，需要对{project_name}可能对{heritage_name}建设控制地带产生的各类影响进行谨慎评估，"
                "并制定全面的保护措施，从而最大程度地减少项目在建设及运营过程中对不可移动文物本体及其周边环境的负面影响。"
                "本次评估涵盖项目方案阶段的合法合规性评估，以及施工及运营阶段的影响预测与评估。"
            )
            self.assertLess(chapter_five.index(expected_first), chapter_five.index("### （一）项目方案设计合规性评估"))
            self.assertLess(chapter_five.index(expected_second), chapter_five.index("### （一）项目方案设计合规性评估"))
            self.assertNotIn("本章先以影响识别矩阵", chapter_five)


if __name__ == "__main__":
    unittest.main()
