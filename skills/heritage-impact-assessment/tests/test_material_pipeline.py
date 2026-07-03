from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPT = SKILL_DIR / "scripts" / "run_material_processing.py"


class MaterialProcessingPipelineTest(unittest.TestCase):
    def test_material_processing_writes_v022_stage_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project_dir = root / "sample_project"
            output_dir = root / "work"
            project_dir.mkdir()

            (project_dir / "项目说明.txt").write_text(
                "海南区觉海寺改扩建配套设施改造项目，建设内容包括山门、念佛堂、寮房。",
                encoding="utf-8",
            )
            (project_dir / "文物调查报告.txt").write_text(
                "涉及桌子山秦长城东风农场七队长城3段，项目位于建设控制地带。",
                encoding="utf-8",
            )
            (project_dir / "项目用地范围.kml").write_text(
                "<kml><Placemark><name>项目用地范围</name></Placemark></kml>",
                encoding="utf-8",
            )
            (project_dir / "项目位置示意图.png").write_bytes(b"not-a-real-png-but-indexable")

            result = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT),
                    "--项目资料目录",
                    str(project_dir),
                    "--输出目录",
                    str(output_dir),
                    "--项目名称",
                    "海南区觉海寺改扩建配套设施改造项目",
                    "--覆盖",
                ],
                text=True,
                capture_output=True,
            )

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue((output_dir / "processing_output" / "manifest.json").exists())
            self.assertTrue((output_dir / "processing_output" / "files.csv").exists())
            self.assertTrue((output_dir / "processing_output" / "text_index.jsonl").exists())
            self.assertTrue((output_dir / "processing_output" / "figure_index.jsonl").exists())
            self.assertTrue((output_dir / "evidence" / "evidence_register.jsonl").exists())
            self.assertTrue((output_dir / "extracted_figures").exists())
            self.assertTrue((output_dir / "module_state.json").exists())
            self.assertTrue((output_dir / "run_state" / "material_processing.module_done.json").exists())
            self.assertTrue((output_dir / "next_prompts" / "next_prompt_fact_extraction.md").exists())

            manifest = json.loads((output_dir / "processing_output" / "manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["project_name"], "海南区觉海寺改扩建配套设施改造项目")
            self.assertIn("files_csv", manifest["outputs"])

            state = json.loads((output_dir / "module_state.json").read_text(encoding="utf-8"))
            self.assertFalse(state["blocked"])
            self.assertIn("material_processing", state["completed_stages"])

            gaps = (output_dir / "processing_output" / "user_blocking_gaps.jsonl").read_text(encoding="utf-8").strip()
            self.assertEqual(gaps, "")


if __name__ == "__main__":
    unittest.main()
