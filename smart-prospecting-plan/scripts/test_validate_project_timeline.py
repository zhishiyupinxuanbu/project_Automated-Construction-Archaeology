#!/usr/bin/env python3
"""Regression tests for the project timeline hard gate."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

import validate_project_timeline as timeline


class TimelineValidatorTests(unittest.TestCase):
    def test_valid_request_plan_and_watermark_dates_pass(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            photos = Path(temp)
            image = photos / "site.jpg"
            image.write_bytes(b"fake")
            image.with_suffix(".ocr.txt").write_text("现场水印 2026年7月13日", encoding="utf-8")

            issues = timeline.validate_timeline(
                {
                    "请示日期": "2026年7月10日",
                    "回函日期": "2026年7月8日",
                    "文物调查结束日期": "2026年7月7日",
                    "开始日期": "2026年7月12日",
                    "结束日期": "2026年7月15日",
                },
                photos_dir=photos,
                require_request_sequence=True,
                require_photos=True,
            )

        self.assertEqual([], issues)

    def test_request_to_planned_start_must_be_two_or_three_days(self) -> None:
        issues = timeline.validate_timeline(
            {
                "请示日期": "2026年7月10日",
                "回函日期": "2026年7月8日",
                "文物调查结束日期": "2026年7月7日",
                "开始日期": "2026年7月11日",
                "结束日期": "2026年7月15日",
            },
            require_request_sequence=True,
        )

        self.assertTrue(any(issue.code == "REQUEST_TO_START_DELTA" for issue in issues))

    def test_request_can_be_same_day_as_reply(self) -> None:
        issues = timeline.validate_timeline(
            {
                "请示日期": "2026年7月10日",
                "回函日期": "2026年7月10日",
                "文物调查结束日期": "2026年7月7日",
                "开始日期": "2026年7月12日",
                "结束日期": "2026年7月15日",
            },
            require_request_sequence=True,
        )

        self.assertFalse(any(issue.code == "REQUEST_BEFORE_REPLY" for issue in issues))

    def test_photo_watermark_date_must_fall_inside_plan_range(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            photos = Path(temp)
            image = photos / "site.jpg"
            image.write_bytes(b"fake")
            image.with_suffix(".ocr.txt").write_text("水印时间 2026-07-16", encoding="utf-8")

            issues = timeline.validate_timeline(
                {
                    "开始日期": "2026年7月12日",
                    "结束日期": "2026年7月15日",
                },
                photos_dir=photos,
                require_photos=True,
            )

        self.assertTrue(any(issue.code == "PHOTO_WATERMARK_OUT_OF_RANGE" for issue in issues))

    def test_dot_separated_watermark_date_is_detected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            photos = Path(temp)
            image = photos / "site.jpg"
            image.write_bytes(b"fake")
            image.with_suffix(".ocr.txt").write_text("时间：2026.06.17 10:37", encoding="utf-8")

            issues = timeline.validate_timeline(
                {
                    "开始日期": "2026年6月5日",
                    "结束日期": "2026年6月30日",
                },
                photos_dir=photos,
                require_photos=True,
            )

        self.assertFalse(timeline.error_issues(issues))

    def test_photo_without_detected_watermark_date_warns_without_error(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            photos = Path(temp)
            image = photos / "site.jpg"
            image.write_bytes(b"fake")

            issues = timeline.validate_timeline(
                {
                    "开始日期": "2026年7月12日",
                    "结束日期": "2026年7月15日",
                },
                photos_dir=photos,
                require_photos=True,
            )

        self.assertFalse(timeline.error_issues(issues))
        self.assertTrue(any(issue.code == "PHOTO_WATERMARK_DATE_UNDETECTED" for issue in issues))

    def test_docx_embedded_image_watermark_date_is_checked(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            docx = root / "plan.docx"
            with ZipFile(docx, "w") as zf:
                zf.writestr("word/document.xml", "<w:document />")
                zf.writestr("word/media/image1.jpg", b"fake")
            sidecar_dir = docx.with_suffix(".docx.ocr")
            sidecar_dir.mkdir()
            (sidecar_dir / "image1.ocr.txt").write_text("现场水印 2026年7月16日", encoding="utf-8")

            issues = timeline.validate_timeline(
                {
                    "开始日期": "2026年7月12日",
                    "结束日期": "2026年7月15日",
                },
                docx_paths=[docx],
                require_photos=True,
            )

        self.assertTrue(any(issue.code == "PHOTO_WATERMARK_OUT_OF_RANGE" for issue in issues))


if __name__ == "__main__":
    unittest.main()
