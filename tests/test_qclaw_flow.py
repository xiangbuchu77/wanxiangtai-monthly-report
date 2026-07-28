from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from runtime.services.report_core.taobao_report_mvp import screenshot_input, web_app


class QClawFlowTests(unittest.TestCase):
    def test_generation_requires_an_explicit_action(self) -> None:
        self.assertTrue(web_app.wants_report_generation("请生成月报"))
        self.assertTrue(web_app.wants_report_generation("制作半月报"))
        self.assertTrue(web_app.wants_report_generation("导出 Excel"))
        self.assertFalse(web_app.wants_report_generation("月报的逻辑是什么"))
        self.assertFalse(web_app.wants_report_generation("帮我看看这个 Excel"))
        route = web_app.classify_qclaw_intent("", [Path("report.xlsx")], [])
        self.assertEqual(route["intent"], "awaiting_report_request")

    def test_daily_cache_does_not_duplicate_current_upload(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "店铺经营核心月报.xlsx"
            source.write_bytes(b"same-report-content")
            with patch.object(web_app, "DAILY_SOURCE_CACHE_DIR", root / "cache"):
                first = web_app.archive_daily_report_files([source], "测试店铺")
                second = web_app.archive_daily_report_files([source], "测试店铺")
                resolved, choice = web_app.resolve_daily_store_files_for_generation(
                    [source], "测试店铺", {"session_key": "test"}, "month", None
                )
                cached_files = web_app.daily_cached_files_for_store("测试店铺")

            self.assertIsNone(choice)
            self.assertEqual(len(first.stores["测试店铺"]), 1)
            self.assertEqual(len(second.stores["测试店铺"]), 1)
            self.assertEqual(len(cached_files), 1)
            self.assertEqual(resolved, [source.resolve()])

    def test_complementary_screenshot_records_merge_without_summing(self) -> None:
        records = [
            screenshot_input.normalize_metric_record(
                {
                    "统计日期": "2026-06-25",
                    "店铺名称": "测试店铺",
                    "访客数": 193,
                    "支付金额": 1062.58,
                }
            ),
            screenshot_input.normalize_metric_record(
                {
                    "统计日期": "2026-06-25",
                    "全站推广花费": 200,
                    "广告引导成交金额": 600,
                    "点击量": 100,
                }
            ),
            screenshot_input.normalize_metric_record(
                {
                    "统计日期": "2026-06-25",
                    "店铺名称": "测试店铺",
                    "访客数": 193,
                }
            ),
        ]
        merged, conflicts = screenshot_input.merge_complementary_records(records)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]["访客数"], 193.0)
        self.assertEqual(merged[0]["全站推广花费"], 200.0)
        self.assertEqual(merged[0]["点击量"], 100.0)
        self.assertEqual(conflicts, [])

    def test_screenshot_batch_creates_one_structured_input(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            images = [root / "core.png", root / "ad.png"]
            for image in images:
                image.write_bytes(b"image")
            payload = {
                "records": [
                    {"统计日期": "2026-06-25", "店铺名称": "测试店铺", "访客数": 193, "支付金额": 1062.58},
                    {"统计日期": "2026-06-25", "全站推广花费": 200, "点击量": 100},
                ]
            }
            with patch.object(screenshot_input, "image_data_url", return_value="data:image/png;base64,eA=="), patch.object(
                screenshot_input, "call_llm_vision_json", return_value=payload
            ) as vision:
                result = screenshot_input.extract_screenshot_report_files(images, root / "out", object())

            self.assertEqual(vision.call_count, 1)
            self.assertEqual(len(result.report_files), 1)
            frame = pd.read_excel(result.report_files[0])
            self.assertEqual(len(frame), 1)
            self.assertEqual(frame.loc[0, "访客数"], 193)
            self.assertEqual(frame.loc[0, "全站推广花费"], 200)

    def test_final_report_is_exposed_as_media_path(self) -> None:
        script = Path("skills/wanxiangtai-monthly-report/scripts/call_local_agent.py").resolve()
        spec = importlib.util.spec_from_file_location("wxt_call_local_agent", script)
        self.assertIsNotNone(spec)
        module = importlib.util.module_from_spec(spec)
        assert spec and spec.loader
        spec.loader.exec_module(module)
        with tempfile.TemporaryDirectory() as temp:
            report = Path(temp) / "测试店铺2026年06月月报.xlsx"
            report.write_bytes(b"xlsx")
            paths = module.final_report_media_paths(
                {"type": "report", "files": [{"kind": "final_report", "path": str(report)}]}
            )
            self.assertEqual(paths, [report.resolve()])


if __name__ == "__main__":
    unittest.main()
