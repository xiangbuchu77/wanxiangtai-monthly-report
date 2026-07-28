from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path
import sys
import types

try:
    import PIL  # noqa: F401
except ModuleNotFoundError:
    pil_module = types.ModuleType("PIL")
    image_module = types.ModuleType("PIL.Image")
    pil_module.Image = image_module
    sys.modules["PIL"] = pil_module
    sys.modules["PIL.Image"] = image_module

from runtime.services.report_core.taobao_report_mvp.archive_utils import (
    ZipArchiveError,
    decode_zip_member_name,
    extract_zip_archive,
    sanitize_windows_filename,
)
from runtime.services.report_core.taobao_report_mvp.web_app import (
    FormPart,
    QCLAW_UPLOAD_SUFFIXES,
    expand_archives,
    save_uploads,
)
from runtime.services.report_core.taobao_report_mvp.report_workflow_agent import expand_zip_inputs


class ZipArchiveTests(unittest.TestCase):
    def test_upload_is_saved_then_extracted_exactly_once(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source_zip = root / "source.zip"
            with zipfile.ZipFile(source_zip, "w") as archive:
                archive.writestr("报表/店铺经营核心月报.xlsx", b"xlsx")
                archive.writestr("报表/营销场景报表.csv", b"a,b\n1,2\n")
            form = {
                "files": [
                    FormPart(
                        name="files",
                        filename="数据包.zip",
                        data=source_zip.read_bytes(),
                    )
                ]
            }

            saved = save_uploads(form, "files", root / "uploads", QCLAW_UPLOAD_SUFFIXES)
            self.assertEqual([path.suffix for path in saved], [".zip"])

            expanded = expand_archives([*saved, *saved], root / "unzipped", {".xlsx", ".csv"})
            self.assertEqual(len(expanded), 2)
            self.assertEqual({path.suffix for path in expanded}, {".xlsx", ".csv"})

    def test_extracts_chinese_names_and_sanitizes_windows_names(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source_zip = root / "reports.zip"
            with zipfile.ZipFile(source_zip, "w") as archive:
                archive.writestr("中文目录/店铺月报.xlsx", b"xlsx")
                archive.writestr("CON.csv", b"a,b\n")
                archive.writestr("bad<name>?.csv", b"a,b\n")
                archive.writestr("__MACOSX/._hidden.xlsx", b"ignored")

            extracted = extract_zip_archive(source_zip, root / "out", {".xlsx", ".csv"})
            names = {path.name for path in extracted}
            self.assertIn("店铺月报.xlsx", names)
            self.assertIn("_CON.csv", names)
            self.assertIn("bad_name__.csv", names)
            self.assertEqual(len(extracted), 3)

    def test_command_line_path_reuses_the_same_zip_extractor(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source_zip = root / "reports.zip"
            with zipfile.ZipFile(source_zip, "w") as archive:
                archive.writestr("店铺经营核心月报.xlsx", b"xlsx")
                archive.writestr("营销场景报表.csv", b"a,b\n")

            expanded = expand_zip_inputs([source_zip, source_zip], root / "out")
            self.assertEqual(len(expanded), 2)

    def test_recovers_common_windows_gbk_member_name(self) -> None:
        expected = "店铺经营核心月报.xlsx"
        mojibake = expected.encode("gbk").decode("cp437")
        info = zipfile.ZipInfo(mojibake)
        info.flag_bits = 0
        self.assertEqual(decode_zip_member_name(info), expected)

    def test_rejects_corrupt_and_empty_archives_with_clear_errors(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            corrupt = root / "corrupt.zip"
            corrupt.write_bytes(b"not-a-zip")
            with self.assertRaisesRegex(ZipArchiveError, "损坏|有效"):
                extract_zip_archive(corrupt, root / "corrupt-out", {".xlsx"})

            empty = root / "empty.zip"
            with zipfile.ZipFile(empty, "w") as archive:
                archive.writestr("说明.txt", "no reports")
            with self.assertRaisesRegex(ZipArchiveError, "没有可处理"):
                extract_zip_archive(empty, root / "empty-out", {".xlsx"})

    def test_safe_filename_is_flat_and_windows_compatible(self) -> None:
        self.assertEqual(sanitize_windows_filename(r"folder\CON.xlsx"), "_CON.xlsx")
        self.assertEqual(sanitize_windows_filename("../../report.xlsx. "), "report.xlsx")
        self.assertLessEqual(len(sanitize_windows_filename("长" * 200 + ".xlsx")), 120)


if __name__ == "__main__":
    unittest.main()
