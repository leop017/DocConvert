import os
import tempfile
import unittest
from io import StringIO
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

from docconvert.cli import main_cli


def _make_xlsx(path, sheet_name="Sheet1", rows=None):
    if rows is None:
        rows = [["A", "B"], [1, 2]]
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in rows:
        ws.append(row)
    wb.save(path)


class TestCLIParsing(unittest.TestCase):

    def test_version_flag(self):
        with self.assertRaises(SystemExit) as cm:
            main_cli(["--version"])
        self.assertEqual(cm.exception.code, 0)

    def test_help_flag(self):
        with self.assertRaises(SystemExit) as cm:
            main_cli(["--help"])
        self.assertEqual(cm.exception.code, 0)

    def test_no_command_prints_help_returns_one(self, capsys=None):
        result = main_cli([])
        self.assertEqual(result, 1)

    def test_unknown_subcommand_returns_one(self):
        with self.assertRaises(SystemExit) as cm:
            main_cli(["unknown"])
        self.assertEqual(cm.exception.code, 2)

    def test_convert_requires_files(self):
        with self.assertRaises(SystemExit) as cm:
            main_cli(["convert"])
        self.assertEqual(cm.exception.code, 2)


class TestCLIErrorPaths(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.out_dir = os.path.join(self.tmp.name, "out")
        os.makedirs(self.out_dir, exist_ok=True)

    def test_missing_file_returns_one(self):
        result = main_cli([
            "convert", os.path.join(self.tmp.name, "missing.xlsx"),
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 1)

    def test_unsupported_format_returns_one(self):
        bad_file = os.path.join(self.tmp.name, "bad.pdf")
        with open(bad_file, "w") as f:
            f.write("not a real file")
        result = main_cli([
            "convert", bad_file,
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 1)

    def test_mixed_valid_and_invalid(self):
        good_xlsx = os.path.join(self.tmp.name, "good.xlsx")
        _make_xlsx(good_xlsx)
        bad_txt = os.path.join(self.tmp.name, "bad.txt")
        with open(bad_txt, "w") as f:
            f.write("hello")

        result = main_cli([
            "convert", good_xlsx, bad_txt,
            "--output", self.out_dir,
        ])
        # At least one failure means return code 1
        self.assertEqual(result, 1)

    def test_unsupported_format_prints_error_to_stderr(self):
        bad_file = os.path.join(self.tmp.name, "bad.pdf")
        with open(bad_file, "w") as f:
            f.write("not a real file")
        with mock.patch("sys.stderr", new_callable=StringIO) as mock_stderr:
            main_cli(["convert", bad_file, "--output", self.out_dir])
        stderr_text = mock_stderr.getvalue()
        self.assertIn("unsupported format", stderr_text)

    def test_missing_file_prints_error_to_stderr(self):
        with mock.patch("sys.stderr", new_callable=StringIO) as mock_stderr:
            main_cli(["convert", "/no/such/file.xlsx", "--output", self.out_dir])
        stderr_text = mock_stderr.getvalue()
        self.assertIn("file not found", stderr_text)


class TestCLIConversion(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.out_dir = os.path.join(self.tmp.name, "out")
        os.makedirs(self.out_dir, exist_ok=True)
        self.xlsx = os.path.join(self.tmp.name, "test.xlsx")
        _make_xlsx(self.xlsx)

    def test_successful_xlsx_html_conversion(self):
        result = main_cli([
            "convert", self.xlsx,
            "--format", "html",
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 0)
        out_path = os.path.join(self.out_dir, "TEST_Sheet1.html")
        self.assertTrue(os.path.exists(out_path))

    def test_successful_xlsx_md_conversion(self):
        result = main_cli([
            "convert", self.xlsx,
            "--format", "md",
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 0)
        out_path = os.path.join(self.out_dir, "TEST_Sheet1.md")
        self.assertTrue(os.path.exists(out_path))

    def test_successful_xlsx_json_conversion(self):
        result = main_cli([
            "convert", self.xlsx,
            "--format", "json",
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 0)
        out_path = os.path.join(self.out_dir, "TEST_Sheet1.json")
        self.assertTrue(os.path.exists(out_path))

    def test_enhanced_flag_accepted(self):
        result = main_cli([
            "convert", self.xlsx,
            "--format", "md",
            "--enhanced",
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 0)

    def test_sheet_selector_accepted(self):
        result = main_cli([
            "convert", self.xlsx,
            "--sheet", "Sheet1",
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 0)

    def test_verbose_flag_accepted(self):
        result = main_cli([
            "--verbose",
            "convert", self.xlsx,
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 0)

    def test_multiple_files(self):
        xlsx2 = os.path.join(self.tmp.name, "test2.xlsx")
        _make_xlsx(xlsx2, rows=[["X", "Y"], [3, 4]])
        result = main_cli([
            "convert", self.xlsx, xlsx2,
            "--format", "html",
            "--output", self.out_dir,
        ])
        self.assertEqual(result, 0)


class TestCLIProgressCallback(unittest.TestCase):

    def test_progress_callback_prints_to_stderr(self):
        from docconvert.cli import _cli_progress
        from docconvert.models import ProgressEvent

        with mock.patch("sys.stderr", new_callable=StringIO) as mock_stderr:
            _cli_progress(ProgressEvent(message="converting", progress=0.5))
            stderr_text = mock_stderr.getvalue()
            self.assertIn("converting", stderr_text)

        with mock.patch("sys.stderr", new_callable=StringIO) as mock_stderr:
            _cli_progress(ProgressEvent(message="done", progress=1.0, done=True))
            stderr_text = mock_stderr.getvalue()
            self.assertIn("done", stderr_text)


if __name__ == "__main__":
    unittest.main()
