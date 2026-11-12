import os
import subprocess
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook


class WorkbookFormatTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.directory = Path(directory.name)

    def assert_invalid_file(self, path, message):
        original = path.read_bytes()
        script = Path(__file__).resolve().parents[1] / "networking" / "followups.py"
        for arguments in [[], ["--due-through"]]:
            with self.subTest(arguments=arguments):
                result = subprocess.run(
                    [sys.executable, str(script), *arguments],
                    input="2026-11-07\n", text=True, capture_output=True,
                    env={**os.environ, "NETWORKING_WORKBOOK": str(path)},
                    timeout=10,
                )
                self.assertEqual(result.returncode, 1)
                self.assertEqual(result.stderr, message + "\n")
                self.assertEqual(
                    result.stdout, "Enter follow-up date (YYYY-MM-DD): "
                )
                self.assertNotIn("Traceback", result.stderr)
                self.assertNotIn(str(path), result.stdout + result.stderr)
                self.assertEqual(path.read_bytes(), original)

    def test_unsupported_file_extensions(self):
        for suffix in [".xls", ".xlsb", ".csv"]:
            with self.subTest(suffix=suffix):
                path = self.directory / ("synthetic contacts" + suffix)
                path.write_bytes(b"Synthetic unsupported file")
                self.assert_invalid_file(
                    path, "Unsupported workbook format. Use an Excel .xlsx workbook."
                )

    def test_non_zip_xlsx(self):
        path = self.directory / "synthetic contacts.xlsx"
        for content in [b"", b"Synthetic text, not an Excel workbook"]:
            with self.subTest(content=content):
                path.write_bytes(content)
                self.assert_invalid_file(
                    path, "Workbook is not a valid Excel ZIP file. "
                    "Check the file or restore a valid backup."
                )

    def test_truncated_xlsx_archive(self):
        path = self.directory / "synthetic contacts.xlsx"
        workbook = Workbook()
        try:
            workbook.save(path)
        finally:
            workbook.close()
        # Remove the ZIP end-of-central-directory record from a real workbook.
        path.write_bytes(path.read_bytes()[:-22])
        self.assert_invalid_file(
            path, "Workbook is not a valid Excel ZIP file. "
            "Check the file or restore a valid backup."
        )


if __name__ == "__main__":
    unittest.main()
