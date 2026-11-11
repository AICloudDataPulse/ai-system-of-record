import os
import subprocess
import sys
import unittest
from datetime import date
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from networking.followups import main


HEADERS = ["Name", "Company", "Next Action", "Follow-up Due"]


class ContactSchemaTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "synthetic contacts.xlsx"

    def make_workbook(self, rows, title="Contacts"):
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = title
            for row in rows:
                sheet.append(row)
            workbook.save(self.path)
        finally:
            workbook.close()

    def run_cli(self, include_earlier=False):
        script = Path(__file__).resolve().parents[1] / "networking" / "followups.py"
        return subprocess.run(
            [sys.executable, str(script), *(["--due-through"] if include_earlier else [])],
            input="2026-11-07\n", text=True, capture_output=True,
            env={**os.environ, "NETWORKING_WORKBOOK": str(self.path)},
            timeout=10,
        )

    def assert_invalid(self, message):
        original = self.path.read_bytes()
        for include_earlier in [False, True]:
            with self.subTest(include_earlier=include_earlier):
                result = self.run_cli(include_earlier)
                self.assertEqual(result.returncode, 1)
                self.assertIn(message, result.stderr)
                self.assertNotIn("Traceback", result.stderr)
                self.assertNotIn("Total follow-ups", result.stdout)
                self.assertNotIn("Name:", result.stdout)
                self.assertEqual(self.path.read_bytes(), original)

    def test_missing_worksheet(self):
        self.make_workbook([HEADERS], title="Other")
        self.assert_invalid("missing the 'Contacts' worksheet")

    def test_empty_worksheet(self):
        self.make_workbook([])
        self.assert_invalid("missing its header row")

    def test_blank_first_row(self):
        self.make_workbook([[None] * 4, HEADERS])
        self.assert_invalid("missing its header row")

    def test_each_required_header_is_required(self):
        for name in HEADERS:
            with self.subTest(header=name):
                self.make_workbook([[header for header in HEADERS if header != name]])
                self.assert_invalid("Missing required headers: " + name)

    def test_each_required_header_must_be_unique(self):
        for name in HEADERS:
            with self.subTest(header=name):
                self.make_workbook([HEADERS + [name]])
                self.assert_invalid("Duplicate required headers: " + name)

    def test_headers_require_exact_spelling_and_spacing(self):
        for replacement in ["name", " Name", "Name "]:
            with self.subTest(replacement=replacement):
                self.make_workbook([[replacement] + HEADERS[1:]])
                self.assert_invalid("Missing required headers: Name")

    def test_header_only_returns_zero(self):
        self.make_workbook([HEADERS])
        for include_earlier in [False, True]:
            with self.subTest(include_earlier=include_earlier):
                result = self.run_cli(include_earlier)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(result.stderr, "")
                self.assertIn("Total follow-ups: 0\n", result.stdout)

    def test_reordered_headers_and_extra_columns(self):
        self.make_workbook([
            ["Extra", "Follow-up Due", "Next Action", "Company", "Name", "Extra"],
            ["ignored", date(2026, 11, 7), "Send update", "Example Labs", "Alex Example", "ignored"],
        ])
        original = self.path.read_bytes()
        for include_earlier in [False, True]:
            with self.subTest(include_earlier=include_earlier):
                result = self.run_cli(include_earlier)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(result.stderr, "")
                self.assertIn("Name: Alex Example\nCompany: Example Labs\nNext Action: Send update", result.stdout)
                self.assertIn("Total follow-ups: 1\n", result.stdout)
                self.assertEqual(self.path.read_bytes(), original)

    def test_workbook_is_closed_after_schema_error(self):
        for rows, title in [([], "Other"), ([], "Contacts"), ([HEADERS[:-1]], "Contacts")]:
            with self.subTest(rows=rows, title=title):
                self.make_workbook(rows, title)
                workbook = load_workbook(self.path, read_only=True, data_only=True)
                self.addCleanup(workbook.close)
                with (
                    patch("networking.followups.load_workbook", return_value=workbook),
                    patch.object(workbook, "close", wraps=workbook.close) as close,
                    patch("builtins.input", return_value="2026-11-07"),
                    patch.dict(os.environ, {"NETWORKING_WORKBOOK": str(self.path)}),
                    patch("sys.stderr", new_callable=StringIO),
                ):
                    with self.assertRaises(SystemExit) as raised:
                        main()
                    self.assertEqual(raised.exception.code, 1)
                    close.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
