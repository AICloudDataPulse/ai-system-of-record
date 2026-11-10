import unittest
import os
import subprocess
import sys
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook

from networking.followups import main


class FollowupQueryTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.workbook_path = Path(directory.name) / "synthetic_contacts.xlsx"
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = "Contacts"
            sheet.append(["Name", "Company", "Next Action", "Follow-up Due"])
            sheet.append([
                "Alex Example", "Example Labs", "Send an update",
                date(2026, 11, 7),
            ])
            sheet.append([
                "Sam Sample", "Sample Systems", "Arrange a meeting",
                datetime(2026, 11, 7, 15, 30),
            ])
            sheet.append([
                "Taylor Test", "Test Company", "Follow up tomorrow",
                date(2026, 11, 8),
            ])
            sheet.append([
                "Casey Demo", "Demo Company", "No date scheduled", None,
            ])
            sheet.append([
                "Morgan Text", "Example Company", "Text date is skipped",
                "2026-11-06",
            ])
            workbook.save(self.workbook_path)
        finally:
            workbook.close()

    def run_query(self, target_date, *, include_earlier=False):
        with (
            patch("builtins.input", return_value=target_date),
            patch.dict(
                "os.environ",
                {"NETWORKING_WORKBOOK": str(self.workbook_path)},
            ),
            patch("sys.stdout", new_callable=StringIO) as stdout,
            patch("sys.stderr", new_callable=StringIO) as stderr,
        ):
            main(include_earlier=include_earlier)

        self.assertEqual(stderr.getvalue(), "")
        return stdout.getvalue()

    def test_matching_contacts_and_total(self):
        output = self.run_query("2026-11-07")
        self.assertEqual(
            output,
            "\nName: Alex Example\nCompany: Example Labs\n"
            "Next Action: Send an update\n"
            "\nName: Sam Sample\nCompany: Sample Systems\n"
            "Next Action: Arrange a meeting\n"
            "\nTotal follow-ups: 2\n",
        )

    def test_other_dates_and_blank_dates_are_excluded(self):
        output = self.run_query("2026-11-08")
        self.assertIn("Name: Taylor Test\n", output)
        for name in ["Alex Example", "Sam Sample", "Casey Demo"]:
            with self.subTest(name=name):
                self.assertNotIn(name, output)
        self.assertIn("Total follow-ups: 1\n", output)

    def test_no_matching_contacts(self):
        self.assertEqual(
            self.run_query("2026-11-09"),
            "\nTotal follow-ups: 0\n",
        )

    def test_query_preserves_workbook_bytes(self):
        original = self.workbook_path.read_bytes()
        self.run_query("2026-11-07")
        self.assertEqual(self.workbook_path.read_bytes(), original)

    def test_due_through_includes_earlier_and_boundary_dates(self):
        output = self.run_query("2026-11-08", include_earlier=True)
        names = [line for line in output.splitlines() if line.startswith("Name:")]
        self.assertEqual(names, [
            "Name: Alex Example", "Name: Sam Sample", "Name: Taylor Test",
        ])
        self.assertIn("Company: Sample Systems\nNext Action: Arrange a meeting", output)
        self.assertIn("Total follow-ups: 3\n", output)

    def test_due_through_excludes_later_blank_and_text_dates(self):
        output = self.run_query("2026-11-07", include_earlier=True)
        self.assertEqual(output, self.run_query("2026-11-07"))

    def test_due_through_no_matches(self):
        self.assertEqual(
            self.run_query("2026-11-06", include_earlier=True),
            "\nTotal follow-ups: 0\n",
        )

    def test_due_through_preserves_workbook_bytes(self):
        original = self.workbook_path.read_bytes()
        self.run_query("2026-11-08", include_earlier=True)
        self.assertEqual(self.workbook_path.read_bytes(), original)

    def run_cli(self, arguments, target_date):
        script = Path(__file__).resolve().parents[1] / "networking" / "followups.py"
        return subprocess.run(
            [sys.executable, str(script), *arguments],
            input=target_date + "\n", text=True, capture_output=True,
            env={**os.environ, "NETWORKING_WORKBOOK": str(self.workbook_path)},
            timeout=10,
        )

    def test_cli_default_and_due_through(self):
        for arguments, include_earlier in [([], False), (["--due-through"], True)]:
            with self.subTest(arguments=arguments):
                result = self.run_cli(arguments, "2026-11-08")
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(result.stderr, "")
                self.assertEqual(
                    result.stdout,
                    "Enter follow-up date (YYYY-MM-DD): "
                    + self.run_query("2026-11-08", include_earlier=include_earlier),
                )

    def test_due_through_cli_rejects_invalid_date(self):
        result = self.run_cli(["--due-through"], "2026-02-30")
        self.assertEqual(result.returncode, 1)
        self.assertIn("Invalid date.", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_cli_rejects_unknown_option(self):
        result = self.run_cli(["--due-thru"], "2026-11-08")
        self.assertEqual(result.returncode, 2)
        self.assertIn("unrecognized arguments", result.stderr)
        self.assertNotIn("Total follow-ups", result.stdout)


if __name__ == "__main__":
    unittest.main()

