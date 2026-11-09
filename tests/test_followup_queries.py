import unittest
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
            workbook.save(self.workbook_path)
        finally:
            workbook.close()

    def run_query(self, target_date):
        with (
            patch("builtins.input", return_value=target_date),
            patch.dict(
                "os.environ",
                {"NETWORKING_WORKBOOK": str(self.workbook_path)},
            ),
            patch("sys.stdout", new_callable=StringIO) as stdout,
            patch("sys.stderr", new_callable=StringIO) as stderr,
        ):
            main()

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


if __name__ == "__main__":
    unittest.main()
