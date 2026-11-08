import unittest
from datetime import date

from pathlib import Path
from unittest.mock import patch
from io import StringIO

from networking.followups import (
    get_workbook_path,
    main,
    parse_followup_date,
)


class ParseFollowupDateTests(unittest.TestCase):
    def test_valid_date(self):
        self.assertEqual(
            parse_followup_date("2026-11-07"),
            date(2026, 11, 7),
        )

    def test_surrounding_whitespace(self):
        self.assertEqual(
            parse_followup_date(" 2026-11-07 "),
            date(2026, 11, 7),
        )

    def test_valid_leap_day(self):
        self.assertEqual(
            parse_followup_date("2024-02-29"),
            date(2024, 2, 29),
        )

    def test_invalid_dates(self):
        invalid_inputs = [
            "",
            "   ",
            "not-a-date",
            "2026-02-30",
            "2026-02-29",
            "11/07/2026",
            "20261107",
            "2026-1-7",
        ]

        for value in invalid_inputs:
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    parse_followup_date(value)

class WorkbookConfigurationTests(unittest.TestCase):
    def test_missing_configuration(self):
        with patch.dict("os.environ", {}, clear=True):
            with self.assertRaisesRegex(ValueError, "NETWORKING_WORKBOOK"):
                get_workbook_path()

    def test_empty_configuration(self):
        for value in ["", "   "]:
            with self.subTest(value=value):
                with patch.dict(
                    "os.environ",
                    {"NETWORKING_WORKBOOK": value},
                    clear=True,
                ):
                    with self.assertRaisesRegex(
                        ValueError, "NETWORKING_WORKBOOK"
                    ):
                        get_workbook_path()

    def test_configured_path(self):
        workbook_path = "/synthetic data/Networking.xlsx"

        with patch.dict(
            "os.environ",
            {"NETWORKING_WORKBOOK": workbook_path},
            clear=True,
        ):
            self.assertEqual(
                get_workbook_path(),
                Path(workbook_path),
            )

class WorkbookFileErrorTests(unittest.TestCase):
    def check_file_error(self, error, expected_message):
        with (
            patch("builtins.input", return_value="2026-11-07"),
            patch.dict(
                "os.environ",
                {"NETWORKING_WORKBOOK": "/synthetic/Networking.xlsx"},
                clear=True,
            ),
            patch(
                "networking.followups.load_workbook",
                side_effect=error,
            ),
            patch("sys.stderr", new_callable=StringIO) as stderr,
        ):
            with self.assertRaises(SystemExit) as raised:
                main()

            self.assertEqual(raised.exception.code, 1)
            self.assertIn(expected_message, stderr.getvalue())

    def test_missing_file(self):
        self.check_file_error(
            FileNotFoundError(),
            "Workbook not found.",
        )

    def test_permission_denied(self):
        self.check_file_error(
            PermissionError(),
            "Workbook cannot be read.",
        )

    def test_other_file_access_error(self):
        self.check_file_error(
            OSError(),
            "Workbook could not be opened.",
        )

if __name__ == "__main__":
    unittest.main()