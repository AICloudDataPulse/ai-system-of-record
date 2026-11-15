import unittest
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import BadZipFile

from openpyxl import Workbook
from networking.followups import ContactSchemaError, get_followups


class FollowupAdapterTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "synthetic contacts.xlsx"
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = "Contacts"
            sheet.append(["Name", "Company", "Next Action", "Follow-up Due"])
            sheet.append(["Boundary", "Example Labs", "Call", datetime(2026, 11, 7, 15, 30)])
            sheet.append(["Earlier", None, 123, date(2026, 11, 6)])
            sheet.append(["Later", "Example Labs", "Call", date(2026, 11, 8)])
            workbook.save(self.path)
        finally:
            workbook.close()

    def test_configured_queries_return_native_records_without_io_and_preserve_bytes(self):
        original = self.path.read_bytes()
        boundary = {
            "name": "Boundary", "company": "Example Labs", "next_action": "Call",
            "followup_due": date(2026, 11, 7),
        }
        earlier = {
            "name": "Earlier", "company": None, "next_action": 123,
            "followup_due": date(2026, 11, 6),
        }
        for include_earlier in (False, True):
            for value in ("2026-11-07", " 2026-11-07 ", "2024-02-29"):
                with (
                    self.subTest(include_earlier=include_earlier, value=value),
                    patch.dict("os.environ", {"NETWORKING_WORKBOOK": str(self.path)}, clear=True),
                    patch("builtins.input") as prompt,
                    patch("sys.stdout", new_callable=StringIO) as stdout,
                    patch("sys.stderr", new_callable=StringIO) as stderr,
                ):
                    records = get_followups(value, include_earlier=include_earlier)
                    expected = [] if value == "2024-02-29" else (
                        [boundary, earlier] if include_earlier else [boundary]
                    )
                    self.assertEqual(records, expected)
                    for record in records:
                        self.assertIs(type(record["followup_due"]), date)
                    prompt.assert_not_called()
                    self.assertEqual(stdout.getvalue(), "")
                    self.assertEqual(stderr.getvalue(), "")
                    self.assertEqual(self.path.read_bytes(), original)

    def assert_invalid_input(self, value, include_earlier, exception):
        with (
            patch("networking.followups.get_workbook_path") as configuration,
            patch("networking.followups.query_followups") as query,
            patch("networking.followups.load_workbook") as loader,
            patch("builtins.input") as prompt,
            patch("sys.stdout", new_callable=StringIO) as stdout,
            patch("sys.stderr", new_callable=StringIO) as stderr,
        ):
            with self.assertRaises(exception):
                get_followups(value, include_earlier=include_earlier)
            configuration.assert_not_called()
            query.assert_not_called()
            loader.assert_not_called()
            prompt.assert_not_called()
            self.assertEqual(stdout.getvalue(), "")
            self.assertEqual(stderr.getvalue(), "")

    def test_invalid_date_types_fail_before_configuration(self):
        for value in (None, 20261107, True, date(2026, 11, 7),
                      datetime(2026, 11, 7), [], {}):
            for include_earlier in (False, True):
                with self.subTest(value=value, include_earlier=include_earlier):
                    self.assert_invalid_input(value, include_earlier, TypeError)

    def test_invalid_date_strings_fail_before_configuration(self):
        for value in ("", "   ", "invalid", "2026-02-30", "2026-02-29",
                      "11/07/2026", "20261107", "2026-1-7"):
            for include_earlier in (False, True):
                with self.subTest(value=value, include_earlier=include_earlier):
                    self.assert_invalid_input(value, include_earlier, ValueError)

    def test_non_boolean_flags_fail_before_configuration(self):
        for value in (None, 0, 1, 0.0, "", "false", "true", [], {}):
            with self.subTest(value=value):
                self.assert_invalid_input("2026-11-07", value, TypeError)

    def test_missing_or_empty_configuration_raises_without_query_or_output(self):
        for value in (None, "", "   "):
            env = {} if value is None else {"NETWORKING_WORKBOOK": value}
            with (
                self.subTest(value=value),
                patch.dict("os.environ", env, clear=True),
                patch("networking.followups.query_followups") as query,
                patch("networking.followups.load_workbook") as loader,
                patch("sys.stdout", new_callable=StringIO) as stdout,
                patch("sys.stderr", new_callable=StringIO) as stderr,
            ):
                with self.assertRaisesRegex(ValueError, "NETWORKING_WORKBOOK"):
                    get_followups("2026-11-07")
                query.assert_not_called()
                loader.assert_not_called()
                self.assertEqual(stdout.getvalue(), "")
                self.assertEqual(stderr.getvalue(), "")

    def test_query_errors_propagate_without_output(self):
        for error in (FileNotFoundError(), PermissionError(), BadZipFile(),
                      ContactSchemaError("Synthetic schema error"), RuntimeError()):
            with (
                self.subTest(error=type(error).__name__),
                patch.dict("os.environ", {"NETWORKING_WORKBOOK": str(self.path)}, clear=True),
                patch("networking.followups.query_followups", side_effect=error) as query,
                patch("sys.stdout", new_callable=StringIO) as stdout,
                patch("sys.stderr", new_callable=StringIO) as stderr,
            ):
                with self.assertRaises(type(error)) as raised:
                    get_followups("2026-11-07")
                self.assertIs(raised.exception, error)
                query.assert_called_once_with(self.path, date(2026, 11, 7), include_earlier=False)
                self.assertEqual(stdout.getvalue(), "")
                self.assertEqual(stderr.getvalue(), "")


if __name__ == "__main__":
    unittest.main()
