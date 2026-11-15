import json
import unittest
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import BadZipFile

from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException
from networking.followups import (
    ContactSchemaError, WorkbookConfigurationError, get_followups,
    get_followups_response, main,
)


class FollowupResponseTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "synthetic contacts.xlsx"
        self.make_workbook()

    def make_workbook(self, title="Contacts"):
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = title
            sheet.append(["Name", "Company", "Next Action", "Follow-up Due"])
            sheet.append(['Zoë "Example"', None, "Line 1\nLine 2", datetime(2026, 11, 7, 15, 30)])
            sheet.append(["Earlier", True, 123, date(2026, 11, 6)])
            workbook.save(self.path)
        finally:
            workbook.close()

    def response(self, value="2026-11-07", **kwargs):
        with (
            patch("builtins.input") as prompt,
            patch("sys.stdout", new_callable=StringIO) as stdout,
            patch("sys.stderr", new_callable=StringIO) as stderr,
        ):
            result = get_followups_response(value, **kwargs)
            prompt.assert_not_called()
            self.assertEqual(stdout.getvalue(), "")
            self.assertEqual(stderr.getvalue(), "")
        self.assertEqual(json.loads(json.dumps(result, allow_nan=False)), result)
        return result

    def test_success_and_empty_results_match_cli_and_preserve_workbook(self):
        original = self.path.read_bytes()
        for include_earlier in (False, True):
            for value in (" 2026-11-07 ", "2024-02-29"):
                with (
                    self.subTest(include_earlier=include_earlier, value=value),
                    patch.dict("os.environ", {"NETWORKING_WORKBOOK": str(self.path)}, clear=True),
                ):
                    contacts = [] if value == "2024-02-29" else [{
                        "name": 'Zoë "Example"', "company": None,
                        "next_action": "Line 1\nLine 2", "followup_due": "2026-11-07",
                    }]
                    if contacts and include_earlier:
                        contacts.append({
                            "name": "Earlier", "company": "True",
                            "next_action": "123", "followup_due": "2026-11-06",
                        })
                    result = self.response(value, include_earlier=include_earlier)
                    self.assertEqual(result, {
                        "ok": True, "query_date": value.strip(),
                        "mode": "due_through" if include_earlier else "exact",
                        "contacts": contacts, "total": len(contacts),
                    })
                    with patch("sys.stdout", new_callable=StringIO) as stdout:
                        main(date_value=value, include_earlier=include_earlier, json_output=True)
                    self.assertEqual(json.loads(stdout.getvalue()), {
                        "contacts": contacts, "total": len(contacts),
                    })
                    self.assertEqual(self.path.read_bytes(), original)

    def test_invalid_inputs_fail_before_configuration_or_access(self):
        cases = [(v, False) for v in (
            None, 123, True, date(2026, 11, 7), [], {},
            "", "private-invalid-value", "2026-02-29", "2026-02-30", "20261107",
        )] + [("2026-11-07", v) for v in (None, 0, 1, "false", [], {})]
        for value, flag in cases:
            with (
                self.subTest(value=value, flag=flag),
                patch("networking.followups.get_workbook_path") as configuration,
                patch("networking.followups.get_followups") as adapter,
                patch("networking.followups.load_workbook") as loader,
            ):
                result = self.response(value, include_earlier=flag)
                message = (
                    "Provide a valid date in YYYY-MM-DD format."
                    if isinstance(value, str) and isinstance(flag, bool)
                    else "Provide a YYYY-MM-DD date string and a Boolean include_earlier."
                )
                self.assertEqual(result, {"ok": False, "error": {
                    "code": "INVALID_INPUT", "message": message,
                }})
                configuration.assert_not_called()
                adapter.assert_not_called()
                loader.assert_not_called()

    def test_missing_configuration_is_distinct_from_invalid_input(self):
        for value in (None, "", "   "):
            with (
                self.subTest(value=value),
                patch.dict("os.environ", {} if value is None else {"NETWORKING_WORKBOOK": value}, clear=True),
                patch("networking.followups.load_workbook") as loader,
            ):
                self.assertEqual(self.response(), {"ok": False, "error": {
                    "code": "CONFIGURATION_ERROR",
                    "message": "NETWORKING_WORKBOOK is not set or is empty.",
                }})
                with self.assertRaises(ValueError):
                    get_followups("2026-11-07")
                loader.assert_not_called()

    def test_real_workbook_failures_return_codes_without_data_or_paths(self):
        for kind, code in (
            ("missing", "WORKBOOK_NOT_FOUND"),
            ("format", "WORKBOOK_FORMAT_ERROR"),
            ("schema", "CONTACT_SCHEMA_ERROR"),
        ):
            with self.subTest(kind=kind):
                self.make_workbook(title="Other" if kind == "schema" else "Contacts")
                if kind == "missing":
                    self.path.unlink()
                elif kind == "format":
                    self.path.write_bytes(b"synthetic invalid workbook")
                original = self.path.read_bytes() if self.path.exists() else None
                with patch.dict("os.environ", {"NETWORKING_WORKBOOK": str(self.path)}, clear=True):
                    result = self.response()
                self.assertEqual(set(result), {"ok", "error"})
                self.assertIs(result["ok"], False)
                self.assertEqual(result["error"]["code"], code)
                self.assertNotIn(str(self.path), json.dumps(result))
                if original is not None:
                    self.assertEqual(self.path.read_bytes(), original)

    def test_expected_exception_mapping_uses_fixed_messages(self):
        cases = [
            (WorkbookConfigurationError, "CONFIGURATION_ERROR", "NETWORKING_WORKBOOK is not set or is empty."),
            (ContactSchemaError, "CONTACT_SCHEMA_ERROR", "The Contacts worksheet or required headers are invalid."),
            (InvalidFileException, "WORKBOOK_FORMAT_ERROR", "The workbook format or Excel ZIP container is invalid."),
            (BadZipFile, "WORKBOOK_FORMAT_ERROR", "The workbook format or Excel ZIP container is invalid."),
            (FileNotFoundError, "WORKBOOK_NOT_FOUND", "The configured workbook was not found."),
            (PermissionError, "WORKBOOK_PERMISSION_DENIED", "The configured workbook cannot be read."),
            (OSError, "WORKBOOK_ACCESS_ERROR", "The configured workbook could not be accessed."),
        ]
        for exception, code, message in cases:
            with (
                self.subTest(exception=exception.__name__),
                patch("networking.followups.get_followups",
                      side_effect=exception("/private/path.xlsx secret contact details")) as adapter,
            ):
                self.assertEqual(self.response(" 2026-11-07 ", include_earlier=True), {
                    "ok": False, "error": {"code": code, "message": message},
                })
                adapter.assert_called_once_with("2026-11-07", include_earlier=True)

    def test_unexpected_exceptions_propagate(self):
        for error in (TypeError("bug"), ValueError("bug"), KeyError("bug"), RuntimeError("bug")):
            with (
                self.subTest(error=type(error).__name__),
                patch("networking.followups.get_followups", side_effect=error),
            ):
                with self.assertRaises(type(error)) as raised:
                    get_followups_response("2026-11-07")
                self.assertIs(raised.exception, error)
        with patch("networking.followups.get_followups", return_value=[{}]):
            with self.assertRaises(KeyError):
                get_followups_response("2026-11-07")

    def test_serialization_does_not_mutate_native_records(self):
        records = [{
            "name": "Example", "company": None, "next_action": 123,
            "followup_due": date(2026, 11, 7),
        }]
        original = [dict(record) for record in records]
        with patch("networking.followups.get_followups", return_value=records):
            result = self.response()
        self.assertEqual(records, original)
        self.assertIs(type(records[0]["followup_due"]), date)
        self.assertIsNot(result["contacts"][0], records[0])


if __name__ == "__main__":
    unittest.main()
