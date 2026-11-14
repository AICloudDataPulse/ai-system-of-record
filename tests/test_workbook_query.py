import unittest
from datetime import date, datetime, timezone
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from networking.followups import ContactSchemaError, main, query_followups


class WorkbookQueryTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "synthetic contacts.xlsx"
        self.headers = ["Name", "Company", "Next Action", "Follow-up Due"]
        self.rows = [
            ["Boundary", "Example Labs", "Send update", datetime(2026, 11, 7, 15, 30)],
            ["Earlier", None, 123, date(2026, 11, 6)],
            ["Later", "Example", "Call", date(2026, 11, 8)],
            ["Blank", "Example", "Call", None],
            ["Text", "Example", "Call", "2026-11-07"],
        ]
        self.write_workbook([self.headers, *self.rows])

    def write_workbook(self, rows, title="Contacts"):
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = title
            for row in rows:
                sheet.append(row)
            workbook.save(self.path)
        finally:
            workbook.close()

    def open_workbook(self):
        workbook = load_workbook(self.path, read_only=True, data_only=True)
        self.addCleanup(workbook.close)
        return workbook

    def test_returns_records_in_both_modes_and_closes_without_console_or_config(self):
        original = self.path.read_bytes()
        first = {"name": "Boundary", "company": "Example Labs", "next_action": "Send update"}
        for include_earlier, expected in [
            (False, [first]),
            (True, [first, {"name": "Earlier", "company": None, "next_action": 123}]),
        ]:
            with self.subTest(include_earlier=include_earlier):
                workbook = self.open_workbook()
                with (
                    patch("networking.followups.load_workbook", return_value=workbook) as loader,
                    patch.object(workbook, "close", wraps=workbook.close) as close,
                    patch("networking.followups.get_workbook_path") as configuration,
                    patch("builtins.input") as prompt,
                    patch("sys.stdout", new_callable=StringIO) as stdout,
                    patch("sys.stderr", new_callable=StringIO) as stderr,
                ):
                    records = query_followups(
                        self.path, date(2026, 11, 7), include_earlier=include_earlier,
                    )
                    close.assert_called_once_with()
                    loader.assert_called_once_with(self.path, read_only=True, data_only=True)
                    configuration.assert_not_called()
                    prompt.assert_not_called()
                    self.assertEqual(stdout.getvalue(), "")
                    self.assertEqual(stderr.getvalue(), "")
                self.assertIsInstance(records, list)
                self.assertEqual(records, expected)
                self.assertEqual(self.path.read_bytes(), original)

    def test_empty_results_and_header_only_workbook(self):
        for rows in ([self.headers, *self.rows], [self.headers]):
            self.write_workbook(rows)
            original = self.path.read_bytes()
            for include_earlier in (False, True):
                with self.subTest(rows=len(rows), include_earlier=include_earlier):
                    self.assertEqual(query_followups(
                        self.path, date(2026, 11, 5), include_earlier=include_earlier,
                    ), [])
                    self.assertEqual(self.path.read_bytes(), original)

    def test_invalid_target_date_is_rejected_before_workbook_access(self):
        invalid_values = (
            "2026-11-07", None, 20261107, True, [], {},
            datetime(2026, 11, 7),
            datetime(2026, 11, 7, 15, 30, tzinfo=timezone.utc),
        )
        for value in invalid_values:
            for include_earlier in (False, True):
                with (
                    self.subTest(value=value, include_earlier=include_earlier),
                    patch("networking.followups.load_workbook") as loader,
                    patch("sys.stdout", new_callable=StringIO) as stdout,
                    patch("sys.stderr", new_callable=StringIO) as stderr,
                ):
                    with self.assertRaises(TypeError) as raised:
                        query_followups(self.path, value, include_earlier=include_earlier)
                    self.assertEqual(
                        str(raised.exception),
                        "target_date must be a datetime.date, not a datetime.datetime.",
                    )
                    loader.assert_not_called()
                    self.assertEqual(stdout.getvalue(), "")
                    self.assertEqual(stderr.getvalue(), "")

    def test_non_boolean_include_earlier_is_rejected_before_workbook_access(self):
        for value in (None, 0, 1, 0.0, 1.0, "", "false", "true", [], {}):
            with (
                self.subTest(value=value),
                patch("networking.followups.load_workbook") as loader,
                patch("sys.stdout", new_callable=StringIO) as stdout,
                patch("sys.stderr", new_callable=StringIO) as stderr,
            ):
                with self.assertRaises(TypeError) as raised:
                    query_followups(self.path, date(2026, 11, 7), include_earlier=value)
                self.assertEqual(str(raised.exception), "include_earlier must be a bool.")
                loader.assert_not_called()
                self.assertEqual(stdout.getvalue(), "")
                self.assertEqual(stderr.getvalue(), "")

    def test_schema_errors_propagate_and_close_workbook(self):
        for rows, title in [
            ([self.headers], "Other"), ([], "Contacts"),
            ([self.headers[:-1]], "Contacts"), ([self.headers + ["Name"]], "Contacts"),
        ]:
            with self.subTest(rows=rows, title=title):
                self.write_workbook(rows, title)
                original = self.path.read_bytes()
                workbook = self.open_workbook()
                with (
                    patch("networking.followups.load_workbook", return_value=workbook),
                    patch.object(workbook, "close", wraps=workbook.close) as close,
                    patch("sys.stdout", new_callable=StringIO) as stdout,
                    patch("sys.stderr", new_callable=StringIO) as stderr,
                ):
                    with self.assertRaises(ContactSchemaError):
                        query_followups(self.path, date(2026, 11, 7))
                    close.assert_called_once_with()
                    self.assertEqual(stdout.getvalue(), "")
                    self.assertEqual(stderr.getvalue(), "")
                self.assertEqual(self.path.read_bytes(), original)

    def test_iteration_failure_propagates_after_partial_selection_and_closes(self):
        error = RuntimeError("Synthetic iteration failure")
        def failing_records(*args, **kwargs):
            yield {"name": "Synthetic", "company": None, "next_action": None}
            raise error
        workbook = self.open_workbook()
        original = self.path.read_bytes()
        with (
            patch("networking.followups.load_workbook", return_value=workbook),
            patch.object(workbook, "close", wraps=workbook.close) as close,
            patch("networking.followups.iter_followups", side_effect=failing_records),
            patch("sys.stdout", new_callable=StringIO) as stdout,
            patch("sys.stderr", new_callable=StringIO) as stderr,
        ):
            with self.assertRaises(RuntimeError) as raised:
                query_followups(self.path, date(2026, 11, 7))
            self.assertIs(raised.exception, error)
            close.assert_called_once_with()
            self.assertEqual(stdout.getvalue(), "")
            self.assertEqual(stderr.getvalue(), "")
        self.assertEqual(self.path.read_bytes(), original)

    def test_load_errors_propagate_without_console_output(self):
        for error in (FileNotFoundError(), PermissionError(), OSError(),
                      InvalidFileException(), BadZipFile()):
            with (
                self.subTest(error=type(error).__name__),
                patch("networking.followups.load_workbook", side_effect=error),
                patch("sys.stdout", new_callable=StringIO) as stdout,
                patch("sys.stderr", new_callable=StringIO) as stderr,
            ):
                with self.assertRaises(type(error)) as raised:
                    query_followups(self.path, date(2026, 11, 7))
                self.assertIs(raised.exception, error)
                self.assertEqual(stdout.getvalue(), "")
                self.assertEqual(stderr.getvalue(), "")

    def test_cli_does_not_treat_unexpected_value_error_as_schema_error(self):
        error = ValueError("Synthetic unexpected failure")
        for json_output in (False, True):
            with (
                self.subTest(json_output=json_output),
                patch("networking.followups.get_workbook_path", return_value=self.path),
                patch("networking.followups.query_followups", side_effect=error),
                patch("sys.stdout", new_callable=StringIO) as stdout,
                patch("sys.stderr", new_callable=StringIO) as stderr,
            ):
                with self.assertRaises(ValueError) as raised:
                    main(date_value="2026-11-07", json_output=json_output)
                self.assertIs(raised.exception, error)
                self.assertEqual(stdout.getvalue(), "")
                self.assertEqual(stderr.getvalue(), "")

    def test_row_stream_is_closed_before_workbook_on_iteration_failure(self):
        events = []
        error = RuntimeError("Synthetic row failure")
        def rows():
            try:
                yield from self.rows
            finally:
                events.append("rows closed")
        def fail_selection(row_iterator, *args, **kwargs):
            next(row_iterator)
            raise error
        workbook = self.open_workbook()
        with (
            patch("networking.followups.load_workbook", return_value=workbook),
            patch("networking.followups.get_contact_rows_and_columns",
                  return_value=(rows(), {name: index for index, name in enumerate(self.headers)})),
            patch("networking.followups.iter_followups", side_effect=fail_selection),
        ):
            # Save the original method so the spy can call it without recursion.
            original_close = workbook.close
            def close_workbook():
                events.append("workbook closed")
                original_close()
            with patch.object(workbook, "close", side_effect=close_workbook):
                with self.assertRaises(RuntimeError) as raised:
                    query_followups(self.path, date(2026, 11, 7))
                self.assertIs(raised.exception, error)
        self.assertEqual(events, ["rows closed", "workbook closed"])


if __name__ == "__main__":
    unittest.main()
