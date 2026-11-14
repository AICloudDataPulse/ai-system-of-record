import unittest
from datetime import date, datetime
from io import StringIO
from unittest.mock import patch

from networking.followups import iter_followups


class FollowupSelectionTests(unittest.TestCase):
    columns = {"Name": 0, "Company": 1, "Next Action": 2, "Follow-up Due": 3}

    def setUp(self):
        self.rows = [
            ("Boundary", "Example Labs", "Send update", datetime(2026, 11, 7, 15, 30)),
            ("Earlier", "Sample Systems", "Arrange meeting", date(2026, 11, 6)),
            ("Later", "Test Company", "Call", date(2026, 11, 8)),
            ("Blank", "Test Company", "Call", None),
            ("Text", "Test Company", "Call", "2026-11-07"),
            ("Number", "Test Company", "Call", 123),
        ]

    def select(self, target=date(2026, 11, 7), *, include_earlier=False):
        return list(iter_followups(
            iter(self.rows), self.columns, target, include_earlier=include_earlier,
        ))

    def test_exact_date_returns_projected_records_without_console_io(self):
        with (
            patch("builtins.input", side_effect=AssertionError("Unexpected prompt")),
            patch("sys.stdout", new_callable=StringIO) as stdout,
            patch("sys.stderr", new_callable=StringIO) as stderr,
        ):
            records = self.select()
        self.assertEqual(records, [{
            "name": "Boundary", "company": "Example Labs", "next_action": "Send update", "followup_due": date(2026, 11, 7),
        }])
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")

    def test_due_through_includes_boundary_and_earlier_in_worksheet_order(self):
        self.assertEqual(self.select(include_earlier=True), [
            {"name": "Boundary", "company": "Example Labs", "next_action": "Send update", "followup_due": date(2026, 11, 7)},
            {"name": "Earlier", "company": "Sample Systems", "next_action": "Arrange meeting", "followup_due": date(2026, 11, 6)},
        ])

    def test_no_matches_and_empty_rows(self):
        for include_earlier in (False, True):
            with self.subTest(include_earlier=include_earlier):
                self.assertEqual(self.select(date(2026, 11, 5), include_earlier=include_earlier), [])
                self.assertEqual(list(iter_followups(
                    iter(()), self.columns, date(2026, 11, 7),
                    include_earlier=include_earlier,
                )), [])

    def test_due_value_is_calendar_date_for_date_and_datetime_inputs(self):
        for due in (date(2024, 2, 29), datetime(2024, 2, 29, 23, 59)):
            for include_earlier in (False, True):
                with self.subTest(due=due, include_earlier=include_earlier):
                    rows = [("Alex", "Example Labs", "Call", due)]
                    records = list(iter_followups(
                        iter(rows), self.columns, date(2024, 2, 29),
                        include_earlier=include_earlier,
                    ))
                    self.assertEqual(len(records), 1)
                    self.assertIs(type(records[0]["followup_due"]), date)
                    self.assertEqual(records[0]["followup_due"], date(2024, 2, 29))

    def test_reordered_columns_only_return_selected_fields(self):
        rows = [("private extra", date(2026, 11, 7), None, "Example Labs", "Alex")]
        columns = {"Follow-up Due": 1, "Next Action": 2, "Company": 3, "Name": 4}
        self.assertEqual(list(iter_followups(iter(rows), columns, date(2026, 11, 7))), [
            {"name": "Alex", "company": "Example Labs", "next_action": None, "followup_due": date(2026, 11, 7)},
        ])


if __name__ == "__main__":
    unittest.main()
