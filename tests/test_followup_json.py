import json
import os
import subprocess
import sys
import unittest
from datetime import date
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook
from networking.followups import main


class FollowupJsonTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "synthetic contacts.xlsx"
        self.script = Path(__file__).resolve().parents[1] / "networking" / "followups.py"
        self.make_workbook()

    def make_workbook(self, title="Contacts"):
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = title
            sheet.append(["Name", "Company", "Next Action", "Follow-up Due"])
            sheet.append(['Zoë "Example"', "Example\\Labs", "Line 1\nLine 2", date(2026, 11, 7)])
            sheet.append(["Earlier", None, None, date(2026, 11, 6)])
            sheet.append(["Later", "Example", "Call", date(2026, 11, 8)])
            workbook.save(self.path)
        finally:
            workbook.close()

    def run_cli(self, arguments, *, configured=True):
        env = dict(os.environ)
        env.pop("NETWORKING_WORKBOOK", None)
        if configured:
            env["NETWORKING_WORKBOOK"] = str(self.path)
        return subprocess.run(
            [sys.executable, str(self.script), *arguments],
            input="", text=True, capture_output=True, env=env, timeout=10,
        )

    def test_both_modes_return_only_json_and_preserve_workbook(self):
        original = self.path.read_bytes()
        first = {"name": 'Zoë "Example"', "company": "Example\\Labs", "next_action": "Line 1\nLine 2"}
        for flags, expected in [
            ([], [first]),
            (["--due-through"], [first, {"name": "Earlier", "company": None, "next_action": None}]),
        ]:
            with self.subTest(flags=flags):
                result = self.run_cli(["--date", "2026-11-07", "--json", *flags])
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(result.stderr, "")
                self.assertEqual(json.loads(result.stdout), {"contacts": expected, "total": len(expected)})
                self.assertEqual(self.path.read_bytes(), original)

    def test_empty_results(self):
        for flags in ([], ["--due-through"]):
            with self.subTest(flags=flags):
                result = self.run_cli(["--json", "--date", "2026-11-05", *flags])
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(result.stderr, "")
                self.assertEqual(json.loads(result.stdout), {"contacts": [], "total": 0})

    def test_json_requires_explicit_date(self):
        result = self.run_cli(["--json"], configured=False)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result.stdout, "")
        self.assertIn("--json requires --date", result.stderr)
        with (
            patch("builtins.input") as prompt,
            patch("networking.followups.get_workbook_path") as configuration,
            patch("networking.followups.load_workbook") as loader,
            patch("sys.stderr", new_callable=StringIO),
        ):
            with self.assertRaises(SystemExit) as raised:
                main(json_output=True)
            self.assertEqual(raised.exception.code, 2)
            prompt.assert_not_called()
            configuration.assert_not_called()
            loader.assert_not_called()

    def test_invalid_dates_produce_no_json(self):
        for value in ("", "2026-02-30"):
            with self.subTest(value=value):
                result = self.run_cli(["--json", "--date", value], configured=False)
                self.assertEqual(result.returncode, 1)
                self.assertEqual(result.stdout, "")
                self.assertIn("Invalid date.", result.stderr)
                self.assertNotIn("Traceback", result.stderr)

    def test_configuration_and_workbook_errors_produce_no_json(self):
        for error in ("configuration", "missing", "format", "schema"):
            for flags in ([], ["--due-through"]):
                with self.subTest(error=error, flags=flags):
                    self.make_workbook()
                    if error == "missing":
                        self.path.unlink()
                    elif error == "format":
                        self.path.write_bytes(b"not a workbook")
                    elif error == "schema":
                        self.make_workbook(title="Other")
                    original = self.path.read_bytes() if self.path.exists() else None
                    result = self.run_cli(
                        ["--json", "--date", "2026-11-07", *flags],
                        configured=error != "configuration",
                    )
                    self.assertEqual(result.returncode, 1)
                    self.assertEqual(result.stdout, "")
                    self.assertTrue(result.stderr)
                    self.assertNotIn("Traceback", result.stderr)
                    self.assertNotIn(str(self.path), result.stderr)
                    if original is not None:
                        self.assertEqual(self.path.read_bytes(), original)

    def test_display_values_are_strings_or_null(self):
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = "Contacts"
            sheet.append(["Name", "Company", "Next Action", "Follow-up Due"])
            sheet.append([123, True, date(2026, 11, 1), date(2026, 11, 7)])
            workbook.save(self.path)
        finally:
            workbook.close()
        result = self.run_cli(["--json", "--date", "2026-11-07"])
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout), {
            "contacts": [{"name": "123", "company": "True", "next_action": "2026-11-01 00:00:00"}],
            "total": 1,
        })


if __name__ == "__main__":
    unittest.main()
