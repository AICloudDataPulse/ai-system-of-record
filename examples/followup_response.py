"""Demonstrate follow-up requests using a temporary fictional workbook.

Run from the repository root: python -m examples.followup_response
"""

import json
import os
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from networking.followups import handle_followup_request


def main():
    with TemporaryDirectory(prefix="followup-example-") as directory:
        path = Path(directory) / "fictional-contacts.xlsx"
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = "Contacts"
            sheet.append(["Name", "Company", "Next Action", "Follow-up Due"])
            sheet.append([
                "Alex Example", "Example Labs", "Send an update", date(2026, 11, 14),
            ])
            sheet.append([
                "Sam Sample", "Sample Works", "Arrange a call", date(2026, 11, 13),
            ])
            workbook.save(path)
        finally:
            workbook.close()

        previous_path = os.environ.get("NETWORKING_WORKBOOK")
        try:
            os.environ["NETWORKING_WORKBOOK"] = str(path)
            requests = {
                "exact_date": {"date": "2026-11-14"},
                "due_through": {"date": "2026-11-14", "include_earlier": True},
                "empty_result": {"date": "2026-11-12"},
                "invalid_date": {"date": "2026-02-30"},
            }
            results = {
                name: handle_followup_request(request) for name, request in requests.items()
            }
        finally:
            if previous_path is None:
                os.environ.pop("NETWORKING_WORKBOOK", None)
            else:
                os.environ["NETWORKING_WORKBOOK"] = previous_path

    # Print after configuration has been restored and temporary files removed.
    print(json.dumps(results, indent=2, ensure_ascii=False, allow_nan=False))


if __name__ == "__main__":
    main()
