
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook
import os
import sys

def parse_followup_date(value):
    value = value.strip()
    parsed_date = date.fromisoformat(value)

    if value != parsed_date.isoformat():
        raise ValueError("Expected YYYY-MM-DD format")

    return parsed_date

def get_workbook_path():
    value = os.environ.get("NETWORKING_WORKBOOK")

    if value is None or not value.strip():
        raise ValueError(
            "NETWORKING_WORKBOOK is not set or is empty. "
            "Set it to the path of your Excel workbook."
        )

    return Path(value)

def main():
    try:
        target_date = parse_followup_date(
            input("Enter follow-up date (YYYY-MM-DD): ")
        )
    except ValueError:
        print(
            "Invalid date. Enter a valid date in YYYY-MM-DD format.",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        workbook_path = get_workbook_path()
    except ValueError as error:
        print(error, file=sys.stderr)
        sys.exit(1)

    workbook = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True
    )

    try:
        sheet = workbook["Contacts"]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows)

        columns = {
            name: headers.index(name)
            for name in ["Name", "Company", "Next Action", "Follow-up Due"]
        }

        count = 0

        for row in rows:
            due = row[columns["Follow-up Due"]]

            if isinstance(due, datetime):
                due = due.date()

            if isinstance(due, date) and due == target_date:
                count += 1
                print(f"\nName: {row[columns['Name']]}")
                print(f"Company: {row[columns['Company']]}")
                print(f"Next Action: {row[columns['Next Action']]}")

        print(f"\nTotal follow-ups: {count}")

    finally:
        workbook.close()


if __name__ == "__main__":
    main()
