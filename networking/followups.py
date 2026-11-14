
import argparse
import json
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile
import os
import sys

REQUIRED_HEADERS = ("Name", "Company", "Next Action", "Follow-up Due")


class ContactSchemaError(ValueError):
    """The Contacts worksheet or its required headers are invalid."""


def get_contact_rows_and_columns(workbook):
    if "Contacts" not in workbook.sheetnames:
        raise ContactSchemaError("Workbook is missing the 'Contacts' worksheet.")

    rows = workbook["Contacts"].iter_rows(values_only=True)
    headers = next(rows, ())
    if not any(value is not None for value in headers):
        rows.close()
        raise ContactSchemaError("The 'Contacts' worksheet is missing its header row.")

    missing = [name for name in REQUIRED_HEADERS if name not in headers]
    duplicates = [name for name in REQUIRED_HEADERS if headers.count(name) > 1]
    errors = []
    if missing:
        errors.append("Missing required headers: " + ", ".join(missing) + ".")
    if duplicates:
        errors.append("Duplicate required headers: " + ", ".join(duplicates) + ".")
    if errors:
        rows.close()
        raise ContactSchemaError("Invalid 'Contacts' header row. " + " ".join(errors))

    return rows, {name: headers.index(name) for name in REQUIRED_HEADERS}


def iter_followups(rows, columns, target_date, *, include_earlier=False):
    """Yield name/company/next_action dictionaries in worksheet order.

    Accept rows and column indexes from get_contact_rows_and_columns() and
    a validated date. The caller owns the workbook and must keep it open
    while consuming this iterator. This function performs no console I/O.
    """
    for row in rows:
        due = row[columns["Follow-up Due"]]
        if isinstance(due, datetime):
            due = due.date()

        if isinstance(due, date) and (
            due <= target_date if include_earlier else due == target_date
        ):
            yield {
                "name": row[columns["Name"]],
                "company": row[columns["Company"]],
                "next_action": row[columns["Next Action"]],
            }


def query_followups(workbook_path, target_date, *, include_earlier=False):
    """Return matching contact dictionaries and close the workbook.

    Accept an explicit workbook path and a validated datetime.date value.
    Preserve cell values and worksheet order. No prompts, configuration
    lookup, console output, or process exits occur here. File/format errors,
    ContactSchemaError, and unexpected query errors propagate to the caller.
    """
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    rows = None
    try:
        rows, columns = get_contact_rows_and_columns(workbook)
        return list(iter_followups(
            rows, columns, target_date, include_earlier=include_earlier,
        ))
    finally:
        try:
            if rows is not None:
                rows.close()
        finally:
            workbook.close()


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

def main(*, include_earlier=False, date_value=None, json_output=False):
    if json_output and date_value is None:
        print("--json requires --date.", file=sys.stderr)
        sys.exit(2)

    try:
        if date_value is None:
            date_value = input("Enter follow-up date (YYYY-MM-DD): ")
        target_date = parse_followup_date(date_value)
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

    try:
        contacts = query_followups(
            workbook_path, target_date, include_earlier=include_earlier,
        )
    except ContactSchemaError as error:
        print(error, file=sys.stderr)
        sys.exit(1)
    except InvalidFileException:
        print(
            "Unsupported workbook format. Use an Excel .xlsx workbook.",
            file=sys.stderr,
        )
        sys.exit(1)
    except BadZipFile:
        print(
            "Workbook is not a valid Excel ZIP file. "
            "Check the file or restore a valid backup.",
            file=sys.stderr,
        )
        sys.exit(1)
    except FileNotFoundError:
        print(
            "Workbook not found. Check NETWORKING_WORKBOOK.",
            file=sys.stderr,
        )
        sys.exit(1)
    except PermissionError:
        print(
            "Workbook cannot be read. Check file access permissions.",
            file=sys.stderr,
        )
        sys.exit(1)
    except OSError:
        print(
            "Workbook could not be opened. Check that the path "
            "points to an accessible Excel file.",
            file=sys.stderr,
        )
        sys.exit(1)

    if json_output:
        # Display fields have a stable string-or-null JSON contract.
        contacts = [
            {key: None if value is None else str(value)
             for key, value in contact.items()}
            for contact in contacts
        ]
        print(json.dumps({"contacts": contacts, "total": len(contacts)}))
        return

    for contact in contacts:
        print(f"\nName: {contact['name']}")
        print(f"Company: {contact['company']}")
        print(f"Next Action: {contact['next_action']}")

    print(f"\nTotal follow-ups: {len(contacts)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Query contact follow-up dates.")
    parser.add_argument(
        "--date",
        metavar="YYYY-MM-DD",
        help="supply the follow-up date without an interactive prompt",
    )
    parser.add_argument(
        "--due-through",
        action="store_true",
        help="include follow-ups due on or before the entered date",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="print JSON instead of text (requires --date)",
    )
    args = parser.parse_args()
    if args.json and args.date is None:
        parser.error("--json requires --date")
    main(
        include_earlier=args.due_through, date_value=args.date,
        json_output=args.json,
    )



