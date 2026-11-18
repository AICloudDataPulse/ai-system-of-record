
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


class WorkbookConfigurationError(ValueError):
    """NETWORKING_WORKBOOK is missing or empty."""


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
    """Yield contact dictionaries including a calendar-date followup_due.

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
                "followup_due": due,
            }


def query_followups(workbook_path, target_date, *, include_earlier=False):
    """Return matching contact dictionaries and close the workbook.

    Accept an explicit workbook path, a datetime.date (not datetime.datetime),
    and a Boolean include_earlier. Invalid argument types raise TypeError
    before workbook access; values are not coerced.
    Preserve display cell values and worksheet order; normalize followup_due
    to a calendar date. No prompts, configuration
    lookup, console output, or process exits occur here. File/format errors,
    ContactSchemaError, and unexpected query errors propagate to the caller.
    """
    if not isinstance(target_date, date) or isinstance(target_date, datetime):
        raise TypeError("target_date must be a datetime.date, not a datetime.datetime.")
    if not isinstance(include_earlier, bool):
        raise TypeError("include_earlier must be a bool.")

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
        raise WorkbookConfigurationError(
            "NETWORKING_WORKBOOK is not set or is empty. "
            "Set it to the path of your Excel workbook."
        )

    return Path(value)

def get_followups(date_value, *, include_earlier=False):
    """Query NETWORKING_WORKBOOK using a YYYY-MM-DD string and Boolean flag.

    Validate inputs before reading configuration or accessing the workbook.
    Return query_followups() records unchanged; propagate exceptions without
    prompting, printing, or exiting. The caller cannot supply a workbook path.
    """
    if not isinstance(date_value, str):
        raise TypeError("date_value must be a str.")
    if not isinstance(include_earlier, bool):
        raise TypeError("include_earlier must be a bool.")
    target_date = parse_followup_date(date_value)
    workbook_path = get_workbook_path()
    return query_followups(
        workbook_path, target_date, include_earlier=include_earlier,
    )


def _json_contacts(contacts):
    """Format native records without mutating them."""
    return [
        {
            **{key: None if contact[key] is None else str(contact[key])
               for key in ("name", "company", "next_action")},
            "followup_due": contact["followup_due"].isoformat(),
        }
        for contact in contacts
    ]


def _error_response(code, message):
    return {"ok": False, "error": {"code": code, "message": message}}


def get_followups_response(date_value, *, include_earlier=False):
    """Return a JSON-compatible success or expected-failure dictionary.

    No console I/O or process exits. Unexpected exceptions propagate.
    Error responses contain fixed messages, never raw exception details.
    """
    if not isinstance(date_value, str) or not isinstance(include_earlier, bool):
        return _error_response(
            "INVALID_INPUT", "Provide a YYYY-MM-DD date string and a Boolean include_earlier.",
        )
    try:
        target_date = parse_followup_date(date_value)
    except ValueError:
        return _error_response("INVALID_INPUT", "Provide a valid date in YYYY-MM-DD format.")

    try:
        contacts = get_followups(target_date.isoformat(), include_earlier=include_earlier)
    except WorkbookConfigurationError:
        return _error_response("CONFIGURATION_ERROR", "NETWORKING_WORKBOOK is not set or is empty.")
    except ContactSchemaError:
        return _error_response("CONTACT_SCHEMA_ERROR", "The Contacts worksheet or required headers are invalid.")
    except (InvalidFileException, BadZipFile):
        return _error_response("WORKBOOK_FORMAT_ERROR", "The workbook format or Excel ZIP container is invalid.")
    except FileNotFoundError:
        return _error_response("WORKBOOK_NOT_FOUND", "The configured workbook was not found.")
    except PermissionError:
        return _error_response("WORKBOOK_PERMISSION_DENIED", "The configured workbook cannot be read.")
    except OSError:
        return _error_response("WORKBOOK_ACCESS_ERROR", "The configured workbook could not be accessed.")

    return {
        "ok": True,
        "query_date": target_date.isoformat(),
        "mode": "due_through" if include_earlier else "exact",
        "contacts": _json_contacts(contacts),
        "total": len(contacts),
    }


def handle_followup_request(request):
    """Accept a decoded request dictionary and return the existing response.

    Require date; allow only the optional include_earlier field. The response
    wrapper validates field types and dates before configuration or access.
    Do not parse JSON text, mutate the request, or catch unexpected exceptions.
    """
    if (
        not isinstance(request, dict)
        or "date" not in request
        or any(key not in ("date", "include_earlier") for key in request)
    ):
        return _error_response(
            "INVALID_INPUT",
            "Provide a request dictionary with date and optional include_earlier only.",
        )
    return get_followups_response(
        request["date"], include_earlier=request.get("include_earlier", False),
    )


def _unique_json_fields(pairs):
    fields = dict(pairs)
    if len(fields) != len(pairs):
        raise ValueError("Duplicate JSON field")
    return fields


def _reject_json_constant(value):
    raise ValueError("Invalid JSON constant")


def handle_followup_json(request_text):
    """Decode one JSON request and return the dictionary handler's response."""
    message = "Provide one JSON document with unique field names."
    if not isinstance(request_text, str):
        return _error_response("INVALID_INPUT", message)
    try:
        request = json.loads(
            request_text,
            object_pairs_hook=_unique_json_fields,
            parse_constant=_reject_json_constant,
        )
    except (ValueError, RecursionError):
        return _error_response("INVALID_INPUT", message)
    return handle_followup_request(request)


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
        contacts = _json_contacts(contacts)
        print(json.dumps({"contacts": contacts, "total": len(contacts)}))
        return

    for contact in contacts:
        print(f"\nName: {contact['name']}")
        print(f"Company: {contact['company']}")
        print(f"Next Action: {contact['next_action']}")
        print(f"Follow-up Due: {contact['followup_due'].isoformat()}")

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



