# AI System of Record

A small Python project exploring secure, controlled integration between AI
assistants and external systems of record. The current capability is a
deterministic, read-only query of a local networking Excel workbook.

Enter a date to list contacts whose follow-up is due on that date. AI assistants,
agents, MCP servers, and workbook updates are future work.

## Setup (macOS / zsh or bash)

Use Python 3.12, the project's validated Python version, and Git.
From your existing repository directory, run:

```bash
python3 --version
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Create the virtual environment once. If `.venv` already exists, skip its
creation and activate it. In each new terminal session, activate it again
before running the script or tests. Confirm `python3` is Python 3.12 before
creating the environment.

Dependencies are pinned in [requirements.txt](requirements.txt).
All commands below assume the repository root is the current directory
and the virtual environment is active.

## Workbook and configuration

Keep your workbook outside the repository. Use a local Excel `.xlsx` file
with a worksheet named exactly `Contacts`. The first row must contain these
headers, with matching spelling, capitalization, and spaces:

| Header | Use |
| --- | --- |
| `Name` | Printed contact name |
| `Company` | Printed company |
| `Next Action` | Printed next action |
| `Follow-up Due` | Excel date or date-time used for matching |

Each required header must appear exactly once. Columns may be in any order;
additional columns are ignored. An empty worksheet or a blank first row is
rejected. A worksheet containing only valid headers returns zero follow-ups.
Other worksheets are not queried. Use actual Excel date values rather than text that looks like
a date. Blank and non-date due values are skipped. For date-time values, only
the calendar date is compared.

Set the workbook location for the current terminal session:

```bash
export NETWORKING_WORKBOOK="/absolute/path/to/Networking.xlsx"
```

Replace the placeholder with your own absolute path. Keep the quotes when the
path contains spaces. The script reads the environment variable directly;
it does not automatically load a `.env` file. Do not commit your local path
or workbook.

## Run a follow-up query

```bash
python networking/followups.py
```

At the prompt, enter a calendar date in `YYYY-MM-DD` format. Surrounding
whitespace is allowed. Example output for a synthetic workbook containing
one matching contact:

```text
Enter follow-up date (YYYY-MM-DD): 2026-11-07

Name: Alex Example
Company: Example Labs
Next Action: Send an update
Follow-up Due: 2026-11-07

Total follow-ups: 1
```

To include follow-ups due on or before the entered date, run:

```bash
python networking/followups.py --due-through
```

For example, entering `2026-11-10` includes all actual Excel due dates up to
and including November 10, 2026. Running without the option still matches
only the entered date. Both modes skip blank and non-date cells and compare
date-time values by calendar date. The cutoff is always entered explicitly;
the script does not use today's date.

To supply the date without an interactive prompt, use `--date` with either mode:

```bash
python networking/followups.py --date 2026-11-13
python networking/followups.py --date 2026-11-13 --due-through
```

Omitting `--date` preserves the interactive prompt. Both input methods use the
same date validation, including surrounding whitespace and leap-day checks.
An invalid date exits with status 1 before workbook access; `--date` without
a value is a command-line usage error and exits with status 2.

For machine-readable output, add `--json` with an explicit `--date`:

```bash
python networking/followups.py --date 2026-11-14 --due-through --json
```

Successful output is one JSON object on standard output, with no prompt or
text summary. For example:

```json
{"contacts": [{"name": "Alex Example", "company": "Example Labs", "next_action": "Send update", "followup_due": "2026-11-14"}], "total": 1}
```

The keys are `contacts` and `total`; each contact contains `name`,
`company`, `next_action`, and `followup_due`, in worksheet order.
The due date is always a `YYYY-MM-DD` string; Excel date-times use their
calendar date without a time component. For the three display fields, blank
cells become JSON `null` and other values become strings using Python's
`str()`, matching their text-mode representation. Unicode and special characters are
JSON-escaped and round-trip when parsed. Zero matches returns
`{"contacts": [], "total": 0}`.

`--json` without `--date` exits with usage status 2 before prompting or opening
the workbook. Handled date, configuration, workbook, and schema errors retain
status 1 and plain-text messages on standard error; standard output is empty.
Consumers should check the exit status before parsing output. JSON mode
collects matching contacts in memory before emitting the response.
Omitting `--json` retains the existing text output and interactive behavior.

This mode filters due dates only; it does not check whether an action has
already been completed. Use `python networking/followups.py --help` for help.
Unknown command-line options exit with status 2 before prompting for a date.

Each text record includes `Follow-up Due: YYYY-MM-DD` after Next Action.
This is an additive output-contract change: Python/JSON records now have a
fourth field, and text output has an extra line. Consumers that validate
exact keys or line counts must account for it.

Results follow worksheet order. No matches produces `Total follow-ups: 0`
and is a successful query.

Invalid date input, missing/empty configuration, invalid worksheet/header
structure, unsupported formats, invalid ZIP containers, and handled file-access
failures print an error to standard error and exit with status 1.
Successful queries exit with status 0. To inspect the status in your shell,
run `echo $?` immediately after the script.

| Error | What to check |
| --- | --- |
| Invalid date | Enter a real date in `YYYY-MM-DD` format. |
| `NETWORKING_WORKBOOK` is not set or is empty | Export the variable in the same terminal session. |
| Missing `Contacts` worksheet | Use a worksheet named exactly `Contacts`. |
| Missing header row | Put the required headers in the first row. |
| Missing or duplicate required headers | Include each required header exactly once, with exact spelling and spacing. |
| Unsupported workbook format | Select an Excel `.xlsx` workbook; renaming another format does not convert it. |
| Workbook is not a valid Excel ZIP file | Check the selected file or restore a valid backup. |
| Workbook not found | Verify that the configured file exists. |
| Workbook cannot be read | Check file access permissions. |
| Workbook could not be opened | Check that the path identifies an accessible Excel file. |

When finished, leave the virtual environment with:

```bash
deactivate
```

## Try the follow-up request handler with fictional data

With the virtual environment active, run from the repository root:

```bash
python -m examples.followup_response
```

No workbook configuration is needed. The example creates a temporary workbook
with two fictional contacts and passes four request dictionaries to
`handle_followup_request()`. A request containing only `date` selects exact
matches. Adding `"include_earlier": True` also selects earlier dates.
The example prints one JSON object with the four named responses.

| Scenario | Expected result |
| --- | --- |
| `exact_date` | Success; Alex Example, total 1 |
| `due_through` | Success; Alex Example and Sam Sample, total 2 |
| `empty_result` | Success; empty contacts, total 0 |
| `invalid_date` | Failure; `INVALID_INPUT` |

Dates are fixed so results are repeatable. The scenario names belong to this
demonstration; each nested response uses the existing wrapper contract.
The invalid-date response is expected, so the completed example exits with
status 0. Unexpected exceptions propagate.

The example temporarily sets `NETWORKING_WORKBOOK` only in its own process,
restores its previous value (including an empty or absent setting), and removes
the temporary workbook on normal completion or a Python exception. It never
opens your configured workbook. Only the synthetic fixture is written; queries
remain read-only. It uses existing dependencies and makes no network calls.

## Try live tool selection with OpenAI

From the repository root, activate the project environment and set
`OPENAI_API_KEY` privately in your shell. Replace `YOUR_MODEL` with an
OpenAI model available to your API account that supports function calling.

```bash
python -m examples.live_followups --model YOUR_MODEL
```

The example makes up to two billable Responses API requests, one for each
fictional question. The model chooses whether to call the published tool.
Python runs the selected call against a temporary workbook and prints the
results locally. The workbook, its path, and query results stay local.

The **on November 14** question should return Alex Example. The **through
November 14** question should also return Sam Sample. Both checks must pass
for exit status 0. A missing call, multiple calls, or an unexpected query
result fails the run. CI exercises this path with simulated responses.
Record live validation separately, including the model used.

## Run the tests

With the virtual environment active, run from the repository root:

```bash
python -m unittest discover -s tests -v
```

The current suite contains 75 tests and should finish with `OK`.
It does not require your personal workbook or `NETWORKING_WORKBOOK` setting.

- [tests/test_followups.py](tests/test_followups.py) covers date validation,
  configuration, explicit-date validation before workbook access, and simulated
  file-access failures.
- [tests/test_followup_queries.py](tests/test_followup_queries.py) creates
  temporary synthetic workbooks and checks matching output, exclusions,
  zero matches, byte-for-byte workbook preservation, and the command-line
  behavior of both query modes, including explicit dates, missing argument
  values, and unchanged interactive input.
- [tests/test_followup_selection.py](tests/test_followup_selection.py) checks returned
  records, both date modes, worksheet order, skipped non-dates, empty results,
  reordered columns, calendar-date normalization (including leap day), and
  selection without console I/O.
- [tests/test_followup_json.py](tests/test_followup_json.py) checks JSON in both
  modes, special characters, nulls, display-value conversion, empty results,
  required explicit dates, error output, and workbook preservation.
- [tests/test_workbook_query.py](tests/test_workbook_query.py) checks the reusable
  query, native cell values, workbook closure on success and failure, error
  propagation, empty results, no console/configuration interaction, unchanged bytes,
  and rejection of invalid query argument types before workbook access.
- [tests/test_followup_adapter.py](tests/test_followup_adapter.py) checks the
  configured adapter, both modes, invalid inputs before configuration/access,
  missing configuration, exception propagation, and workbook preservation.
- [tests/test_followup_request.py](tests/test_followup_request.py) checks request
  shape, unknown fields, defaults, validation before access, unchanged requests,
  synthetic workbook results, safe errors, and unexpected-error propagation.
- [tests/test_followup_response.py](tests/test_followup_response.py) checks response
  shapes, safe error mapping, JSON compatibility, native-record preservation,
  CLI equivalence, and propagation of unexpected errors.
- [tests/test_contact_schema.py](tests/test_contact_schema.py) checks worksheet
  and header errors through both CLI modes, exact header matching, reordered
  and extra columns, header-only sheets, workbook preservation, and closure
  after schema errors.
- [tests/test_workbook_formats.py](tests/test_workbook_formats.py) checks unsupported
  extensions and invalid ZIP containers through both CLI modes, including empty,
  text, and truncated files, safe error output, and unchanged file bytes.

Temporary test workbooks are cleaned up automatically.

## Automated validation on GitHub

[Python validation](.github/workflows/python-validation.yml) runs on pull
requests targeting `main` and pushes to `main`. One Ubuntu 24.04 job uses
Python 3.12, installs `requirements.txt`, runs the full unittest suite, and
runs the synthetic response example. Resource warnings are shown in the logs;
they are not configured as a separate failing gate. The example is a smoke
check; response assertions are covered by the tests.

Review the **Tests and synthetic example** check on the pull request before
merging. A failed dependency installation, test, or example command fails the
job. The workflow has a ten-minute timeout. It uses read-only repository
permissions, does not persist checkout credentials, and pins GitHub actions
to full commit SHAs with release comments. It requires no custom secrets or
personal workbook configuration; all workbook data is synthetic.

This adds CI validation only. It does not deploy the application or change
branch protection. Requiring a passing check before GitHub permits merging
would be a separate repository-settings decision.

## Current architecture and data protection

`main()` in [networking/followups.py](networking/followups.py) validates the
entered date, reads configuration, calls `query_followups()`, and formats the
returned records as text or JSON. It retains CLI error reporting and exit codes.

`query_followups(workbook_path, target_date, *, include_earlier=False)` accepts
an explicit path, a `datetime.date` (excluding `datetime.datetime`), and a
Boolean `include_earlier` (default `False`). Invalid date/flag types raise
`TypeError` before opening the workbook. Strings are not parsed, datetimes
are not truncated, and truthy/falsy values such as 1, 0, or "false" are not
coerced to Booleans. The CLI continues to parse `--date` strings separately.
The function opens the workbook with
`read_only=True` and `data_only=True`, validates the worksheet/headers, and
collects the records selected by `iter_followups()`. Cleanup closes the row iterator and workbook before returning or propagating
an error after a successful open. Schema validation also closes its row
iterator before raising a known header error.

Each returned dictionary contains `name`, `company`, `next_action`, and
`followup_due`, preserving worksheet order. The three display fields retain
native cell values (including None); `followup_due` is a native
`datetime.date`, with date-time cells normalized to their calendar date.
The CLI formats due dates using `isoformat()` and converts display fields
to strings/null for JSON. The reusable query performs no
prompting, environment lookup, printing, or process exits. Invalid worksheet
structure raises `ContactSchemaError`, a `ValueError` subclass; file/format
and unexpected errors propagate to the caller.

Example from Python, run from the repository root:

```python
from datetime import date
from pathlib import Path
from networking.followups import query_followups

contacts = query_followups(
    Path("/absolute/path/to/Networking.xlsx"),
    date(2026, 11, 14),
    include_earlier=True,
)
```

For callers using the configured workbook, use the adapter in the same module:

```python
from networking.followups import get_followups

contacts = get_followups("2026-11-14", include_earlier=True)
```

`get_followups(date_value, *, include_earlier=False)` requires a date string
in `YYYY-MM-DD` format and a Boolean flag. Surrounding whitespace and valid
leap days follow the existing date parser's rules. Invalid types raise
`TypeError`; invalid date strings raise `ValueError`. Both are rejected
before configuration lookup or workbook access.

The adapter reads `NETWORKING_WORKBOOK` at call time and calls
`query_followups()`. It accepts no workbook-path argument. Missing/empty
configuration raises `ValueError`; query exceptions propagate unchanged.
The returned list and native field types match `query_followups()`, including
the calendar-date `followup_due`. It performs no prompting, printing, process
exits, or JSON conversion. Configure the environment before calling it.
The CLI continues using its existing input and error-handling flow.

For a JSON-compatible response dictionary, call the separate wrapper:

```python
from networking.followups import get_followups_response

response = get_followups_response(" 2026-11-14 ", include_earlier=True)
```

A successful response has exactly these top-level fields:

```json
{"ok": true, "query_date": "2026-11-14", "mode": "due_through", "contacts": [], "total": 0}
```

`query_date` is normalized to YYYY-MM-DD; `mode` is `exact` or
`due_through`. Contacts use the same four fields and conversion rules as
CLI JSON. No matches is a success with an empty list and total 0.

An expected failure has only `ok` and `error`:

```json
{"ok": false, "error": {"code": "WORKBOOK_NOT_FOUND", "message": "The configured workbook was not found."}}
```

| Code | Condition | Fixed message |
| --- | --- | --- |
| `INVALID_INPUT` | Wrong date/flag type | Provide a YYYY-MM-DD date string and a Boolean include_earlier. |
| `INVALID_INPUT` | Invalid date string | Provide a valid date in YYYY-MM-DD format. |
| `CONFIGURATION_ERROR` | Missing/empty workbook configuration | NETWORKING_WORKBOOK is not set or is empty. |
| `CONTACT_SCHEMA_ERROR` | Missing worksheet or invalid required headers | The Contacts worksheet or required headers are invalid. |
| `WORKBOOK_FORMAT_ERROR` | Unsupported format or invalid ZIP container | The workbook format or Excel ZIP container is invalid. |
| `WORKBOOK_NOT_FOUND` | FileNotFoundError | The configured workbook was not found. |
| `WORKBOOK_PERMISSION_DENIED` | PermissionError | The configured workbook cannot be read. |
| `WORKBOOK_ACCESS_ERROR` | Other OSError | The configured workbook could not be accessed. |

The wrapper validates inputs before configuration/access, calls
`get_followups()`, and returns a dictionary without printing or exiting.
Expected errors contain fixed messages without raw input, file paths,
exception text, or partial records. Unexpected exceptions propagate;
formatting errors are outside the expected-error handler.

Missing configuration now raises `WorkbookConfigurationError`, a
`ValueError` subclass, allowing the wrapper to distinguish it from unrelated
ValueErrors. Existing callers catching ValueError and the CLI keep their
behavior. The native adapter still raises exceptions and returns native
records. The CLI JSON format remains `contacts`/`total`; it does not adopt
this wrapper's envelope.

For a decoded request dictionary, use the narrow request handler:

```python
from networking.followups import handle_followup_request

response = handle_followup_request({
    "date": "2026-11-14",
    "include_earlier": True,
})
```

`handle_followup_request(request)` requires a dictionary with a `date` field.
Only `include_earlier` is optional; it defaults to `False`. Unknown fields
(including workbook paths), missing dates, and non-dictionary requests return
`INVALID_INPUT`. JSON text is not parsed by this function.

Field validation uses the existing response wrapper: date must be a valid
YYYY-MM-DD string (surrounding whitespace allowed), and include_earlier must
be a Boolean. Numbers and strings are not coerced. Invalid requests fail before
configuration lookup or workbook access. Shape errors use a fixed message:
"Provide a request dictionary with date and optional include_earlier only."

Valid requests return the existing success/error response unchanged.
The request dictionary is not modified. The handler performs no console I/O,
accepts no workbook path, and does not catch unexpected exceptions.
Workbook selection remains controlled by `NETWORKING_WORKBOOK`.

For a request supplied as JSON text, call the Python adapter:

```python
from networking.followups import handle_followup_json

response = handle_followup_json('{"date":"2026-11-14","include_earlier":true}')
```

The adapter accepts one JSON document as a Python string and returns the same
response dictionary as `handle_followup_request()`. Malformed JSON, duplicate
field names, and non-string inputs return a fixed `INVALID_INPUT` error.

Both CLI formats now collect matching records before printing, using memory
proportional to the matching results. Query failures produce no partial
contact output. The lower-level `iter_followups()` remains available; callers
using it directly must keep their workbook open while consuming the iterator.

### Tool definition for assistant integration

[get_followups.tool.json](networking/get_followups.tool.json) describes the
read-only lookup and its arguments using JSON Schema. A future assistant
integration can use this provider-neutral definition to describe the tool to
a model. For example, “Who is due through November 17, 2026?” maps to
`{"date":"2026-11-17","include_earlier":true}`. For “Who is due on November
17, 2026?”, the arguments are `{"date":"2026-11-17"}`.

A host can pass the tool name and serialized arguments to the dispatcher.
For example, this simulates a call with the workbook configured locally:

```python
from networking.followups import dispatch_tool_call

result = dispatch_tool_call("get_followups", '{"date":"2026-11-17","include_earlier":true}')
```

The dispatcher accepts the exact name `get_followups` and returns the response
from `handle_followup_json()`. Any other name yields `UNKNOWN_TOOL` before
argument parsing or workbook access. The example runs entirely in Python.

The local workbook is the authoritative working record. OneDrive is used for
backup and recovery, not continuous synchronization managed by this tool.
The script has no workbook-save operation and makes no network calls.
Configuration, source code, and personal data remain separate.

Excel files, virtual environments, and local environment files are excluded
by `.gitignore`. Review staged changes before committing; ignore rules do not
remove files that were already tracked. Console output can contain personal
contact information, so use synthetic examples in shared logs and screenshots.

## Current limitations

- Queries support an exact due date or an inclusive due-through cutoff on
  `Contacts`; there are no bounded date ranges, completion-status filters,
  or queries against other worksheets.
- The CLI handles ContactSchemaError, InvalidFileException, BadZipFile, and
  OSError subclasses raised by the query, using its existing error messages.
  Other failures, including malformed XML or missing workbook components,
  may still produce tracebacks. No broad exception handler is used.
- Text dates are not parsed from workbook cells.
- Formula cells use their stored cached values; the script does not calculate
  formulas. Missing or stale caches can affect query results.
- No record updates, AI integration, MCP interface, audit trail, or application
  authorization layer is implemented. Local operating-system permissions
  control file access.

## Development and roadmap

Follow GitHub Flow: start a short-lived branch from updated `main`, make a
small change, validate it, and open a pull request. Review the PR's
**Files changed** tab before merging. Keep `main` stable.

The planned progression is deterministic data access, controlled updates with
validation/backups/explicit authorization, AI integration, MCP integration,
and broader governance safeguards. These are roadmap items, not current features.

See [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md) for architectural decisions,
milestone evidence, and the current stopping point.



