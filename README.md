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

This mode filters due dates only; it does not check whether an action has
already been completed. Use `python networking/followups.py --help` for help.
Unknown command-line options exit with status 2 before prompting for a date.

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

## Run the tests

With the virtual environment active, run from the repository root:

```bash
python -m unittest discover -s tests -v
```

The current suite contains 33 tests and should finish with `OK`.
It does not require your personal workbook or `NETWORKING_WORKBOOK` setting.

- [tests/test_followups.py](tests/test_followups.py) covers date validation,
  configuration, and simulated file-access failures.
- [tests/test_followup_queries.py](tests/test_followup_queries.py) creates
  temporary synthetic workbooks and checks matching output, exclusions,
  zero matches, byte-for-byte workbook preservation, and the command-line
  behavior of both query modes.
- [tests/test_contact_schema.py](tests/test_contact_schema.py) checks worksheet
  and header errors through both CLI modes, exact header matching, reordered
  and extra columns, header-only sheets, workbook preservation, and closure
  after schema errors.
- [tests/test_workbook_formats.py](tests/test_workbook_formats.py) checks unsupported
  extensions and invalid ZIP containers through both CLI modes, including empty,
  text, and truncated files, safe error output, and unchanged file bytes.

Temporary test workbooks are cleaned up automatically.

## Current architecture and data protection

[networking/followups.py](networking/followups.py) validates the entered date,
reads the workbook path from the environment, and opens the workbook through
`openpyxl` with `read_only=True` and `data_only=True`. It locates the required
columns by their headers, scans the rows, prints matching records and a count,
and closes the workbook in a `finally` block.

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
- Unsupported file extensions and invalid ZIP containers detected while opening
  the workbook produce friendly errors. Deeper corruption inside a valid ZIP
  archive (such as malformed XML or missing workbook components), and errors
  during row iteration, remain outside this handling and may produce tracebacks.
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

