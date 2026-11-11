# AI System of Record — Project Context

**Repository:** `AICloudDataPulse/ai-system-of-record`  
**Stable branch:** `main`
**Current documentation branch:** `docs/milestone-8-status`
**Active feature development:** None; next increment not yet agreed.
**Development environment:** macOS, VS Code, Python 3.12.3
**Status:** Milestones 1–8 completed and merged into `main`.
Milestone 8 merged through PR #8 on November 11, 2026.
All 30 automated tests passed in the execution environment before merge.
Local macOS validation has not been separately confirmed.
The Milestone 7 status update was merged through PR #7 on November 10, 2026.

## 1. Project objective

Develop a secure, controlled integration between AI assistants and external systems of record.

The initial use case is a personal networking and career-management Excel workbook.

The project will progress incrementally from deterministic Python utilities toward AI-assisted workflows and potentially MCP-based integration.

## 2. Architecture decisions

- The local Excel workbook is the authoritative working statement of record.
- OneDrive is used for backup and recovery, not continuous synchronization.
- The workbook must never be committed to GitHub.
- Initial operations are strictly read-only.
- Future write operations will require validation, backups, and explicit authorization.
- Keep the implementation simple and avoid unnecessary infrastructure.
- Use synthetic data for public examples and tests.

## 3. Repository and configuration management

**GitHub account:** AICloudDataPulse  
**Repository:** `ai-system-of-record`

Branching strategy: GitHub Flow.

- `main`: stable, tested code
- `feature/*`: new functionality
- `fix/*`: bug fixes
- `docs/*`: documentation
- `chore/*`: maintenance

Use short-lived branches, meaningful commits, pull requests, and merge completed features into `main`.

SSH authentication is configured using a dedicated ACDP identity (`github-acdp`).

## 4. Current implementation

**Script:** `networking/followups.py`  
**Library:** `openpyxl`  
**Python environment:** `.venv`

The script:

1. Opens the local Excel workbook in read-only mode.
2. Validates the `Contacts` worksheet and its required first-row headers,
   then reads its records.
3. Accepts a date in `YYYY-MM-DD` format.
4. Matches records using the `Follow-up Due` column: exact date by default,
   or on/before the entered date with `--due-through`.
5. Prints contact name, company, and next action.
6. Reports the number of matching contacts.

The workbook currently contains four worksheets:

- Contacts
- Referals
- Applications
- Interviews

### Date validation and testing

- `parse_followup_date()` validates calendar dates and requires
  `YYYY-MM-DD` format, allowing surrounding whitespace.
- Invalid input produces a clear error and exits with status 1 before
  opening the workbook.
- Workbook configuration is read inside `main()`, allowing the parser
  to be imported without setting `NETWORKING_WORKBOOK`.
- Automated tests use Python's built-in `unittest` framework.
- Validation tests: `tests/test_followups.py`.
- Query tests: `tests/test_followup_queries.py`, using temporary synthetic
  workbooks and the real read-only workbook loader.
- Schema tests: `tests/test_contact_schema.py`, covering both CLI modes and
  workbook closure after schema errors.

Run tests from the repository root with the project virtual
environment activated:

    python -m unittest discover -s tests -v

Dependencies must be installed in the virtual environment from
`requirements.txt`.

## 5. Verified milestones

**Milestone 1 — Completed**

Successfully opened the workbook and listed its worksheets.

**Milestone 2 — Completed**

Successfully queried contacts requiring follow-up on November 7, 2026.

Result: three matching contacts.

The workbook was not modified.

**Milestone 3 — Completed**

Removed the hardcoded workbook path and introduced the
NETWORKING_WORKBOOK environment variable.

Verified that the November 7, 2026 query still returns
the same three contacts.

The workbook remains read-only.

**First GitHub Flow cycle — Completed**

- Implemented and tested Milestones 1–3.
- Created and reviewed the first pull request.
- Squash-merged `feature/networking-followups` into `main`.
- Synchronized the local `main` branch with GitHub.
- Deleted the completed local feature branch.
- Created `feature/followup-validation` for Milestone 4.

**Milestone 4 — Completed and merged**

Implemented:
- Strict YYYY-MM-DD date validation with surrounding whitespace allowed.
- Clear error handling for invalid date input.
- Missing or empty NETWORKING_WORKBOOK configuration handling.
- Missing workbook, permission-denied, and other file-access error handling.
- Automated tests using Python's built-in unittest framework.

Verified locally:
- All 10 automated tests pass.
- Date tests cover valid dates, surrounding whitespace, leap-day
  handling, and eight invalid-input cases.
- Configuration tests cover missing, empty, whitespace-only,
  and configured workbook paths.
- File-access tests simulate missing files, permission failures,
  and other operating-system errors.
- Manual invalid-date and missing-configuration checks passed:
  clear errors without tracebacks and exit status 1.
- Manual missing-file check passed: clear error without a traceback
  and exit status 1.
- The normal November 7, 2026 query still returns three follow-ups.

Workbook access remains read-only. Tests do not require personal
workbook data.

Malformed workbook contents and worksheet-schema validation are
outside this milestone's scope.

Merged into `main` through PR #2 on November 8, 2026
(merge commit `9e08e26178070f7be3b387f8c4c79ec9846c249d`).

**Milestone 5 — Completed and merged**

Added four automated tests in `tests/test_followup_queries.py`:
- Matching contact details and total, including a date-time value with
  a non-midnight time.
- Exclusion of contacts due on other dates and contacts with blank dates.
- Zero matching contacts.
- Byte-for-byte workbook preservation after a successful query.

Tests generate synthetic Excel workbooks in `TemporaryDirectory` and
exercise `main()` with the real workbook loader. Input, configuration,
and console output are controlled within each test. Temporary files are
cleaned up automatically; no workbook fixtures or personal data are committed.

Verified on November 9, 2026 in the development execution environment:
- Python 3.12.14 and openpyxl 3.1.5.
- `python -m unittest discover -s tests -v`: all 14 tests passed.
- No production code or dependency changes.

Merged into `main` through [PR #3](https://github.com/AICloudDataPulse/ai-system-of-record/pull/3)
on November 9, 2026
(merge commit `9129e5f3c2e453a82cb57c0e44ce108233d0a2c6`).

Local validation also completed on macOS on November 9, 2026, as confirmed
by the developer's terminal output:
- Project `.venv` activated.
- Dependencies from `requirements.txt` satisfied:
  openpyxl 3.1.5 and et_xmlfile 2.0.0.
- `python -m unittest discover -s tests -v`: all 14 tests passed,
  ending with `OK`.

Validation used temporary synthetic workbooks; the personal workbook
was not required.

**Milestone 6 — Completed and merged**

Expanded `README.md` with environment setup, workbook configuration,
query/test commands, required worksheet headers, architecture, data protection,
current limitations, and the GitHub Flow review process.

Validation on November 9, 2026 in the execution environment:
- Reviewed the documentation against the current source and pinned dependencies.
- Activated the existing virtual environment and checked requirements installation.
- All 14 automated tests passed with Python 3.12.14 and openpyxl 3.1.5.
- Ran the documented script entry point using a temporary synthetic workbook
  at a path containing spaces; output matched the README example and exit status was 0.
- Checked invalid-date and missing-configuration invocations: clear errors
  and exit status 1.
- No application code, tests, dependencies, or infrastructure changed.

The existing execution virtual environment uses installed runtime packages.
Fresh dependency downloads and a new macOS setup were not revalidated.
No personal workbook was accessed.

The preceding context-only update was merged through PR #4
(commit `281eac089e4e9fc9e10f7df7ed308739496a1391`).

Merged through [PR #5](https://github.com/AICloudDataPulse/ai-system-of-record/pull/5)
on November 9, 2026 (commit `c8ec6b40d97ad4521af7df9320fa69f90869f1ef`).
The developer's November 10 terminal output confirmed local `main` at this
commit with a clean working tree and the virtual environment active.

**Milestone 7 — Completed and merged**

Added the optional `--due-through` command-line flag. The existing interactive
date prompt remains; without the flag, exact-date behavior is unchanged.
With the flag, records due on or before the entered date are included.
Results retain worksheet order and the existing output fields and total.
Blank and non-date values are skipped. Date-times use their calendar date.
This query does not inspect completion status or infer today's date.

Validation on November 10, 2026 in the execution environment:
- The original 14-test baseline passed before changes.
- All 21 tests passed with Python 3.12.14 and openpyxl 3.1.5.
- Seven additional tests cover inclusive boundaries and earlier dates,
  later/blank/text exclusions, zero matches, workbook preservation, actual
  CLI invocations for both modes, invalid dates, and unknown CLI options.
- Tests use temporary synthetic workbooks; no personal data was accessed.
- Read-only loading, configuration, and pinned dependencies are unchanged.

The execution virtual environment uses installed runtime packages;
fresh dependency downloads were not revalidated.

The developer confirmed successful local testing and merge on November 10,
2026. Detailed local test output was not supplied for this increment; the
21-test execution-environment result above is independently recorded.

Merged into `main` through [PR #6](https://github.com/AICloudDataPulse/ai-system-of-record/pull/6)
on November 10, 2026
(merge commit `9616ae3e2b9aeea613a8ac0da564d5cb6651cc2c`).
GitHub confirmed that `main` points to this merge commit. Synchronization of
the developer's local `main` after this merge has not been confirmed.

**Milestone 8 — Completed and merged**

Added validation before scanning contact records:
- Require a worksheet named exactly `Contacts` and a nonblank first row.
- Require exactly one occurrence of each header: `Name`, `Company`,
  `Next Action`, and `Follow-up Due`.
- Report missing and duplicate required headers without printing workbook
  records or the configured path. Schema errors go to standard error and
  exit with status 1, without a traceback.
- Preserve exact header matching, arbitrary column order, ignored extra
  columns, both date-query modes, and read-only access.
- A valid header-only worksheet returns zero follow-ups successfully.
- Close the workbook on validation failure through the existing `finally` block.

Validation on November 11, 2026 in the execution environment:
- Original 21-test baseline passed before changes.
- All 30 tests passed with Python 3.12.14 and openpyxl 3.1.5.
- Nine new tests cover missing worksheet, empty worksheet, blank first row,
  each missing/duplicate required header, exact header spelling/spacing,
  header-only sheets, reordered/extra columns, and closure after schema errors.
- CLI cases exercise both query modes with temporary synthetic workbooks;
  error cases verify status 1, no traceback or query results, and unchanged
  workbook bytes. Valid reordered-column queries also preserve workbook bytes.
- No personal workbook was accessed; no dependencies or infrastructure changed.
- The execution virtual environment uses installed runtime packages;
  fresh dependency downloads were not revalidated.

Merged into `main` through [PR #8](https://github.com/AICloudDataPulse/ai-system-of-record/pull/8)
on November 11, 2026
(merge commit `be6452389a34e1efd2fac70673605c532d4046b2`).
The developer reported the merge, and GitHub confirmed PR #8 is merged
and `main` points to this commit at verification time.
Local macOS validation and synchronization of the developer's local `main`
have not been separately confirmed; the merge report alone is not test evidence.
Malformed workbook and unsupported-file-format handling remain out of scope.

## 6. Next development tasks

1. Review and merge this documentation-only status update.
2. Synchronize local `main` after merging.
3. Confirm local macOS validation if it has not already been performed.
4. Review the current source and agree on one small increment before implementation.

No next feature has been selected or started.

Complete and validate each increment before starting the next.
Include focused automated tests with behavior changes and record
validation results in this context.

Keep personal workbook data and local configuration out of Git.

## 7. Longer-term roadmap

**Phase 1 — Deterministic data access**

Read Excel records and support queries.

**Phase 2 — Controlled updates**

Implement validated changes with backups and explicit authorization.

**Phase 3 — AI integration**

Allow an AI assistant to invoke narrowly scoped operations.

**Phase 4 — MCP integration**

Expose selected operations through an MCP server, subject to appropriate security controls.

**Phase 5 — Governance**

Introduce audit trails, authorization boundaries, error recovery, and operational safeguards.

## 8. Engineering principles

- Take small, independently testable steps.
- Validate each milestone before proceeding.
- Separate data, configuration, and source code.
- Apply least privilege.
- Keep personal information out of public repositories.
- Avoid introducing AI where deterministic code is sufficient.
- Preserve Git configuration-management discipline.
- Include focused automated tests with changes to parsing, validation,
  configuration handling, and query behavior.
- Run the existing test suite before committing an increment.
- Use synthetic data and temporary workbooks for automated tests;
  never depend on personal workbook data or a developer-specific path.
- Supplement automated tests with manual checks of the user-facing
  script where appropriate.
- Record what was tested and distinguish verified results from
  pending checks.

## 9. Current stopping point

Milestones 1–8 are completed and merged into `main`.
All 30 automated tests passed in the execution environment before the
Milestone 8 merge. Local macOS validation and local synchronization have
not been separately confirmed. See Section 5 for the validation evidence.

Workbook access remains read-only. Both exact-date and `--due-through`
queries retain their existing behavior for valid worksheets.

This documentation-only update on `docs/milestone-8-status` records the
confirmed merge and changes only `PROJECT_CONTEXT.md`.

**Next immediate task:** Review and merge this status update, then synchronize
local `main`. Agree on the next small increment before implementing it.

Continue making small, independently testable changes and validate
each increment before proceeding.

Do not introduce MCP, agents, or additional infrastructure yet.
