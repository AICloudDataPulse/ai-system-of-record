# AI System of Record — Project Context

**Repository:** `AICloudDataPulse/ai-system-of-record`  
**Stable branch:** `main`
**Active development branch:** `feature/followup-date-argument`
**Active feature development:** Milestone 10 — optional command-line date; implemented, pending review and merge.
**Development environment:** macOS, VS Code, Python 3.12.3
**Status:** Milestones 1–9 completed and merged into `main`.
Milestone 10 is implemented on the feature branch; review and merge are pending.
The Milestone 9 status update was merged through PR #11 on November 12, 2026
(commit `4e13620bf295591c7a4cb46eeefc1af21dd5d593`).
Milestone 9 merged through PR #10 on November 12, 2026.
All 33 automated tests passed in the execution environment before merge.
Local macOS validation has not been separately confirmed.
The Milestone 7 status update was merged through PR #7 on November 10, 2026.
The Milestone 8 status update was merged through PR #9 on November 11, 2026
(commit `18d26ad5bf8f2888e943fece8da859a435d17f7b`).

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
3. Accepts a date in `YYYY-MM-DD` format, through `--date` or the existing
   interactive prompt when the option is omitted.
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
- File-format tests: `tests/test_workbook_formats.py`, covering unsupported
  extensions and invalid ZIP containers detected during workbook loading.

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

**Milestone 9 — Completed and merged**

Added targeted handling around workbook loading for `InvalidFileException`
(unsupported extensions) and `BadZipFile` (invalid Excel ZIP containers).
Errors go to standard error with status 1, without raw exception details,
configured paths, tracebacks, or query results. Existing valid query behavior,
read-only loading, configuration, and dependencies are unchanged.

Validation on November 12, 2026 in the execution environment:
- Original 30-test baseline passed before changes.
- All 33 tests passed with Python 3.12.14 and openpyxl 3.1.5.
- Three new tests exercise both CLI modes with unsupported `.xls`, `.xlsb`,
  and `.csv` files; empty/text `.xlsx` files; and a truncated workbook ZIP.
- Cases verify status 1, exact error output, no path disclosure or results,
  and byte-for-byte file preservation.
- Tests use temporary synthetic files only; no personal workbook was accessed.
- The project virtual environment uses installed runtime packages;
  fresh dependency downloads were not revalidated.
- Local macOS validation has not been separately confirmed.

Merged into `main` through [PR #10](https://github.com/AICloudDataPulse/ai-system-of-record/pull/10)
on November 12, 2026
(merge commit `aa3b1e3d0b0e7a6001320465fc3a9fe6594e8803`).
The developer reported review and merge; GitHub confirmed the merge and
`main` at this commit. The merged tree matches the version that passed all
33 tests. Local macOS validation and synchronization of the developer's
local `main` have not been separately confirmed.

Deeper corruption within a valid ZIP archive and failures during row iteration
remain outside this increment. No broad exception handler was introduced.

**Milestone 10 — Implemented; review and merge pending**

Added optional `--date YYYY-MM-DD` to both exact-date and `--due-through` queries.
Omitting the option retains the interactive prompt. Both paths reuse
`parse_followup_date()`. An explicit empty/invalid date exits with status 1
before configuration lookup or workbook access; a missing argument value exits
with argparse usage status 2. Output fields, ordering, and read-only access remain
unchanged. No dependencies or infrastructure were added.

Validation on November 13, 2026 in the execution environment:
- Original 33-test baseline passed before changes.
- All 37 tests passed with Python 3.12.14 and openpyxl 3.1.5.
- Four additional tests cover explicit dates in both modes, identical results
  without prompting, surrounding whitespace, leap days, invalid/empty dates,
  missing option values, validation before workbook access, and unchanged bytes.
- Existing interactive CLI and workbook error tests continue to pass.
- Synthetic temporary workbooks only; no personal workbook accessed.
- The project virtual environment uses installed runtime packages; pinned
  requirements were satisfied, but fresh dependency downloads were not tested.
- Local macOS validation and merge remain pending.

## 6. Next development tasks

1. Review the Milestone 10 pull request and validate locally on macOS.
2. Merge after successful review and validation, then synchronize local `main`.
3. Record the confirmed merge and validation evidence in this document.
4. Agree on the next small increment before further implementation.

Complete and validate each increment before starting the next.
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

Milestones 1–9 are completed and merged into `main`, including the Milestone 9
status update through PR #11. Milestone 10 is implemented on
`feature/followup-date-argument` and awaits review, local macOS validation,
and merge. All 37 tests passed in the execution environment on November 13, 2026.

Workbook access remains read-only. Exact-date and `--due-through` queries now
accept an optional explicit `--date`; omitting it retains interactive input.

**Next immediate task:** Review and validate the Milestone 10 pull request.
No subsequent increment has been selected or started.

Do not introduce MCP, agents, or additional infrastructure yet.
