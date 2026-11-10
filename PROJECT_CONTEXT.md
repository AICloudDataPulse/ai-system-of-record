# AI System of Record — Project Context

**Repository:** `AICloudDataPulse/ai-system-of-record`  
**Stable branch:** `main`
**Current development branch:** `feature/followups-due-through`
**Development environment:** macOS, VS Code, Python 3.12.3
**Status:** Milestones 1–6 completed and merged into `main`.
Milestone 6 merged through PR #5 on November 9, 2026.
Milestone 7 due-through queries are implemented and validated on the feature
branch; PR review, local macOS validation, and merge remain pending.

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
2. Reads the `Contacts` worksheet.
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

**Milestone 7 — Implemented and validated; review and merge pending**

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
fresh dependency downloads and local macOS validation of this increment
remain unverified.

## 6. Next development tasks

1. Review the due-through query PR from `feature/followups-due-through`.
2. Run the 21-test suite and try both query modes locally on macOS.
3. Address any findings, then merge and synchronize local `main`.
4. Record the merge with the next agreed increment.

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

Milestones 1–6 are completed and merged into `main`.
Milestone 7 is implemented on `feature/followups-due-through` and ready for
PR review. It changes `networking/followups.py`, its query tests, `README.md`,
and this context file. No dependencies or infrastructure were added.

All 21 automated tests pass in the execution environment. Local macOS
validation and PR merge remain pending. The earlier local baseline of
14 passing tests remains recorded under Milestone 5.

**Next immediate task:** Review the feature PR and validate locally before
merging. No subsequent implementation has started.

Continue making small, independently testable changes and validate
each increment before proceeding.

Do not introduce MCP, agents, or additional infrastructure yet.
