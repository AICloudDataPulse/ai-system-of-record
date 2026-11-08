# AI System of Record — Project Context

**Repository:** `AICloudDataPulse/ai-system-of-record`  
**Current branch:** `feature/followup-validation`
**Development environment:** macOS, VS Code, Python 3.12.3
**Status:** Milestones 1–3 completed and merged into `main`. Milestone 4 in progress — date validation implemented and four automated tests passing. Manual validation of the complete script is pending.

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
4. Matches records using the `Follow-up Due` column.
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
- Test file: `tests/test_followups.py`.

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

**Milestone 4 — In progress**

Implemented date-input validation, workbook configuration validation,
and automated tests.

Verified locally:
- Seven automated tests pass.
- Date coverage includes valid dates, surrounding whitespace, a valid
  leap day, and eight invalid-input cases.
- Configuration coverage includes missing, empty, whitespace-only,
  and configured workbook paths.
- Manual missing-configuration check passed: clear error without
  a traceback and exit status 1.

Manual validation pending:
- Confirm a valid date still returns the expected workbook results.
- Confirm invalid date input displays a clear error without a traceback.
- Confirm invalid date input exits with status 1.

Remaining Milestone 4 work:
- Handle missing or inaccessible workbook files.
- Add automated tests for workbook-file error handling.

Milestone 4 is not yet complete.

## 6. Next development tasks

## 6. Next development tasks

1. Complete manual validation of the date-validation increment.
2. Review the diff and commit the validated code, tests, and context
   update on `feature/followup-validation`.
3. Add missing or empty workbook-configuration handling and tests.
4. Add missing or inaccessible workbook-file handling and tests.
5. Add query tests using temporary workbooks with synthetic data.
6. Document usage and architecture.
7. Expand read-only query capabilities.

Complete and validate each increment before starting the next.
Merge completed, tested work into `main` through a pull request.

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

Milestones 1–3 are completed and merged into `main`.

Current working branch: `feature/followup-validation`.

The date-validation increment is implemented locally.
All four automated tests pass.

**Next immediate task:** Manually validate the complete script with
a valid date and an invalid date, including the invalid-input exit
status. Then review and commit this increment with its tests and
updated project context.

After that, address missing or empty workbook configuration as the
next small, independently testable change.

Do not introduce MCP, agents, or additional infrastructure yet.
