# AI System of Record — Project Context

**Repository:** `AICloudDataPulse/ai-system-of-record`  
**Current branch:** `feature/followup-validation`
**Development environment:** macOS, VS Code, Python 3.12.3
**Status:** Milestones 1–3 completed and merged into `main`.
Milestone 4 implementation and local validation completed, with
10 automated tests passing. Code, tests, and documentation are
committed and pushed to `feature/followup-validation`.
Ready for pull request review and merge.

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

**Milestone 4 — Implementation and local validation completed**

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

Code, tests, and documentation are committed and pushed to
`feature/followup-validation`.

Pull request creation, review, and merge remain outstanding.

## 6. Next development tasks

1. Open a pull request from `feature/followup-validation` into `main`,
   recording automated and manual validation results.
2. Review the complete pull request diff and address any findings.
3. Merge the reviewed changes and synchronize local `main`.
4. Update this context to record the merge and next working branch.
5. Create a new short-lived feature branch for query tests using
   temporary workbooks with synthetic data.
6. Document usage and architecture.
7. Expand read-only query capabilities.

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

Milestones 1–3 are completed and merged into `main`.

Current working branch: `feature/followup-validation`.

Milestone 4 implementation and local validation are complete:
- Date-input validation.
- Missing or empty workbook-configuration handling.
- Missing or inaccessible workbook-file handling.
- All 10 automated tests passing.
- Manual error checks passed with clear messages and exit status 1.
- Successful query reconfirmed with three follow-ups for November 7, 2026.

The implementation, tests, and validation documentation are committed
and pushed to the feature branch.

**Next immediate task:** Open a pull request from
`feature/followup-validation` into `main`, review the complete diff,
and address any findings before merging.

Milestone 4 has not yet been merged.

After review and merge, synchronize local `main`, create a new
feature branch for automated query tests using synthetic workbooks,
and update this context.

Continue making small, independently testable changes and validate
each increment before proceeding.

Do not introduce MCP, agents, or additional infrastructure yet.
