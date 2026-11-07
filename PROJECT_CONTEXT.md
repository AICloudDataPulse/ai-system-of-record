# AI System of Record — Project Context

**Repository:** `AICloudDataPulse/ai-system-of-record`  
**Current branch:** `feature/followup-validation`
**Development environment:** macOS, VS Code, Python 3.12.3
**Status:** Milestones 1–3 completed and merged into `main`. Milestone 4 in preparation.

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

## 6. Next development tasks

1. Milestone 4 — Improve input validation and error handling.
   - Handle invalid date input.
   - Handle missing `NETWORKING_WORKBOOK` configuration.
   - Handle missing or inaccessible workbook files.
2. Add automated tests using synthetic workbook data.
3. Document usage and architecture.
4. Expand read-only query capabilities.

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

## 9. Current stopping point

Milestones 1–3 have been completed, tested, and merged into `main`.

Current working branch: `feature/followup-validation`.

**Next immediate task:** Implement and manually test invalid date
input handling in `networking/followups.py`.

Proceed incrementally, making one small, independently testable
change at a time.

Do not introduce MCP, agents, or additional infrastructure yet.
