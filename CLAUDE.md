# CLAUDE.md

## About This Workspace

Project Equities is an equity research workspace for tracking and modeling public companies. Work here includes financial modeling (Excel), data analysis (CSV), written research, and building custom tools and skills for use by other analysts.

## Directory Layout

- `skills/` — skill files defining how to perform specific workflows
- `commands/` — slash command files for repeatable tasks
- `templates/` — Excel model templates used as bases for new models
- `references/` — project-wide context files and guidelines
- `<TICKER>/` — one folder per company (e.g. `TTI/`, `XYZ/`), using `Template folder dir/` as the starting point for new companies

All of the above live under `.claude/` except company folders which are at project root.

### Company Folder Structure
Each ticker folder contains:
- **Models/** — Excel financial models (`<Ticker> Model.xlsx`)
- **Broker notes/** — Third-party analyst reports
- **Company reports/** — Filings and investor presentations
- **PW notes/** — Internal research notes
- **context-log.md** — rolling session log for model status, open items, and key decisions

## Skill Files

- Each skill lives in `.claude/skills/<skill-name>/SKILL.md`
- Supplementary reference docs for a skill live in `.claude/skills/<skill-name>/references/`
- When a modeling task is requested, identify the relevant skill and load it before proceeding

## Financial Modeling Conventions

- Always read the relevant skill file before beginning any modeling task
- Excel templates are in `.claude/templates/` — reference via `@.claude/templates/<filename>.xlsx`
- Template filenames use hyphens, no spaces (e.g. `restaurant-lbo-template.xlsx`)
- Never modify a file in `templates/` directly — always copy to the relevant company `Models/` folder first
- Always read the Excel file fresh from disk before making any modifications, even if you have worked with it earlier in the same session. Never rely on a previously loaded version of the file.
- Never overwrite existing rows unless explicitly instructed — insert new rows instead
- Files prefixed with `~$` are Excel lock files from open workbooks — ignore them

## Context Logs

- Every company folder contains a `context-log.md`
- At the start of every session, read the relevant company `context-log.md` before doing any work
- At the end of every session, update `context-log.md` as follows:
  - Overwrite the Last Session section entirely
  - Update Model Status to reflect current state
  - Append to Key Decisions only if a significant modeling choice was made
  - Append to Open Items if anything is unresolved or flagged
- Keep `context-log.md` under 100 lines — summarise and compress older entries rather than appending indefinitely
- Never delete entries from Key Decisions — summarise them if space is needed
- When creating a new company folder, copy `.claude/references/context-log-template.md` into the new company folder as `context-log.md`
- If a company folder does not contain a `context-log.md`, create one immediately using the template at `.claude/references/context-log-template.md` before proceeding with any work

## Data Accuracy

- Never fabricate or estimate financial data. Use only what is provided in source files or confirmed by the user.
- When presenting numbers, state the source (file name, row, column) so the user can verify.
- Flag any assumptions explicitly — distinguish between hard data and inferred values.
- If data appears inconsistent or incomplete, raise it rather than working around it silently.
- Never guess. If any input, assumption, or formula is unclear, ask before making changes.