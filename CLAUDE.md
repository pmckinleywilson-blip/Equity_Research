# CLAUDE.md

## About This Workspace

Project Equities is an equity research workspace for tracking and modeling public companies. Work here includes financial modeling (Excel), data analysis (CSV), written research, and building custom tools and skills for use by other analysts.

## References

- Available project specific skills located \\wsl.localhost\Ubuntu\home\pmwilson\Project_Equities\.claude\skills
- Available project specific slash commands located \\wsl.localhost\Ubuntu\home\pmwilson\Project_Equities\.claude\commands
- Available project reference files for context and guidelines located \\wsl.localhost\Ubuntu\home\pmwilson\Project_Equities\.claude\references

## Working with Excel Models

- Claude has full access to create and modify `.xlsx` files.
- **Never guess.** If any input, assumption, or formula is unclear, ask the user before making changes.
- Files prefixed with `~$` are Excel lock files from open workbooks — ignore them.
- Always work from the latest excel file to preserve any manual changes that I've made.
- Never overwrite existing rows except if explicity instructed to, insert new rows instead.

## Data Accuracy

- Never fabricate or estimate financial data. Use only what is provided in source files or confirmed by the user.
- When presenting numbers, state the source (file name, row, column) so the user can verify.
- Flag any assumptions explicitly — distinguish between hard data and inferred values.
- If data appears inconsistent or incomplete, raise it rather than working around it silently.

## Folder Structure

Each company has its own top-level folder named by ticker symbol (e.g., `TTI/`, `XYZ/`). Use `Template folder dir/` as the starting point when adding a new company.

Every company folder follows this structure:

- **Models/** — Excel financial models (`<Ticker> Model.xlsx`)
- **Broker notes/** — Third-party analyst reports
- **Company reports/** — Filings and investor presentations
- **PW notes/** — Internal research notes

## Important Notes

