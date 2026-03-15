Download ASX earnings reports for company: $ARGUMENTS

First, ask the user how far back they want to download reports using AskUserQuestion with these options:
- "Last results day" (period: last)
- "1 year" (period: 1)
- "3 years" (period: 3)
- "5 years" (period: 5)
- "10 years" (period: 10)

Then run the following command with the chosen period, ensuring the CLAUDECODE environment variable is unset so the nested Claude CLI call works:

```
unset CLAUDECODE && python scripts/asx_reports.py $ARGUMENTS <period>
```

After the script completes:
1. Report how many files were downloaded, skipped, and failed
2. List the contents of the company's `Company reports/` folder
