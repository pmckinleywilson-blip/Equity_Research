---
description: >
  Save the build decisions from a completed /build-model run as a test fixture.
  The fixture enables automated re-testing via /test-build. Run after a
  successful /build-model: /save-fixture [TICKER]
---

# /save-fixture

Save the build decisions for: $ARGUMENTS

---

## Step 1 — Locate the build artifacts

1. Read `[TICKER]/context-log.md` — this contains the key decisions, segment
   structure, and modelling choices from the build session.
2. Identify which skill file was used — check `.claude/skills/` for the skill
   that corresponds to the template used.
3. Check if a test fixture already exists at
   `.claude/skills/[skill-name]/test-fixture-[TICKER].yml`. If it does, ask the
   user whether to overwrite or keep the existing one.

---

## Step 2 — Extract decisions from context log

Read the context log's Key Decisions section and extract:

- **Company details**: name, ticker, exchange, fiscal year end, currency,
  starting period
- **Segment structure**: segment names, what each includes, P&L depth, KPIs
  reported per segment
- **Forecast driver structure**: the exact Phase 3b build for each segment
  (revenue drivers, cost drivers, EBITDA derivation)
- **Group-level structure**: how revenue, expenses, and EBITDA are composed
  at the group level; whether expenses are segment-driven residuals or
  independently forecast
- **Key modelling decisions**: any company-specific accounting treatments,
  BS items, or formula routing decisions
- **Data limitations**: which periods have incomplete data, what's estimated

Also extract:
- **Source documents used**: list the files in `[TICKER]/Company reports/`
- **Actual periods**: which annual and sub-period columns have actuals
- **Forecast start**: which column is the first forecast period

---

## Step 3 — Build the fixture

Write a YAML file with the following structure:

```yaml
template: [template filename]
skill: [skill folder name]

ingest:
  frequency: [half-year / quarterly / annual]
  retain_all_analytical: [true/false]
  retain_all_valuations: [true/false]

build:
  company:
    name: [company name]
    ticker: [TICKER.EXCHANGE]
    exchange: [exchange name]
    fiscal_year_end: [date]
    currency: [currency code]
    start_period: [earliest actual period]
    macro_indicators: [none / list]

  source_documents:
    - [filename 1]
    - [filename 2]
    ...

  segments:
    - name: [segment name]
      includes: [list of geographies/divisions]
      depth: [Revenue / GP / EBITDA / EBIT]
      kpis:
        - [KPI 1]
        - [KPI 2]
        ...

  forecast_drivers:
    [segment_name]:
      [section]:
        - [driver description: formula logic]
        ...

  group_level:
    revenue: [composition]
    other_revenue: [composition]
    cogs: [composition]
    expenses: [list of expense lines]
    ebitda_approach: [segment_driven / independent]
    ebitda_bridge: [bridge formula]
    expense_forecast_method: [description]
    da_components: [list]
    finance_components:
      income: [list]
      costs: [list]
    company_specific_bs_items: [list]

  decisions:
    - [decision 1]
    - [decision 2]
    ...

validation:
  annual_actuals: [list of periods]
  hy_actuals: [list of periods]
  forecast_start_annual: [period]
  forecast_start_hy: [period]
  limitations:
    - [limitation 1]
    ...
  expected_rows:
    [sheet_name]: [range]
    ...
  ratio_checks:
    [ratio_name]: [min, max]
    ...
```

---

## Step 4 — Write and confirm

1. Write the fixture to `.claude/skills/[skill-name]/test-fixture-[TICKER].yml`
2. Confirm to the user: "Test fixture saved to [path]. This captures the
   build decisions for [TICKER] so that future builds can be automatically
   tested."
3. Prompt the user with next steps:
   - **To validate the current model:** run `/validate-model [TICKER]`
   - **To run the automated test loop:** run `/test-build [skill-name]`
   - **To build a model for another company:** run `/build-model [TICKER2]`
