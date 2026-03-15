---
name: company-excel-model
description: Build and maintain Excel equity research financial models using a standardised 3-sheet template (Annual, Value, HY & Segments)
---

# Company Excel Model Skill

## Critical Principles

1. **Read fresh** — always read the workbook from disk before any operation; never rely on stale data.
2. **No overwrites** — never overwrite existing data without explicit instruction.
3. **Formulas over hardcodes** — preserve formula cells; only hardcode where the template expects hard inputs.
4. **Never fabricate data** — if a value is unknown, leave the cell blank or flag it; do not invent numbers.
5. **Column A codes are lookup keys** — every INDEX/MATCH cross-sheet reference resolves via the Column A coded label, not by row number.

## Template Overview

| Sheet | Purpose | Columns |
|-------|---------|---------|
| Annual | Full-year P&L, EPS, KPIs, Balance Sheet, Cash Flow, ROIC | A-P (13 years) |
| Value | Market snapshot, WACC, DCF, SOTP | B-R |
| HY & Segments | Half-year P&L, segment forecasts, HY KPIs | A-AC (26 half-year periods) |

## Row Layout Convention

| Column | Content |
|--------|---------|
| A | Coded label (lookup key, e.g. `Rev-Total Revenue`) |
| B | Display name |
| C | Units (e.g. `[CCY]m`, `%`, `#m`, `x`) |
| D+ | Period data (years on Annual/Value, halves on HY & Segments) |

## Column A Coding Convention

Prefix codes identify line-item groups:

| Prefix | Group | Examples |
|--------|-------|----------|
| `Rev-` | Revenue | `Rev-[Segment] Revenue`, `Rev-Total Revenue` |
| `COGS-` | Cost of Goods Sold | `COGS-[Segment] COGS`, `COGS-Total COGS` |
| `GP-` | Gross Profit | `GP-[Segment] GP`, `GP-Gross Profit` |
| `OPEX-` | Operating Expenses | `OPEX-Employee Benefits`, `OPEX-Total OpEx` |
| `EBITDA-` | EBITDA | `EBITDA-[Segment] EBITDA`, `EBITDA-Underlying EBITDA` |
| `Stat-` | Statutory Adjustments | `Stat-SBP`, `Stat-Significant Items` |
| `DA-` | Depreciation & Amortisation | `DA-Depreciation PPE`, `DA-ROU Amortisation`, `DA-Total DA` |
| `EBIT-` | EBIT | `EBIT-Underlying EBIT` |
| `Int-` | Interest / Finance | `Int-Interest Income`, `Int-Net Finance Costs` |
| `PBT-` | Profit Before Tax | `PBT-PBT` |
| `Tax-` | Tax | `Tax-Tax Expense` |
| `NPAT-` | Net Profit After Tax | `NPAT-Underlying NPAT`, `NPAT-Statutory NPAT` |
| `EPS-` | Earnings Per Share | `EPS-WASO Diluted`, `EPS-Underlying EPS` |
| `Div-` | Dividends | `Div-DPS`, `Div-Total Dividends` |
| `KPI-` | Operating Metrics | `KPI-[Segment] Volume`, `KPI-Headcount` |
| `BS-` | Balance Sheet | `BS-Cash`, `BS-Trade Receivables`, `BS-Total Banking Debt` |
| `CF-` | Cash Flow | `CF-EBITDA`, `CF-Capex PPE`, `CF-Net OCF` |

## Cross-Sheet Formula Patterns

### INDEX/MATCH (Value ← Annual)

All Value sheet projections pull from Annual via:

```
=INDEX(Annual!$D:$P, MATCH("[Col A Code]", Annual!$A:$A, 0), MATCH([FY label], Annual!$D$3:$P$3, 0))
```

### HY-to-Annual Linkage (2H derivation)

For forecast 2H periods, derive as full year minus 1H:

```
=INDEX(Annual!$A:$R, MATCH($A[row], Annual!$A:$A, 0), MATCH([year], Annual!$A$1:$R$1, 0)) - [1H value]
```

### WACC

```
Cost of Equity = Risk-free Rate + ERP × Beta
After-tax Cost of Debt = Pre-tax CoD × (1 - Tax Rate)
WACC = CoE × (1 - D/(D+E)) + ATCoD × D/(D+E)
```

### DCF Discount Factors

```
Stub Period = (FYE Date - Valuation Date) / 365.25
Discount Factor = 1 / (1 + WACC) ^ (Stub + n)    where n = 0, 1, ..., 9
```

### Terminal Value (Gordon Growth)

```
Terminal Value = Normalised FCFF × (1 + g) / (WACC - g)
```

Normalised FCFF assumes capex = D&A in the terminal year: `NOPAT + WC Change`.

### FCFF Build

```
FCFF = NOPAT + D&A + Capex + WC Change
     where NOPAT = EBIT × (1 - Tax Rate)
```

### DCF Equity Bridge

```
Enterprise Value = Σ PV(FCFF) + PV(Terminal Value)
Equity Value = EV - Net Debt - Lease Liabilities
Per Share Value = Equity Value / Shares Outstanding
```

### SOTP (EV/EBITDA)

```
Segment EV = Segment EBITDA × Multiple
Group EV = Σ Segment EVs
Equity Value = Group EV - Net Debt - Lease Liabilities
```

Corporate segment multiple = blended: `(Seg1 EBITDA × Mult1 + Seg2 EBITDA × Mult2) / (Seg1 + Seg2)`

## Workflow: New Company Model

1. **Copy template** — duplicate `HY_model_template.xlsx`, rename to `[Ticker]_model.xlsx`.
2. **Define segments** — replace `[Segment]` placeholders in Column A codes across all three sheets with actual segment names.
3. **Set currency** — update unit labels from `[CCY]m` to actual currency (e.g. `AUDm`, `USDm`).
4. **Set FYE** — adjust period-end dates in row 4 of Annual and HY sheets to match the company's fiscal year end.
5. **Enter historicals** — populate hard-input cells on the Annual sheet with reported financials (at least 2-3 years).
6. **Enter HY historicals** — populate half-year actuals on HY & Segments sheet.
7. **Build segment forecasts** — fill segment forecast blocks (volume, pricing, margins) on HY & Segments; verify HY totals flow to Annual via linkage formulas.
8. **Complete BS & CF forecasts** — populate forecast balance sheet and cash flow assumptions on Annual.
9. **Set valuation inputs** — on Value sheet: share price, WACC inputs, SOTP multiples.
10. **Run integrity checks** — verify BS Check = 0, HY = Annual reconciliation, DCF decomposition.

## Workflow: Model Update

1. **Read workbook** — load fresh from disk.
2. **Identify update scope** — new actuals? revised forecasts? valuation refresh?
3. **Update inputs** — enter new hard-input values only in cells designated for hard inputs.
4. **Verify formulas intact** — spot-check key calculated cells haven't been overwritten.
5. **Run integrity checks** — BS Check, cash reconciliation, HY = Annual.
6. **Update valuation date & price** — refresh Value sheet market snapshot if needed.

## Generalisation Notes

- **Currency**: All monetary units use `[CCY]m` — replace with actual currency code.
- **Segments**: Template supports N segments. Each segment needs Revenue, COGS, GP, EBITDA rows with matching Column A codes. Add or remove segment rows as needed.
- **Fiscal Year End**: Adjust period-end dates and period labels (e.g. `FY25A`, `1H25`) to match the company's reporting calendar.
- **KPIs**: Operating metrics are industry-specific. Replace volume/pricing KPIs with relevant metrics (e.g. subscribers, stores, tonnes, MWh).
- **Leases**: IFRS 16 lease treatment (ROU assets, lease liabilities, lease interest) is baked into the template. Adjust if the company uses a different standard or has immaterial leases.
