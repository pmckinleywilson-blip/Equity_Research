---
name: HY Equity Model (New Template)
description: Half-year (1H/2H) equity model template for companies reporting semi-annually. Use this skill when building or modifying financial models for HY-reporting companies. Do not use for quarterly-reporting companies.
---

# HY Equity Model — Skill File

## Template Reference

- **File**: `.claude/templates/HY_model_template_NEW.xlsx`
- **Reporting frequency**: Half-year (1H/2H) with annual roll-up
- **Currency**: Template default is NZD — replace with the target company's reporting currency at build time

---

## Sheet Architecture

5 sheets in this order:

| # | Sheet Name | Purpose |
|---|-----------|---------|
| 1 | **Value** | DCF valuation and EV/EBITDA Sum-of-the-Parts |
| 2 | **Annual** | Full annual P&L, Balance Sheet, Cash Flow, Operating FCF, ROIC (rows 1–194) |
| 3 | **HY & Segments** | Half-yearly consolidated P&L plus segment-level forecast engine (rows 1–116) |
| 4 | **Inputs** | **Cleared at build time** — leave blank. Do not create linkages to this sheet. |
| 5 | **Charts** | **Cleared at build time** — delete all chart objects and clear all cells. Leave blank. |

---

## Sheet Zone Architecture

### Annual Sheet Zones

| Zone | Rows | Purpose | Source / Dependent |
|------|------|---------|--------------------|
| Header | 1–4 | Years, period labels, dates | — |
| P&L | 5–76 | Full income statement with segments | **DEPENDENT** on HY & Segments |
| EPS & Dividends | 78–92 | Share counts, EPS, DPS, yields | DEPENDENT on P&L |
| Operating Metrics | 94–106 | Volumes, per-tonne, headcount | DEPENDENT on HY & P&L |
| Balance Sheet | 108–144 | Full BS with roll-forwards | Mix: **SOURCE** for BS assumptions; DEPENDENT on CF |
| Cash Flow | 146–177 | Full CF statement | DEPENDENT on P&L + BS |
| Operating FCF | 179–186 | FCF calculations and yields | DEPENDENT on CF |
| ROIC | 189–194 | Return metrics | DEPENDENT on P&L + BS |

### HY & Segments Sheet Zones

| Zone | Rows | Purpose | Source / Dependent |
|------|------|---------|--------------------|
| Header | 1–4 | Half-year period labels | — |
| Consolidated P&L | 5–90 | Half-yearly consolidated P&L, EPS, dividends | **DEPENDENT** on Segment Forecast (2H forecast+) |
| Segment Forecast — Seg 1 | 92–103 | First segment volume/price/margin/opex engine | **SOURCE** — forecasts originate here |
| Segment Forecast — Seg 2 | 105–116 | Second segment volume/price/margin/opex engine | **SOURCE** — forecasts originate here |

**Dependency direction**: The Segment Forecast zone is the **source** of all P&L forecasts. The Consolidated P&L zone's 2H+ forecast cells **must only contain references to the Segment Forecast zone's output rows**, never independent forecast logic.

Representative formula showing how Consolidated P&L pulls from Segment Forecast:
```
Steel Revenue (K7) = K97    [references Segment Forecast - Steel Revenue output]
Steel COGS (K13) = -K7*(1-K98)   [derived from revenue and GP margin assumption]
```

### Data Flow Direction
```
HY Segment Forecast zone (rows 92–116) → HY Consolidated P&L (rows 7–77)
HY Consolidated P&L → Annual P&L (via INDEX/MATCH summing 1H+2H)
Annual P&L + BS assumptions → Annual CF → Annual BS roll-forward
Annual BS/P&L → Value sheet (DCF + SOTP)
```

---

## Column Layout

### Annual Sheet
- **Col A**: Row key (hidden column). Used for cross-sheet INDEX/MATCH lookups.
- **Col B**: Human-readable label.
- **Col C**: Units (e.g. `NZDm`, `%`, `% YoY`, `NZDps`, `#m`, `000t`, `NZD`, `x`, `years`).
- **Col D**: First data column = first actual year.
- **Cols D–F**: Actuals (3 years).
- **Cols G–P**: Forecasts (10 years).
- **Row 1**: Calendar year integers (e.g. 2023, 2024, ...).
- **Row 2**: Zone labels (see Zone Label Positioning below).
- **Row 3**: Period headers (e.g. `FY23A`, `FY24A`, `FY26E`).
- **Row 4**: Period-end dates (e.g. 30-Jun-2023).

### HY & Segments Sheet
- **Col A**: Same row key system as Annual (hidden column). Identical keys for shared rows.
- **Col B**: Label.
- **Col C**: Units.
- **Cols D–I**: Actuals (6 half-year periods: 1H, 2H alternating).
- **Col J onwards**: Forecasts. Columns alternate 1H/2H.
- **Row 1**: Calendar year integers, each repeated twice (e.g. D1=E1=2023).
- **Row 2**: Zone labels.
- **Row 3**: Period headers (e.g. `1H23`, `2H23`, `1H24`).
- **Row 4**: Period-end dates (Dec 31 for 1H, Jun 30 for 2H — assumes June fiscal year-end; adjust to target company).

### Value Sheet
- No Col A keys. Layout starts at Col B.
- **Col B**: Labels.
- **Col C**: Values/inputs for DCF assumptions (rows 4–22) and summary outputs.
- **Cols D–M**: FCF projections for forecast years.

---

## Row Identification System

The model uses a **Column A key system** on both the Annual and HY & Segments sheets. Column A is hidden (width=30) and contains structured keys.

### Key Format

`SectionPrefix-RowName`

| Prefix | Section | Examples |
|--------|---------|----------|
| `Rev-` | Revenue | `Rev-Steel Revenue`, `Rev-Total Revenue` |
| `COGS-` | Cost of Goods Sold | `COGS-Steel COGS`, `COGS-Total COGS` |
| `GP-` | Gross Profit | `GP-Steel GP`, `GP-Gross Profit` |
| `OPEX-` | Operating Expenses | `OPEX-Employee Benefits`, `OPEX-Total OpEx` |
| `EBITDA-` | EBITDA | `EBITDA-Steel EBITDA`, `EBITDA-Underlying EBITDA` |
| `Stat-` | Statutory adjustments | `Stat-SBP`, `Stat-Significant Items`, `Stat-Statutory EBITDA` |
| `DA-` | Depreciation & Amort | `DA-Depreciation PPE`, `DA-ROU Amortisation`, `DA-Total DA` |
| `EBIT-` | EBIT | `EBIT-Underlying EBIT` |
| `Int-` | Interest | `Int-Interest Income`, `Int-Lease Interest`, `Int-Bank Interest`, `Int-Net Finance Costs` |
| `PBT-` | PBT | `PBT-PBT` |
| `Tax-` | Tax | `Tax-Tax Expense` |
| `NPAT-` | Net Profit | `NPAT-NCI`, `NPAT-Underlying NPAT`, `NPAT-Sig Items AT`, `NPAT-Statutory NPAT` |
| `EPS-` | EPS & Shares | `EPS-YE Shares`, `EPS-WASO Basic`, `EPS-Underlying EPS` |
| `Div-` | Dividends | `Div-DPS`, `Div-Total Dividends` |
| `KPI-` | Operating Metrics | `KPI-Steel Volume`, `KPI-Steel Rev/t` |
| `BS-` | Balance Sheet | `BS-Cash`, `BS-Trade Receivables`, `BS-PPE` |
| `CF-` | Cash Flow | `CF-EBITDA`, `CF-WC Change`, `CF-Net OCF` |

**Rows without keys**: Analytical/ratio rows (growth %, margins, rates), subtotals that don't need cross-sheet lookup (Total Assets, Total Liabilities, Total Equity, Working Capital, etc.), and section headers have no Column A key.

When building a new model, update the key names to match the new company's line items (e.g. `Rev-Steel Revenue` → `Rev-Hardware Revenue`), keeping the prefix convention intact.

---

## Row Lookup System

Cross-sheet formulas use **INDEX/MATCH** on Column A keys and Row 3 period headers.

### Annual ← HY & Segments (forecast columns)

Each Annual forecast cell sums the two half-year columns for that year:

```excel
=INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("1H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))+INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("2H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))
```

- `$A7` = Column A key of the current row
- `G$1` = Calendar year integer in Row 1
- `RIGHT(G$1,2)` extracts the 2-digit year to build `"1H26"` / `"2H26"`
- The formula matches both half-year columns and sums them

This pattern applies to all P&L rows that have Column A keys. Rows without keys (growth rates, margins) use local formulas on the Annual sheet.

---

## Color Coding

### Cell Fill Colors

| Color | Hex | Usage |
|-------|-----|-------|
| Dark Navy Blue | `FF002060` | Row 2 zone label cells ("Actual" zone) |
| Medium Blue | `FF0070C0` | Row 2 zone label cells ("Forecast" zone, Annual only) |
| Light Blue | `FFC5D9F1` | Section header rows (P&L, Balance Sheet, Cash Flow, Operating FCF, ROIC) |
| Light Gray | `FFD9D9D9` | Operating Metrics header, Segment Forecast headers |

### Font Colors

| Color | Hex | Meaning |
|-------|-----|---------|
| Dark Red / Maroon | `FFC00000` | **Forecast assumption inputs** — values the user should change when forecasting |
| Blue | `FF0000CC` | **Hardcoded reported data** on HY & Segments sheet (1H actuals and first 1H of near-term forecast) |
| Red | `FFFF0000` | Valuation Date on Value sheet |
| White | `FFFFFFFF` | Text on dark-fill zone label cells |
| Default (theme:1) | — | Standard formula-driven or regular text cells |

### Color Rules
- Any cell with dark red font (`FFC00000`) is a user-editable assumption input
- Any cell with blue font (`FF0000CC`) contains hardcoded reported data
- No fill differentiation between actual and forecast data cells (both have no fill)
- Actual vs forecast zones are distinguished by Row 2 zone labels only

---

## Row Formatting Rules

### Bold Pattern
- **Section headers** (P&L, Revenue, COGS, etc.): Bold, no data in data columns, light blue or gray fill
- **Subtotals/Totals** (Total Revenue, Gross Profit, Underlying EBITDA, etc.): Bold with **thin top + thin bottom borders**
- **Component data rows** (individual revenue lines, cost lines): Not bold, no borders
- **Analytical/ratio rows** (Growth %, Margin %): Not bold, no borders
- **Per-unit metric rows** (Rev/tonne, GP/tonne, EPS): Bold, no borders

### Header Row Conventions
- **Row 1**: Calendar year integers, no formatting
- **Row 2**: Zone labels — dark navy fill with white font (Actual zone), medium blue fill with white font (Forecast zone on Annual)
- **Row 3**: Period headers (e.g. `FY23A`, `1H24`) — light blue fill (`FFC5D9F1`) on Forecast zone columns, dark navy fill on Actual zone columns
- **Row 4**: Period-end dates, same fill pattern as Row 3

### Border Convention
- **Thin top + thin bottom** = subtotal row (sums the rows above it)
- No double-bottom borders in the template
- Section headers have no borders

---

## Number Format Conventions

| Data Type | Format String | Example Rows |
|-----------|--------------|-------------|
| Monetary (main currency, millions) | `#,##0.0` | Revenue, COGS, GP, EBITDA, D&A, EBIT, Interest, PBT, NPAT, BS items, CF items |
| Percentage | `0.0%` | Growth rates, margins, tax rates, yields, payout ratio |
| Per-share | `0.000` | EPS, DPS, FCF per Share |
| Share count (millions) | `0.0` | Shares outstanding, WASO, Dilution |
| Volume (thousands of tonnes) | `#,##0.0` | Volume rows |
| Revenue/cost per unit (whole currency) | `#,##0` | Rev/tonne, GP/tonne, Opex/tonne |
| Integer count | `#,##0` | Customers, Headcount |
| Ratio (x) | `0.0\x` | ND/EBITDA, P/B (escaped `\x` suffix) |
| Years | `0.0` | Avg Lease Life |
| Debt funding % | `0%` | Debt funding % of CFI |
| Date | `dd\-mmm\-yy` | Valuation Date (Value sheet) |
| Share price | `0.00` | Current Share Price, Beta |
| Discount factor | `0.00` | DCF discount factors |

These formats are consistent across all sheets.

---

## Zone Label Positioning

Row 2 contains zone labels that distinguish Actual from Forecast periods.

**Rule**: The "Actual" label is placed at the **first actual data column** (Col D). The "Forecast" label is placed at the **first forecast data column**.

- **Annual**: D2 = `"Actual ---------->"`, G2 = `"Forecast ----->"` (Forecast starts at Col G)
- **HY & Segments**: D2 = `"Actual ---------->"`, K2 = `"Forecast ----->"` (Forecast starts at Col K for the first pure formula-driven 2H forecast; Col J = 1H near-term may contain hardcoded data)

When building a new model, position the zone labels at the first column of each zone after adjusting for the number of actual periods.

---

## Blank Row Convention

**One blank row** is used as a separator between each major section throughout the Annual and HY sheets. Pattern:

- One blank row between Revenue Growth and COGS
- One blank row between Total COGS and Gross Profit
- One blank row between GP Margin and Operating Expenses
- One blank row between OpEx Growth and EBITDA
- One blank row between EBITDA Margin and Statutory Adjustments
- One blank row between Statutory EBITDA and D&A
- One blank row between EBIT Margin and Interest
- One blank row between Bank Interest Rate and PBT
- One blank row between NPAT Margin and EPS
- One blank row between WASO Diluted and EPS values
- One blank row between EPS Growth and DPS
- One blank row between Dividend Growth and Operating Metrics
- One blank row between last KPI and Balance Sheet
- And so on for each section transition

**No blank rows within a section** — only between sections. Sub-headers (e.g. "Revenue", "COGS") appear immediately after the blank row.

---

## Cross-Sheet Structural Correspondence

The **Annual** and **HY & Segments** sheets share the same P&L row structure from Revenue through Statutory NPAT. They use **identical Column A keys** in the same order.

**Rule**: Any change to a P&L line item on one sheet (adding a row, renaming a key, changing the order) **must be reflected on the other sheet** to preserve the INDEX/MATCH linkage.

The HY sheet additionally contains:
- Operating Metrics rows (volumes, per-tonne) that mirror the Annual Operating Metrics section (same keys)
- Segment Forecast zones (rows 92–116) that exist **only on HY** — these have no Annual counterpart

---

## Sign Conventions

| Item | Sign |
|------|------|
| Revenue | Positive |
| COGS | **Negative** |
| Gross Profit | Positive (Revenue + negative COGS) |
| Operating Expenses | **Negative** |
| EBITDA | Positive |
| Significant Items (costs) | **Negative** |
| Share-based Payments | **Negative** (or zero) |
| Depreciation & Amortisation | **Negative** |
| EBIT | Positive |
| Interest Income | Positive |
| Interest Expense (Lease & Bank) | **Negative** |
| Tax Expense | **Negative** |
| NCI (deduction) | **Negative** |
| NPAT | Positive |
| BS Assets | Positive |
| BS Liabilities | Positive |
| CF — EBITDA | Positive |
| CF — Interest/Tax Paid | **Negative** |
| CF — Capex | **Negative** |
| CF — Dividends Paid | **Negative** |
| CF — Lease Principal | **Negative** |
| CF — Change in Debt | Positive = borrowing, Negative = repayment |
| Working Capital Change | Negative = cash outflow (increase in WC) |

---

## Cross-Sheet Formula Rules

### Annual ← HY & Segments
All Annual P&L forecast cells (rows with Column A keys) use the INDEX/MATCH pattern documented in the Row Lookup System section. This sums the 1H + 2H values from the HY sheet.

Annual actual columns (Cols D–F) are hardcoded — they do not reference the HY sheet.

### HY Consolidated ← HY Segment Forecast
The HY Consolidated P&L zone (rows 5–77) pulls from the Segment Forecast zone (rows 92–116) for 2H+ forecast columns. Representative formulas:
```
Steel Revenue (K7) = K97
Steel COGS (K13) = -K7*(1-K98)
Corporate EBITDA (K40) = -K9*K41
Total OpEx (K33) = K42-K22   [backed out: EBITDA - GP]
Individual Opex (K29) = IF(I33=0,0,K33*I29/I33)   [proportional to prior period]
Total D&A (K54) = -K9*K55   [D&A as % of revenue]
Depr PPE split (K52) = IF(I54=0,0,K54*I52/I54)   [proportional]
ROU Amort (K53) = K54-K52   [residual]
```

### HY Interest ← Annual BS
HY interest formulas reference the Annual Balance Sheet for start/end period balances:
```
Interest Income (K63) = INDEX(Annual!$110:$110,MATCH(K$1-1,Annual!$1:$1,0))*K67-J63
```
Gets prior-year-end Cash from Annual BS row 110, multiplies by interest rate, subtracts 1H to get 2H.

```
Bank Interest (K65) = -((INDEX(Annual!$128:$128,MATCH(K$1-1,Annual!$1:$1,0))+INDEX(Annual!$128:$128,MATCH(K$1,Annual!$1:$1,0)))/2)*K69-J65
```
Average of start/end year Total Banking Debt from Annual BS row 128, multiplied by bank rate, minus 1H.

### Value ← Annual
DCF and SOTP formulas reference Annual sheet rows for EBITDA, D&A, EBIT, tax, capex, WC change, net debt, lease liabilities, shares. References use direct cell addresses (e.g. `Annual!H35`).

---

## Sub-Period Derivation

Half-year periods on the HY & Segments sheet are **independently forecast** in the Segment Forecast zone (not derived from annual). The Annual sheet then **aggregates** them.

### 1H Forecast Convention
The first 1H of forecast (Col J on HY) typically contains **hardcoded values** (blue font `FF0000CC`) representing the most recently reported or known half. As more actuals become available, this column transitions from forecast to actual.

### 2H Forecast Formulas (Segment zone)
Each 2H+ forecast column uses growth-on-prior-corresponding-period logic:

```excel
Volume (K93) = I93*(1+K94)       [prior 2H volume * (1 + volume growth %)]
Rev/Tonne (K95) = I95*(1+K96)    [prior 2H rev/t * (1 + price growth %)]
Revenue (K97) = K93*K95/1000     [volume * price / 1000]
GP (K99) = K97*K98               [revenue * GP margin %]
Opex (K100) = I100*(1+K101)      [prior 2H opex * (1 + opex growth %)]
Segment EBITDA (K103) = K99+K100 [GP + Opex, where opex is negative]
```

Growth assumptions reference the **prior corresponding period** (2H references prior 2H, 1H references prior 1H), not the immediately prior half.

---

## Assumption Input Placement

The template enforces **full assumption visibility**: ALL forecast assumptions are surfaced on dedicated rows. No assumptions are embedded within formulas.

### Convention
- Each assumption row sits **adjacent to the row it drives** (typically immediately below or within the same section)
- The driven row references the assumption row by direct cell reference
- Assumption rows are identified by **dark red font** (`FFC00000`)
- No forecast formula contains a hardcoded assumption value — every driver is a cell reference

### Assumption Rows on Annual Sheet

| Row | Label | What it drives | Default pattern |
|-----|-------|---------------|-----------------|
| 52 | Avg Lease Life | Lease Principal Payments (R171) | Flat-lined: `=F52` |
| 64 | Lease Interest Rate | Lease Interest (via HY) | Flat-lined: `=F64` |
| 79 | YE Basic Shares Outstanding | WASO, EPS, per-share metrics | Hardcoded input |
| 90 | Payout Ratio | DPS (R88) | Flat-lined: `=F90` |
| 118 | Receivables / Revenue | Trade Receivables (R111) | Flat-lined: `=F118` |
| 119 | Inventory / Revenue | Inventories (R112) | Flat-lined: `=F119` |
| 121 | Payables / Revenue | Trade Payables (R125) | Flat-lined: `=F121` |
| 122 | New Lease Additions | ROU Assets (R115), Lease Liabilities (R127) | Flat-lined: `=F122` |
| 161 | Capex / Sales | Capex PPE (R160) | Flat-lined: `=F161` |
| 162 | Capex (Intangibles) | Intangibles BS (R114) | Input: `=0` |
| 163 | Acquisitions | Intangibles BS (R114) | Hardcoded input |
| 164 | Asset Sales | CFI | Input: `=0` |
| 165 | Other CFI | CFI | Input: `=0` |
| 170 | Share Issues / Buybacks | Issued Capital (R137) | Hardcoded input |
| 173 | Other CFF | CFF | Input: `=0` |
| 175 | Debt funding % of CFI | Change in Debt (R172) | Hardcoded input (e.g. 30%) |

### Assumption Rows on HY & Segments Sheet (Segment Forecast zone)

Per segment (Steel = rows 93–103, Metals = rows 106–116):

| Row offset | Label | Font color | Default pattern |
|------------|-------|-----------|-----------------|
| +1 | Volume Growth | Dark red (`FFC00000`) | Hardcoded % |
| +3 | Rev/Tonne Growth | Dark red | Hardcoded % |
| +5 | GP Margin | Dark red | Flat-lined: `=I__` |
| +8 | Opex Growth | Dark red | Hardcoded % |

Additional HY assumptions in the Consolidated P&L zone:
- Corp EBITDA / Revenue (flat-lined)
- D&A / Revenue (flat-lined)
- Interest rates (Income, Lease, Bank — flat-lined)
- Tax Rate (flat-lined)

---

## Flow vs Point-in-Time

- **P&L items** are **flow** (period totals). Annual = sum of 1H + 2H.
- **Balance Sheet items** are **point-in-time** (period-end balances). Annual BS uses year-end values, not sums of halves.
- **Cash Flow items** are **flow** (period totals). Annual = sum of 1H + 2H equivalent.

The Annual sheet aggregates P&L from HY via the 1H+2H INDEX/MATCH formula. BS is calculated directly on the Annual sheet using roll-forward formulas. CF is calculated on the Annual sheet using links to P&L and BS.

---

## Cash Flow Section Structure

**Format**: EBITDA-based (indirect method starting from Underlying EBITDA)

### CFO (rows 147–157)

| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 148 | `CF-EBITDA` | Underlying EBITDA | Component | `=G38` (links to P&L EBITDA) |
| 149 | `CF-WC Change` | Working Capital Change | Component | `=-(G111-F111)-(G112-F112)+(G125-F125)` |
| 150 | `CF-Significant Items` | Significant Items/Non-Cash Items | Component | `=G44` (links to P&L sig items) |
| 151 | *(none)* | Gross Operating Cash Flow | **Subtotal** | `=SUM(G148:G150)` |
| 152 | `CF-Int Received` | Interest Received | Component | `=G59` (links to P&L) |
| 153 | `CF-Interest Paid` | Interest Paid | Component | `=G61` (links to P&L) |
| 154 | `CF-Lease Int Paid` | Lease Interest Paid | Component | `=G60` (links to P&L) |
| 155 | `CF-Tax Paid` | Tax Paid | Component | `=G69+(G73-G44)` (tax expense + sig items AT - sig items) |
| 156 | `CF-Net OCF` | Net Operating Cash Flow | **Subtotal** | `=G151+G152+G153+G154+G155` |
| 157 | *(none)* | OCF Growth | Analytical | `=IF(F156=0,"",G156/F156-1)` |

**WC Change formula breakdown**: `=-(ΔReceivables)-(ΔInventories)+(ΔPayables)` — increase in receivables/inventories = cash outflow (negative); increase in payables = cash inflow (positive).

**Tax Paid formula breakdown**: Tax expense + (Significant Items After Tax − Significant Items) adjusts for cash tax on significant items.

### CFI (rows 159–166)

| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 160 | `CF-Capex PPE` | Capex (PPE) | Component | `=G161*G9` (capex/sales × revenue; result is negative) |
| 161 | *(none)* | Capex / Sales | **Assumption** | Flat-lined (dark red) |
| 162 | `CF-Capex Intang` | Capex (Intangibles) | Input | `=0` (dark red) |
| 163 | `CF-Acquisitions` | Acquisitions | Input | Hardcoded (dark red) |
| 164 | `CF-Asset Sales` | Asset Sales | Input | `=0` (dark red) |
| 165 | `CF-Other CFI` | Other | Input | `=0` (dark red) |
| 166 | *(none)* | Total Investing Cash Flow | **Subtotal** | `=SUM(G160,G162:G165)` |

**Note**: Capex/Sales is a negative ratio (e.g. -3.5%), so `Capex/Sales × Revenue` produces a negative capex value.

### CFF (rows 168–175)

| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 169 | `CF-Dividends` | Dividends Paid | Component | `=-G89` (negative of total dividends from P&L) |
| 170 | `CF-Share Issues` | Share Issues / Buybacks | Input | Hardcoded (dark red) |
| 171 | `CF-Lease Principal` | Lease Principal Payments | Component | `=-F127/G52` (prior lease liability ÷ avg lease life, negative) |
| 172 | `CF-Debt Change` | Change in Debt | Component | `=-G175*G166` (debt funding % × total investing CF) |
| 173 | `CF-Other CFF` | Other | Input | `=0` (dark red) |
| 174 | *(none)* | Total Financing Cash Flow | **Subtotal** | `=SUM(G169:G173)` |
| 175 | *(none)* | Debt funding % of CFI | **Assumption** | Hardcoded (dark red, e.g. 30%) |

**Debt Change formula**: `=-Debt_Funding_% × Total_Investing_CF`. Since Total Investing CF is negative (cash outflow), the double negative produces a positive number (borrowing to fund investment).

### Net Cash (row 177)
```
Net Change in Cash = Net OCF + Total Investing CF + Total Financing CF
```

---

## Balance Sheet Projection Methods

### Assets

| Row | Label | Method | Formula |
|-----|-------|--------|---------|
| 110 | Cash | CF-linked | `=F110+G177` (prior + net change in cash) |
| 111 | Trade Receivables | Revenue ratio | `=G9*G118` (revenue × receivables/revenue %) |
| 112 | Inventories | Revenue ratio | `=G9*G119` (revenue × inventory/revenue %) |
| 113 | PPE | Roll-forward | `=F113-G160+G48` (prior − capex + depreciation) |
| 114 | Intangibles | Roll-forward | `=F114-G162-G163` (prior − intang capex − acquisitions) |
| 115 | ROU Assets | Roll-forward | `=F115+G122+G49` (prior + new leases + ROU amort) |
| 116 | Other Assets | Flat | `=F116` |
| 117 | Total Assets | Sum | `=SUM(G110:G116)` |

### Liabilities

| Row | Label | Method | Formula |
|-----|-------|--------|---------|
| 125 | Trade Payables | Revenue ratio | `=G9*G121` (revenue × payables/revenue %) |
| 126 | Other Liabilities | Flat | `=F126` |
| 127 | Lease Liabilities | Roll-forward | `=F127+G122+G171` (prior + new leases + principal payments) |
| 128 | Total Banking Debt | Roll-forward | `=F128+G172` (prior + change in debt) |
| 129 | Total Liabilities | Sum | `=SUM(G125:G128)` |

### Equity

| Row | Label | Method | Formula |
|-----|-------|--------|---------|
| 137 | Issued Capital | Roll-forward | `=F137+G170` (prior + share issues) |
| 138 | Retained Profits | Roll-forward | `=F138+G74-G89` (prior + statutory NPAT − total dividends) |
| 139 | Reserves | Flat | `=F139` |
| 140 | Minorities | Roll-forward | `=F140-G71` (prior − NCI) |
| 141 | Total Equity | Sum | `=SUM(G137:G140)` |

### BS Check
Row 144: `=G117-G129-G141` — must equal zero (Total Assets − Total Liabilities − Total Equity).

---

## Return Metrics (rows 189–194)

| Row | Label | Formula |
|-----|-------|---------|
| 190 | Invested Capital | `=G141+G131` (Total Equity + Net Banking Debt) |
| 191 | Underlying EBIT | `=G54` |
| 192 | ROFE (Return on Funds Employed) | `=IF(G190=0,"",G191/G190)` |
| 193 | NOPAT | `=G191*(1-G70)` (EBIT × (1 − tax rate)) |
| 194 | ROIC | `=IF(G190=0,"",G193/G190)` |

---

## Valuation Methods

### Method 1: DCF (Value sheet, rows 2–49)

**Location**: Value sheet, B2:M49

**Structure**:

1. **Market Data Inputs** (rows 4–9):
   - C4: Current Share Price (dark red = input)
   - C5: Shares Outstanding (from Annual)
   - C6: Market Cap (`=C4*C5`)
   - C7: Net Debt (from Annual BS)
   - C8: Market EV (`=C6+C7`)
   - C9: Valuation Date (red font = input)

2. **WACC Calculation** (rows 11–22):
   - C12: Risk-free Rate (input)
   - C13: Equity Risk Premium (input)
   - C14: Beta (input)
   - C15: Cost of Equity (`=C12+C14*C13`)
   - C16: Pre-tax Cost of Debt (input)
   - C17: Tax Rate (input)
   - C18: After-tax Cost of Debt (`=C16*(1-C17)`)
   - C19: Debt Weight (input)
   - C20: WACC (`=C15*(1-C19)+C18*C19`)
   - C21: Terminal Growth Rate (input)
   - C22: Stub period adjustment (fraction of year from valuation date to first forecast year-end)

3. **FCF Projection** (rows 24–35, Cols D–M):
   - Row 25: EBITDA (from Annual)
   - Row 26: D&A (from Annual)
   - Row 27: EBIT
   - Row 28: Tax on EBIT (`=EBIT × tax rate`)
   - Row 29: NOPAT
   - Row 30: Add back D&A
   - Row 31: Capex (from Annual CF)
   - Row 32: WC Change (from Annual CF)
   - Row 33: FCFF (`=NOPAT + D&A + Capex + WC Change`)
   - Row 34: Normalised FCFF (terminal year: capex = D&A)
   - Row 35: Terminal Value (`=M34*(1+$C$21)/($C$20-$C$21)`) — Gordon Growth Model

4. **Discounting** (rows 37–42):
   - Row 37: Discount factors: `=1/(1+$C$20)^($C$22+N)` where N=0..9, with stub period
   - Row 38: PV of each year's FCFF
   - Row 39: Sum of PV of FCFFs
   - Row 40: PV of Terminal Value
   - Row 41: Enterprise Value (`=R39+R40`)

5. **Equity Bridge** (rows 43–49):
   - Row 43: Less Net Debt
   - Row 44: Less Lease Liabilities
   - Row 47: Equity Value (`=EV − Net Debt − Leases`)
   - Row 48: Shares
   - Row 49: **DCF Value Per Share** (`=Equity Value / Shares`)

**All WACC inputs** use dark red font = user-editable.

### Method 2: EV/EBITDA Sum-of-the-Parts (Value sheet, rows 52–66)

**Location**: Value sheet, B52:E66

**Structure**:
1. **Segment EBITDA** (rows 55–57): Pulled from Annual forecast year (e.g. FY27E)
   - `C55 = Annual!H35` (Segment 1 EBITDA)
   - `C56 = Annual!H36` (Segment 2 EBITDA)
   - `C57 = Annual!H37` (Corporate EBITDA)

2. **Multiples** (Col D, dark red = input):
   - D55: Segment 1 multiple (e.g. 8x)
   - D56: Segment 2 multiple (e.g. 7x)
   - D57: Corporate multiple — implied weighted average: `=IF((C55+C56)=0,"",(C55*D55+C56*D56)/(C55+C56))`

3. **Implied EV** (Col E): `=EBITDA × Multiple` per segment
   - E59: Group EV = `SUM(E55:E57)`

4. **Equity Bridge** (rows 60–66):
   - Less Net Debt, Less Lease Liabilities
   - Equity Value / Shares = **SOTP Value Per Share**
   - E66: Implied Group EV/EBITDA = `IF(Annual!H38=0,"",E59/Annual!H38)`

**When repurposing**: Update segment names and the number of segments. The Corporate/overhead segment is valued using a weighted average of the operating segment multiples.

---

## Line Item Retention Policy

When repurposing this template, each row is classified as either **RETAIN** (keep the row, its formulas, and its formatting exactly as-is) or **REPLACE** (delete the company-specific content and rebuild for the new company).

### RETAIN rows

These rows are structural — their formulas and layout are preserved across all models built from this template:

**Annual — Statutory Adjustments (rows 42–45)**:
SBP, Significant Items, Statutory EBITDA

**Annual — D&A (rows 47–52)**:
Depreciation PPE, ROU Amortisation, Total D&A, D&A/Revenue, Avg Lease Life

**Annual — Interest (rows 58–65)**:
Interest Income, Lease Interest, Bank Interest, Net Finance Costs, Interest Income Rate, Lease Interest Rate, Bank Interest Rate

**Annual — PBT/Tax/NPAT (rows 67–76)**:
PBT, Tax Expense, Tax Rate, NCI, Underlying NPAT, Sig Items AT, Statutory NPAT, NPAT Growth, NPAT Margin

**Annual — EPS & Dividends (rows 78–92)**:
YE Shares, WASO Basic, Dilution, WASO Diluted, Underlying EPS, Statutory EPS, EPS Growth, DPS, Total Dividends, Payout Ratio, Dividend Yield, Dividend Growth

**Annual — BS Analytics (rows 118–122, 131–134, 142–144)**:
Receivables/Revenue, Inventory/Revenue, Working Capital, Payables/Revenue, New Lease Additions, Net Banking Debt, Adj Net Debt, ND/EBITDA, Gearing, ROE, P/B, BS Check

**Annual — Full Cash Flow (rows 146–177)**:
All CFO, CFI, CFF rows as documented above

**Annual — Operating FCF (rows 179–186)**:
Net OCF, Net Capex, Lease Principal, Operating FCF, FCF per Share, FCF Yield, FCF Margin

**Annual — ROIC (rows 189–194)**:
Invested Capital, EBIT, ROFE, NOPAT, ROIC

**HY — Below-EBITDA structure (rows 42–90)**:
Statutory adjustments, D&A, Interest, Tax/NPAT, EPS & Dividends (same structure as Annual, at half-yearly granularity)

**Value — Full sheet (rows 2–66)**:
DCF and SOTP structures retained

### REPLACE rows

These rows are company-specific and must be rebuilt:

**Annual & HY — Revenue through EBITDA (rows 5–41)**:
All segment revenue lines, COGS lines, GP lines, OpEx lines, EBITDA by segment. Replace segment names, add/remove segments as needed.

**Annual — Operating Metrics (rows 94–106)**:
All KPI rows. Replace with new company's operating metrics.

**Annual — BS component rows (rows 109–117, 124–129, 136–141)**:
Asset, Liability, and Equity line items. Adjust to match new company's BS structure. RETAIN rows (analytics, ratios) sit below/within these and should not be moved.

**HY — Segment Forecast zones (rows 92–116)**:
Replace segment names, number of segments, and forecast driver structure for the new company.

**Inputs sheet**: Cleared entirely. Do not create linkages.

**Charts sheet**: Cleared entirely. Leave blank.

---

## Template Preservation Method

When repurposing this template for a new company:

1. **Copy** the template to the new company's `Models/` folder — never modify the original template
2. **Modify in place** — do not rebuild sheets from scratch. Insert new rows where needed, delete company-specific rows that don't apply, and overwrite cell values. This preserves formatting, formulas, and structural integrity.
3. **RETAIN rows** must not be deleted, reordered, or have their formulas altered (except updating cell references when rows are inserted/deleted above them)
4. **REPLACE rows** should have their labels, Column A keys, and data values replaced for the new company. The structural pattern (section headers, blank rows, subtotals) should be preserved.

---

## Repurposing Checklist

When building a new model from this template:

- [ ] Copy template to `[TICKER]/Models/[Ticker] Model.xlsx`
- [ ] Clear the Inputs sheet (delete all data and formulas)
- [ ] Clear the Charts sheet (delete all charts and clear all cells)
- [ ] Update Row 2 zone labels: replace company name, adjust Actual/Forecast column positions
- [ ] Update Row 3 period headers for new company's reporting periods and fiscal year-end
- [ ] Update Row 4 dates for new company's fiscal year-end dates
- [ ] Update Row 1 calendar year integers
- [ ] Replace all segment-specific Revenue, COGS, GP, EBITDA rows (insert/delete as needed)
- [ ] Update Column A keys for all replaced rows (maintain `Prefix-Label` convention)
- [ ] Ensure Annual and HY sheets have matching keys for all P&L rows
- [ ] Replace Operating Metrics with new company's KPIs
- [ ] Adjust BS line items for new company's balance sheet structure
- [ ] Replace HY Segment Forecast zones for new company's segments
- [ ] Verify INDEX/MATCH range covers all new columns (expand `$A:$AG` if needed)
- [ ] Update Value sheet: segment names in SOTP, verify Annual references point to correct rows
- [ ] Update currency labels in Col C if not the template default
- [ ] Populate actual data from source documents
- [ ] Set forecast assumptions
- [ ] Verify BS Check = 0 for all forecast periods
- [ ] Verify Net Change in Cash = Cash movement on BS for all forecast periods

---

## Quality Gates

Before considering a model complete, verify:

1. **BS Check**: Row 144 = 0 for every period (Total Assets − Total Liabilities − Total Equity)
2. **Cash Reconciliation**: Net Change in Cash (R177) = Cash (R110) − Prior Cash for every forecast period
3. **Annual ↔ HY Consistency**: For every P&L row with a Column A key, Annual forecast = 1H + 2H from HY sheet
4. **Sign Conventions**: All costs/expenses negative, all income positive, BS items positive
5. **No Circular References**: The model must not contain circular references
6. **Formatting Consistency**: All new rows follow the same bold/border/number format/font color conventions
7. **Column A Keys**: Every new component row has a unique key; keys match between Annual and HY
8. **Assumption Visibility**: No hardcoded assumptions embedded in formulas — every driver on its own row with dark red font
9. **Zone Labels**: Actual/Forecast zone labels correctly positioned at the first column of each zone

---

## Critical Error Checklist

Common formula errors to watch for when building from this template:

1. **INDEX/MATCH range too narrow**: If columns are added beyond the original range (`$A:$AG`), the INDEX/MATCH formulas on the Annual sheet will fail silently. Verify the range covers all HY columns.
2. **Sign errors in CF**: Capex should be negative, WC change sign depends on direction. Double-check that `CF-Capex PPE = Capex/Sales × Revenue` produces a negative number (Capex/Sales ratio should be negative).
3. **Lease Principal formula**: `=-F127/G52` divides prior lease liability by average lease life. If Avg Lease Life is zero or blank, this produces a #DIV/0! error.
4. **Debt Change formula**: `=-G175*G166` relies on Total Investing CF being negative. If Total Investing CF is positive (unusual), debt change will be negative (repayment), which may not be intended.
5. **Tax Paid formula**: `=G69+(G73-G44)` assumes significant items flow through both pre-tax (R44) and after-tax (R73). If either row is removed, adjust this formula.
6. **HY Interest references Annual BS rows by absolute row number** (e.g. `Annual!$110:$110` for Cash, `Annual!$128:$128` for Debt). If BS rows are inserted/deleted, these references must be updated.
7. **Retained Profits roll-forward** uses Statutory NPAT (R74), not Underlying NPAT. Ensure statutory adjustments flow correctly.
8. **SOTP references specific Annual rows** (e.g. `Annual!H35` for segment EBITDA). After adding/removing P&L rows, verify these still point to the correct rows.
