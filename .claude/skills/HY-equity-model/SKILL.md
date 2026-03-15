---
name: HY Equity Model Template
description: >
  Skill file for a half-year equity research model template with 3-sheet cascade
  (HY & Segments → Annual → Value). Use this skill when building models for companies
  that report on a half-year (interim) cycle. Contains two valuation methods (DCF and
  EV/EBITDA SOTP), EBITDA-based cash flow, full BS roll-forward, and ROIC metrics.
---

## Template Reference

- File: `.claude/templates/HY_model_template.xlsx`
- Reporting frequency: Half-year (1H / 2H)
- Forecast horizon: 10 years (FY26E–FY35E in template)
- Actuals columns: 3 full years in template

## Sheet Architecture

Three sheets in this order:

1. **Value** — DCF valuation and EV/EBITDA SOTP. Pulls all data from Annual via INDEX/MATCH.
2. **Annual** — Full-year P&L, BS, CF, OFCF, ROIC. Actuals are hardcoded (blue). Forecast flow items pull from HY & Segments (1H + 2H) via INDEX/MATCH. Forecast BS items use roll-forward formulas on this sheet.
3. **HY & Segments** — Half-year P&L and operating metrics with two spatial zones. Contains the primary forecast driver logic.

Data flows: HY & Segments (source) → Annual (aggregation) → Value (valuation)

## Sheet Zone Architecture

**HY & Segments** has two distinct zones:

- **Zone 1 (rows 5–90): P&L Summary** — Half-year P&L from Revenue through NPAT, plus operating metrics (KPIs). For actual periods, 1H values are hardcoded; 2H values are derived from Annual FY minus 1H. For forecast periods, Zone 1's segment rows (Revenue, COGS, GP, EBITDA) reference Zone 2 outputs. Group-level rows (Total Revenue, Total COGS, Gross Profit, Total OpEx, Underlying EBITDA, EBIT, PBT, NPAT) contain their own calculation formulas (sums, cascades).

- **Zone 2 (rows 92–116): Segment Forecast Drivers** — Contains the primary forecast logic for each segment. Structure per segment: Volume → Volume Growth → Revenue/Unit → Rev/Unit Growth → Revenue → GP Margin → Gross Profit → Opex → Opex Growth → Opex per Unit → Segment EBITDA. For actual periods, Zone 2 rows reference Zone 1 (e.g. `=F81` for Volume, `=F7` for Revenue). For forecast periods, Zone 2 contains independent driver logic (e.g. `=I93*(1+K94)` for Volume, `=K93*K95/1000` for Revenue).

**Dependency direction:** Zone 2 is the **source zone** for forecast periods. Zone 1's forecast segment cells must only contain references to Zone 2 outputs (e.g. `=K97` for Revenue, `=K103` for EBITDA), never independent forecast logic. Zone 1's group-level subtotal and ratio rows retain their own calculation formulas.

Representative formulas:
- Zone 1 Revenue forecast: `=K97` (references Zone 2 Steel Revenue output)
- Zone 1 EBITDA forecast: `=K103` (references Zone 2 Steel Segment EBITDA)
- Zone 2 Revenue forecast: `=K93*K95/1000` (Volume × Rev/Tonne ÷ 1000)
- Zone 2 GP forecast: `=K97*K98` (Revenue × GP Margin)
- Zone 2 EBITDA forecast: `=K99+K100` (Gross Profit + Opex)

**Annual** and **Value** sheets are single-zone — no spatial zones.

## Column Layout

### Annual
| Column | Purpose |
|--------|---------|
| A | Row keys (PREFIX-Label format) |
| B | Row labels |
| C | Units (e.g. NZDm, %, NZDps, 000t) |
| D–F | Actuals (3 years) |
| G–P | Forecasts (10 years) |

- Row 1: Calendar year numbers (e.g. 2023, 2024, ...)
- Row 2: Zone labels — Col D: "Actual ---------->" (dark blue fill FF002060, white text), Col G: "Forecast ----->" (medium blue fill FF0070C0, white text)
- Row 3: Period labels (e.g. FY23A, FY24A, FY26E). Col B: Company name, Col C: "Units"
- Row 4: Period end dates (e.g. 2024-06-30)

### HY & Segments
| Column | Purpose |
|--------|---------|
| A | Row keys (same PREFIX-Label format as Annual) — present in Zone 1 only |
| B | Row labels |
| C | Units |
| D–I | Actuals (3 years × 2 halves = 6 columns) |
| J–AC | Forecasts (10 years × 2 halves = 20 columns) |

- Row 1: Calendar year numbers (pairs — e.g. 2023, 2023, 2024, 2024)
- Row 2: Zone labels — same fill colors as Annual
- Row 3: Period labels (e.g. 1H23, 2H23, 1H24, 2H24)
- Row 4: Period end dates

Columns are paired: each fiscal year has two adjacent columns (1H, 2H). 1H columns are odd-positioned (D, F, H, J, ...), 2H columns are even-positioned (E, G, I, K, ...).

### Value
| Column | Purpose |
|--------|---------|
| B | Row labels |
| C | Valuation inputs and summary outputs |
| D–M | FCF projection years (10 years, aligned with Annual forecast columns) |

- Row 2: "DCF Valuation" section header (bold, light blue fill FFC5D9F1)
- Row 24: "FCF Projection" header with period labels in D–M (e.g. FY26E, FY27E, ...)
- Row 52: "EV/EBITDA SOTP" section header (bold)

## Row Identification System

Both Annual and HY & Segments (Zone 1) use a Column A key system. Keys follow the format `PREFIX-Label` where:

- `Rev-` — Revenue lines
- `COGS-` — Cost of goods sold lines
- `GP-` — Gross profit lines
- `OPEX-` — Operating expense lines
- `EBITDA-` — EBITDA lines
- `Stat-` — Statutory adjustment lines
- `DA-` — Depreciation & amortisation lines
- `EBIT-` — EBIT lines
- `Int-` — Interest lines
- `PBT-` — Profit before tax
- `Tax-` — Tax lines
- `NPAT-` — Net profit lines
- `EPS-` — Earnings per share and share count lines
- `Div-` — Dividend lines
- `KPI-` — Operating metric lines
- `BS-` — Balance sheet lines
- `CF-` — Cash flow lines

**Purpose:** Keys provide durable identifiers for (1) INDEX/MATCH cross-sheet lookups and (2) data continuity across company restatements. When a company restates prior year figures, the key system ensures values are matched by semantic identity rather than row position, preventing data corruption.

**Keys are company-specific** — when repurposing the template for a new company, all keys must be rebuilt based on the new company's reported line items, using the same prefix conventions.

Zone 2 on HY & Segments does NOT have Column A keys — its rows are identified by position relative to the segment header.

## Row Lookup System

Cross-sheet formulas use INDEX/MATCH with Column A keys.

**Annual → HY & Segments (forecast flow items):**
```
=INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("1H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))+INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("2H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))
```
This sums 1H + 2H by: matching the Column A key on HY & Segments, then matching the period label constructed from "1H"/"2H" + last two digits of the year.

**HY & Segments → Annual (2H derivation for actuals):**
```
=INDEX(Annual!$A:$R,MATCH($A7,Annual!$A:$A,0),MATCH(E$1,Annual!$A$1:$R$1,0))-D7
```
This computes 2H = FY Annual value (looked up by key and year) minus 1H value.

**HY & Segments → Annual (interest rate computation for actuals):**
```
=IFERROR(IF(INDEX(Annual!$110:$110,MATCH(F$1-1,Annual!$1:$1,0))=0,"",INDEX(Annual!$59:$59,MATCH(F$1,Annual!$1:$1,0))/INDEX(Annual!$110:$110,MATCH(F$1-1,Annual!$1:$1,0))),"")
```
This uses direct row references (e.g. `Annual!$110:$110` for Cash row, `Annual!$59:$59` for Interest Income row) with year matching via row 1.

**Value → Annual:**
```
=INDEX(Annual!$D:$P,MATCH("EBITDA-Underlying EBITDA",Annual!$A:$A,0),MATCH(D$24,Annual!$D$3:$P$3,0))
```
Matches by literal key string and period label.

## Color Coding

| Color | RGB Code | Meaning | Used Where |
|-------|----------|---------|------------|
| Blue | FF0000CC | Hardcoded actual values | Data cells in actual periods (Annual, HY Zone 1) |
| Maroon/Dark Red | FFC00000 | Forecast assumption inputs | Zone 2 assumption rows (Volume Growth, Rev/Tonne Growth, GP Margin, Opex Growth), HY interest/tax rates, Annual Avg Lease Life, Value WACC inputs |
| Red | FFFF0000 | Special input (date) | Value sheet Valuation Date |
| Light blue fill | FFC5D9F1 | Section header | P&L, Balance Sheet, Cash Flow, Operating Free Cash Flow, ROIC section headers (Annual); DCF Valuation header (Value) |
| Grey fill | FFD9D9D9 | Section header (secondary) | Operating Metrics header (Annual) |
| Dark blue fill | FF002060 | Actual zone label | Row 2 "Actual ---------->" with white text |
| Medium blue fill | FF0070C0 | Forecast zone label | Row 2 "Forecast ----->" with white text |
| Black (default) | — | Calculated values | Formula-driven cells in forecast periods |

## Row Formatting Rules

**Bold patterns:**
- Section headers (P&L, Balance Sheet, Cash Flow, etc.): Bold, sometimes with fill color
- Sub-section headers (Revenue, COGS, Gross Profit, EBITDA, etc.): Bold, no fill
- Subtotal rows (Total Revenue, Total COGS, Gross Profit, Underlying EBITDA, etc.): Bold with top and bottom thin borders
- Calculated ratio/analytical rows in some sections: Bold (e.g. Revenue per Employee, EBITDA per Employee, KPI derived rows like Rev/t, GP/t)
- Component/input rows: Not bold

**Border patterns:**
- Subtotal rows have thin top and thin bottom borders on all cells
- No other rows have borders

**Header row conventions:**
- Row 2: Zone labels ("Actual ---------->" and "Forecast ----->") with colored fills
- Row 3: Period labels (FY23A, 1H24, etc.) — plain formatting, Col B has company name
- Row 4: Period end dates — not displayed prominently

## Sign Conventions

All values follow natural sign convention:
- **Revenue, income, assets:** Positive
- **Costs, expenses, D&A, interest expense, tax expense:** Negative
- **Liabilities (BS):** Positive (Trade Payables, Debt, Lease Liabilities stored as positive values)
- **Cash outflows (CF):** Negative (Capex, Dividends Paid, Interest Paid, Tax Paid, Lease Principal)
- **Cash inflows (CF):** Positive (Net OCF, Interest Received, Asset Sales)
- **NCI:** Shown as the P&L charge (negative when NCI takes profit, zero or positive when reversed)

Key implication: GP = Revenue + COGS (where COGS is negative), EBIT = EBITDA + D&A (where D&A is negative).

## Cross-Sheet Formula Rules

1. **Annual forecast flow items** = 1H + 2H from HY & Segments, looked up via INDEX/MATCH using Column A keys and constructed period labels ("1H"/"2H" + year suffix).
2. **Annual forecast BS/CF items** use roll-forward formulas directly on the Annual sheet — they do NOT pull from HY & Segments.
3. **HY & Segments 2H actuals** = Annual FY value (INDEX/MATCH by key and year) minus 1H value.
4. **HY & Segments forecast interest** uses direct row references to Annual BS rows for balance lookups (e.g. `Annual!$110:$110` for Cash, `Annual!$127:$127` for Lease Liabilities, `Annual!$128:$128` for Banking Debt).
5. **Value sheet** pulls from Annual only, via INDEX/MATCH using literal key strings and period labels.
6. No sheet references any other sheet by direct cell address for P&L items — all P&L cross-sheet references use Column A keys.

## Sub-Period Derivation

**Actuals — deriving 2H from Annual:**
For all flow items with Column A keys, 2H is derived as:
```
=INDEX(Annual!$A:$R,MATCH($A7,Annual!$A:$A,0),MATCH(E$1,Annual!$A$1:$R$1,0))-D7
```
Pattern: `FY value (from Annual, matched by key and year) - 1H value (adjacent column)`

This applies to: Revenue, COGS, OpEx, D&A, Interest, Tax, EBITDA (segment), Significant Items, NCI, Volume, and all other flow items with keys.

**Subtotal/calculated rows** derive 2H from their own component rows on the HY sheet (e.g. `=E7+E8` for Total Revenue, `=E42+E54` for EBIT).

**Forecasts — 2H derivation:**
2H forecast values are NOT derived from Annual. They are independently calculated on the HY & Segments sheet using the forecast driver logic. The Annual sheet then sums 1H + 2H to get the full year.

## Assumption Input Placement

The template enforces **full assumption visibility** — every forecast assumption is surfaced on its own dedicated row with maroon (FFC00000) font color. No forecast formula contains a hardcoded assumption value.

**Pattern:** The assumption row sits adjacent to (typically directly below) the row it drives. The driven row's forecast formula references the assumption row by cell reference.

Examples from Zone 2:
| Driven Row | Assumption Row | Forecast Formula |
|-----------|---------------|-----------------|
| Volume (row 93) | Volume Growth (row 94) | `=I93*(1+K94)` |
| Revenue/Tonne (row 95) | Rev/Tonne Growth (row 96) | `=I95*(1+K96)` |
| Gross Profit (row 99) | GP Margin (row 98) | `=K97*K98` |
| Opex (row 100) | Opex Growth (row 101) | `=I100*(1+K101)` |

Examples from Zone 1:
| Driven Row | Assumption Row | Forecast Formula |
|-----------|---------------|-----------------|
| Total D&A (row 54) | D&A/Revenue (row 55) | `=-K9*K55` |
| Corporate EBITDA (row 40) | Corp EBITDA/Revenue (row 41) | `=-K9*K41` |
| Tax Expense (row 72) | Tax Rate (row 73) | `=-K71*K73` |
| Interest Income (row 63) | Interest Income Rate (row 67) | `=prior_yr_cash_balance*K67-J63` |
| Bank Interest (row 65) | Bank Interest Rate (row 69) | `=-avg_debt_balance*K69-J65` |

For forecast periods, assumption rows are flatlined from the prior corresponding period (pcp):
- `=I67` (Interest Income Rate), `=I68` (Lease Interest Rate), `=I73` (Tax Rate), `=I41` (Corp EBITDA/Rev), `=I55` (D&A/Rev)

**Critical rule:** No forecast formula may contain a hardcoded assumption. E.g. `=PCP*1.025` is prohibited — must be `=PCP*(1+GrowthRow)` where GrowthRow is a dedicated maroon input cell.

## Flow vs Point-in-Time

The template distinguishes between:

**Flow items** (sum 1H + 2H for annual): Revenue, COGS, GP, OpEx, EBITDA, D&A, Interest, Tax, NPAT, Capex, Dividends, all CF items, Volume, Significant Items.

Annual forecast formula for flow items:
```
=INDEX('HY & Segments'!...,MATCH(key)...,MATCH("1H"&yr)) + INDEX('HY & Segments'!...,MATCH(key)...,MATCH("2H"&yr))
```

**Point-in-time items** (annual = year-end value, not sum): All BS items (Cash, Receivables, Inventories, PPE, Debt, Equity, etc.), Shares Outstanding, Headcount.

Annual forecast for point-in-time items uses roll-forward formulas on the Annual sheet directly — they do not reference HY & Segments.

## Cash Flow Section Structure

EBITDA-based format. Exact row order on Annual sheet:

### CFO (rows 147–158)
| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 148 | CF-EBITDA | Underlying EBITDA | Component | `=G38` (links to EBITDA subtotal) |
| 149 | CF-WC Change | Working Capital Change | Component | `=-(G111-F111)-(G112-F112)+(G125-F125)` (negative of change in Receivables + Inventory, plus change in Payables) |
| 150 | CF-Significant Items | Significant Items/Non-Cash Items | Component | `=G44` (links to Significant Items from P&L) |
| 151 | — | Gross Operating Cash Flow | **Subtotal** | `=SUM(G148:G150)` |
| 152 | CF-Int Received | Interest Received | Component | `=G59` (links to Interest Income from P&L) |
| 153 | CF-Interest Paid | Interest Paid | Component | `=G61` (links to Bank Interest Expense from P&L) |
| 154 | CF-Lease Int Paid | Lease Interest Paid | Component | `=G60` (links to Lease Interest from P&L) |
| 155 | CF-Tax Paid | Tax Paid | Component | `=G69+(G73-G44)` (Tax Expense + Sig Items pre-tax adjustment) |
| 156 | CF-Net OCF | Net Operating Cash Flow | **Subtotal** | `=G151+G152+G153+G154+G155` |
| 157 | — | OCF Growth | **Analytical** | `=IF(D156=0,"",E156/D156-1)` |
| 158 | — | EBITDA Cashflow conversion | **Analytical** | `=E151/E148` (Gross OCF / EBITDA) |

### CFI (rows 160–167)
| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 161 | CF-Capex PPE | Capex (PPE) | Component | `=G162*G9` (Capex/Sales ratio × Revenue) |
| 162 | — | Capex / Sales | **Analytical/Input** | `=IF(E9=0,"",(E161+E163)/E9)` (historical); flatlined in forecast |
| 163 | CF-Capex Intang | Capex (Intangibles) | Component | `=0` |
| 164 | CF-Acquisitions | Acquisitions | Component | `=0` |
| 165 | CF-Asset Sales | Asset Sales | Component | `=0` |
| 166 | CF-Other CFI | Other | Residual | `=0` |
| 167 | — | Total Investing Cash Flow | **Subtotal** | `=SUM(G161,G163:G166)` |

### CFF (rows 169–177)
| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 170 | CF-Dividends | Dividends Paid | Component | `=-G89` (negative of Total Dividends from P&L) |
| 171 | CF-Share Issues | Share Issues / Buybacks | Component | `=G137-F137` (change in Issued Capital from BS) |
| 172 | CF-Lease Principal | Lease Principal Payments | Component | `=-F127/G52` (prior Lease Liabilities ÷ Avg Lease Life) |
| 173 | CF-Debt Change | Change in Debt | Component | `=G128-F128` (change in Total Banking Debt from BS) |
| 174 | CF-Other CFF | Other | Residual | `=0` |
| 175 | — | Total Financing Cash Flow | **Subtotal** | `=SUM(G170:G174)` |
| 177 | — | Net Change in Cash | **Subtotal** | `=G156+G167+G175` |

### Operating Free Cash Flow (rows 179–186)
| Row | Label | Type | Formula |
|-----|-------|------|---------|
| 180 | Net OCF | Component | `=E156` |
| 181 | Net Capex | Component | `=E161+E163` |
| 182 | Lease Principal | Component | `=E172` |
| 183 | Operating Free Cash Flow | **Subtotal** | `=E180+E181+E182` |
| 184 | FCF per Share | **Analytical** | `=IF(E82=0,"",E183/E82)` |
| 185 | FCF Yield | **Analytical** | `=IF(Value!$C$4=0,"",E184/Value!$C$4)` |
| 186 | FCF Margin | **Analytical** | `=IF(E9=0,"",E183/E9)` |

## Balance Sheet Projection Methods

All BS items are on the Annual sheet. Forecast method for each:

### Assets
| Row | Key | Item | Projection Method | Formula |
|-----|-----|------|------------------|---------|
| 110 | BS-Cash | Cash | CF-linked roll-forward | `=F110+G177` (prior Cash + Net Change in Cash) |
| 111 | BS-Trade Receivables | Trade Receivables | Revenue ratio | `=G9*G118` (Revenue × Receivables/Revenue %) |
| 112 | BS-Inventories | Inventories | Revenue ratio | `=G9*G119` (Revenue × Inventory/Revenue %) |
| 113 | BS-PPE | PPE | Movement schedule | `=F113-G161+G48` (prior PPE − negative Capex + negative Depreciation = prior + Capex − Depn) |
| 114 | BS-Intangibles | Intangibles | Flatline | `=F114` (prior period) |
| 115 | BS-ROU Assets | ROU Assets | Movement schedule | `=F115+G122+G49` (prior ROU + New Lease Additions + ROU Amortisation) |
| 116 | BS-Other Assets | Other Assets | Flatline | `=F116` |

Supporting rows:
- Row 118: Receivables/Revenue % — `=F118` (flatlined from last actual)
- Row 119: Inventory/Revenue % — `=F119` (flatlined)
- Row 121: Payables/Revenue % — `=F121` (flatlined)
- Row 122: New Lease Additions — `=F122` (flatlined)

### Liabilities
| Row | Key | Item | Projection Method | Formula |
|-----|-----|------|------------------|---------|
| 125 | BS-Trade Payables | Trade & Other Payables | Revenue ratio | `=G9*G121` (Revenue × Payables/Revenue %) |
| 126 | BS-Other Liabilities | Other Liabilities | Flatline | `=F126` |
| 127 | BS-Lease Liabilities | Lease Liabilities | Movement schedule | `=F127+G122+G172` (prior + New Lease Additions + Lease Principal Payments; note G172 is negative) |
| 128 | BS-Total Banking Debt | Total Banking Debt | CF-linked | `=F128+G173` (prior + Change in Debt) |

### Equity
| Row | Key | Item | Projection Method | Formula |
|-----|-----|------|------------------|---------|
| 137 | BS-Issued Capital | Issued Capital | CF-linked | `=F137+G171` (prior + Share Issues/Buybacks) |
| 138 | BS-Retained Profits | Retained Profits | Earnings roll-forward | `=F138+G74-G89` (prior + Statutory NPAT − Total Dividends) |
| 139 | BS-Reserves | Reserves | Flatline | `=F139` |
| 140 | BS-Minorities | Minorities | NCI roll-forward | `=F140-G71` (prior − NCI charge) |

### Derived BS rows
- Row 117: Total Assets = `=SUM(G110:G116)`
- Row 120: Working Capital = `=G111+G112-G125`
- Row 129: Total Liabilities = `=SUM(G125:G128)`
- Row 131: Net Banking Debt = `=G128-G110`
- Row 132: Adj Net Debt (incl leases) = `=G131+G127`
- Row 133: ND/EBITDA = `=IF(G38=0,"",G131/G38)`
- Row 134: Gearing = `=IF((G131+G141)=0,"",G131/(G131+G141))`
- Row 141: Total Equity = `=SUM(G137:G140)`
- Row 142: ROE = `=IF(G141=0,"",G72/G141)`
- Row 143: P/B = `=IF(OR(G141=0,Value!$C$4=0),"",Value!$C$4*G79/G141)`
- Row 144: BS Check = `=G117-G129-G141`

### Special BS-linked formulas
- Row 49 (ROU Amortisation forecast): `=-F115/G52` (prior ROU Assets ÷ Avg Lease Life)
- Row 52 (Avg Lease Life): `=F52` (flatlined; maroon input)
- Row 172 (Lease Principal): `=-F127/G52` (prior Lease Liabilities ÷ Avg Lease Life)

## Return Metrics

ROIC section on Annual sheet (rows 189–194):

| Row | Label | Formula |
|-----|-------|---------|
| 190 | Invested Capital | `=E141+E131` (Total Equity + Net Banking Debt) |
| 191 | Underlying EBIT | `=E54` (links to EBIT row) |
| 192 | ROFE | `=IF(E190=0,"",E191/E190)` (EBIT / Invested Capital) |
| 193 | NOPAT | `=E191*(1-E70)` (EBIT × (1 − Tax Rate)) |
| 194 | ROIC | `=IF(E190=0,"",E193/E190)` (NOPAT / Invested Capital) |

## Valuation Methods

### DCF (Value sheet, rows 2–49)

**Location:** Value sheet, rows 2–49
**Method:** 10-year FCFF DCF with terminal value

**Input assumptions (maroon FFC00000 inputs):**
- C4: Current Share Price
- C9: Valuation Date (red FFFF0000)
- C12: Risk-free Rate
- C13: Equity Risk Premium
- C14: Beta
- C16: Pre-tax Cost of Debt
- C17: Tax Rate (for WACC and NOPAT)
- C19: Debt Weight (D/(D+E))
- C21: Terminal Growth Rate

**Calculated WACC:**
- C15: Cost of Equity = `=C12+C13*C14`
- C18: After-tax Cost of Debt = `=C16*(1-C17)`
- C20: WACC = `=C15*(1-C19)+C18*C19`

**Stub period:** `=(INDEX(Annual!$D:$P,4,MATCH($D$24,Annual!$D$3:$P$3,0))-C9)/365.25`
Calculates the fraction of a year from valuation date to the first forecast period end.

**FCFF build (rows 25–33, columns D–M):**
- Row 25: EBITDA — `=INDEX(Annual!$D:$P,MATCH("EBITDA-Underlying EBITDA",Annual!$A:$A,0),MATCH(D$24,Annual!$D$3:$P$3,0))`
- Row 26: less D&A — same INDEX/MATCH for "DA-Total DA"
- Row 27: EBIT — same for "EBIT-Underlying EBIT"
- Row 28: less Tax on EBIT — `=-D27*$C$17`
- Row 29: NOPAT — `=D27+D28`
- Row 30: plus D&A — `=-D26`
- Row 31: less Capex — `=INDEX(...)` for "CF-Capex PPE" + "CF-Capex Intang"
- Row 32: less WC Change — `=INDEX(...)` for "CF-WC Change"
- Row 33: FCFF — `=D29+D30+D31+D32`

**Terminal value (last forecast year only):**
- M34: Normalised FCFF (capex=D&A) — `=M29+M32` (NOPAT + WC Change, excluding capex)
- M35: Terminal Value — `=M34*(1+$C$21)/($C$20-$C$21)` (Gordon Growth)

**Discounting:**
- Row 37: Discount Factor — `=1/(1+$C$20)^($C$22+N)` where N = 0,1,2,...,9
- Row 38: PV of FCFF — `=D33*D37`
- M39: PV of Terminal Value — `=M35*M37`

**Equity bridge:**
- C41: Sum of PV of FCFs — `=SUM(D38:M38)`
- C42: PV of Terminal Value — `=M39`
- C43: Enterprise Value — `=C41+C42`
- C45: less Net Debt — `=-C7`
- C46: less Lease Liabilities — `=-INDEX(Annual!$D:$P,MATCH("BS-Lease Liabilities",...),MATCH($D$24,...)-1)`
- C47: Equity Value — `=C43+C45+C46`
- C48: Per Share Value — `=C47/C5`
- C49: Upside/Downside — `=IF(C4=0,"",C48/C4-1)`

**Data sources:** All from Annual sheet via INDEX/MATCH using Column A keys. Net Debt and Lease Liabilities use the period before the first forecast year (i.e. most recent actual).

### EV/EBITDA SOTP (Value sheet, rows 52–68)

**Location:** Value sheet, rows 52–68
**Method:** Sum-of-the-parts using segment EBITDA × multiple

**Input assumptions:**
- C54: Select FY (e.g. "FY27E") — dropdown/input to choose which forecast year
- D57, D58: Segment multiples (maroon input implied by editability; currently 8x and 7x)
- D59: Corporate multiple — auto-calculated as weighted average: `=IF((C57+C58)=0,"",(C57*D57+C58*D58)/(C57+C58))`

**Segment EBITDA sourcing:**
- C57: `=INDEX(Annual!$D:$P,MATCH("EBITDA-Steel EBITDA",Annual!$A:$A,0),MATCH($C$54,Annual!$D$3:$P$3,0))`
- C58: Same pattern for "EBITDA-Metals EBITDA"
- C59: Same pattern for "EBITDA-Corporate EBITDA"

**EV calculation:**
- E57–E59: Segment EV = EBITDA × Multiple (e.g. `=C57*D57`)
- E61: Group EV = `=SUM(E57:E59)`

**Equity bridge:**
- E62: less Net Debt = `=-C7`
- E63: less Lease Liabilities = `=-INDEX(Annual!$D:$P,MATCH("BS-Lease Liabilities",...),MATCH($D$24,...)-1)`
- E64: Equity Value = `=E61+E62+E63`
- E65: Per Share Value = `=E64/C5`
- E66: Upside/Downside = `=IF(C4=0,"",E65/C4-1)`
- E68: Implied Group EV/EBITDA = `=IF(Group EBITDA=0,"",E61/Group EBITDA)`

**SOTP segment rows are company-specific** and must be replaced when repurposing (see Template Preservation Method).

## Line Item Retention Policy

When repurposing this template, the following items require special attention:

**Segment-specific items** — these carry company-specific segment names and must be replaced for each new company. The number of segment rows may increase or decrease.

**Group-level items** — these are generic (Total Revenue, Underlying EBITDA, EBIT, NPAT, etc.) and must be retained. Their formulas (sums, cascades) will need range adjustments if the number of segment rows changes.

**Analytical/ratio rows** — these must always be retained. They include: Revenue Growth, GP Margin, EBITDA Margin, EBIT Margin, D&A/Revenue, NPAT Margin, EPS Growth, Dividend Yield, Payout Ratio, OCF Growth, EBITDA Cash Conversion, Capex/Sales, FCF per Share, FCF Yield, FCF Margin, ROE, P/B, ND/EBITDA, Gearing, ROFE, ROIC.

## Template Preservation Method

### Modify-in-place rule
The repurposing method is **INSERT and DELETE**, not clear and rebuild. Group-level rows (subtotals, headers, ratios, analytical rows) remain in place — never delete them, never recreate them. To add new segment rows, insert blank rows directly above the relevant subtotal row and populate. To remove old segment rows that no longer apply, delete them. Then update the subtotal formula's range to cover the new segment rows.

### Retain/Replace map

**REPLACE** (company-specific — delete old, insert new):

*Annual:*
- Rows 7–8: Segment revenue lines (Rev-Steel Revenue, Rev-Metals Revenue)
- Rows 13–14: Segment COGS lines (COGS-Steel COGS, COGS-Metals COGS)
- Rows 19–21: Segment GP lines (GP-Steel GP, GP-Metals GP, GP-Corporate GP)
- Rows 35–37: Segment EBITDA lines (EBITDA-Steel/Metals/Corporate)
- Rows 95–101: Segment KPI rows (Steel/Metals Volume, Rev/t, GP/t)
- Row 102: Active Customers (KPI-Customers)
- Row 104: DIFOT (KPI-DIFOT)

*HY & Segments:*
- Zone 1 segment rows: Same segment revenue, COGS, GP, EBITDA rows as Annual (rows 7–8, 13–14, 19–21, 38–40), plus segment margin ratios (rows 25–26, 45–46)
- Zone 1 KPI rows: Same as Annual (rows 81–82, 84–87, 88, 90)
- Zone 2 (rows 92–116): Entire section — replace with new company's segment forecast driver structure

*Value:*
- SOTP segment rows (57–59): Replace with new company's segments

**RETAIN** (keep as-is, do not delete or recreate):

All other rows, including:
- All section headers and sub-headers
- All subtotal rows (Total Revenue, Total COGS, Gross Profit, Total OpEx, Underlying EBITDA, Statutory EBITDA, Total D&A, EBIT, Net Finance Costs, PBT, NPAT, Total Assets, Total Liabilities, Total Equity, etc.)
- All growth/margin/ratio analytical rows
- All Statutory EBITDA adjustment rows (SBP, Significant Items)
- All D&A rows
- All Interest rows and interest rate assumption rows
- All Tax/NPAT rows
- All EPS/Dividend rows
- All BS rows and BS ratio rows
- All CF rows and CF analytical rows (Gross OCF, Cash Conversion, OFCF section)
- All ROIC rows
- All DCF valuation rows (rows 2–49)
- SOTP equity bridge rows (61–68)

## Repurposing Checklist

1. Copy template to `[TICKER]/Models/[TICKER] Model.xlsx`
2. Update Row 2 labels (company name in B2, adjust actual/forecast zone label positions)
3. Update Row 3 labels (company name in B3, period labels for company's fiscal year)
4. Update Row 4 dates to match company's period end dates
5. Delete old segment rows, insert new segment rows above subtotals
6. Rebuild Column A keys for all new/modified rows using PREFIX-Label convention
7. Update subtotal formulas to cover new segment row ranges
8. Build Zone 2 segment forecast driver sections for new company's segments
9. Wire Zone 1 forecast segment rows to reference Zone 2 outputs
10. Enter historical actuals (blue FF0000CC font for hardcoded values)
11. Populate forecast assumption rows (maroon FFC00000 font)
12. Update SOTP segment rows with new segments and multiples
13. Update Value sheet share price, valuation date, WACC inputs
14. Verify BS Check = 0 for all periods
15. Verify Annual forecast flow items = HY 1H + 2H

## Quality Gates

1. **Zero formula errors** — no #REF!, #DIV/0!, #VALUE!, #N/A across all sheets
2. **BS identity** — BS Check row (144) = 0 for every column (tolerance ±0.2)
3. **P&L cascade** — Total Revenue = sum of segments; GP = Revenue + COGS; EBITDA = sum of segment EBITDAs; EBIT = EBITDA + D&A; PBT = EBIT + Net Finance Costs; NPAT = PBT + Tax + NCI
4. **Annual = 1H + 2H** — for every flow item with a Column A key, Annual forecast value must equal the sum of the corresponding 1H and 2H values from HY & Segments
5. **Sign consistency** — all values follow the sign conventions documented above
6. **Cross-sheet key integrity** — every Column A key referenced in an INDEX/MATCH formula must exist on the target sheet
7. **No hidden assumptions** — no forecast formula contains a hardcoded assumption value; all assumptions are on dedicated maroon rows
8. **Retained row completeness** — all rows marked RETAIN in the preservation map must be present with their original formulas intact
9. **Zone dependency** — Zone 1 forecast segment rows must reference Zone 2 outputs only; no independent forecast logic in Zone 1

## Critical Error Checklist

1. **Formatting attached to row numbers, not fields** — when inserting/deleting rows, formatting must move with the row content, not stay at fixed row positions. Use insert/delete (not clear/rebuild) to preserve formatting automatically.
2. **Missing analytical rows** — OCF Growth, EBITDA Cash Conversion, OFCF section (FCF/Share, FCF Yield, FCF Margin), and ROIC section must never be deleted during repurposing.
3. **Hidden assumptions in formulas** — e.g. `=PCP*1.025` instead of `=PCP*(1+GrowthRow)`. Every assumption must be on a dedicated maroon row.
4. **Zone 1 forecast not wired to Zone 2** — Zone 1 segment forecast cells must reference Zone 2 outputs, never contain independent forecast logic.
5. **Interest/tax forecast by flatline amount** — interest and tax must be forecast using rate × balance pattern (rate on a dedicated assumption row, computed from historical actuals, then flatlined), not by flatlining the dollar amount.
6. **2H derivation using wrong method** — 2H actuals must be FY−1H (from Annual via INDEX/MATCH), not independently hardcoded or formula-estimated.
7. **BS roll-forward broken by row insertion** — when inserting rows above BS or CF sections, verify all cell references in roll-forward formulas still point to the correct rows.
