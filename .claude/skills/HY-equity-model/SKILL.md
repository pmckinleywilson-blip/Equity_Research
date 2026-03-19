---
description: "Half-year equity model template for companies reporting on a semi-annual (1H/2H) cycle. Use this skill when building or modifying financial models for HY-reporting companies. Do not use for quarterly-reporting companies."
---

# Half-Year Equity Model Template

## Template Reference

- **File:** `.claude/templates/HY_model_template.xlsx`
- **Frequency:** Half-year (1H/2H) rolling up to Annual
- **Sheets:** Value, Annual, HY & Segments (3 sheets)
- **Flow:** HY & Segments -> Annual -> Value (bottom-up)

Always read the template fresh from disk before making any modifications. Never rely on a previously loaded version.

## Sheet Architecture

### Value Sheet (68 rows, cols A-R)

No Column A keys. Labels in Column B, data in C-M.

Two valuation blocks:
- **DCF (rows 4-49):** Market data (rows 4-9), WACC inputs (rows 12-22), FCF projection (rows 25-35 across cols D-M for 10 forecast years), discount factors (rows 37-39), EV bridge (rows 41-49)
- **EV/EBITDA SOTP (rows 52-68):** FY selector (C54), segment EBITDA x Multiple = EV (rows 57-59), EV bridge (rows 61-68)

### Annual Sheet (200 rows, cols A-P)

- Row 1: Year integers in D-P
- Row 2: B=sheet title, zone labels in data columns ("Actual ---------->" and "Forecast ----->")
- Row 3: B=company name, C="Units", D-P=period labels (e.g. FY23A-FY35E)
- Row 4: Period-end dates
- Column A = structured keys (format: `Section-Line Item`)
- Column B = display labels
- Column C = units
- Data in D-P (13 fiscal years)
- Actuals in first 3 columns (D-F), Forecasts from G onward

### HY & Segments Sheet (116 rows, cols A-AC)

- Row 1: Year integers (paired, e.g. 2023,2023,2024,2024...) in D-AC
- Row 2: B=sheet title, zone labels in data columns
- Row 3: B=company name, C="Units", D-AC=period labels (1H23, 2H23... 1H35, 2H35)
- Row 4: Period-end dates (31 Dec for 1H, 30 Jun for 2H)
- Column A keys match Annual sheet keys (same Section-Line format)
- Column B = display labels, Column C = units
- Data in D-AC (26 half-year periods)

## Sheet Zone Architecture

### HY & Segments has two distinct zones:

1. **Consolidated P&L zone (rows 5-78):** Contains the half-year P&L with the same line items and Column A keys as the Annual sheet. This is a DEPENDENT zone -- its forecast cells reference the Segment Forecast zone outputs, not independent logic.

2. **Segment Forecast zone (rows 80-116):** Contains the primary forecast driver logic. This is the SOURCE zone. It has:
   - KPIs section (rows 80-90)
   - Segment Forecast - [Segment 1] (rows 92-103): Volume, Volume Growth, Rev/Unit, Rev/Unit Growth, Revenue, GP Margin, GP, Opex, Opex Growth, Opex/Unit, Segment EBITDA
   - Segment Forecast - [Segment 2] (rows 105-116): same structure

**Dependency direction:** Consolidated zone forecast cells reference source zone outputs. The dependent zone's forecast cells must only contain references to the source zone, never independent forecast logic.

Representative formula: HY Row 7 (Segment 1 Revenue) forecast = `=K97` where K97 is the Revenue row in the Segment 1 forecast zone.

### Annual sheet is a DEPENDENT sheet

Most P&L forecast cells use INDEX/MATCH to sum 1H+2H from HY & Segments.

### Segment Forecast Driver Structure

Each segment follows this driver structure:
1. **Volume** -- base input (hardcoded for first forecast, then growth-driven)
2. **Volume Growth** -- assumption input (maroon), hardcoded per period
3. **Rev/Unit** -- derived from prior period x (1 + Rev/Unit Growth)
4. **Rev/Unit Growth** -- assumption input (maroon)
5. **Revenue** = Volume x Rev/Unit / 1000
6. **GP Margin** -- assumption input (maroon), rolls forward from last actual
7. **Gross Profit** = Revenue x GP Margin
8. **Opex** -- derived from prior period x (1 + Opex Growth)
9. **Opex Growth** -- assumption input (maroon)
10. **Opex/Unit** -- analytical (Opex / Volume)
11. **Segment EBITDA** = GP + Opex

### How segment drivers flow to consolidated P&L

On HY & Segments:
- Segment Revenue forecast = reference to driver Revenue row (e.g. `=K97`)
- Segment COGS forecast = `=-K7*(1-K98)` (Revenue x (1 - GP Margin))
- Segment GP forecast = reference to driver GP row (e.g. `=K99`)
- Segment EBITDA forecast = reference to driver EBITDA row (e.g. `=K103`)
- Total OpEx forecast = `=K42-K22` (backed out: EBITDA - GP)
- OpEx line items forecast: proportional split based on prior period mix: `=IF(I33=0,0,K33*I29/I33)`
- Corporate EBITDA = `=-K9*K41` (Revenue x Corporate EBITDA/Revenue ratio)
- Underlying EBITDA = `=SUM(K38:K40)` (sum of segment EBITDAs)

### Other HY forecast mechanics

- **D&A:** Total D&A = Revenue x D&A/Revenue ratio. Split into components using prior period proportions.
- **Interest:** Uses Annual BS balances x interest rate assumptions, minus 1H already booked for 2H.
  - Interest Income: `=INDEX(Annual!$[Cash row]:$[Cash row],MATCH(K$1-1,Annual!$1:$1,0))*K67-J63`
  - Lease Interest: `=-INDEX(Annual!$[Lease Liab row]:$[Lease Liab row],MATCH(K$1-1,Annual!$1:$1,0))*K68-J64`
  - Bank Interest: `=-((INDEX(Annual!$[Debt row]:$[Debt row],MATCH(K$1-1,Annual!$1:$1,0))+INDEX(Annual!$[Debt row]:$[Debt row],MATCH(K$1,Annual!$1:$1,0)))/2)*K69-J65`
- **Tax:** `=-K71*K73` (PBT x Tax Rate)
- **NCI:** 1H rolls forward from prior, 2H scales with EBITDA ratio

## Column Layout

| Sheet | Col A | Col B | Col C | Data starts |
|-------|-------|-------|-------|-------------|
| Value | (unused) | Labels (30 width) | Data/inputs (12 width) | Col C for inputs, Col D for projections |
| Annual | Keys (30 width) | Labels (32 width) | Units (8 width) | Col D (12 width per column) |
| HY & Segments | Keys (30 width) | Labels (32 width) | Units (8 width) | Col D (11 width per column) |

## Row Identification System

Column A contains structured lookup keys in format `Section-Line Item`. Examples:
- `Rev-[Segment] Revenue`, `Rev-Total Revenue`
- `COGS-[Segment] COGS`, `GP-Gross Profit`
- `EBITDA-Underlying EBITDA`, `DA-Total DA`
- `EBIT-Underlying EBIT`, `Int-Net Finance Costs`
- `BS-Cash`, `BS-Trade Receivables`, `BS-PPE`
- `CF-EBITDA`, `CF-WC Change`, `CF-Net OCF`

Subtotal/analytical rows (Revenue Growth, GP Margin, EBITDA Margin, etc.) have NO Column A key -- only a Column B label.

Keys are identical across Annual and HY & Segments sheets for corresponding rows.

## Row Lookup System

Cross-sheet formulas use INDEX/MATCH on Column A keys.

**Annual to HY & Segments (P&L aggregation):**
```
=INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("1H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))+INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("2H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))
```

**Value to Annual:**
```
=INDEX(Annual!$D:$P,MATCH("EBITDA-Underlying EBITDA",Annual!$A:$A,0),MATCH(D$24,Annual!$D$3:$P$3,0))
```

## Color Coding

### Font Colors

| Color | RGB | Meaning |
|-------|-----|---------|
| Dark blue | FF0000CC | Historical actual data (hard-coded values) |
| Maroon/dark red | FFC00000 | User-editable forecast assumptions |
| Bright red | FFFF0000 | Special user input (valuation date only -- yellow fill + red font) |
| Black / Theme(1) | Theme(1, tint=0.0) | Calculated/formula cells in forecast columns |
| White | FFFFFFFF | Zone label font on dark background |

### Fill/Background Colors

| Fill | RGB | Meaning |
|------|-----|---------|
| Light blue | FFC5D9F1 | Major section headers |
| Light grey | FFD9D9D9 | Sub-section headers |
| Dark navy | FF002060 | "Actual" zone label bar |
| Medium blue | FF0070C0 | "Forecast" zone label bar |
| Yellow | FFFFFF00 | Key user inputs requiring attention |
| No fill | -- | All standard data rows |

## Row Formatting Rules

### Major Section Headers (e.g. "P&L", "Balance Sheet", "Cash Flow")
- Col B: bold, light blue fill (FFC5D9F1)
- Data cols: empty

### Sub-section Headers (e.g. "Operating Metrics", "Segment Forecast - [Segment]")
- Col B: bold, grey fill (FFD9D9D9)
- Data cols: empty

### Category Sub-headers (e.g. "Revenue", "COGS", "Gross Profit")
- Col B: bold + single underline
- Data cols: empty

### Subtotal/Total Rows (e.g. "Total Revenue", "Gross Profit", "EBITDA")
- Col B and all data cells: bold
- Borders: top=thin, bottom=thin on B and all data columns

### Final Total Rows (e.g. "Equity Value")
- Borders: top=thin, bottom=medium (thick)

### Line Item Rows (standard data rows)
- Regular weight, no borders, no fill

### Analytical/Ratio Rows (e.g. "Revenue Growth", "GP Margin")
- Regular weight, no borders, no fill
- No Column A key

### Header Row Conventions
- Row 2: Sheet title in B2 (14pt bold), zone labels in data columns (bold, white font on navy/blue fill)
- Row 3: Company name in B3, "Units" in C3, period labels in data columns (bold)
- Row 4: Period-end dates (regular weight)

## Number Format Conventions

| Data Type | Format String | Example |
|-----------|--------------|---------|
| Monetary amounts (local currency, millions) | `#,##0.0` | 596.3 |
| Percentages | `0.0%` | 4.0% |
| Per-share values (EPS, DPS) | `0.000` | 0.456 |
| Share price | `0.00` | 7.91 |
| Multiples | `0.0\x` | 8.0x |
| Whole numbers (headcount, volume) | `#,##0` | 1,234 |
| Dates (row 4) | `d/mm/yy` | 30/06/23 |
| Valuation date | `dd\-mmm\-yy` | 31-Dec-25 |
| Ratios/years | `0.0` | 5.2 |
| Shares outstanding | `0.0` | 100.5 |

These formats are consistent across all sheets for the same data type.

## Zone Label Positioning

- The "Actual ---------->" label is placed at the FIRST actual data column (Annual: D2, HY: D2)
- The "Forecast ----->" label is placed at the FIRST forecast data column (Annual: G2, HY: K2)
- Zone labels span only their own cell -- they do not merge across columns

## Blank Row Convention

One blank spacer row appears between each major section (e.g. between Revenue and COGS, between COGS and Gross Profit, etc.). No blank rows within a section. Blank rows have no formatting.

Key blank row positions:
- Annual: rows 11, 17, 25, 33, 41, 46, 57, 66, 77, 83, 87, 93, 107, 123, 130, 135, 145, 159, 168, 176, 178, 187-188
- HY: rows 11, 17, 27, 36, 47, 50, 56, 61, 70, 79, 91, 104

## Cross-Sheet Structural Correspondence

The Annual sheet and HY & Segments sheet share the same row structure for P&L line items (Revenue through NPAT). The Column A keys, line item order, and labels are identical across both sheets. Any change to a P&L line item on one sheet must be reflected on the other to maintain the INDEX/MATCH linkage.

## Sign Conventions

- Revenue: positive
- COGS: negative
- Gross Profit: positive (Revenue + COGS, since COGS is negative)
- OpEx: negative
- EBITDA: positive
- D&A: negative
- EBIT: positive (EBITDA + D&A, since D&A negative)
- Interest expense: negative
- Tax: negative
- NPAT: positive
- Capex: negative
- Cash inflows: positive, cash outflows: negative throughout CF
- BS items: positive (assets and liabilities both positive)
- Working Capital change in CF: derived from BS movements with sign adjustments

## Cross-Sheet Formula Rules

1. **Annual -> HY & Segments (P&L aggregation):** Most P&L forecast rows on Annual use INDEX/MATCH to sum 1H+2H from HY sheet.
   ```
   =INDEX('HY & Segments'!$A:$AG,MATCH($A[row],'HY & Segments'!$A:$A,0),MATCH("1H"&RIGHT([col]$1,2),'HY & Segments'!$3:$3,0))+INDEX('HY & Segments'!$A:$AG,MATCH($A[row],'HY & Segments'!$A:$A,0),MATCH("2H"&RIGHT([col]$1,2),'HY & Segments'!$3:$3,0))
   ```

2. **HY & Segments -> Annual (2H actual derivation):** For 2H actual periods, HY derives 2H by subtracting known 1H from Annual full-year:
   ```
   =INDEX(Annual!$A:$R,MATCH($A[row],Annual!$A:$A,0),MATCH([year],Annual!$A$1:$R$1,0))-[1H cell]
   ```

3. **HY & Segments -> Annual (interest calculations):** HY interest rows in forecast reference Annual BS balances:
   - Interest Income: `=INDEX(Annual!$[Cash row]:$[Cash row],MATCH(K$1-1,Annual!$1:$1,0))*K[rate row]-J[this row]`
   - Lease Interest: `=-INDEX(Annual!$[Lease Liab row]:$[Lease Liab row],MATCH(K$1-1,Annual!$1:$1,0))*K[rate row]-J[this row]`
   - Bank Interest: `=-((INDEX(Annual!$[Debt row]:$[Debt row],MATCH(K$1-1,Annual!$1:$1,0))+INDEX(Annual!$[Debt row]:$[Debt row],MATCH(K$1,Annual!$1:$1,0)))/2)*K[rate row]-J[this row]`

4. **Value -> Annual:** Uses INDEX/MATCH with string key lookup:
   ```
   =INDEX(Annual!$D:$P,MATCH("[key string]",Annual!$A:$A,0),MATCH($D$24,Annual!$D$3:$P$3,0)-1)
   ```

5. **Annual -> Value:** Dividend yield = `=IF(Value!$C$4=0,"",G88/Value!$C$4)`, YE Shares = `=G79+IF(Value!$C$4=0,0,H171/Value!$C$4)`

## Sub-Period Derivation

Half-year periods are the primary forecast input level. Annual figures are derived by summing 1H+2H.

**Annual forecast formula (sums half-years):**
```
=INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("1H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))+INDEX('HY & Segments'!$A:$AG,MATCH($A7,'HY & Segments'!$A:$A,0),MATCH("2H"&RIGHT(G$1,2),'HY & Segments'!$3:$3,0))
```

**2H actual derivation (for historical periods where only full-year and 1H are known):**
```
=INDEX(Annual!$A:$R,MATCH($A7,Annual!$A:$A,0),MATCH(E$1,Annual!$A$1:$R$1,0))-D7
```

## Assumption Input Placement

The template enforces assumption visibility -- ALL forecast assumptions are surfaced on dedicated rows adjacent to the rows they drive. No forecast formula contains a hardcoded assumption value (except structural zeros for items like Acquisitions, Other CFI, etc.).

**Convention:** The assumption row sits immediately below (or adjacent to) the row it drives. The driven row references the assumption row by cell reference. Assumption rows use maroon (FFC00000) font color to distinguish them from calculated rows.

**Examples:**
- Capex/Sales (assumption) sits below Capex PPE (driven: `=G[Capex/Sales]*G[Revenue]`)
- Receivables/Revenue (assumption) sits below Trade Receivables (driven: `=G[Revenue]*G[Recv/Rev]`)
- Volume Growth (assumption) sits below Volume (driven: `=I[Vol]*(1+K[Growth])`)
- GP Margin (assumption) sits above Gross Profit (driven: `=K[Revenue]*K[Margin]`)

**Forecast assumptions that roll forward from last actual** use `=F[row]` pattern (prior column value). These are editable -- the user changes individual period values as needed.

## Flow vs Point-in-Time

- **P&L items (Revenue through NPAT):** Flow items. Half-year values are independent; annual = sum of halves.
- **Balance Sheet items:** Point-in-time. The Annual sheet holds the year-end (2H) snapshot. BS roll-forward formulas on Annual reference the prior year-end balance.
- **Cash Flow items:** Flow items derived from P&L and BS movements. Annual CF = sum of movements.

## Cash Flow Section Structure

CF format: EBITDA-based (starts with Underlying EBITDA, not receipts/payments).

### CFO (rows 147-156)

| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 148 | CF-EBITDA | Underlying EBITDA | Component | `=G38` (links to P&L EBITDA) |
| 149 | CF-WC Change | WC Change | Component | `=-(G111-F111)-(G112-F112)+(G125-F125)` |
| 150 | CF-Significant Items | Sig Items | Component | `=G44` (links to P&L Sig Items) |
| 151 | -- | Gross OCF | Subtotal | `=SUM(G148:G150)` |
| 152 | CF-Int Received | Interest Received | Component | `=G59` (links to P&L Interest Income) |
| 153 | CF-Interest Paid | Interest Paid | Component | `=G61` (links to P&L Bank Interest) |
| 154 | CF-Lease Int Paid | Lease Interest Paid | Component | `=G60` (links to P&L Lease Interest) |
| 155 | CF-Tax Paid | Tax Paid | Component | `=G69+(G73-G44)` |
| 156 | CF-Net OCF | Net OCF | Subtotal | `=G151+G152+G153+G154+G155` |

WC Change formula logic: negative of receivables change, negative of inventory change, plus payables change. An increase in receivables is a cash outflow.

Tax Paid formula logic: Tax Expense + (Sig Items AT - Sig Items BT).

### CFI (rows 160-167)

| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 161 | CF-Capex PPE | Capex PPE | Component | `=G162*G9` (Capex/Sales ratio x Revenue) |
| 162 | -- | Capex/Sales | Assumption | `=F162` (rolls forward from last actual) |
| 163 | CF-Capex Intang | Capex Intangibles | Component | `=0` (hardcoded zero) |
| 164 | CF-Acquisitions | Acquisitions | Component | `=0` (hardcoded zero) |
| 165 | CF-Asset Sales | Asset Sales | Component | `=0` (hardcoded zero) |
| 166 | CF-Other CFI | Other CFI | Component | `=0` (hardcoded zero) |
| 167 | -- | Total ICF | Subtotal | `=SUM(G161,G163:G166)` |

### CFF (rows 169-175)

| Row | Key | Label | Type | Forecast Formula |
|-----|-----|-------|------|-----------------|
| 170 | CF-Dividends | Dividends Paid | Component | `=-G89` (negative of Total Dividends) |
| 171 | CF-Share Issues | Share Issues | Component | Hardcoded (assumption input) |
| 172 | CF-Lease Principal | Lease Principal | Component | `=-F127/G52` (Prior Lease Liabilities / Avg Lease Life) |
| 173 | CF-Debt Change | Change in Debt | Component | `=0` (hardcoded zero, assumption input) |
| 174 | CF-Other CFF | Other CFF | Component | `=0` (hardcoded zero) |
| 175 | -- | Total FCF | Subtotal | `=SUM(G170:G174)` |

| Row | Label | Formula |
|-----|-------|---------|
| 177 | Net Change in Cash | `=G156+G167+G175` (OCF + ICF + FCF) |

### Operating FCF (rows 179-186)

| Row | Label | Type | Formula |
|-----|-------|------|---------|
| 180 | Net OCF | Component | `=G156` |
| 181 | Net Capex | Component | `=G161+G163` (PPE Capex + Intang Capex) |
| 182 | Lease Principal | Component | `=G172` |
| 183 | Operating FCF | Subtotal | `=G180+G181+G182` |
| 184 | FCF per Share | Analytical | `=IF(G82=0,"",G183/G82)` |
| 185 | FCF Yield | Analytical | `=IF(Value!$C$4=0,"",G184/Value!$C$4)` |
| 186 | FCF Margin | Analytical | `=IF(G9=0,"",G183/G9)` |

## Balance Sheet Projection Methods

| Row | Key | Label | Method | Forecast Formula |
|-----|-----|-------|--------|-----------------|
| 110 | BS-Cash | Cash | CF-linked | `=F110+G177` (Prior + Net Change in Cash) |
| 111 | BS-Trade Receivables | Trade Receivables | Revenue ratio | `=G9*G118` (Revenue x Receivables/Revenue) |
| 112 | BS-Inventories | Inventories | Revenue ratio | `=G9*G119` (Revenue x Inventory/Revenue) |
| 113 | BS-PPE | PPE | Roll-forward | `=F113-G161+G48` (Prior - Capex PPE [neg] + Depreciation [neg]) |
| 114 | BS-Intangibles | Intangibles | Flat | `=F114` (rolls forward unchanged) |
| 115 | BS-ROU Assets | ROU Assets | Roll-forward | `=F115+G122+G49` (Prior + New Lease Additions + ROU Amort [neg]) |
| 116 | BS-Other Assets | Other Assets | Flat | `=F116` |
| 117 | -- | Total Assets | Sum | `=SUM(G110:G116)` |
| 118 | -- | Receivables/Revenue | Assumption | `=F118` (rolls forward) |
| 119 | -- | Inventory/Revenue | Assumption | `=F119` (rolls forward) |
| 120 | -- | Working Capital | Analytical | `=G111+G112-G125` |
| 121 | -- | Payables/Revenue | Assumption | `=F121` (rolls forward) |
| 122 | -- | New Lease Additions | Assumption | `=F122` (rolls forward) |
| 125 | BS-Trade Payables | Trade Payables | Revenue ratio | `=G9*G121` (Revenue x Payables/Revenue) |
| 126 | BS-Other Liabilities | Other Liabilities | Flat | `=F126` |
| 127 | BS-Lease Liabilities | Lease Liabilities | Roll-forward | `=F127+G122+G172` (Prior + New Leases + Lease Principal [neg]) |
| 128 | BS-Total Banking Debt | Banking Debt | CF-linked | `=F128+G173` (Prior + Change in Debt) |
| 131 | -- | Net Banking Debt | Analytical | `=G128-G110` |
| 132 | -- | Adj Net Debt | Analytical | `=G131+G127` |
| 137 | BS-Issued Capital | Issued Capital | CF-linked | `=F137+G171` (Prior + Share Issues) |
| 138 | BS-Retained Profits | Retained Profits | Roll-forward | `=F138+G74-G89` (Prior + Stat NPAT - Total Dividends) |
| 139 | BS-Reserves | Reserves | Flat | `=F139` |
| 140 | BS-Minorities | Minorities | Roll-forward | `=F140-G71` (Prior - NCI) |
| 141 | -- | Total Equity | Sum | `=SUM(G137:G140)` |
| 144 | -- | BS Check | Check | `=G117-G129-G141` (Assets - Liabilities - Equity, should = 0) |

## Return Metrics

Rows 189-194 on Annual:

| Row | Label | Formula |
|-----|-------|---------|
| 190 | Invested Capital | `=G141+G131` (Total Equity + Net Banking Debt) |
| 191 | Underlying EBIT | `=G54` |
| 192 | ROFE | `=IF(G190=0,"",G191/G190)` (EBIT / Invested Capital) |
| 193 | NOPAT | `=G191*(1-G70)` (EBIT x (1 - Tax Rate)) |
| 194 | ROIC | `=IF(G190=0,"",G193/G190)` (NOPAT / Invested Capital) |

## Valuation Methods

### 1. DCF Valuation (Value sheet, rows 2-49)

**Structure:**
- **Market data (rows 4-9):** Share Price (C4, maroon user input), Shares Outstanding (C5, INDEX/MATCH from Annual YE Shares prior year), Market Cap (C6=C4xC5), Net Debt (C7, INDEX/MATCH from Annual BS), Market EV (C8=C6+C7), Valuation Date (C9, yellow fill + red font user input)
- **WACC inputs (rows 12-22):** Risk-free Rate (C12), ERP (C13), Beta (C14), Cost of Equity (C15=C12+C13xC14), Pre-tax CoD (C16), Tax Rate (C17), After-tax CoD (C18=C16x(1-C17)), Debt Weighting (C19), WACC (C20=C15x(1-C19)+C18xC19), Terminal Growth (C21), Stub Period (C22, formula using valuation date and first forecast period-end)
- **FCF projection (rows 25-35, cols D-M = 10 forecast years):** EBITDA, D&A, EBIT pulled via INDEX/MATCH from Annual. Tax on EBIT = -EBIT x Tax Rate. NOPAT = EBIT + Tax on EBIT. Plus D&A, less Capex (PPE+Intang from Annual CF), less WC Change (from Annual CF). FCFF = sum. Terminal year: Normalised FCFF (capex=D&A, so =NOPAT+WC Change), Terminal Value = Normalised FCFF x (1+TGR)/(WACC-TGR).
- **Discounting (rows 37-39):** Discount Factor = 1/(1+WACC)^(stub+n) where n increments 0,1,2...9 across cols D-M. PV of FCFF = FCFF x DF. PV of TV = TV x DF (terminal year only).
- **EV bridge (rows 41-49):** Sum of PV of FCFs + PV of TV = Enterprise Value. Less Net Debt, less Lease Liabilities (from Annual BS) = Equity Value. Per Share Value = Equity Value / Shares. Upside/Downside = Per Share / Share Price - 1.

**Key formulas:**
- FCFF: `=D29+D30+D31+D32`
- Normalised FCFF: `=M29+M32` (terminal year only, assumes capex=D&A)
- Terminal Value: `=M34*(1+$C$21)/($C$20-$C$21)`
- Discount Factor: `=1/(1+$C$20)^($C$22+0)` (increment 0 through 9 across cols)

### 2. EV/EBITDA SOTP (Value sheet, rows 52-68)

**Structure:**
- FY selector (C54, yellow fill) -- user picks which forecast year's EBITDA to use
- Row 56: Header showing selected FY EBITDA label (`=$C$54&" EBITDA"`)
- Segment rows (57-59): Each has EBITDA (col C, INDEX/MATCH from Annual by segment EBITDA key), Multiple (col D, maroon user input, yellow fill), Implied EV (col E, =C x D). Corporate segment multiple is blended: `=IF((C57+C58)=0,"",(C57*D57+C58*D58)/(C57+C58))`
- EV bridge (rows 61-66): Group EV = SUM of segment EVs. Less Net Debt (=-C7). Less Lease Liabilities (INDEX/MATCH from Annual BS). Equity Value = EV + adjustments. Per Share Value = Equity Value / Shares. Upside/Downside.
- Row 68: Implied Group EV/EBITDA = Group EV / Group EBITDA

Segment names are company-specific -- they get replaced when repurposing. The structure (EBITDA x Multiple = EV, blended corporate, EV bridge) is retained.

## Line Item Retention Policy

### Value Sheet
- **RETAIN ALL:** Rows 4-49 (entire DCF section)
- **RETAIN:** Rows 54, 56, 61-68 (SOTP framework: FY selector, header, EV bridge, implied multiple)
- **REPLACE:** Rows 57-59 (segment EBITDA x Multiple rows -- company-specific segment names)

### Annual Sheet
- **RETAIN (subtotals & analytics):** Row 9 (Total Revenue), Row 10 (Revenue Growth), Row 16 (Total COGS), Row 22 (Gross Profit), Row 24 (GP Margin), Row 31 (Total OpEx), Row 38 (Underlying EBITDA), Row 40 (EBITDA Margin), Row 45 (Stat EBITDA), Row 50 (Total D&A), Row 54 (Underlying EBIT), Row 56 (EBIT Margin), Row 62 (Net Finance Costs), Row 65 (Bank Int Rate), Row 68 (PBT), Row 70 (Tax Rate), Row 72 (Underlying NPAT), Row 74 (Statutory NPAT), Row 76 (NPAT Margin)
- **REPLACE (P&L line items):** Rows 7-8 (segment revenues), 13-15 (segment COGS), 19-21 (segment GP), 27-30 (OpEx items), 35-37 (segment EBITDA), 43-44 (SBP, Sig Items), 48-49 (D&A items), 52 (Avg Lease Life), 59-61 (interest items), 64 (Lease Int Rate), 69 (Tax), 71 (NCI), 73 (Sig Items AT)
- **RETAIN:** Rows 79-91 (entire EPS & Dividends section)
- **REPLACE:** Rows 95-106 (Operating Metrics / KPIs)
- **RETAIN:** Rows 108-144 (entire Balance Sheet)
- **RETAIN:** Rows 146-194 (entire Cash Flow & Returns)

### HY & Segments Sheet
- **RETAIN (subtotals & analytics):** Same pattern as Annual -- Total Revenue, Revenue Growth, Total COGS, Gross Profit, GP Margin, segment GP margins, Total OpEx, Cost-to-Income, Underlying EBITDA, EBITDA Margin, segment EBITDA margins, Stat EBITDA, Total D&A, D&A/Revenue, EBIT, EBIT Margin, Net Finance, interest rates, PBT, Tax Rate, Underlying NPAT, Stat NPAT
- **REPLACE (P&L line items):** Same pattern as Annual -- segment revenues, segment COGS, segment GP, OpEx items, segment EBITDA, corporate EBITDA/Rev, SBP, Sig Items, D&A items, interest items, Tax, NCI, Sig Items AT
- **REPLACE:** Rows 80-116 (entire KPIs and Segment Forecast sections)

## Template Preservation Method

When repurposing for a new company:
1. **Copy** the template to the company's Models/ folder -- never modify the template directly.
2. **RETAIN rows** keep their row position, Column A key, formulas, and formatting. Clear company-specific actual data from retained rows but preserve forecast formulas.
3. **REPLACE rows** get their Column B labels, Column A keys, and data overwritten for the new company. Row positions may shift if the new company has more or fewer segments/line items.
4. **Insert/delete rows** as needed to match the new company's structure. Never rebuild a section from scratch -- modify in place.
5. After any row insertions/deletions, verify all cross-sheet INDEX/MATCH formulas still resolve correctly (Column A keys must match between Annual and HY sheets).

## Repurposing Checklist

1. Copy template to `[TICKER]/Models/[Ticker] Model.xlsx`
2. Update sheet titles (Row 2 B2) and company name (Row 3 B3) on all sheets
3. Determine the new company's segment structure and reporting line items
4. Replace segment-specific P&L rows on both Annual and HY & Segments sheets (maintaining identical Column A keys across both)
5. Replace Operating Metrics / KPIs section with company-relevant KPIs
6. Replace Segment Forecast driver section on HY & Segments with new company's segments
7. Update SOTP segment rows on Value sheet to match new segments
8. Rebuild all HY consolidated P&L forecast formulas to reference new driver section
9. Rebuild all Annual P&L forecast formulas to use 1H+2H INDEX/MATCH aggregation
10. Populate actual data (blue font) for historical periods
11. Set forecast assumption inputs (maroon font) for all projection periods
12. Update Value sheet market data (share price, valuation date)
13. Update WACC inputs for new company
14. Verify BS Check row = 0 for all periods
15. Verify all cross-sheet INDEX/MATCH formulas resolve (no #N/A errors)

## Quality Gates

1. **BS Check = 0** for every period (row 144 on Annual). Any non-zero value indicates a broken linkage.
2. **No #N/A or #REF! errors** anywhere in the model -- all INDEX/MATCH lookups must resolve.
3. **Column A key consistency** -- every key on Annual must have an exact match on HY & Segments (and vice versa for P&L items).
4. **Sign conventions preserved** -- COGS negative, D&A negative, interest expense negative, capex negative, tax negative.
5. **Formatting compliance** -- subtotals bold with thin borders, section headers with correct fill colors, assumption inputs in maroon, actuals in blue.
6. **No hidden assumptions** -- every forecast assumption must be on a dedicated maroon-font row, not embedded in a formula.
7. **Zone labels positioned correctly** -- Actual label at first actual data column, Forecast label at first forecast data column.

## Critical Error Checklist

1. **Broken INDEX/MATCH from Annual to HY:** If a Column A key on HY & Segments does not exactly match the corresponding key on Annual, the 1H+2H sum formula returns #N/A. Always verify key spelling after any row changes.
2. **BS roll-forward sign errors:** PPE formula is `=Prior - Capex + Depreciation`. Since Capex and Depreciation are both negative values, this means `Prior - (negative) + (negative)` = `Prior + Capex_abs - Dep_abs`. Misunderstanding the sign convention here breaks the BS.
3. **WC Change sign in CF:** The formula `=-(Receivables change) - (Inventory change) + (Payables change)` must maintain correct signs. An increase in receivables is a cash outflow (negative WC change).
4. **Tax Paid adjustment:** Tax Paid in CF = `Tax Expense + (Sig Items AT - Sig Items BT)`. Omitting the significant items adjustment breaks the CF.
5. **Interest 2H calculation:** The HY interest formulas subtract 1H already booked (e.g. `-J63`). If the 1H formula changes or row positions shift, the 2H calculation breaks.
6. **Lease mechanics circularity risk:** ROU Amort depends on ROU Assets (via Avg Lease Life), which depends on New Lease Additions and ROU Amort. Ensure the chain is non-circular (it uses prior period balances).
7. **Discount factor stub period:** The DCF discount factor starts at stub+0, not stub+1. The stub period itself accounts for the partial-year offset. Getting this wrong shifts all PVs by one year.
8. **SOTP corporate multiple:** The blended multiple formula divides by total segment EBITDA. If a segment has negative EBITDA, the blended multiple can produce nonsensical results.
