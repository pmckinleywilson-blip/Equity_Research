# context-log.md — GYG.AX

## Model Status
- Model file: `GYG Model.xlsx`
- Last modified: 2026-03-18
- Sheets in use: Value, Annual, HY & Segments
- Build stage: Initial build complete. Structure, historicals (FY23-1H26), and forecast formulas all wired. Requires Excel review of forecast outputs and assumption refinement.

## Last Session
- Date: 2026-03-18
- Work completed:
  - Populated HY segment driver actuals (rows 88-97 Australia, 100-106 US) for cols F-J (1H24 through 1H26)
  - Australia drivers: corp restaurants, new openings, corp revenue, revenue growth, AUV, franchise restaurants, franchise revenue, total revenue, EBITDA/SegRev margin, segment EBITDA
  - US drivers: restaurants, openings, revenue, growth, AUV, EBITDA
  - Wired HY segment driver forecasts (cols K-AC, 19 half-year periods) with maroon assumptions
  - Wired HY consolidated P&L forecasts (rows 7-66): revenue from drivers, EBITDA bridge, expenses (back-solved), D&A (ratio-based), finance, 30% statutory tax, NPAT
  - Wired Annual P&L (rows 7-65) via INDEX/MATCH "1H"/"2H"&RIGHT(col$1,2) from HY
  - Wired Annual EPS & Dividends (rows 69-82): shares flat, payout flat
  - Wired Annual BS (rows 105-141): cash from CF, receivables/inventory/payables from revenue ratios, PPE from capex+dep, ROU from leases+dep, lease liabilities from principal+new
  - Wired Annual CF (rows 145-174): EBITDA, WC, SBP add-back, interest, tax, capex (ratio-based), lease principal (1/8 of prior balance), dividends (lagged 1yr)
  - Wired Annual OpFCF (rows 177-183) and ROIC (rows 187-191)
  - Updated Value sheet: tax rate 0.28->0.30, fixed SOTP implied multiple formula key
  - Relabelled HY row 96 to "EBITDA / Segment Revenue" (was EBITDA Margin)
- Files modified: GYG/Models/GYG Model.xlsx

## Open Items
- Populate remaining HY P&L actuals for 1H23-2H23 (cols D-E) if data available
- Update WACC inputs (currently template defaults: RFR 4%, ERP 6%, Beta 1.0)
- Update share price and valuation date on Value sheet
- Verify BS Check = 0 when opened in Excel (circular refs may need iterative calc enabled)
- Refine US segment forecast (currently EBITDA flat from PCP; needs improvement trajectory)
- Add dedicated franchise opening assumption row to avoid hardcoded +5/half in row 93
- Review D&A ratio approach (currently flat from PCP) — may understate as new store openings increase
- Consider adding HY balance sheet for half-year granularity

## Key Decisions
- [2026-03-18] P&L uses EBITDA Bridge structure. Expenses as memo (back-solved). EBITDA key = "EBITDA-Statutory EBITDA".
- [2026-03-18] SOTP uses Australia (25x) and US (15x) segment EBITDA multiples as placeholders.
- [2026-03-18] Segment EBITDA derived as Total Segment Revenue x EBITDA/Segment Revenue margin (row 96). This ratio replaces EBITDA/Network Sales for modeling simplicity since network sales are not modeled.
- [2026-03-18] Forecast tax rate set to 30% statutory (vs elevated actual ETR driven by US losses and SBP non-deductibility). Simplifies to normalised rate.
- [2026-03-18] Franchise restaurant growth hardcoded at +5 per half (row 93). Needs dedicated assumption row for proper flexibility.
- [2026-03-18] Lease principal repayment = 1/8 of prior balance (approx 8yr avg lease life).
