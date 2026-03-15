# context-log.md — OCL

## Model Status
- Model file: `OCL Model.xlsx`
- Last modified: 2026-03-15
- Sheets in use: Annual, HY & Segments, Value
- Build stage: Initial build complete — structure, actuals (FY21–FY25 + 1H21–1H26), formulas, and forecasts wired

## Last Session
- Date: 2026-03-15
- Work completed:
  - Copied HY_model_template.xlsx → OCL/Models/OCL Model.xlsx
  - Rebuilt all three sheets for OCL (Australian GovTech SaaS, FYE 30 June, AUD)
  - Annual: 10 years FY21A–FY30E, 179 rows (P&L → EPS → KPIs → BS → CF → OFCF → ROIC)
  - HY & Segments: 20 half-year columns 1H21–2H30E, 113 rows (P&L → KPIs → 3 segment forecast blocks → group inputs)
  - Value: DCF (5-year FCFF, 8.7% WACC) + EV/EBITDA SOTP
  - Revenue: 3 product groups (Info Intelligence, Planning & Building, Regulatory Solutions) + interest income
  - Single-segment COGS, GP, OpEx (Distribution, R&D net of cap, Admin)
  - R&D capitalisation from FY24 modelled (Amort Dev Costs in D&A, Capitalised Dev in CFI)
  - No external debt — Net Cash replaces Net Banking Debt
  - All INDEX/MATCH cross-sheet formulas: flow = 1H+2H, point-in-time = 2H only
  - 2H back-calculations via INDEX/MATCH to Annual for historical periods
  - Forecast assumptions seeded (maroon): ARR growth II 11%, PB 25%, RS 20%; COGS 6%; Dist 35%; R&D 25% at 50% cap rate; Admin 7.5%; ETR 14%
  - Column A keys: 77 on Annual, 37 on HY — all unique, all matching between sheets
- Files modified: OCL/Models/OCL Model.xlsx, OCL/context-log.md

## Open Items
- Historical data entered from plan assumptions — should be verified against actual ASX filings (PDFs in OCL/Company reports/)
- Interest income forecast uses simplified carry-forward; could be improved with cash-balance-driven calculation
- D&A forecast (Amort Dev Costs) uses carry-forward; could model capitalised pool with useful life amortisation
- Lease principal payment forecast uses simplified 4-year average life assumption
- BS Check needs verification in Excel (formula-only check here; actual values depend on Excel calculation)
- DPS forecast uses payout ratio × EPS; may want to switch to explicit DPS input
- Value sheet SOTP is group-level only (single operating segment) — may want to add revenue multiples by product group

## Key Decisions
- [2026-03-15] Single-segment P&L structure (no segment EBITDA rows) — OCL reports as one operating segment with revenue disaggregated into 3 product groups
- [2026-03-15] R&D capitalisation discontinuity: FY21–FY23 all R&D expensed (Amort Dev = $0); FY24+ net of capitalisation. Gross R&D tracked in KPIs.
- [2026-03-15] Interest income included in Total Revenue (material ~$3.4m FY25) but excluded from GP calculation (GP = contract revenue − COGS)
- [2026-03-15] No external debt — removed Total Banking Debt, Net Banking Debt. Net Cash shown instead.
- [2026-03-15] Removed NCI row (100% owned subsidiaries), Inventories row (SaaS company)
- [2026-03-15] Added Contract Assets, Contract Liabilities, Deferred Tax, Provisions to BS
- [2026-03-15] Stat-Significant Items → Stat-M&A Costs + Stat-FX; NPAT-Sig Items AT → NPAT-Other Items AT
