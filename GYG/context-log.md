# context-log.md — GYG

## Model Status
- Model file: `GYG Model.xlsx`
- Last modified: 2026-03-15
- Sheets in use: Annual, HY & Segments, Value
- Build stage: Complete — historical actuals, forecast formulas, BS roll-forwards, CF linkages, and Value sheet (DCF + SOTP) all built

## Last Session
- Date: 2026-03-15
- Work completed:
  - Downloaded 18 ASX reports, read prospectus, FY24 AR, FY25 AR, 1H25 interim, 1H26 interim
  - Built complete 3-sheet model: Segments → Annual → Value
  - Annual: FY22-FY25 actuals, FY26E-FY29E forecasts with INDEX/MATCH from Segments
  - Segments: 1H23-1H26 actuals, 2H26E-2H29E forecasts with PCP-based escalation
  - Value: DCF (4-year FCFF + terminal value) and EV/EBITDA SOTP (Australia + US)
  - BS roll-forwards: Cash, term deposits, receivables, PPE, ROU, intangibles, lease liabilities, provisions, equity
  - CF: OCF from receipts/payments, CFI with capex and term deposits, CFF with dividends, buyback, lease payments
  - Verification: P&L cascade PASS (FY23-FY25), sign conventions PASS, segment EBITDA PASS
- Files modified: GYG Model.xlsx, context-log.md

## Open Items
- BS Check needs verification in Excel (openpyxl can't evaluate formulas)
- Forecast formulas need validation by opening in Excel and checking INDEX/MATCH resolution
- FY22 BS not entered (limited data from prospectus; only FY23+ BS available)
- 1H23 segment data from prospectus pro forma — may differ slightly from statutory
- US segment margin assumptions aggressive (-40% improving to -15%) — review against management guidance
- Capex assumptions ($70-85m pa) and new lease additions ($40-48m) need review against pipeline guidance
- SOTP Australia multiple of 25x needs benchmarking against QSR peers (CMG, CAVA, etc.)

## Key Decisions
- [2026-03-15] Starting from FY22 (prospectus) not FY21 (no data). 4 years of history.
- [2026-03-15] Two operating segments: Australia (incl. Singapore/Japan) and US — matching CODM.
- [2026-03-15] Network sales build by format (DT/Strip/Other × AUV) at network level in KPIs.
- [2026-03-15] No macro indicators — no commodity/FX exposures identified.
- [2026-03-15] Forecast method: PCP-based escalation with maroon assumption inputs for growth rates, margins, ratios.
- [2026-03-15] Cash rent add-back approach used in EBITDA bridge (matching company's segment reporting).
- [2026-03-15] Effective tax rate ~45% reflecting non-deductible SBP and unrecognised US tax losses.