# context-log.md — GYG

## Model Status
- Model file: `GYG Model.xlsx`
- Template: HY_model_template.xlsx
- Build stage: Phase 6 — All forecast formulas wired (HY Zone 2, Zone 1, Annual, Value sheet updated)

## Last Session
- Date: 2026-03-15
- Work completed:
  - Wired HY Zone 2 forecast formulas (cols I=2H26E through AA=2H35E): AU network build (DT/Strip/Other counts, AUVs, network sales), corp/franchise build (count, sales, margins, royalties), SG/JP international, G&A, segment EBITDA; US segment (count, sales, margins, G&A, EBITDA)
  - Wired HY Zone 1 forecast formulas: Revenue rows reference Zone 2 outputs, COGS/OpEx maintain PCP ratios to revenue, D&A proportional split from PCP, interest PCP flatline, tax at 30%, segment EBITDA from Zone 2
  - Wired Annual forecast formulas: all flow items use INDEX/MATCH 1H+2H from HY, BS roll-forwards (Cash=CF-linked, Rec/Inv/Pay=revenue ratios, PPE=capex-depn, ROU=lease additions-amort, Lease Liab=additions+principal)
  - Annual CF: EBITDA-based format, Capex=Capex/Sales×Revenue, Lease Principal=-prior lease liab/avg lease life, WC change, interest/tax linkages
  - Updated Value sheet: AUD currency, share price $35, WACC inputs (Rf=4.2%, ERP=6%, Beta=1, no debt, TGR=2.5%), SOTP segments (AU Segment 25x, US Segment 0x, Corporate residual), all INDEX/MATCH refs updated for $D:$Q range
  - All verification checks passed (0 empty forecast cells, 0 missing keys)
- Files modified: GYG Model.xlsx, context-log.md
- Scripts: scripts/gyg_forecast_build.py

## Data Notes
- FY22 Other Revenue: total 13.3 hardcoded (no sub-item split); FY22 D&A: total -14.4 hardcoded
- 2H24 Franchise Fee -0.349 suspect (likely reclass/timing in FY25 annual vs 1H25 report)
- 1H24 segment EBITDA uses pro forma adjusted figures (pre-IPO adjustments)
- CF working capital and non-cash items back-calculated from reported OCF

## Open Items
- FY22 BS/CF not available (Prospectus had only P&L pro forma)
- HY Zone 2: DT/Strip/Other format counts blank for 1H24-1H25 (only FY25 and 1H26 actual)
- HY Zone 2: SG/JP network sales for 1H26 not entered (only Asia total ~42m available)
- US Corp Margin % forecast starts at -50% improving 5pp/half — needs review against management commentary
- AU new restaurant openings seeded at +9 DT, +3 Strip per half — calibrate to latest guidance
- Avg Lease Life set to 9 years (derived from ROU/ROU Depn) — verify against lease note disclosures
- New Lease Additions seeded at 45.5m/year (FY25 implied) — may need adjustment for accelerating rollout
- Capex/Sales at -13.1% — check if this holds as network matures
- SOTP AU segment multiple 25x needs benchmarking against QSR peers (current Domino's, Collins Foods comps)
- Share buyback program ($100m total, $27m used in 1H26) not yet modelled — CF-Share Issues set to 0
- Interest forecast uses PCP flatline on HY; should ideally reference Annual BS via INDEX/MATCH for rate×balance approach
- Annual analytical rows (Rev Growth, GP Growth, EBITDA Growth, NPAT Growth) now have formulas for forecast periods

## Key Decisions
- [2026-03-15] Template: HY_model_template.xlsx (3-sheet cascade: HY & Segments → Annual → Value)
- [2026-03-15] Two segments: Australia (incl SG/JP) and US
- [2026-03-15] Revenue build: AU format-level (DT×AUV + Strip×AUV + Other×AUV) → network sales → corp/franchise split
- [2026-03-15] HY sheet: D=1H24 through H=1H26 (actuals), I=2H26E onward (forecast)
- [2026-03-15] Forecast methodology: Zone 2 drives revenue via count growth + AUV growth, margins flatlined from PCP, Corp sales derived as (corp count/total count)×AU network sales
- [2026-03-15] D&A forecast: total D&A as ratio of revenue (PCP flatline), split proportionally into PPE/ROU/Reacq/Other
- [2026-03-15] DCF uses Statutory EBITDA (not Group Segment EBITDA) as the EBITDA input
