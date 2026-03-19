# Validation Gaps — Lessons Learned

## Gap 1: BS/CF forecast formatting not checked (found 2026-03-18)

**What happened:** The /validate-model script's formatting checks (number format, bold, borders) only sampled cells in the P&L section and actual columns. It completely missed that forecast columns (G-P) in the BS/CF/OpFCF/ROIC sections (rows 105-191) had `General` format, no bold on subtotals, and no borders.

**Root cause in build:** The forecast formula wiring script wrote formulas using `cell.value = "=formula"` without copying formatting from adjacent actual cells. The row insertion at row 108 preserved formatting on existing (actual) cells but newly-written formula cells were bare.

**Fix applied to validation:** The validation must check number format, bold, and borders on ALL data cells across ALL rows and ALL columns — not just samples from specific sections. Specifically:
- CHECK 3e (number format): scan every data cell in every data column, not just actuals
- CHECK 3a (bold on subtotals): check every data column on every subtotal row
- CHECK 3d (borders on subtotals): same — every column, not just B and a few samples

**Fix applied to build:** After writing forecast formulas, copy `number_format`, bold, and border attributes from the last actual column to all forecast columns in the same row. This should be a standard post-processing step in the build script.
