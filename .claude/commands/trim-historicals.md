Trim historical periods from a financial model: $ARGUMENTS

## Overview

Remove old historical periods from an Excel financial model built with the standard architecture (Annual, Segments, Inputs tabs). The model file and the number of years of history to keep are specified in `$ARGUMENTS`.

**Example usage:** `/trim-historicals VSL/Models/VSL Model_AI.xlsx keep=3`

If `$ARGUMENTS` doesn't specify a file path or how many years to keep, ask the user using AskUserQuestion.

---

## Phase 1 — Analyse the model

1. Load the workbook with openpyxl and identify all sheets.
2. Read row 1 (year numbers) and row 3 (period labels like FY23A, 1H23) on each data sheet to determine:
   - Which columns are actuals vs forecasts
   - Which columns to delete vs keep
3. Identify the first kept actual column on each sheet — this is the "boundary column" where formulas may reference deleted prior-year data.

---

## Phase 2 — Catalog formulas

Before deleting anything, catalog every formula in the workbook:

```python
# For each sheet, save: {(row, col): formula_string}
```

This catalog is used after deletion to rewrite formulas with correct column references.

---

## Phase 3 — Delete columns

Use `ws.delete_cols(first_col, count)` on each sheet to remove the unwanted periods.

**Important:** openpyxl's `delete_cols` moves cells but does NOT update formula references. All formulas will still contain the old column letters after this step. That's expected — Phase 4 fixes them.

Column deletions by sheet:
- **Annual:** 2 columns per removed financial year
- **Segments (half-year):** 2 columns per removed financial year (1H + 2H)
- **Inputs:** Match the same half-year periods as Segments

---

## Phase 4 — Rewrite all formulas

For each formula that was in a kept column, translate it to the correct new column position.

### 4a. Same-sheet references

Use openpyxl's `Translator` to shift formula references:

```python
from openpyxl.formula.translate import Translator

t = Translator(original_formula, origin=old_cell_coordinate)
translated = t.translate_formula(new_cell_coordinate)
```

The Translator correctly handles same-sheet relative references (e.g., shifting `=F7+F8` from F9 to D9 produces `=D7+D8`).

### 4b. Cross-sheet reference corrections

The Translator shifts ALL column references by the same delta, but cross-sheet references may need a different shift if the referenced sheet had a different number of columns deleted.

**Fix pattern:** After Translator, adjust cross-sheet column references using regex:

```python
# Pattern: SheetName![$]COL[$]ROW
pattern = r"(SheetName!)(\$?)([A-Z]{1,3})(\$?\d+)"
```

Apply the correction = (target_sheet_shift - source_sheet_shift) to each cross-sheet column reference.

Common cases:
- **Segments formulas → Inputs references:** If Segments shifted by -4 but Inputs shifted by -2, correct Inputs refs by +2
- **Value formulas → Annual references:** If Value didn't shift (0) but Annual shifted by -2, shift Annual refs by -2

### 4c. Formulas on non-shifted sheets

For sheets that had no columns deleted (e.g., Value), formulas aren't processed by the Translator. Manually find and shift any cross-sheet references to sheets that DID have columns deleted.

---

## Phase 5 — Fix boundary column formulas

The first remaining actual column has formulas that referenced the now-deleted prior year. Fix these:

### Growth rates (YoY comparisons)
Set to blank (`""`) — there's no prior year to compare against.

Typical rows: Revenue Growth, GP Growth, EBITDA Growth, NPAT Growth, OpEx Growth, Operating Cash Flow Growth, and any segment-level YoY growth rows.

### Cash flow items referencing prior-year balance sheet
Hardcode the computed values. Read these from the file (using `data_only=True`) BEFORE deleting columns:
- **Working capital change:** `=-(Receivables_change)-(Inventory_change)+(Payables_change)`
- **Change in issued capital:** `=Current_year - Prior_year`
- **Change in banking debt:** `=Current_year - Prior_year`

### Special ratios referencing prior-year balance sheet
Set to blank:
- Depreciation / prior-year assets
- Interest expense / average debt (if it averages current + prior year)

### Segments tab boundary
Apply the same treatment to the first 1H and 2H columns: blank YoY growth rates that referenced deleted periods.

---

## Phase 6 — Fix labels

The "Actual ---------->" label in row 2 gets deleted with column D. Re-set it on the new column D for each sheet.

The "Forecast ----->" label shifts automatically with the columns — verify it's still in the correct position.

---

## Phase 7 — Verify

Run these checks on the saved file:

1. **No #REF! errors:** Scan every formula cell for `#REF!`
2. **Header consistency:** Verify row 1 (years), row 3 (period labels), and row 4 (dates) are correct and contiguous
3. **Cross-sheet alignment:** For each INDEX/MATCH formula, verify the lookup value exists on the target sheet
4. **Spot-check formulas:** Print key formulas from the first actual, second actual, and first forecast columns to confirm correct column references
5. **Cross-sheet references:** Verify Segments→Inputs, Segments→Annual, Annual→Segments, and Value→Annual references point to the correct columns

Report all results to the user.
