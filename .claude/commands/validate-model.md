---
description: >
  Validate a built financial model against its source template. Compares
  formatting, structure, formulas, and conventions. Diagnoses root causes
  and proposes fixes. Run after /build-model: /validate-model [TICKER]
---

# /validate-model

Validate the model for: $ARGUMENTS

---

## Step 1 — Identify files

1. Find the built model at `[TICKER]/Models/[TICKER] Model.xlsx`. If it does
   not exist, stop and tell the user to run `/build-model [TICKER]` first.
2. Identify the source template — check the skill files in `.claude/skills/`
   to find which template was used (documented in the skill file's Template
   Reference section). Load the template from `.claude/templates/`.
3. Load the corresponding skill file and keep it in context.

---

## Step 2 — Run comparison

Write and execute a Python script using openpyxl that compares the built model
against the template. The script must check every item below and produce a
structured report.

**IMPORTANT**: All formatting checks (2c, 2d) must scan EVERY data cell in
EVERY row across EVERY data column — both actual and forecast columns. Do not
sample a few cells or sections. Formatting errors frequently occur in forecast
columns where formulas were written without copying formatting from actuals,
or in BS/CF sections that were not checked because only the P&L was sampled.

### 2a. Sheet structure
- Same sheets exist in both files (same names, same order)
- Row count is reasonable (model may have more or fewer rows due to segment
  changes, but should be within ~20% of template)

### 2b. Row-by-row structure comparison
For each sheet, compare every row:
- Extract row identifier (Column A key or equivalent per skill file), Column B
  label, Column C units from both template and model
- Flag any template row that is missing from the model (particularly retained
  rows from the skill file's Template Preservation Method)
- Flag any model row that has no template equivalent and is not a legitimate
  new segment or KPI row
- Verify row order matches — retained rows should appear in the same relative
  order as the template

### 2c. Formatting comparison
For EVERY row in the model, across ALL data columns (actual AND forecast),
compare against the template row it corresponds to (or the skill file's Row
Formatting Rules if no direct template row exists):

- **Font bold**: Subtotal rows must be bold across ALL cells in the row —
  check every data column, not just the label column (column B). A subtotal
  row with bold on B but not on data columns is a failure.
- **Font color**: Actual data cells should use the actuals color (per skill
  file Color Coding). Forecast assumption cells should use the assumption
  color. Formula cells should use default (black). Check every data cell.
- **Fill color**: Section headers, zone labels, and other filled rows must
  match the template's fill colors. Fills must extend across ALL data
  columns, not just the label. Check every column in the row.
- **Borders**: Subtotal rows must have borders on ALL data cells, not just
  the label. Compare border style (thin/medium, top/bottom) against
  template. Check every data column.
- **Number format**: Every data cell must have the correct number format per
  the skill file's Number Format Conventions. Flag any cells with "General"
  format in data columns. Check monetary, percentage, per-share, count, and
  ratio formats separately. Scan every cell in every data column — do not
  skip forecast columns or BS/CF sections.

### 2d. Zone label positioning
- Check row 2 (or whichever row the skill file documents for zone labels)
- The "Actual" zone label must be at the first actual data column with the
  correct fill extending across ALL actual columns
- The "Forecast" zone label must be at the first forecast data column with
  the correct fill extending across ALL forecast columns
- Check on every sheet that has zone labels

### 2e. Blank row spacing
- Count blank rows between each major section
- Compare against the skill file's Blank Row Convention
- Flag extra or missing blank rows
- Flag trailing empty rows at the bottom of the sheet
- Flag trailing empty columns beyond the data range

### 2f. Cross-sheet structural correspondence
- If the skill file documents that two sheets share the same row structure,
  extract the row identifiers from the corresponding sections on both sheets
- Verify they are identical in content and order
- Flag any identifier present on one sheet but missing from the other

### 2g. Formula structural validation
For every formula cell in forecast columns:
- If it contains a cross-sheet lookup, extract the lookup value and confirm
  it exists on the target sheet
- If it contains a period label match, confirm the period label exists in
  the target sheet's header row
- If it contains a direct cell reference, confirm the referenced cell exists
  and is within the data range
- Flag any formula that would produce #N/A, #REF!, or #VALUE! based on
  structural analysis

### 2h. Skill file compliance
Check each convention documented in the skill file:
- Sign conventions (spot-check actual values against expected signs)
- Cross-sheet formula method matches skill file
- Assumption input placement follows skill file pattern
- CF, BS, and return metric formulas follow the patterns documented in the
  skill file's structure tables

### 2i. Value sheet integrity
- All labels updated from template currency to company currency
- All cross-sheet references use correct keys and column ranges
- DCF and SOTP structure matches template (same rows, same formula patterns)
- Input cells (share price, WACC components, multiples) are populated

---

## Step 3 — Produce report

Present the results as a categorized issue list:

**Categories:**
- `CRITICAL` — formula errors, missing sheets, broken cross-sheet references
  (model will not calculate correctly in Excel)
- `MAJOR` — formatting failures that affect usability (missing bold/borders
  on subtotals, wrong number formats, zone labels misplaced)
- `MODERATE` — structural issues (missing analytical rows, extra blank rows,
  trailing empty columns, inconsistent font colors)
- `MINOR` — cosmetic issues (label wording differences, minor fill
  discrepancies)

For each issue, state:
- Sheet name and row/column location
- What was expected (from template or skill file)
- What was found in the model
- Suggested fix

---

## Step 4 — Diagnose root causes

For each issue found in Step 3, determine the root cause:

### Category A: Skill file gap
The skill file does not document a convention that the template uses. The
ingest-model command failed to capture it.
- **Evidence**: The template has a pattern (formatting, formula, structure)
  that is not mentioned anywhere in the skill file
- **Fix**: Propose a specific addition to the ingest-model command (so it
  captures this convention for any template), AND propose the corresponding
  skill file section content for the current template

### Category B: Build command gap
The build-model command does not instruct the builder to follow a convention
that IS documented in the skill file.
- **Evidence**: The skill file clearly documents the convention, but the
  build-model command doesn't reference it or enforce it
- **Fix**: Propose a specific addition or change to the build-model command

### Category C: Execution error
Both the skill file and build command correctly specify the convention, but
the builder did not follow them.
- **Evidence**: The skill file documents it, the build command references it,
  but the output model doesn't comply
- **Fix**: Add the specific failure to the skill file's Critical Error
  Checklist. Consider whether the build command needs stronger language or
  a more explicit verification step.

### Category D: Validation gap
The validation script itself failed to detect an issue that the user found
manually, or detected a false positive.
- **Evidence**: The user reported an issue that the script missed, or the
  script flagged something that is actually correct
- **Fix**: Propose a specific change to this validate-model command to
  catch the gap in future runs. Document the gap in
  `.claude/skills/[skill-name]/references/validation-gaps.md` for reference.

### Generalizability check
Before proposing any fix to ingest-model or build-model, verify it is NOT
company-specific or template-specific:
- Does the proposed change reference any company name, ticker, segment name,
  sheet name, column letter, row number, or formula pattern specific to one
  template?
- If yes: the change belongs in the skill file, not the command
- If no: it can go in the command

---

## Step 5 — Present proposed fixes

Group all proposed fixes by file:

```
PROPOSED CHANGES:

ingest-model.md:
  [Issue #]: [specific change]
  ...

build-model.md:
  [Issue #]: [specific change]
  ...

validate-model.md:
  [Issue #]: [specific change]
  ...

[skill-name]/SKILL.md:
  [Issue #]: [specific change]
  ...

[skill-name]/references/validation-gaps.md:
  [Issue #]: [gap description]
  ...
```

Ask the user: "Would you like me to apply these fixes? After applying, you
can re-run `/build-model [TICKER]` and `/validate-model [TICKER]` to test
whether the fixes resolved the issues."

---

## Step 6 — Apply fixes and iterate (if user approves)

If the user approves:
1. Apply all proposed changes to the command files, skill file, and
   validation gaps document
2. Prompt the user with next steps:
   - **To rebuild and revalidate:** "Run `/build-model [TICKER]` to rebuild
     with the updated commands, then `/validate-model [TICKER]` again to
     check if the issues are resolved."
   - **To save decisions for automated testing:** "Run
     `/save-fixture [TICKER]`"
   - **To run automated rebuild + validate loop:** "Run
     `/test-build [skill-name]`"

If the user declines:
1. Save the issue report to `[TICKER]/validation-report.md` for reference
2. Prompt: "The report has been saved. You can review and apply fixes
   manually, or re-run `/validate-model [TICKER]` after making changes."
