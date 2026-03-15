---
name: hy-equity-model
description: >
  Use this skill whenever building, repurposing, or modifying a half-yearly equity
  financial model using the 3-sheet cascade architecture. This skill is designed for
  companies that report on an annual basis with a single half-year interim result (1H
  and 2H). It is not appropriate for quarterly reporters. Triggers include: "build a
  model for [company]", "repurpose the HY template", "create an equity model", "set up
  a new company model", or any request involving HY_model_template.xlsx or models
  derived from it. Always load this skill before touching any such file.
---

# HY Equity Model Skill

## Template Reference
Base template: `@.claude/templates/HY_model_template.xlsx`
Always read this file fresh from disk before making any modifications.

---

## Sheet Architecture

Three sheets in a strict cascade — never reverse the data flow:

| Sheet | Purpose | Receives from |
|---|---|---|
| **Segments** | Half-year detail, primary forecast driver | Source of truth |
| **Annual** | Annual P&L, BS, CF, ROIC | Segments (forecasts via INDEX/MATCH) |
| **Value** | DCF + EV/EBITDA SOTP | Annual |

---

## Column Layout

**Annual:** Columns D+. Row 1 = year integer. Row 3 = period label (FY25A, FY26E). Row 4 = period end date.

**Segments:** Columns D+, paired 1H/2H per year. Row 1 = year. Row 3 = half label (1H25, 2H25). Row 4 = period end date.

---

## Column A Key System

Column A on both Annual and Segments holds machine-readable keys used by all
INDEX/MATCH lookups. Keys are **company-specific** — built fresh from the company's
reported financials for each new model. The format is fixed.

### Format: `Section-Item Name`

| Prefix | Section |
|---|---|
| `Rev-` | Revenue |
| `COGS-` | Cost of goods sold |
| `GP-` | Gross profit |
| `OPEX-` | Operating expenses |
| `EBITDA-` | EBITDA items |
| `Stat-` | Statutory adjustments |
| `DA-` | Depreciation and amortisation |
| `EBIT-` | EBIT |
| `Int-` | Interest items |
| `Tax-` | Tax |
| `EPS-` | Per share items |
| `BS-` | Balance sheet |
| `CF-` | Cash flow |
| `KPI-` | Operating metrics and non-financial KPIs |

### How to build the key list for a new company
1. Read the company's most recent annual report
2. Walk P&L → BS → CF → operating metrics in order
3. Assign a key to every reported line item that will appear as a model row
4. For segmented reporters: create individual segment keys plus a consolidated total
5. Add calculated rows that don't appear in filings but are needed (e.g. `GP-Gross Profit`)
6. Write out the complete key list and confirm with the user before building any sheet

### Key rules
- Keys must be unique across the workbook
- Keys must match exactly between Annual and Segments for every cross-referenced row
- Rows that exist only on Annual (BS, CF) use local formulas — no INDEX/MATCH to Segments
- Never rename a key after the model is built

---
### Why Column A keys matter beyond INDEX/MATCH

The key schema serves two distinct purposes:

**1. Formula lookups** — INDEX/MATCH uses keys to find the correct row across sheets
regardless of row position. This is the mechanical function.

**2. Data continuity across reporting changes** — when a company restates prior year
figures or restructures its segments, keys provide a secondary matching mechanism
beyond value comparison. Old segment keys continue to anchor historical data, new
keys anchor forward data, and group-level totals remain continuous throughout. This
is what makes the model resilient to the reporting changes that every company
eventually makes.

This is why keys must never be renamed after the model is built — renaming a key
silently breaks the historical continuity it was anchoring.
---

## Color Coding

| Color | Meaning |
|---|---|
| Blue text (FF0000CC) | Hardcoded actuals — never overwrite with formulas |
| Black text | Formulas and calculated values |
| Maroon text (FFC00000) | User-editable assumption inputs |
| Yellow fill (FFFFFF00) | Key assumption cells |
| Light blue fill (FFC5D9F1) | Section headers |
| Dark blue banner (FF002060) | "Actual" period label |
| Mid blue banner (FF0070C0) | "Forecast" period label |

---

## Sign Conventions

Errors here cascade through the entire model.

**P&L:** Revenue positive. COGS negative. OpEx negative. D&A negative. Interest expense negative. Tax expense negative. NPAT positive.

**Balance sheet:** Assets positive. Liabilities positive (not negative — sign applied in formulas). Equity positive.

**Cash flow:** EBITDA positive. Capex negative. WC build negative. Interest paid negative. Tax paid negative. Dividends negative. Lease principal negative. Debt drawdown positive.

---

## Cross-Sheet Formula Rules

Always use INDEX/MATCH for cross-sheet references — never direct cell references. Direct references break when columns are inserted or reordered.

### Flow vs point-in-time — critical distinction

When Annual pulls from Segments via INDEX/MATCH, the lookup method depends on the nature of the row:

**Flow items** (accumulated over the period — P&L lines, CF lines, volumes): **sum 1H + 2H**

**Point-in-time items** (a count or balance at period end — BS balances, headcount, active customers, rates like DIFOT): **2H only**

The unit column is a reliable guide: NZDm P&L and CF lines are flows; # counts, % rates, and NZDm BS balances are point-in-time. When uncertain, ask: would adding the two halves together produce a meaningful annual figure? If not, use 2H only.

### Historical 2H back-calculation
For historical years on Segments: 2H = Annual full year − 1H, via INDEX/MATCH back to Annual. Applies to flow items only — point-in-time items have 1H and 2H entered independently as actuals.

### Rows that exist only on Annual
BS and CF rows use local projection formulas and do not exist on Segments. Never attempt INDEX/MATCH to Segments for these rows.

---
## Restatements and Segment Restructures

### Restatements of prior year figures
- Never overwrite an original reported value
- Note the restatement to the user — do not update without explicit instruction
- If instructed to update, add a cell comment noting the restatement date and source

### Segment restructures
- Never overwrite prior years with a new segment structure
- Preserve prior segment rows and their Column A keys exactly
- Add new rows with new Column A keys for the new structure
- Overlapping old and new segment rows in transition years is intentional and correct
- Group-level totals must not double-sum — use period-conditional logic to switch
  the summation basis at the restructure date
- Document the restructure date and old-to-new segment mapping in a cell comment
  on the first affected row
---

## Balance Sheet and Cash Flow

**BS roll-forwards:** Each BS item rolls forward each period using the appropriate driver (revenue ratio, prior balance + movement, CF-linked). BS identity (Assets = Liabilities + Equity) must hold every column — include a BS Check row. Acceptable rounding tolerance ±0.2 due to $000 to $m conversion.

**ROU Assets and Lease Liabilities:** Both must roll forward using a New Lease Additions input row that adds equally to both. Without this, both lines decline to zero as leases unwind. Never carry either flat while including the corresponding P&L or CF charges.

**CF architecture:** Each section (OCF, CFI, CFF) has explicit component rows, a hardcoded total row for actuals, and an "Other" residual row calculated as `Total − SUM(components)`. The Other row absorbs items not explicitly modelled. For forecasts, totals are formula-driven from components. Every P&L item that is not a non-cash charge must have a corresponding CF line. CF movements must feed BS roll-forwards.

---

## Line Item Retention Policy

When repurposing for a new company, apply this rule to every row:

**Retain unchanged:** All group-level line items across P&L, Balance Sheet, Cash Flow,
and Value sheet — including all derived rows, subtotals, margins, growth rates, ratios,
and valuation bridges. Critically, this includes the "Other" residual rows in each CF
section and the "Other Assets" / "Other Liabilities" rows on the BS. These Other rows
are what allow the template to absorb any company's specific line items without
restructuring — reported totals always reconcile regardless of how a company
disaggregates them.

**Replace:** All segment-level line items. Remove the existing segments and rebuild
from the new company's reported segment structure using the Column A key conventions.

**Replace:** The entire Operating Metrics section. Remove existing KPIs and rebuild
from the new company's reported operational disclosures.

When in doubt about whether a row is group-level or segment-level, check whether it
would appear on the face of the consolidated financial statements. If yes, retain it.

### Mapping reported line items to template rows

When populating a new company's actuals, map each reported line item to the most
semantically appropriate dedicated row first. Only route to an Other row if no
dedicated row is a reasonable match. Never force an item into a mismatched dedicated
row just to avoid using Other.

If a reported item is material enough to warrant its own dedicated row that doesn't
currently exist in the template, flag it to the user before adding it rather than
silently absorbing it into Other.

---

## Repurposing Checklist

1. Build the Column A key list from the new company's filings — confirm with user
2. Update company name, ticker, and currency label (column C)
3. Replace segment-level rows using new company's reported segment structure
4. Replace Operating Metrics section with new company's reported KPIs
5. Enter actuals (blue hardcoded) for all reported periods, mapping to dedicated rows
   first and Other rows only where no dedicated row is a reasonable match
6. Flag any material items that warrant a new dedicated row — confirm with user before adding
7. Seed maroon assumption rows with initial estimates
8. Update fiscal year end dates (row 4, both sheets)
9. Update valuation date and DCF period headers on Value sheet
10. Update SOTP segment names and seed multiples

---

## Quality Gates

Before delivery, verify:
- [ ] Zero formula errors across all three sheets
- [ ] Annual forecast: flow items = 1H + 2H; point-in-time items = 2H only
- [ ] BS Check row = 0 (tolerance ±0.2)
- [ ] No blue hardcodes in formula cells; no formulas in actual-period blue cells
- [ ] All maroon assumption cells populated in forecast columns

---

## Critical Error Checklist

- WC change in DCF: pull directly from CF, do not negate again
- Interest income in P&L must also appear in CF — omitting it creates a permanent BS imbalance