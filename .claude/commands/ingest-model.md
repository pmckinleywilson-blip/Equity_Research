# /ingest-model

You are about to analyse an Excel financial model template and produce a skill file
that describes its architecture, conventions, and rules. The skill file must be
generalizable — it should describe how the template works, not anything specific to
the company the template was originally built for.

## Step 0 — Select template and skill target

Before doing anything else, ask the user two questions using AskUserQuestion:

### 0a — Which Excel model template?

List every `.xlsx` file in `.claude/templates/` (excluding files prefixed with `~$`).
Present them as numbered options, plus an additional option:
- **"Other — point me to a file"**

If the user selects an existing template, use that file for the ingest.

If the user selects "Other", ask them for the file path. Once provided:
1. Copy the file into `.claude/templates/` (use the original filename, converted to
   snake_case with underscores, no spaces).
2. Confirm the copy to the user, then proceed with the ingest using the new copy.

### 0b — Skill file destination

List every existing skill folder in `.claude/skills/` by name.
Present them as numbered options, plus an additional option:
- **"Create a new skill"**

If the user selects an existing skill, the ingest will overwrite that skill's `SKILL.md`.
Confirm this with the user before proceeding.

If the user selects "Create a new skill", ask them for a skill name (this becomes the
folder name, e.g. `quarterly-equity-model`). Validate that the name uses lowercase and
hyphens only, no spaces. Create the directory `.claude/skills/[skill-name]/` if it does
not exist.

Store the chosen template path and skill destination for use in later steps.

## Step 1 — Read the model

Load the uploaded Excel file fresh. Work through every sheet systematically:
- Note sheet names, count, and order
- Note column layout conventions (where data starts, what rows contain headers/dates/labels)
- Note row identification system — does the model use a Column A key system, named
  ranges, or direct cell references for cross-sheet lookups? What format do the keys
  use? What lookup method do cross-sheet formulas use (INDEX/MATCH, direct cell refs,
  VLOOKUP, named ranges)? Capture a representative formula literally. If no key system
  exists, document how references work.
- Note color coding conventions — what colors are used and what do they mean?
- Note formatting beyond color — which rows are bold? Which have borders? Identify the
  pattern. Note header row conventions (e.g. row 2 zone labels, row 3 period label fill
  colors, any other header formatting patterns observed).
- Note number format conventions — for each sheet, examine the `number_format` property
  on data cells. Document the exact Excel format string used for: (a) monetary data
  cells (e.g. `#,##0.0`), (b) percentage cells (e.g. `0.0%`), (c) per-share cells,
  (d) count/volume cells, (e) ratio cells (e.g. `0.0x`). If different sheets use
  different formats for the same type of data, document each separately.
- Note zone label positioning — if the model uses zone labels (e.g. "Actual" and
  "Forecast" labels on row 2), document where each label is positioned relative to
  the data columns. Determine the rule: does the label always appear at the first
  data column of that zone? Capture the exact column relationship.
- Note blank row conventions — are there blank spacer rows between sections? If so,
  how many and where? Document the pattern (e.g. "one blank row between each major
  section, no blank rows within a section"). If there are no spacer rows, state this.
- Where multiple sheets share similar row structures (e.g. a sub-period sheet and an
  annual sheet both containing P&L line items), note whether the line items, their
  order, and their Column A keys are identical across sheets. Document any structural
  correspondence observed — this is important for ensuring consistency when
  repurposing.
- Note sign conventions from the actual data values
- Note formula patterns — how do sheets reference each other?
- If sub-annual periods exist, examine how each sub-period's values are determined:
  independently hardcoded, derived from annual (what formula?), or something else.
  Capture the literal formula from a representative cell.
- Note what is hardcoded vs formula-driven
- Note any assumption input rows and how they are distinguished. Are inputs adjacent
  to calculated rows or centralized? Document the convention.
- Determine whether the template enforces assumption visibility — are ALL forecast
  assumptions surfaced on their own dedicated rows (e.g. a "Volume Growth" row below
  the "Volume" row it drives), or are some assumptions embedded within formulas
  (e.g. `=PCP*1.025` where 1.025 is a hidden assumption)? If the template uses
  dedicated rows for all assumptions, document this as a core architectural principle.
  Capture the exact pattern: assumption row sits adjacent to the row it drives, the
  driven row references the assumption row by cell reference, and no forecast formula
  contains a hardcoded assumption value.
- Note CF architecture — walk every CF row. For each: note key, label, whether it's
  a component/subtotal/analytical/residual. Note the CF format (EBITDA-based or
  Receipts/Payments). **For each forecast formula in the CF:** document what it links
  to and how it derives its value (e.g. "Capex = Capex/Sales % × Revenue",
  "WC Change = -(Receivables change) - (Inventory change) + (Payables change)",
  "Tax Paid = Tax expense + adjustment for prior year tax payable"). Note any
  analytical rows after the main CF (Gross OCF, Cash Conversion, Operating FCF,
  FCF per Share, FCF Yield, FCF Margin).
- Note BS structure — walk every BS row. For each forecast formula, document the
  roll-forward logic (e.g. "PPE = Prior PPE + Capex − Depreciation",
  "Lease Liabilities = Prior + New Lease Additions − Principal Payments",
  "Cash = Prior + Net Change in Cash from CF"). Note which BS items are driven by
  revenue ratios, which roll forward with a movement schedule, and which are CF-linked.
- Note any ratio/return sections after the CF (ROIC, Invested Capital, ROFE, NOPAT).
  Capture row structure and formulas.
- Note sheet spatial architecture — for each sheet, identify distinct zones (e.g.
  consolidated summary at top, segment driver sections below). Note row ranges, zone
  purposes, and how zones reference each other in forecasts. For each zone, determine
  whether it is a **source zone** (contains the primary forecast driver logic) or a
  **dependent zone** (its forecast cells reference the source zone's outputs rather
  than containing independent forecast logic). Document the dependency direction
  explicitly — a dependent zone's forecast cells must never contain independent
  forecast formulas; they must only contain references to the source zone's output
  rows. Capture a representative formula showing how the dependent zone pulls from
  the source zone.
- Note valuation methods — identify every valuation method present in the template.
  These may be on a dedicated Value/Valuation sheet or embedded within other sheets.
  Common methods include DCF, EV/EBITDA SOTP, trading comps, precedent transactions,
  DDM, NAV, or combinations. For each valuation method found, document:
  - Where it is located (sheet name, row range)
  - The structure: input assumptions (WACC, multiples, growth rates), calculation
    mechanics (discount factors, terminal value, equity bridge), and outputs
    (per-share value, upside/downside)
  - The literal formulas for key calculation steps (e.g. how FCFF is built, how
    the discount factor is calculated, how the equity bridge works from EV to
    per-share value)
  - What data it pulls from other sheets and via what formula method (INDEX/MATCH,
    direct refs)
  - Any user-editable inputs (multiples, WACC components, valuation date) and how
    they are distinguished (maroon text, yellow fill, etc.)
  - If the template has an SOTP: note segment names, how segment EBITDA is sourced,
    how multiples are applied, how segment EVs aggregate, and how the equity bridge
    flows

## Step 2 — Ask clarifying questions

Before writing the skill file, ask the user:
1. What reporting frequency is this template designed for? (Half-year, quarterly,
   annual only)
2. Are there any conventions in the model that are not visible from the file itself?
3. I have identified the following analytical and ratio sections [list them — e.g.
   Gross OCF, Cash Conversion, OpFCF, FCF Yield, ROIC, etc.]. Are there any you
   would NOT like to retain when repurposing?
4. I have identified the following valuation methods: [list them — e.g. DCF, SOTP,
   etc.]. Are there any you would NOT like to retain, or any additional methods
   to add?
5. Are there any critical errors from your experience with this template that should
   be added to the error checklist?

## Step 2.5 — Confirm row retention

Present all rows from each sheet grouped by section. All rows start **unchecked**
(default = REPLACE). The user **selects** any rows they want to RETAIN.

Show: row number, Column A key (if any), Column B label, and a flag for
"analytical/ratio" rows that the ingest has identified.

Frame the question as: *"Below are all rows grouped by section. By default every row
will be marked as REPLACE (company-specific, to be rebuilt for each new company).
Select any rows that should be RETAINED as-is when repurposing the template.
If none need retaining in a section, select 'None — replace all in this section'."*

Use AskUserQuestion with multiSelect per section group. Include a **"None — replace
all in this section"** option at the top of each group so that the user can
explicitly confirm they want everything replaced without leaving the question
unanswered. If the user selects "None", treat all rows in that section as REPLACE.
If the user selects specific rows, those become RETAIN; everything else is REPLACE.

The output becomes the explicit retain/replace list in the skill file.

## Step 3 — Produce the skill file

Write a SKILL.md using only the sections that apply to this model. Use these section
headers where relevant, omitting any that do not apply:

- Template Reference
- Sheet Architecture
- Sheet Zone Architecture (only if distinct spatial zones exist on any sheet —
  document which zone is the source and which is dependent, the dependency
  direction, and a representative formula showing how the dependent zone
  references the source zone. State explicitly that the dependent zone's forecast
  cells must only contain references to the source zone, never independent
  forecast logic)
- Column Layout
- Row Identification System (only if the model uses a key-based lookup system)
- Row Lookup System (document the key format and lookup method observed, with a
  representative formula)
- Color Coding (only if the model uses a consistent color convention)
- Row Formatting Rules (bold/border pattern, header row conventions, per-section
  formatting if different sections follow different patterns)
- Number Format Conventions (document the exact Excel number_format string for each
  cell type — monetary, percentage, per-share, count — per sheet if they differ)
- Zone Label Positioning (document the rule for where zone labels appear relative to
  data columns — e.g. "the Forecast label is always placed at the first forecast
  data column")
- Blank Row Convention (document the spacer row pattern between sections)
- Cross-Sheet Structural Correspondence (only if multiple sheets share the same row
  structure — document which sheets must have identical line items and keys, and the
  rule that changes to one must be reflected in the other)
- Sign Conventions
- Cross-Sheet Formula Rules (only if sheets reference each other)
- Sub-Period Derivation (document the exact formula pattern observed for deriving
  sub-periods — literal Excel formula included)
- Assumption Input Placement (document the observed convention, including whether
  the template enforces that all assumptions must be on dedicated rows with no
  hidden assumptions embedded in formulas)
- Flow vs Point-in-Time (only if the model has sub-annual periods that roll up)
- Cash Flow Section Structure (exact row order, CF format type, and for each
  forecast formula: what it links to and how the projection works)
- Balance Sheet Projection Methods (for each BS line item: the roll-forward formula
  and what drives it — revenue ratios, movement schedules, CF linkages)
- Return Metrics (ROIC etc. if they exist — rows and formulas)
- Valuation Methods (for each method found: location, structure, input assumptions,
  calculation mechanics with literal formulas, data sources, user-editable inputs.
  If SOTP: segment structure, multiple application, equity bridge)
- Line Item Retention Policy
- Template Preservation Method (the retain/replace list from Step 2.5, plus the
  modify-in-place rule)
- Repurposing Checklist
- Quality Gates
- Critical Error Checklist (only include errors the user has confirmed or that are
  clearly evidenced in the model's formula structure)

## Skill file rules

- Every section must reflect what is actually in the model — do not invent conventions
- Nothing company-specific (no ticker names, segment names, currencies unless the
  currency is a template default the user confirms should be retained)
- The YAML description must clearly state what reporting frequency the template is
  designed for and when to use this skill vs others
- The skill file must contain enough structural detail that someone repurposing the
  template could reproduce its exact layout, formatting, and formula patterns without
  opening the original template.
- Do not assume any convention is standard. Document what this template does,
  including things that might seem obvious.
- CF and BS mechanics (formula linkages, roll-forward logic, projection methods)
  must be documented in the skill file — these are template-specific and must be
  preserved when repurposing. P&L forecast methods (growth rates, segment drivers)
  are NOT documented — those are determined at build time.
- Include literal Excel formulas where the pattern is model-specific.
- After producing the draft, ask the user to confirm before finalising

## Step 4 — Confirm and write

Present the draft skill file to the user. Ask:
- Does this accurately reflect your template's conventions?
- Are there any sections that should be added, removed, or amended?
- Are there any critical errors from your experience building with this template
  that should be added to the checklist?

Incorporate feedback, then:
1. Use the skill destination chosen in Step 0b (the directory already exists)
2. Write the final skill file to `.claude/skills/[skill-name]/SKILL.md`
3. Confirm to the user that the file has been written and provide the full path
4. Prompt the user with the next step:
   - **To build a model using this template:** run `/build-model [TICKER]`
     where `[TICKER]` is the company ticker with exchange suffix (e.g. BHP.AX)
