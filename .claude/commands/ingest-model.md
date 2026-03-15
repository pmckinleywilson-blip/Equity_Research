# /ingest-model

You are about to analyse an Excel financial model template and produce a skill file
that describes its architecture, conventions, and rules. The skill file must be
generalizable — it should describe how the template works, not anything specific to
the company the template was originally built for.

## Step 1 — Read the model

Load the uploaded Excel file fresh. Work through every sheet systematically:
- Note sheet names, count, and order
- Note column layout conventions (where data starts, what rows contain headers/dates/labels)
- Note row identification system — does the model use a Column A key system, named
  ranges, or direct cell references for cross-sheet lookups?
- Note color coding conventions — what colors are used and what do they mean?
- Note sign conventions from the actual data values
- Note formula patterns — how do sheets reference each other?
- Note what is hardcoded vs formula-driven
- Note any assumption input rows and how they are distinguished
- Note CF architecture — are there "Other" residual rows?
- Note BS structure — are there roll-forward mechanics?
- Note any valuation sheet structure

## Step 2 — Ask clarifying questions

Before writing the skill file, ask the user:
1. What reporting frequency is this template designed for? (Half-year, quarterly, annual only)
2. Are there any conventions in the model that are not visible from the file itself?
3. Are there any sections of the model that are company-specific and should be flagged
   as "replace" rather than "retain" when repurposing?
4. Are there any critical errors the user has encountered when using this template that
   should be added to the error checklist?

## Step 3 — Produce the skill file

Write a SKILL.md using only the sections that apply to this model. Use these section
headers where relevant, omitting any that do not apply:

- Template Reference
- Sheet Architecture
- Column Layout
- Row Identification System (only if the model uses a key-based lookup system)
- Color Coding (only if the model uses a consistent color convention)
- Sign Conventions
- Cross-Sheet Formula Rules (only if sheets reference each other)
- Flow vs Point-in-Time (only if the model has sub-annual periods that roll up)
- Balance Sheet and Cash Flow (only if the model contains a full 3-statement structure)
- Line Item Retention Policy
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
- Keep it lean — only include what Claude would not know from general financial
  modelling knowledge. Do not document standard finance or standard Excel behaviour
- After producing the draft, ask the user to confirm before finalising

## Step 4 — Confirm and write

Present the draft skill file to the user. Ask:
- Does this accurately reflect your template's conventions?
- Are there any sections that should be added, removed, or amended?
- Are there any critical errors from your experience building with this template
  that should be added to the checklist?

Incorporate feedback, then:
1. Ask the user to confirm the skill name (this becomes the folder name, e.g.
   `hy-equity-model`)
2. Create the directory `.claude/skills/[skill-name]/` if it does not exist
3. Write the final skill file to `.claude/skills/[skill-name]/SKILL.md`
4. Confirm to the user that the file has been written and provide the full path