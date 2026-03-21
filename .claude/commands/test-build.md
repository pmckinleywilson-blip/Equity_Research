---
description: >
  Automated test loop: rebuild a model from template using saved fixture
  decisions, validate against template, diagnose issues, and suggest fixes
  to the ingest-model or build-model commands.
  Run: /test-build [skill-name] [TICKER]
  If TICKER is omitted and multiple fixtures exist, you will be prompted to choose.
---

# /test-build

Run the automated test loop for skill: $ARGUMENTS

---

## Step 1 — Load prerequisites

1. Load the skill file from `.claude/skills/[skill-name]/SKILL.md`
2. Locate the test fixture:
   - If a TICKER was provided as the second argument, load
     `.claude/skills/[skill-name]/test-fixture-[TICKER].yml`
   - If no TICKER was provided, scan for all `test-fixture-*.yml` files in
     `.claude/skills/[skill-name]/`
   - If exactly one fixture exists, use it automatically
   - If multiple fixtures exist, list them and ask the user to choose
   - If no fixtures exist, stop and tell the user: "No test fixtures found.
     Run `/build-model [TICKER]` to build a model interactively, then
     `/save-fixture [TICKER]` to capture the decisions."
3. Read the template file path from the skill file's Template Reference.
4. Read the test company's ticker and details from the fixture.
5. Confirm to the user: "Ready to test [skill-name] using [TICKER] as the
   test company. This will rebuild the model from scratch, validate it against
   the template, and report any issues. Proceed?"

---

## Step 2 — Rebuild from scratch

Using the fixture's saved decisions (do NOT ask the user any interactive
questions — all answers come from the fixture):

1. Delete the existing model file at `[TICKER]/Models/[TICKER] Model.xlsx`
   if it exists (confirm with user first)
2. Copy the template to `[TICKER]/Models/[TICKER] Model.xlsx`
3. Load the skill file for all template conventions
4. Execute the Phase 5 build process from `/build-model`, using:
   - Segment structure from `fixture.build.segments`
   - Forecast drivers from `fixture.build.forecast_drivers`
   - Group-level structure from `fixture.build.group_level`
   - Key decisions from `fixture.build.decisions`
   - Source documents from `[TICKER]/Company reports/`
   - All actual periods from `fixture.validation.annual_actuals` and
     `fixture.validation.hy_actuals`
5. Execute Phase 6 verification checks and record results
6. Save the model

---

## Step 3 — Validate

Run the full `/validate-model` comparison:
- Compare the rebuilt model against the source template
- Check all formatting, structure, formula, and skill file compliance items
- Produce the categorized issue list (CRITICAL / MAJOR / MODERATE / MINOR)

---

## Step 4 — Diagnose

For each issue found, determine the root cause:

### Category A: Skill file gap
The skill file does not document a convention that the template uses. The
ingest-model command failed to capture it.
- **Evidence**: The template has a pattern (formatting, formula, structure)
  that is not mentioned anywhere in the skill file
- **Fix**: Propose a specific addition to the ingest-model command that would
  cause it to capture this convention, then propose the corresponding skill
  file section content

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
  Checklist so future builds are warned. Consider whether the build command
  needs stronger language or a more explicit verification step.

### Generalizability check
Before proposing any fix, verify it is NOT company-specific or template-
specific:
- Does the proposed change reference any company name, ticker, segment name,
  sheet name, column letter, row number, or formula pattern that is specific
  to one template?
- If yes, the change belongs in the skill file, not the command
- If no, it can go in the command

---

## Step 5 — Report and propose fixes

Present the results to the user in this format:

```
=== TEST BUILD REPORT ===
Skill: [skill-name]
Test company: [TICKER]
Template: [template file]

VALIDATION RESULTS:
- Critical: [count]
- Major: [count]
- Moderate: [count]
- Minor: [count]

DIAGNOSED ISSUES:
[For each issue:]
  Issue: [description]
  Category: [A/B/C]
  Location: [sheet, row, col]
  Expected: [what template/skill file says]
  Found: [what model has]
  Proposed fix: [specific change to ingest-model, build-model, or skill file]

PROPOSED COMMAND CHANGES:
[List all proposed changes grouped by file:]
  ingest-model.md:
    - [change 1]
    - [change 2]
  build-model.md:
    - [change 1]
  [skill-name]/SKILL.md:
    - [change 1]
```

Ask the user: "Would you like me to apply these fixes and re-run the test?"

---

## Step 6 — Iterate (if user approves)

If the user approves:
1. Apply the proposed changes to the command and/or skill files
2. Go back to Step 2 and rebuild
3. Maximum 3 iterations to prevent infinite loops
4. After each iteration, report only NEW issues (don't re-report fixed ones)

If the user declines or after max iterations:
1. Present a final summary of all issues resolved and remaining
2. List all changes made to command files and skill files
3. Prompt the user:
   - **To test with a different company** (generalizability check): "Place
     source documents in `[TICKER2]/Company reports/`, create a fixture with
     `/save-fixture [TICKER2]`, then run `/test-build [skill-name]` again."
   - **To build a production model:** run `/build-model [TICKER]`
