---
name: fix-root-cause-not-symptoms
description: When model output has issues, fix the upstream command/skill that produces the output, not the output itself. Follow the chain: ingest-model → skill file → build-model → model output.
type: feedback
---

When the model output has structural or formatting issues, the user wants to fix the root cause in the command/skill chain, not patch the output directly. The chain is: ingest-model command → skill file → build-model command → model output. Fix at the earliest point in the chain where the problem originates.

**Why:** The ingest-model command reads any user's template and produces a skill file. If the ingest command doesn't capture enough detail, the skill file will be incomplete, and the build-model command won't have sufficient instructions to produce correct output. Fixing the skill file manually is a symptom fix; fixing the ingest command is the root cause fix.

**How to apply:** When model output issues arise, trace back to which command in the chain should have captured the missing information, and fix that command first.
