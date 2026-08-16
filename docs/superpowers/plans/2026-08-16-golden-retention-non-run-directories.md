# Golden Retention Non-Run Directory Fix Plan

**Goal:** Prevent pytest and other non-Golden diagnostics directories from displacing and deleting real Golden reports.

**Design:** `docs/superpowers/specs/2026-08-16-golden-retention-non-run-directories-design.md`

## Task 1: Add and prove the regression test

**Files:**

- Modify: `tests/unit/test_codex_automation_retention.py`

Add a test with two reported failed runs and a newer non-run pytest directory. Run only that test and verify RED because the current implementation keeps the pytest directory and deletes the latest real failure.

## Task 2: Filter retention candidates

**Files:**

- Modify: `scripts/codex_automation/retention.py`

Construct candidates only from child directories containing `golden_report.json`, cache parsed reports, and use those candidates for all classification and deletion. Do not delete or return non-run directories.

Run the focused test and the complete retention test file. Inspect `git diff --check` and the complete diff.

## Task 3: Reverify and repeat Golden

Run the full local pytest suite. If green, rerun `RUN_GOLDEN_CODEX.ps1`, read the resulting report before any checkpoint, and compare field-related G2 findings with the 114-finding baseline recorded in the field-format design.

Commit/push the field fix, retention repair, tests, specs, and plans only if the valid Golden report proves measurable fidelity improvement and no previously passing gate regresses.
