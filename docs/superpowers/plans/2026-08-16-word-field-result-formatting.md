# Word Field-Result Formatting Implementation Plan

> **For Codex:** Execute under `AGENTS.md`: observed RED, smallest evidence-backed fix, observed GREEN, full pytest, then one clean real-Word Golden comparison.

**Goal:** Stop Word REF results from inheriting target-bookmark formatting and remove the diagnosed field-related G2 run splits.

**Architecture:** Keep compilation, field instructions, cached text, and update policy unchanged. Reuse the interactive renderer's existing post-insert run formatter on `field.Result` before collapsing that result into the continuation range.

**Tech stack:** Python 3.12, pytest, pywin32/Microsoft Word COM, local fake Word objects, PowerShell Golden harness.

**Design:** `docs/superpowers/specs/2026-08-16-word-field-result-formatting-design.md`

## Task 1: Strengthen the cached-field regression test and prove RED

**Files:**

- Modify: `tests/unit/test_interactive_word_executor.py`

### Step 1: Model Word's inherited field formatting

Extend the local `ResultRange` in the cached-field test so it supports Word font properties and begins with `Bold = -1` and `NameBi = "Times New Roman"`.

Use a formatting-capable active range. Execute `ApplyRunProperties` with at least `bold: false`, `font_ascii: Times New Roman`, and `font_cs: Cambria` before `CreateField`.

### Step 2: Assert the source formatting contract

Keep the current assertions that cached text is seeded, the field is not updated, and the active range is the field result. Add assertions that the result has Word false bold and `NameBi == "Cambria"`.

Rename the test if needed so its name covers both cached text and formatting.

### Step 3: Run the focused test and prove RED

Run:

`C:\WordReplica-Automation\.venv\Scripts\python.exe -m pytest tests/unit/test_interactive_word_executor.py -q -p no:cacheprovider -k "create_field_seeds_cached_result"`

Expected: failure specifically because the result remains bold and/or retains Times New Roman complex-script font. Correct fixture wiring if necessary, but do not change production code before that behavioral RED is observed.

## Task 2: Reapply active run properties to the field result

**Files:**

- Modify: `src/word_replica/renderers/interactive_word.py`
- Test: `tests/unit/test_interactive_word_executor.py`

### Step 1: Make the result range the formatting target

In `_event_CreateField`, after the existing cached-result assignment:

1. duplicate `field.Result` as today;
2. when cached text and active run properties are both present, assign that duplicate to `self.active_range`;
3. call `_apply_post_insert_run_properties(self._active_run_properties, cached_text)`;
4. collapse the same result range and retain it as `self.active_range`.

Do not call `field.Update`, change `PreserveFormatting`, alter instructions, or add a save.

### Step 2: Run focused GREEN

Run the Task 1 focused command. Expected: PASS.

### Step 3: Run the entire interactive executor file

Run:

`C:\WordReplica-Automation\.venv\Scripts\python.exe -m pytest tests/unit/test_interactive_word_executor.py -q -p no:cacheprovider`

Expected: all tests pass, including field continuation and prior image/page-break protections.

### Step 4: Inspect the complete diff

Run `git diff --check` and inspect the full production/test diff. Confirm the handler still collapses exactly the field result range used for formatting and that no unrelated event path changed.

## Task 3: Run the full local regression suite

**Files:**

- Verify: repository test suite

Run:

`C:\WordReplica-Automation\.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider --basetemp C:\WordReplica-Automation\diagnostics\pytest-full-field-format`

Expected: all collected tests pass with only established skips. Do not run real Word if the suite is red.

## Task 4: Run and analyze one clean real-Word Golden

**Files:**

- Run: `RUN_GOLDEN_CODEX.ps1`
- Read: newest `C:\WordReplica-Automation\diagnostics\*\golden_report.json`
- Read: run-local trace, parsed G2 findings, performance profile, and source integrity evidence

### Step 1: Confirm preconditions

Confirm branch `automation-dev`, intended worktree scope, unchanged source hash, and no active owned Golden harness. Do not inspect or close unrelated Word documents beyond process ownership metadata.

### Step 2: Run the official harness

From the repository root run:

`.\RUN_GOLDEN_CODEX.ps1`

### Step 3: Compare against the accepted baseline

Baseline report:

`C:\WordReplica-Automation\diagnostics\20260816T023116Z_85b73369\golden_report.json`

Required evidence:

- G2 is below 114 and first divergence is no longer `173/runs/0/text_len`;
- diagnose all 17 field-containing paragraph groups, not only the first one;
- no new G0, G1, G5, or G7 regression;
- source remains unchanged with the required hash;
- performance profile shows no material runtime regression.

Expected best case: up to 69 field-related findings disappear, moving G2 from 114 toward 45.

## Task 5: Checkpoint only accepted improvement

**Files:**

- Include: design, plan, regression test, minimal renderer fix

If the Golden evidence meets Task 4 acceptance, run final verification, commit the exact intended files on `automation-dev`, and push only `origin/automation-dev`. Record commit hash and diagnostic path.

If it does not improve as predicted, do not commit. Follow fail-safe accounting and inspect the field-result OOXML/event evidence before considering another fix.

After a successful checkpoint, continue from the next first concrete divergence until seminar and GOLDEN #1 satisfy their full fidelity and runtime completion gates.
