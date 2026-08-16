# Word Image-Run Complex-Script Font Implementation Plan

> **For Codex:** Execute this plan under `AGENTS.md`: one proven cause, observed RED, smallest fix, observed GREEN, full pytest, then a clean real-Word Golden comparison.

**Goal:** Preserve `font_cs` on drawing-only Word runs so the four diagnosed G2 divergences at paragraphs 24, 491, 494, and 497 disappear without changing image behavior or materially affecting runtime.

**Architecture:** Keep the blueprint and event stream unchanged. Extend only the interactive `InsertImage` event handler: after Word creates an inline image, transfer the active run's complex-script font to the image range's `Font.NameBi`, then continue through the existing range and geometry flow.

**Tech stack:** Python 3.12, pytest, pywin32/Microsoft Word COM, local fake Word objects, PowerShell Golden harness.

**Design:** `docs/superpowers/specs/2026-08-16-word-image-run-complex-font-design.md`

## Task 1: Add the focused regression test and prove RED

**Files:**

- Modify: `tests/unit/test_interactive_word_executor.py`
- Test: `tests/unit/test_interactive_word_executor.py`

### Step 1: Make the fake inserted image observable

Update `FakeInlineShapes` so it retains the `FakeInlineImage` returned by `AddPicture`. The fake image's `FormattingRange.Font.NameBi` should start as `Times New Roman`, matching Word's observed default for the failing Golden runs.

Keep existing image tests compatible and do not add production-only behavior to the fake.

### Step 2: Add one regression test

Add `test_insert_image_preserves_active_complex_script_font_on_image_range`:

1. Create the controller and fake image document.
2. Execute `ApplyRunProperties` with `font_cs: Cambria`.
3. Execute an inline `InsertImage` event.
4. Assert the inserted image range has `Font.NameBi == "Cambria"`.

### Step 3: Run only the new test and verify RED

Run:

`C:\WordReplica-Automation\.venv\Scripts\python.exe -m pytest tests/unit/test_interactive_word_executor.py -q -p no:cacheprovider -k "insert_image_preserves_active_complex_script_font_on_image_range"`

Expected: one failure showing the image range remains `Times New Roman`. If it fails for fixture wiring or another reason, correct the test until it fails specifically on the missing `NameBi` transfer.

Do not change production code before this RED evidence is captured.

## Task 2: Apply the smallest production fix and prove focused GREEN

**Files:**

- Modify: `src/word_replica/renderers/interactive_word.py`
- Test: `tests/unit/test_interactive_word_executor.py`

### Step 1: Transfer only the diagnosed property

In `_event_InsertImage`, immediately after `InlineShapes.AddPicture` returns:

1. Read `font_cs` from `self._active_run_properties`.
2. If it is non-empty, obtain `inline.Range.Font` through `_retry_getattr`.
3. Set `NameBi` through `_retry_setattr` inside `suppress(Exception)`.

Do not replay the full run-property set. Do not move the existing range duplication, collapse, floating conversion, or geometry logic.

### Step 2: Run the new test

Run the focused command from Task 1.

Expected: PASS.

### Step 3: Run the entire interactive executor unit file

Run:

`C:\WordReplica-Automation\.venv\Scripts\python.exe -m pytest tests/unit/test_interactive_word_executor.py -q -p no:cacheprovider`

Expected: all tests pass, including the existing image geometry ordering test.

### Step 4: Inspect the diff

Run `git diff --check` and inspect the complete diff. Confirm the production change is limited to one optional image-range property transfer and the test support needed to prove it.

## Task 3: Run the full local regression suite

**Files:**

- Verify: repository test suite

### Step 1: Run authoritative pytest

Run the full suite using the repository's established writable temporary directory and no pytest cache:

`C:\WordReplica-Automation\.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider --basetemp C:\WordReplica-Automation\diagnostics\pytest-temp`

Expected: all collected local tests pass; skips are allowed only where already expected.

If any test fails, diagnose it before proceeding. Do not run the real Word Golden with a red suite.

## Task 4: Run a clean real-Word Golden and compare G0-G9

**Files:**

- Run: `RUN_GOLDEN_CODEX.ps1`
- Read: newest `C:\WordReplica-Automation\diagnostics\*\golden_report.json`
- Read: relevant diagnostics in the same run directory

### Step 1: Confirm safety preconditions

Confirm the branch is `automation-dev`, the worktree contains only this iteration's intended files, and no owned Golden process is already running. Do not terminate Word globally or touch unrelated documents.

### Step 2: Run the official harness

From the repository root run:

`.\RUN_GOLDEN_CODEX.ps1`

Allow the harness to manage only its own Word session. Do not modify the source DOCX.

### Step 3: Read and compare the report

Read the entire newest `golden_report.json`, not only console output. Compare it with:

`C:\WordReplica-Automation\diagnostics\20260816T012001Z_d930b2c0\golden_report.json`

Acceptance for this iteration:

- G2 mismatch count falls from 118 to at most 114.
- The four paths `24/runs/0/font_cs`, `491/runs/0/font_cs`, `494/runs/0/font_cs`, and `497/runs/0/font_cs` are absent.
- G0, G1, G5, and G7 remain PASS.
- No gate that previously passed regresses.
- Source hash remains `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`.
- Runtime has no material regression.

If the expected mismatches remain, inspect the local event trace and output OOXML before considering another code change.

## Task 5: Checkpoint only proven improvement

**Files:**

- Include: design, plan, regression test, minimal renderer fix

### Step 1: Verify final repository state

Inspect `git diff`, `git diff --check`, `git status --short`, and the current branch. Exclude unrelated user changes if any appear.

### Step 2: Commit and push only on accepted Golden evidence

If Task 4 meets acceptance, commit the intended files on `automation-dev` with a focused message and push only that branch. Record the commit hash and diagnostic report path.

If Task 4 does not improve a gate or mismatch count as predicted, do not checkpoint. Follow the fail-safe accounting in `AGENTS.md` and return to the first concrete diagnostic cause.

### Step 3: Continue the autonomous fidelity loop

This checkpoint is not completion. Re-read the accepted report, identify the next first concrete root cause, and repeat RED → minimal fix → GREEN → full suite → real Word Golden until the seminar and GOLDEN #1 satisfy their full two-run gates and runtime targets.
