# Word Inherited Color Transition Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Prevent an uncolored Word run from inheriting the previous run's explicit font color without slowing unrelated insertions.

**Architecture:** Detect only explicit-to-default color transitions while run properties are applied. Consume a one-shot reset after Word expands the range over inserted text, then reuse the existing post-insert property repair path.

**Tech Stack:** Python 3, pytest, pywin32 Word COM renderer, PowerShell Golden harness.

## Global Constraints

- Work only on `automation-dev`; never develop on `main`.
- Do not modify the Golden source DOCX.
- Do not globally terminate `WINWORD.EXE` or close unrelated Word documents.
- Follow RED -> minimal fix -> GREEN -> full pytest -> real Word Golden.
- Commit and push only after measurable Golden gate improvement.

---

### Task 1: Reproduce inherited caption color

**Files:**
- Modify: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- Consumes: `InteractiveWordController.execute_event(ReconstructionEvent(...))`
- Produces: a regression test proving only the inserted uncolored suffix loses inherited direct color

- [x] **Step 1: Add a Word-like expanded-range fake and regression test**

```python
def test_uncolored_run_clears_explicit_color_inherited_from_previous_run_after_insert():
    target = WordLikeInheritedColorRange()
    controller = InteractiveWordController.for_testing(active_range=target)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "prefix", {
        "bold": True,
        "color": "1B2F4B",
    }))
    controller.execute_event(ReconstructionEvent("InsertText", "prefix", {"text": "Tablica 1."}))
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "suffix", {"bold": True}))
    controller.execute_event(ReconstructionEvent("InsertText", "suffix", {"text": " Zakonski okvir"}))

    assert target.inserted_colors == [0x4B2F1B, None]
```

- [x] **Step 2: Run the targeted test and verify RED**

Run: `python -m pytest tests/unit/test_interactive_word_executor.py::test_uncolored_run_clears_explicit_color_inherited_from_previous_run_after_insert -q`

Expected: FAIL because the second recorded color is still `0x4B2F1B`.

### Task 2: Reset only diagnosed color transitions

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Test: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- Consumes: previous and current run-property dictionaries
- Produces: `_reset_inherited_run_color_after_insert: bool` consumed by `_apply_post_insert_run_properties`

- [x] **Step 1: Initialize one-shot transition state**

Add `_reset_inherited_run_color_after_insert = False` next to `_active_run_properties`.

- [x] **Step 2: Detect explicit-to-default color transitions**

Before replacing `_active_run_properties`, classify a color as explicit only when it is present, non-empty, and not `auto` or `none`. Set the one-shot flag when the previous properties have an explicit color and the new properties do not.

- [x] **Step 3: Consume the reset on the expanded inserted range**

At the start of `_apply_post_insert_run_properties`, clear the flag, invoke `target.Font.Reset()` once when required, and then continue through the existing explicit-property reapplication logic.

- [x] **Step 4: Run the targeted test and verify GREEN**

Run: `python -m pytest tests/unit/test_interactive_word_executor.py::test_uncolored_run_clears_explicit_color_inherited_from_previous_run_after_insert -q`

Expected: PASS.

- [x] **Step 5: Run the interactive renderer unit file**

Run: `python -m pytest tests/unit/test_interactive_word_executor.py -q`

Expected: all tests pass.

### Task 3: Verify against the complete system

**Files:**
- Verify: all tracked source and test files
- Read: latest retained `C:\WordReplica-Automation\diagnostics\<run>\golden_report.json`

**Interfaces:**
- Consumes: green local implementation and unchanged Golden source
- Produces: G0-G9 comparison and an eligible automation-dev checkpoint only if improved

- [x] **Step 1: Run the complete pytest suite**

Run: `python -m pytest -q`

Expected: all tests pass.

- [x] **Step 2: Run the real Word Golden harness**

Run: `.\RUN_GOLDEN_CODEX.ps1`

Expected: the source hash is unchanged, all previously passing gates remain passing, and G2 is below 45 findings.

- [x] **Step 3: Compare the new report with `20260816T052406Z_af539637`**

Record every G0-G9 boolean and finding count, runtime, source SHA-256, and first remaining divergence. Stop immediately on an unexplained regression of a passing gate.

- [x] **Step 4: Commit and push only if Golden measurably improves**

Stage only the design, plan, regression test, and minimal renderer fix. Commit on `automation-dev` and push `automation-dev` to `origin`.

## Verification Evidence

- RED: the suffix retained `4927259` instead of `None`.
- Targeted GREEN: 1 passed.
- Interactive renderer: 57 passed.
- Full regression: 415 passed, 52 skipped.
- Golden run: `20260816T055626Z_42089d19`.
- G2: 45 findings -> 1 finding.
- Preserved PASS gates: G0, G1, G5, G7.
- Source SHA-256 unchanged: `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`.
- Interactive runtime: 2298.56 seconds (38 minutes 18.56 seconds), below 45 minutes.

## Review Disposition

The read-only review found no critical issue. It identified broader theoretical cases involving full-font restoration, tab/line-break insertion, and range-context transitions. They are not present in the diagnosed caption event sequences, and the real Word audit introduced no new G2 finding. Under the AGENTS rule against speculative fixes, those cases are deferred until a retained diagnostic demonstrates a concrete divergence and can drive a separate RED test.
