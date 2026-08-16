# Word Field Exit Range Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the active Word insertion range outside each newly created field so field-end markers remain in their source paragraphs.

**Architecture:** Preserve the existing cached-result and formatting workflow. After collapsing the duplicated field-result range, move that range one Word character past the field-end marker and continue reconstruction there.

**Tech Stack:** Python 3, pytest, pywin32 Word COM renderer, PowerShell Golden harness.

## Global Constraints

- Work only on `automation-dev`; never develop on `main`.
- Do not modify the Golden source DOCX.
- Do not globally terminate `WINWORD.EXE` or close unrelated Word documents.
- Follow RED -> minimal fix -> GREEN -> full pytest -> real Word Golden.
- Commit and push only after measurable Golden gate improvement.

---

### Task 1: Reproduce the field-exit cursor bug

**Files:**
- Modify: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- Consumes: `InteractiveWordController.execute_event(ReconstructionEvent("CreateField", ...))`
- Produces: a test proving the active range moves exactly one character beyond the field result

- [x] **Step 1: Add a Word-like field-result range test**

```python
def test_create_field_moves_active_range_past_field_end_marker():
    result = WordLikeFieldResult(start=10, end=14)
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = WordLikeFieldDocument(result)

    controller.execute_event(ReconstructionEvent("CreateField", "field", {
        "instruction": "REF ref_tab_1 \\h",
        "cached_result": "1",
    }))

    assert result.collapse_calls == [0]
    assert result.move_calls == [(1, 1)]
    assert (controller.active_range.Start, controller.active_range.End) == (15, 15)
```

- [x] **Step 2: Run the targeted test and verify RED**

Run: `python -m pytest tests/unit/test_interactive_word_executor.py::test_create_field_moves_active_range_past_field_end_marker -q`

Expected: FAIL because the current renderer never calls `Move` and leaves the range at position 14.

### Task 2: Exit the field after cached-result formatting

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Test: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- Consumes: collapsed duplicate of `field.Result`
- Produces: active collapsed range one `WD_CHARACTER` beyond the result end

- [x] **Step 1: Add the Word character unit constant**

Add `WD_CHARACTER = 1` beside the existing Word constants.

- [x] **Step 2: Move the collapsed result range one character**

After `result.Collapse(WD_COLLAPSE_END)`, retrieve callable `result.Move` and invoke it with positional arguments `(WD_CHARACTER, 1)`. Raise a clear runtime error if `Move` is unavailable.

- [x] **Step 3: Run the targeted test and verify GREEN**

Run: `python -m pytest tests/unit/test_interactive_word_executor.py::test_create_field_moves_active_range_past_field_end_marker -q`

Expected: PASS.

- [x] **Step 4: Run field and renderer regression tests**

Run: `python -m pytest tests/unit/test_interactive_word_executor.py -q`

Expected: all tests pass, including cached-result formatting.

### Task 3: Verify the real document

**Files:**
- Verify: all tracked source and test files
- Read: the new local `C:\WordReplica-Automation\diagnostics\<run>\golden_report.json`

**Interfaces:**
- Consumes: green field-exit implementation and unchanged Golden source
- Produces: a G0-G9 comparison and eligible automation-dev checkpoint only if improved

- [x] **Step 1: Run the full pytest suite**

Run: `python -m pytest -q`

Expected: all tests pass.

- [x] **Step 2: Run the real Word Golden harness**

Run: `.\RUN_GOLDEN_CODEX.ps1`

Expected: G2 passes, G7 stays passing, existing PASS gates do not regress, and runtime stays below 45 minutes.

- [x] **Step 3: Inspect field-end distribution and compare G0-G9**

Parse source and output with `DocxParser`, confirm output field-end markers no longer all accumulate in paragraph 1369, and record every gate boolean, finding count, runtime, and source hash.

- [x] **Step 4: Commit and push only on measurable improvement**

Stage only this design, plan, regression test, and minimal renderer fix. Commit and push only to `automation-dev`.

## Verification Evidence

- RED: `move_calls` was empty and the collapsed range remained before the field-end marker.
- Targeted GREEN: 2 passed, including cached-result formatting.
- Interactive renderer: 58 passed.
- Full regression: 416 passed, 52 skipped.
- Golden run: `20260816T065317Z_253a8e71`.
- G2: 1 finding -> PASS with 0 findings.
- G0, G1, G5, and G7 remained PASS; automation score increased from 4 to 5.
- Page count: 11 output pages -> 73 output pages, now matching the 73-page source.
- Field-end distribution: source and output both have 26 markers at paragraph indexes `[173, 222, 300, 356, 419, 490, 502, 503, 517, 524, 528, 593, 594, 633, 711, 774, 1104]`.
- Interactive runtime: 2298.56 seconds -> 339.54 seconds (5 minutes 39.54 seconds).
- Source SHA-256 unchanged: `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`.
