# Word Table Property Topology Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Restore exact source table property topology after Word serializes equivalent table formatting into different OOXML nodes.

**Architecture:** Extend the existing post-Word deterministic table-layout cleanup. Pair structurally equivalent source/output tables and deep-copy only table, grid, row, and cell property containers while retaining all reconstructed content.

**Tech Stack:** Python 3, lxml, zipfile/DOCX OPC packages, pytest, Microsoft Word Golden harness.

## Global Constraints

- Work only on `automation-dev`; never develop on `main`.
- Keep the Golden source DOCX read-only.
- Do not globally terminate Word or close unrelated documents.
- Follow RED -> minimal fix -> GREEN -> full pytest -> real Word Golden.
- Do not weaken G3 comparison policy to hide structural differences.
- Commit and push only after measurable Golden improvement.

---

### Task 1: Reproduce Word-materialized table property topology

**Files:**
- Modify: `tests/unit/test_interactive_rebuild_service.py`

**Interfaces:**
- Consumes: `InteractiveRebuildService._restore_source_table_layout(output_path, source_path)`
- Produces: a package-level test proving exact property containers are restored without replacing output text

- [x] **Step 1: Add synthetic source/output table packages**

Build source XML with explicit `tblPr/tblBorders`, `tblGrid`, `trPr`, and `tcPr` values. Build output XML with missing nil borders, generated auto shading/cell borders, changed grid widths, and the literal text `OUTPUT CONTENT`.

- [x] **Step 2: Assert exact topology restoration**

After cleanup, serialize and compare the paired `tblPr`, `tblGrid`, `trPr`, and `tcPr` nodes. Assert the output still contains `OUTPUT CONTENT` and does not contain source text.

- [x] **Step 3: Run the targeted test and verify RED**

Run: `python -m pytest tests/unit/test_interactive_rebuild_service.py::test_source_table_property_topology_is_restored_after_word_materializes_defaults -q`

Expected: FAIL because the current cleanup does not restore property nodes that exist in both packages but differ.

Observed: FAIL with `assert 0 == 4` before the production change.

### Task 2: Restore exact property containers

**Files:**
- Modify: `src/word_replica/services/interactive_rebuild.py`
- Test: `tests/unit/test_interactive_rebuild_service.py`

**Interfaces:**
- Consumes: paired source/output OOXML table elements
- Produces: exact source `tblPr`, `tblGrid`, `trPr`, and `tcPr` in reconstructed output

- [x] **Step 1: Add a local optional-child synchronizer**

Use `copy.deepcopy` and `etree.tostring` to compare, remove, replace, or insert one property child. Return `1` only when the output changes.

- [x] **Step 2: Synchronize table, row, and cell property containers**

For each paired table synchronize `tblPr` and `tblGrid`; pair rows and synchronize `trPr`; pair cells and synchronize `tcPr`.

- [x] **Step 3: Run targeted and existing cleanup tests**

Run: `python -m pytest tests/unit/test_interactive_rebuild_service.py::test_source_table_property_topology_is_restored_after_word_materializes_defaults tests/unit/test_interactive_rebuild_service.py::test_source_autofit_table_layout_is_restored -q`

Expected: both PASS.

Observed: 2 passed.

- [x] **Step 4: Run the full service test file**

Run: `python -m pytest tests/unit/test_interactive_rebuild_service.py -q`

Expected: all tests pass.

Observed: 26 passed.

### Task 3: Verify G3 in real Word

**Files:**
- Verify: all tracked source and test files
- Read: new local `golden_report.json`

**Interfaces:**
- Consumes: green package cleanup and unchanged Golden source
- Produces: G0-G9 comparison and checkpoint only on improvement

- [x] **Step 1: Run full pytest**

Run: `python -m pytest -q`

Expected: all tests pass.

Observed: 417 passed, 52 skipped.

- [x] **Step 2: Run `RUN_GOLDEN_CODEX.ps1`**

Expected: G3 reaches PASS or materially decreases; all prior PASS gates remain PASS; source/output stay 73/73 pages; runtime remains below 45 minutes.

Observed in run `20260816T071657Z_93328a24`: G3 PASS with zero findings; G0/G1/G2/G5/G7 remained PASS; source/output remained 73/73 pages; interactive runtime was 251.50 seconds.

- [x] **Step 3: Compare report and inspect first remaining G3 divergence**

Record all gate booleans, counts, page counts, runtime, source hash, and any remaining table-property path.

Observed gates: G0 PASS, G1 PASS, G2 PASS, G3 PASS, G4 FAIL (23), G5 PASS, G6 FAIL (3), G7 PASS, G8 FAIL, G9 FAIL. Source hash remained `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`; `source_unchanged=true`; first divergence moved to G4 `0/width_emu`.

Review follow-up: two additional regressions were observed RED for mismatched nested topology and `trPr` insertion order, then GREEN after adding a pre-mutation recursive topology guard and schema-correct insertion after `tblPrEx`. Service tests: 28 passed. Full suite: 419 passed, 52 skipped. A second real Word run, `20260816T073042Z_93328a24`, again produced G3 PASS with zero findings, retained 73/73 pages and the unchanged source hash, and kept G4 as the first divergence.

- [x] **Step 4: Commit and push on measurable improvement only**

Stage only the design, plan, regression test, and minimal cleanup implementation. Commit and push only to `automation-dev`.
