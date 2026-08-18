# Unexpected Header Shape Defaults Cleanup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove only Word-generated `w:hdrShapeDefaults` elements absent from the source so GOLDEN #1 G9 improves without new COM calls or seminar runtime regression.

**Architecture:** Add one deterministic post-Word package cleanup to `InteractiveRebuildService`. The helper compares source and output settings topology, rewrites only `word/settings.xml` when required, and is wired into existing final checkpoint and audit bookkeeping.

**Tech Stack:** Python 3.12, `zipfile`, `tempfile`, `lxml`, pytest, Microsoft Word COM Golden audit.

## Global Constraints

- Work only on `automation-dev`; never modify `main`.
- Never modify the Golden or seminar source DOCX.
- Never globally terminate Word or close unrelated Word documents.
- Add no Word COM calls and weaken no G0-G9 threshold.
- Follow RED -> minimal fix -> GREEN -> full pytest -> real Word Golden.
- Commit and push only after the real Golden measurably improves.

---

### Task 1: Prove and implement conditional settings cleanup

**Files:**
- Modify: `tests/unit/test_interactive_rebuild_service.py`
- Modify: `src/word_replica/services/interactive_rebuild.py`

**Interfaces:**
- Consumes: source and output DOCX paths after the owned Word renderer closes.
- Produces: `InteractiveRebuildService._remove_unexpected_header_shape_defaults(output_path: Path, source_path: Path) -> int`.

- [x] **Step 1: Write the failing package behavior tests**

Add tests that build minimal source/output ZIP packages with `word/settings.xml`. The first source omits `w:hdrShapeDefaults`; its output contains one plus `w:compat`. Assert the helper returns `1`, removes only the unexpected element, and preserves `w:compat`. The second source contains `w:hdrShapeDefaults`; assert return `0` and unchanged output bytes.

- [x] **Step 2: Run the targeted tests and verify RED**

Run:

```powershell
python -m pytest tests/unit/test_interactive_rebuild_service.py -k "header_shape_defaults" -q -p no:cacheprovider --basetemp C:\Users\PC\Documents\Codex\2026-08-12\read-agents-md-completely-before-doing\work\pytest-red-header-shape-defaults
```

Expected: FAIL because `_remove_unexpected_header_shape_defaults` does not exist.

- [x] **Step 3: Implement the minimal package cleanup**

In `InteractiveRebuildService`, read both settings parts, return zero when the source contains a direct `w:hdrShapeDefaults`, remove all direct output occurrences otherwise, and use the established temporary-ZIP replacement pattern only when the removed count is nonzero.

- [x] **Step 4: Run targeted tests and verify GREEN**

Run the Step 2 command again. Expected: both tests PASS.

### Task 2: Wire cleanup into finalization bookkeeping

**Files:**
- Modify: `tests/unit/test_interactive_rebuild_service.py`
- Modify: `src/word_replica/services/interactive_rebuild.py`

**Interfaces:**
- Consumes: `_remove_unexpected_header_shape_defaults(...) -> int` from Task 1.
- Produces: final output without unexpected header shape defaults, refreshed checkpoint SHA, and `UNEXPECTED_HEADER_SHAPE_DEFAULTS_REMOVED` audit evidence.

- [x] **Step 1: Write a failing finalization test**

Extend a real `_finalize_qa` fixture with a renderer that saves an otherwise valid source copy and injects one `w:hdrShapeDefaults` into output settings. Assert the finalized output lacks the element and the audit contains `UNEXPECTED_HEADER_SHAPE_DEFAULTS_REMOVED` with count `1`.

- [x] **Step 2: Verify RED**

Run the single finalization regression and confirm the unexpected element remains before wiring.

- [x] **Step 3: Wire the helper**

Call the helper immediately after the renderer closes. Include its count in the condition that refreshes `final_checkpoint.output_sha256`, and append the audit event only when the count is nonzero.

- [x] **Step 4: Verify GREEN and focused regressions**

Run all new tests plus `tests/unit/test_interactive_rebuild_service.py`. Expected: all PASS.

### Task 3: Full verification and real Word Golden

**Files:**
- Verify: entire repository
- Read: newest `C:\WordReplica-Automation\diagnostics\<run_id>\golden_report.json`

**Interfaces:**
- Consumes: green implementation from Tasks 1-2.
- Produces: authoritative G0-G9 and runtime evidence.

- [x] **Step 1: Run the full suite**

```powershell
python -m pytest -q -p no:cacheprovider --basetemp C:\Users\PC\Documents\Codex\2026-08-12\read-agents-md-completely-before-doing\work\pytest-full-header-shape-defaults
```

Expected: no failures.

- [x] **Step 2: Request read-only code review**

Require no Critical or Important findings before the real Word run.

- [x] **Step 3: Run the Golden**

```powershell
.\RUN_GOLDEN_CODEX.ps1
```

Read the complete generated report. Require unchanged source SHA, G0-G8 PASS, runtime below 45 minutes, and G9 metrics measurably better than run `20260816T093134Z_687e3788`.

- [x] **Step 4: Checkpoint only on measured improvement**

If the Golden improves, commit source, tests, spec, and plan on `automation-dev`, then push that branch. If it does not improve or any earlier gate regresses, revert the failed hypothesis and stop under AGENTS.md.

- [ ] **Step 5: Continue the fidelity loop**

Read the next first G9 divergence and repeat RED -> minimal fix -> GREEN -> full suite -> Golden until two consecutive G0-G9 FULL PASS runs occur on the same commit.

Checkpoint evidence: full suite `441 passed, 52 skipped`; read-only review found no Critical or Important issues. Real Word run `20260816T104156Z_a0dd263b` preserved source SHA-256, passed G0-G8, completed reconstruction in `252.1037s`, removed one unexpected `w:hdrShapeDefaults`, and moved the first G9 divergence from page 7 to page 21 (`0.026680374649406043`, MAE `1.004742616336698`).
