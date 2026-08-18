# WordReplica Golden #1 Full Fidelity and Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make GOLDEN #1 reproduce the source document with G0-G9 FULL PASS twice consecutively on the same commit and unchanged source SHA-256, while reducing one complete real-Microsoft-Word Golden run to 45 minutes or less on this machine.

**Architecture:** Keep the existing event-driven Word COM renderer and the G0-G9 audit as the acceptance authority. Use retained event traces to isolate performance work in short real-Word benchmarks, then return to the mandatory first-divergence TDD loop for fidelity. No optimization is accepted unless its benchmark improves and its reconstructed DOCX remains equivalent under the relevant fidelity projections.

**Tech Stack:** Python 3.12 automation virtual environment, pytest, pywin32/Microsoft Word COM, OOXML/DOCX parsing, PowerShell launchers, local JSON diagnostics.

## Global Constraints

- Work only on `automation-dev`; never develop directly on `main`.
- Never modify the Golden source DOCX.
- Never globally terminate `WINWORD.EXE` or close unrelated Word documents.
- Terminate a Word PID only when ownership PID and process-creation FILETIME are proven.
- Every production behavior change requires observed RED, smallest fix, observed GREEN, full pytest, and a real Word Golden comparison.
- Stop on the fail-safe rules in `AGENTS.md`.
- Do not commit failed hypotheses.
- Create and push a checkpoint only after the full suite is green and Golden evidence measurably improves.
- Do not promote to `main` until two consecutive G0-G9 FULL PASS runs occur on the same commit with unchanged source SHA-256.

## Definition of Done

All conditions below are mandatory:

1. G0-G9 are PASS in two consecutive real-Microsoft-Word runs.
2. Both runs use the same commit and source SHA-256 `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8` (or a deliberately re-baselined hash approved by the user; the source itself is never edited by automation).
3. Source and output page counts agree and G8 page text partitioning passes.
4. Bold, font, run boundaries, paragraph formatting, tables, images, fields, headers/footers, sections, pagination, and visual comparison pass their existing gates without weakening tolerances to hide defects.
5. Each complete run, including reconstruction and audit, finishes in at most 45 minutes on this machine.
6. The full local regression suite is green immediately before both acceptance runs.
7. `golden_report.json` says `automation_decision.promotion_ready=true` before any promotion to `main`.

## Measured Baseline (Run `20260813T204758Z_5ed0c311`)

- Full pytest: 344 passed, 51 skipped.
- Reconstruction: 8,870.94 seconds (147.85 minutes).
- Audit: 31.24 seconds.
- G0 PASS, G1 PASS, G2-G9 FAIL.
- Pages: source 73, output 11.
- First divergence: G2 `1/runs/0/character_spacing`, expected `18`, actual `17`.
- G2 mismatches: 1,601; G3 mismatches: 1,453.
- Source SHA-256 unchanged.
- Trace: 12,351 paired events.
- `InsertText`: 2,310 calls, 6,818.25 seconds, 76.9% of reconstruction time.
- `BeginParagraph`: 1,372 calls, 951.14 seconds, 10.7% of reconstruction time.
- More than five seconds: 741 events; more than thirty seconds: 3 events.

---

### Task 1: Make retained Golden traces produce a stable performance profile

**Files:**
- Create: `scripts/codex_automation/trace_profile.py`
- Create: `tests/unit/test_codex_automation_trace_profile.py`
- Modify: `scripts/codex_automation/cli.py`

**Interfaces:**
- Produces `profile_event_trace(path: Path) -> dict`.
- Returned dictionary contains `paired_event_count`, `total_seconds`, `by_event_type`, `slowest_events`, and threshold counts for 1, 5, and 30 seconds.
- CLI option `--profile-report PATH` writes the profile next to `golden_report.json` without adding heavy artifacts to Git.

- [ ] **Step 1: Write the failing pairing and aggregation test**

  Build a temporary JSONL trace containing two complete before/after pairs and one unmatched `before`. Assert that only complete pairs count, durations are grouped by event type, and the unmatched row is reported without crashing.

- [ ] **Step 2: Run the focused test and verify RED**

  Run: `C:\WordReplica-Automation\.venv\Scripts\python.exe -m pytest tests/unit/test_codex_automation_trace_profile.py -q -p no:cacheprovider`

  Expected: FAIL because `profile_event_trace` does not exist.

- [ ] **Step 3: Implement deterministic trace profiling**

  Pair records by `(event_index, event_type, source_element_id)`, parse timestamps as timezone-aware UTC, ignore malformed lines while counting them, and sort slowest events by duration descending then event index ascending.

- [ ] **Step 4: Run focused tests and verify GREEN**

  Run the same focused command and expect all tests to pass.

- [ ] **Step 5: Profile the retained baseline**

  Run the profiler against `C:\WordReplica-Automation\diagnostics\20260813T204758Z_5ed0c311\interactive\event_trace.jsonl` and assert the report reproduces 12,351 paired events and approximately 6,818 seconds for `InsertText` (allow one-second rounding tolerance).

### Task 2: Add short real-Word performance fixtures with fidelity guards

**Files:**
- Create: `scripts/codex_automation/word_microbenchmark.py`
- Create: `tests/unit/test_word_microbenchmark.py`
- Create: `tests/integration/word/test_interactive_word_performance.py`
- Use: `tests/fixtures/build_fixtures.py`

**Interfaces:**
- Produces `build_text_growth_fixture(path: Path, *, prefix_paragraphs: int, measured_paragraphs: int, table_cells: int) -> Path`.
- Produces `run_word_microbenchmark(source: Path, output_dir: Path) -> dict` with elapsed time, event profile, and L0-L3 results.
- Integration test is skipped unless `WORD_REPLICA_RUN_WORD_PERF=1`.

- [ ] **Step 1: Write a failing fixture-shape test**

  Generate a compact document with a large unmeasured text prefix, ordinary paragraphs, and a table containing alternating normal/bold runs. Assert exact source text, paragraph count, table/cell count, and bold span count.

- [ ] **Step 2: Run the fixture test and verify RED**

  Expected: FAIL because the benchmark builder is absent.

- [ ] **Step 3: Implement the deterministic fixture builder**

  Use fixed text and fixed formatting only; do not copy or alter the Golden DOCX. Ensure the output is byte-stable except for package timestamps already normalized by existing fixture helpers.

- [ ] **Step 4: Write the real-Word benchmark test**

  Reconstruct the fixture, reopen the result through the normal parser, and assert L0-L3 PASS plus exact alternating bold states. Save elapsed/profile JSON only under the test's temporary output directory.

- [ ] **Step 5: Establish three baseline measurements**

  Run the benchmark three times. Record median reconstruction seconds and `InsertText` total. A production optimization proceeds only if all three runs are evaluable and preserve L0-L3.

### Task 3: Eliminate redundant post-insertion run-format application

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `tests/unit/test_interactive_word_executor.py`
- Modify: `tests/integration/word/test_interactive_word_performance.py`

**Interfaces:**
- `ApplyRunProperties` remains the single formatting operation before its following `InsertText`.
- `_insert_text(text: str)` inserts and collapses the range without calling `_event_ApplyRunProperties` a second time.
- Mixed normal/bold/italic/font/size/spacing spans must survive save and reopen unchanged.

- [ ] **Step 1: Write a failing duplicate-format regression test**

  Execute `ApplyRunProperties` then `InsertText` with a logging fake COM range. Assert `Font.Reset` and each font property are assigned exactly once. Current behavior must be RED because `_insert_text` reapplies the active properties.

- [ ] **Step 2: Add a real-Word fidelity guard**

  Use adjacent runs with normal, bold, italic, explicit font, nonzero character spacing, and normal text after them. Save/reopen and assert every span's text and effective properties exactly match the source projection.

- [ ] **Step 3: Make the smallest renderer change**

  Remove only the post-`InsertAfter` `_event_ApplyRunProperties(...)` call and its cache invalidation. Keep pre-insert `ApplyRunProperties`, `_collapse_end()`, page-break state, and post-table state unchanged.

- [ ] **Step 4: Verify GREEN locally and in Word**

  Run the focused unit test, executor test file, and opt-in Word benchmark. Reject the change if any bold/font/spacing span drifts.

- [ ] **Step 5: Apply the performance acceptance gate**

  Require at least 30% lower median `InsertText` time and at least 20% lower total benchmark time across three runs. If the threshold is not met, revert this hypothesis without committing it.

### Task 4: Measure and control Word repagination during construction

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `tests/unit/test_interactive_word_executor.py`
- Modify: `tests/integration/word/test_interactive_word_performance.py`

**Interfaces:**
- Controller records original `Application.ScreenUpdating` and `Application.Options.Pagination` values when available.
- Background reconstruction disables both during event execution.
- Before final save/audit, original values are restored and Word repaginates once.
- Visible mode keeps the user-visible behavior requested by `-VisibleWord`.

- [ ] **Step 1: Write RED lifecycle tests**

  Assert background open disables screen updating/pagination, finalization restores them even after an exception, and visible mode does not unexpectedly hide updates.

- [ ] **Step 2: Implement scoped state restoration**

  Store the original values on the controller, restore them in the existing close/final-save lifecycle, and guard unsupported Word properties with the current retry/suppression conventions.

- [ ] **Step 3: Verify formatting and pagination fidelity**

  Run the real-Word benchmark, reopen the output, and require L0-L3 PASS and the same page count as the benchmark source after the final repagination.

- [ ] **Step 4: Apply the performance acceptance gate**

  Require an additional 20% median total-time reduction over Task 3 across three runs. Reject the change if the gain is below threshold or if visible mode/regression tests drift.

### Task 5: Resolve the current first fidelity divergence (G2 character spacing)

**Files:**
- Create: `tests/integration/word/test_word_character_spacing_roundtrip.py`
- Modify only after RED evidence: `src/word_replica/renderers/interactive_word.py` or `src/word_replica/qa/formatting.py`
- Test: `tests/unit/test_qa_levels.py`

**Interfaces:**
- Diagnostic test writes OOXML spacing values 0 through 40 twentieths of a point through the supported COM path, saves, reparses, and records the Word 2010 round-trip mapping.
- Renderer correction is allowed only when Word can serialize the exact source value through a demonstrated COM input.
- Evaluator normalization is allowed only when two values are proven by the mapping to represent the same Word-effective spacing.

- [ ] **Step 1: Write and run the Word mapping diagnostic**

  Include the current failing value `18`; assert the test initially exposes the `18 -> 17` round trip seen in Golden run `20260813T204758Z_5ed0c311`.

- [ ] **Step 2: Select the evidence-backed correction path**

  If a nearby COM input serializes exactly to 18, add a unit RED test for that conversion and change only the renderer conversion. If no COM input serializes to 18 but Word renders a stable equivalent class, add a unit RED test for that exact equivalence and normalize only that demonstrated class in G2.

- [ ] **Step 3: Verify focused GREEN and full regression**

  Run the spacing integration test, G2 unit tests, and full pytest with an explicit writable `--basetemp`.

- [ ] **Step 4: Run one real Golden and compare**

  Require G0/G1 to remain PASS, G2 mismatch count to decrease, source SHA-256 to remain unchanged, and total runtime to move toward the 45-minute target. Stop immediately on an unexplained passing-gate regression.

### Task 6: Continue first-divergence fidelity repair through G9

**Files:**
- Read each run's local `golden_report.json` and retained diagnostics first.
- Modify only the source/test files implicated by that run's first concrete root cause.

**Interfaces:**
- One production root cause per TDD cycle.
- Every cycle emits a comparable report containing G0-G9, mismatch counts, page counts, source hash, runtime, and trace profile.

- [ ] **Step 1: Select only the first divergent gate and first concrete root cause**

  Do not work on downstream G3-G9 symptoms while an earlier unexplained G2 divergence remains.

- [ ] **Step 2: Execute mandatory RED -> minimal fix -> GREEN**

  Record the focused failing command and failure, make the smallest change, and rerun the same command to PASS.

- [ ] **Step 3: Run full pytest and one real Golden**

  Compare gate booleans, mismatch counts, source/output pages, runtime, and source hash to the immediately previous retained run.

- [ ] **Step 4: Checkpoint only measurable improvements**

  Commit and push to `automation-dev` only when tests are green and a gate turns green, a mismatch count materially decreases without regression, or runtime materially improves while fidelity is unchanged.

- [ ] **Step 5: Execute the promotion proof**

  When all gates pass, rerun the unmodified commit. Promote only after the second consecutive FULL PASS and `promotion_ready=true`.

## Milestones

- **M1 - Fast feedback:** deterministic trace profiler and three-run Word microbenchmark exist.
- **M2 - Practical runtime:** full Golden reconstruction plus audit is <=45 minutes with no loss of G0/G1 or increase in any earlier gate's mismatch count.
- **M3 - Layout convergence:** output page count reaches 73 and G2-G8 pass.
- **M4 - Visual convergence:** G9 passes without loosening its tolerance to hide defects.
- **M5 - Stable release candidate:** two consecutive G0-G9 FULL PASS runs on one commit; promotion gate true.
