# Word-Native Table Fast Path Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace hundreds of per-cell Word text insertions in eligible tables with one native Word text-to-table transaction per table while preserving exact G0-G9 fidelity, restart safety, and local diagnostics.

**Architecture:** Compile the existing resolved legacy table events into one immutable `InsertTableBatch` event only when their grammar proves the table is rectangular and text-only. The Word controller inserts one tab/paragraph-delimited payload, calls Word's native `Range.ConvertToTable`, and applies the already-resolved table, cell, paragraph, and run properties to exact ranges. Ineligible tables keep the legacy event stream unchanged.

**Tech Stack:** Python 3.12, pytest, pywin32/Microsoft Word COM, python-docx fixture generation, existing WordReplica parser/blueprint/checkpoint/QA infrastructure, local JSONL event traces.

## Global Constraints

- Work only on `automation-dev`; never develop directly on `main`.
- Never modify either source DOCX.
- Never globally terminate `WINWORD.EXE` or close unrelated Word documents.
- Every production behavior change requires observed RED, smallest fix, observed GREEN, full pytest, and real-Word evidence.
- Do not weaken Word state snapshots, process ownership, G0-G9 gates, or visual tolerances.
- The legacy table path remains the automatic fallback for every ineligible table.
- The batch transaction has one resumable boundary immediately before it and one only after successful table verification.
- Do not commit intermediate task changes. The repository contract permits a checkpoint only after tests are green and Golden evidence measurably improves.
- Acceptance requires at least 50% lower median table reconstruction time across three paired benchmark runs with unchanged G0-G3.
- Seminar acceptance requires two consecutive G0-G9 FULL PASS runs on the same commit, 17 pages, and reconstruction below five minutes.
- Golden acceptance for this iteration requires unchanged source SHA-256 `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`, no passing-gate regression, and material runtime improvement.

---

## File Structure

- Create `src/word_replica/interactive/table_batch.py`: pure conversion of resolved legacy table events into an immutable batch event.
- Create `tests/unit/test_table_batch.py`: event-grammar, offset, delimiter, and fallback tests.
- Modify `src/word_replica/interactive/blueprint.py`: compile legacy events once, replace eligible tables with one batch event, and preserve semantic counts.
- Modify `src/word_replica/domain/reconstruction.py`: count visible text carried by `InsertTableBatch`.
- Modify `src/word_replica/config.py`: expose an internal fidelity-preserving benchmark toggle `enable_table_fast_path`.
- Modify `src/word_replica/renderers/interactive_word.py`: execute native text-to-table conversion and exact range formatting.
- Modify `src/word_replica/interactive/checkpoints.py`: save immediately before a batch and after its verified completion.
- Modify `src/word_replica/interactive/verification.py`: account for atomic table text and completed-table progress.
- Modify `src/word_replica/services/interactive_rebuild.py`: compile with the option, verify batch boundaries, and forward phase metrics.
- Modify `scripts/remote_harness/child_runner.py`: add a benchmark-only legacy toggle.
- Modify `scripts/remote_harness/event_trace.py`: retain table phase timing records.
- Modify `scripts/codex_automation/word_microbenchmark.py`: create a deterministic multi-table fixture and run paired legacy/fast measurements.
- Create `tests/integration/word/test_interactive_table_fast_path.py`: opt-in real-Word fidelity and performance guard.
- Modify the directly corresponding unit test files named in each task below.

---

### Task 1: Convert resolved legacy table events into an exact batch plan

**Files:**
- Create: `src/word_replica/interactive/table_batch.py`
- Create: `tests/unit/test_table_batch.py`

**Interfaces:**
- Consumes: `Sequence[ReconstructionEvent]` beginning with `BeginTable` and ending with `EndTable`.
- Produces: `build_table_batch_event(events: Sequence[ReconstructionEvent]) -> ReconstructionEvent | None`.
- Produces event type: `InsertTableBatch` with `rows`, `columns`, table/row/cell properties, `cells`, `paragraph_count`, `text_projection`, and `legacy_event_count`.

- [ ] **Step 1: Write the eligible one-cell RED test**

```python
def test_resolved_text_only_table_becomes_one_batch_event():
    events = (
        ReconstructionEvent("BeginTable", "t1", {"rows": 1, "columns": 1}),
        ReconstructionEvent("SetTableProperties", "t1", {"layout": "fixed"}),
        ReconstructionEvent("SetColumnWidth", "t1", {"column": 1, "width_twips": "2400"}),
        ReconstructionEvent("SetCellProperties", "c1", {"row": 1, "column": 1, "properties": {"vertical_alignment": "top"}}),
        ReconstructionEvent("EnterCell", "c1", {"row": 1, "column": 1}),
        ReconstructionEvent("BeginParagraph", "p1", {}),
        ReconstructionEvent("ApplyParagraphProperties", "p1", {"style_id": "Normal", "alignment": "both"}),
        ReconstructionEvent("ApplyRunProperties", "r1", {"bold": False, "font_ascii": "Times New Roman"}),
        ReconstructionEvent("InsertText", "r1", {"text": "Alpha "}),
        ReconstructionEvent("ApplyRunProperties", "r2", {"bold": True, "font_ascii": "Times New Roman"}),
        ReconstructionEvent("InsertText", "r2", {"text": "Beta"}),
        ReconstructionEvent("EndParagraph", "p1", {}),
        ReconstructionEvent("LeaveCell", "c1", {"row": 1, "column": 1}),
        ReconstructionEvent("EndTable", "t1", {}),
    )

    batch = build_table_batch_event(events)

    assert batch is not None
    assert batch.event_type == "InsertTableBatch"
    assert batch.payload["paragraph_count"] == 1
    assert batch.payload["text_projection"] == "Alpha Beta"
    assert batch.payload["cells"][0]["text"] == "Alpha Beta"
    assert batch.payload["cells"][0]["runs"] == (
        {"source_element_id": "r1", "start": 0, "end": 6, "properties": {"bold": False, "font_ascii": "Times New Roman"}},
        {"source_element_id": "r2", "start": 6, "end": 10, "properties": {"bold": True, "font_ascii": "Times New Roman"}},
    )
```

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
python -m pytest -p no:cacheprovider --basetemp C:\WordReplica-Automation\work\pytest-table-batch-red tests\unit\test_table_batch.py::test_resolved_text_only_table_becomes_one_batch_event -q
```

Expected: import failure because `word_replica.interactive.table_batch` does not exist.

- [ ] **Step 3: Add fallback RED tests for unsupported grammar**

Add parameterized cases where the resolved event slice contains `MergeCells`, `InsertTab`, `InsertLineBreak`, `InsertPageBreak`, `CreateField`, `BookmarkStart`, `CreateBookmark`, `InsertImage`, `CreateListBinding`, nested `BeginTable`, two `BeginParagraph` events in one cell, or source text containing `\t`, `\r`, or `\x07`. Assert `build_table_batch_event(events) is None` for every case.

- [ ] **Step 4: Implement the pure parser and offset builder**

Create these constants and function:

```python
ALLOWED_STRUCTURE_EVENTS = {
    "BeginTable", "SetTableProperties", "SetColumnWidth", "SetRowProperties",
    "SetCellProperties", "EnterCell", "BeginParagraph",
    "ApplyParagraphProperties", "ApplyRunProperties", "InsertText",
    "EndParagraph", "LeaveCell", "EndTable",
}
FORBIDDEN_DELIMITERS = {"\t", "\r", "\x07"}


def build_table_batch_event(
    events: Sequence[ReconstructionEvent],
) -> ReconstructionEvent | None:
    """Return one atomic Word-native table event or None without mutating input."""
```

Parse the exact grammar in row/column order. Accumulate each run's `start` before appending text and `end` afterward. Reject duplicate/missing cells, missing paragraph properties, missing run/text pairs, out-of-order coordinates, dimension mismatches, and any unsupported event before constructing the immutable payload.

- [ ] **Step 5: Run the complete pure unit test file and verify GREEN**

Run:

```powershell
python -m pytest -p no:cacheprovider --basetemp C:\WordReplica-Automation\work\pytest-table-batch-green tests\unit\test_table_batch.py -q
```

Expected: all event-grammar, offset, and fallback tests pass.

---

### Task 2: Emit atomic events only for eligible tables and preserve blueprint truth

**Files:**
- Modify: `src/word_replica/interactive/blueprint.py`
- Modify: `src/word_replica/domain/reconstruction.py`
- Modify: `src/word_replica/config.py`
- Modify: `src/word_replica/services/interactive_rebuild.py`
- Modify: `tests/unit/test_blueprint_compiler.py`
- Modify: `tests/unit/test_interactive_config.py`
- Modify: `tests/unit/test_interactive_rebuild_service.py`

**Interfaces:**
- `BlueprintCompiler.__init__(*, enable_table_fast_path: bool = True)`.
- `InteractiveOptions.enable_table_fast_path: bool = True`.
- Eligible table: one `InsertTableBatch` event.
- Ineligible/disabled table: byte-for-byte-equivalent legacy event sequence.

- [ ] **Step 1: Write RED compiler tests for eligible and disabled modes**

```python
def test_text_only_table_compiles_to_one_atomic_batch_event():
    model = DocumentModel(
        source_sha256="a" * 64,
        body=[Table("t", [TableRow("row", [TableCell("c", [Paragraph("p", [Run("r", "Alpha")])])])])],
    )
    blueprint = BlueprintCompiler(enable_table_fast_path=True).compile(model)
    assert [event.event_type for event in blueprint.events] == ["InsertTableBatch"]
    assert blueprint.semantic_counts["tables"] == 1
    assert blueprint.semantic_counts["paragraphs"] == 1
    assert blueprint.total_visible_characters == 5


def test_disabled_fast_path_keeps_legacy_table_events():
    blueprint = BlueprintCompiler(enable_table_fast_path=False).compile(model)
    event_types = [event.event_type for event in blueprint.events]
    assert event_types[0] == "BeginTable"
    assert event_types[-1] == "EndTable"
    assert "InsertTableBatch" not in event_types
```

- [ ] **Step 2: Verify focused RED**

Run the two tests and expect `TypeError` because `BlueprintCompiler` has no such constructor and `InteractiveOptions` has no toggle.

- [ ] **Step 3: Refactor without changing legacy compilation**

Rename the current implementation to:

```python
def _compile_table_legacy(
    self,
    table: Table,
    events: list[ReconstructionEvent],
    location: SemanticLocation,
) -> None:
```

Make `_compile_table` compile into a temporary list once, call `build_table_batch_event`, and append either one batch event or the original temporary list. Do not duplicate effective style resolution.

- [ ] **Step 4: Preserve semantic and visible-character accounting**

Count `InsertTableBatch` as one table and count its cells as paragraphs. In `ReconstructionBlueprint.build`, include `len(event.payload["text_projection"])` when `event.event_type == "InsertTableBatch"`.

- [ ] **Step 5: Thread the option through prepare and resume**

Both compiler call sites in `InteractiveRebuildService` must use:

```python
BlueprintCompiler(
    enable_table_fast_path=options.interactive.enable_table_fast_path,
).compile(model)
```

Include the boolean in settings serialization and restoration so checkpoint fingerprints remain deterministic.

- [ ] **Step 6: Verify compiler/config/service GREEN**

Run:

```powershell
python -m pytest -p no:cacheprovider --basetemp C:\WordReplica-Automation\work\pytest-table-compiler tests\unit\test_table_batch.py tests\unit\test_blueprint_compiler.py tests\unit\test_interactive_config.py tests\unit\test_interactive_rebuild_service.py -q
```

Expected: all selected tests pass and legacy fallback tests retain their existing event ordering.

---

### Task 3: Execute one native Word text-to-table transaction

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- Controller handler: `_event_InsertTableBatch(event: ReconstructionEvent) -> None`.
- Controller metric accessor: `consume_table_batch_metrics() -> dict[str, object] | None`.
- Word operation: one `Range.InsertAfter(payload_text)` followed by one `Range.ConvertToTable(Separator=1, NumRows=rows, NumColumns=columns)`.

- [ ] **Step 1: Write a RED fake-COM test for call count and shape**

Build a fake range/document where `InsertAfter` records one string, `Document.Range(start, end)` returns the inserted range, and `ConvertToTable` returns a fake table with addressable cells. Execute a 2x2 batch and assert:

```python
assert root.insert_after_calls == ["A\tB\rC\tD"]
assert inserted.convert_calls == [{"Separator": 1, "NumRows": 2, "NumColumns": 2}]
assert fake_document.tables_add_calls == []
assert controller.active_range is fake_table.after_range
```

- [ ] **Step 2: Verify RED**

Run the single test and expect `ValueError: unsupported reconstruction event: InsertTableBatch`.

- [ ] **Step 3: Implement insertion and conversion only**

Capture the collapsed insertion start, insert the delimited payload once, create an exact `document.Range(start, start + len(payload_text))`, and call `ConvertToTable`. Wrap every COM boundary in the existing rejected-call retry helpers. Do not call `Tables.Add` for a batch event.

- [ ] **Step 4: Add RED tests for exact run ranges**

For a cell containing `Alpha ` plus bold `Beta`, assert paragraph formatting targets the cell range excluding the terminal cell marker and run formatting targets offsets `[0, 6)` and `[6, 10)` relative to the cell's Word start.

- [ ] **Step 5: Reuse existing property handlers for GREEN**

Create a temporary table context with the converted table and cell map. Apply existing table/column/row/cell handlers from payload records. For each cell, duplicate `cell.Range`, remove the final cell marker, apply paragraph properties, then create exact document ranges for every run and apply run properties. Restore the post-table range and `_post_table_paragraph_active=True` only after all cells succeed.

- [ ] **Step 6: Record deterministic phase metrics**

Measure `insert_seconds`, `convert_seconds`, `geometry_seconds`, `formatting_seconds`, `verification_seconds`, `total_seconds`, `row_count`, `cell_count`, and `run_count` with `time.perf_counter`. Store one completed metric dictionary only after success; include `failed_phase` before re-raising a COM failure.

- [ ] **Step 7: Verify the complete controller test file GREEN**

Run:

```powershell
python -m pytest -p no:cacheprovider --basetemp C:\WordReplica-Automation\work\pytest-table-controller tests\unit\test_interactive_word_executor.py -q
```

Expected: all existing controller tests and new batch tests pass.

---

### Task 4: Preserve checkpoints, live verification, progress, and trace evidence

**Files:**
- Modify: `src/word_replica/interactive/checkpoints.py`
- Modify: `src/word_replica/interactive/verification.py`
- Modify: `src/word_replica/services/interactive_rebuild.py`
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `scripts/remote_harness/event_trace.py`
- Modify: `tests/unit/test_interactive_checkpoints.py`
- Modify: `tests/unit/test_interactive_verification.py`
- Modify: `tests/unit/test_interactive_rebuild_service.py`
- Modify: `tests/unit/test_remote_harness_event_trace.py`

**Interfaces:**
- `InteractiveCheckpointCoordinator.event_started(index, event, snapshot)` checkpoints `index - 1` before `InsertTableBatch`.
- `InsertTableBatch` is both a completed-table boundary and a safe live-verification boundary.
- Trace observer callback: `table_batch_profile(index, event, metrics)` writes status `table_batch_profile`.

- [ ] **Step 1: Write checkpoint RED tests**

Assert an `InsertTableBatch` started at index 12 saves a checkpoint at index 11 with reason `before_table_batch`, and its completion saves the normal table boundary when `checkpoint_after_tables=True`. Assert repeated observer callbacks do not duplicate the same pre-batch checkpoint.

- [ ] **Step 2: Write progress/live-verification RED tests**

Complete one batch event whose payload has `text_projection="AB"`, four cells, and `paragraph_count=4`. Assert completed characters increase by two, completed tables by one, table index by one, and expected body prefix receives `AB` before live verification.

- [ ] **Step 3: Implement atomic boundary accounting**

Treat `InsertTableBatch` as the equivalent of legacy `BeginTable` plus `EndTable` for block context, completed tables, and checkpoint policy. Add it to `_InteractiveServiceObserver.SAFE_VERIFY_EVENTS` and append only `text_projection` to expected body text.

- [ ] **Step 4: Write and implement phase-trace RED/GREEN**

After a batch handler succeeds, `InteractiveWordRenderer.execute_blueprint` consumes its metric dictionary and calls `observer.table_batch_profile`. If the handler raises, consume and emit the failed-phase metric before forwarding `event_failed`. Forward the callback through `_CompositeObserver` and `_InteractiveServiceObserver`. `JsonlEventTraceObserver` writes table ID, phase times, row/cell/run counts, and success/failure phase without including cell text.

- [ ] **Step 5: Verify all safety and trace tests GREEN**

Run:

```powershell
python -m pytest -p no:cacheprovider --basetemp C:\WordReplica-Automation\work\pytest-table-safety tests\unit\test_interactive_checkpoints.py tests\unit\test_interactive_verification.py tests\unit\test_interactive_rebuild_service.py tests\unit\test_remote_harness_event_trace.py -q
```

Expected: pre-batch restart checkpoint, post-batch verification, progress, and profile rows all pass without weakening ordinary state snapshots.

---

### Task 5: Prove real-Word fidelity and median speed before Golden

**Files:**
- Modify: `scripts/remote_harness/child_runner.py`
- Modify: `scripts/codex_automation/word_microbenchmark.py`
- Modify: `tests/unit/test_remote_harness_child_runner.py`
- Modify: `tests/unit/test_word_microbenchmark.py`
- Create: `tests/integration/word/test_interactive_table_fast_path.py`

**Interfaces:**
- Child option: `--disable-table-fast-path` sets `InteractiveOptions.enable_table_fast_path=False` for paired measurements only.
- Fixture builder: `build_table_batch_fixture(path: Path, *, tables: int, rows: int, columns: int) -> Path`.
- Comparison runner: `run_table_batch_comparison(source: Path, output_dir: Path, *, repo_root: Path, runs: int = 3) -> dict`.

- [ ] **Step 1: Write the deterministic fixture RED test**

Generate two 12x4 tables. Every cell contains normal text, one bold span, and trailing normal text. Alternate vertical alignment and fixed cell widths. Assert exact table/cell/paragraph counts, exact cell text, exact bold run positions, and no unsupported batch events in the compiled blueprint.

- [ ] **Step 2: Implement the fixture and verify GREEN**

Use python-docx only to create the independent benchmark source. Do not copy Golden or seminar content. Save it under the caller's output directory.

- [ ] **Step 3: Add the benchmark-only legacy toggle**

Thread `--disable-table-fast-path` through `_options_for` in the child runner. Default production behavior remains enabled; the option exists only to produce paired evidence from the same working tree.

- [ ] **Step 4: Write the opt-in real-Word fidelity test**

Skip unless `WORD_REPLICA_RUN_WORD_PERF=1`. Run one legacy and one fast reconstruction, reopen both, and require:

```python
assert fast["l0_l3"] == legacy["l0_l3"]
assert all(level["passed"] for level in fast["l0_l3"].values())
assert fast_model.plain_text() == source_model.plain_text()
assert table_projection(fast_model) == table_projection(source_model)
assert any(
    row["event_type"] == "InsertTableBatch"
    for row in fast["performance_profile"]["by_event_type"]
)
```

- [ ] **Step 5: Run three paired real-Word benchmarks**

Run:

```powershell
$env:WORD_REPLICA_RUN_WORD_PERF='1'
python -m pytest -p no:cacheprovider --basetemp C:\WordReplica-Automation\work\pytest-table-word tests\integration\word\test_interactive_table_fast_path.py -q -s
```

Retain reports under `C:\WordReplica-Automation\work\table-fast-path-benchmark-<run>`. Do not run Golden at this stage.

- [ ] **Step 6: Apply the benchmark acceptance gate**

Accept only if all six reconstructions are evaluable, all fast outputs match exact G0-G3 projections, and:

```text
100 * (legacy_median_table_seconds - fast_median_table_seconds)
      / legacy_median_table_seconds >= 50.0
```

If fidelity differs or improvement is below 50%, revert the production hypothesis without committing and use the retained phase profile to choose the next evidenced design.

---

### Task 6: Run full regression and two seminar acceptance passes

**Files:**
- Test all changed production and test files.
- Store heavy artifacts only under `C:\WordReplica-Automation\work`.

**Interfaces:**
- Full pytest is the local regression authority.
- `seminar_report.json` is the seminar G0-G9 authority.

- [ ] **Step 1: Run full pytest**

```powershell
python -m pytest -p no:cacheprovider --basetemp C:\WordReplica-Automation\work\pytest-table-full -q
```

Expected: all tests pass; Word-only tests remain skipped unless explicitly enabled.

- [ ] **Step 2: Recheck seminar source SHA-256**

Expected SHA-256:

```text
47948159867f4fb6c668f0bba066380591c142995443bb9c8a5f17d6984bff30
```

- [ ] **Step 3: Run seminar pass 1 and full audit**

Use `scripts\remote_harness\child_runner.py --stage interactive_maximum --defer-l4-qa`, then `scripts\codex_automation\audit_cli.py` against its local output. Require G0-G9 true, 17/17 pages, unchanged source hash, and reconstruction below 300 seconds.

- [ ] **Step 4: Run seminar pass 2 on the unchanged working commit**

Repeat into a distinct run directory. Require the same full-pass conditions and confirm both reports carry the same commit SHA.

- [ ] **Step 5: Stop on any seminar drift**

Do not run Golden if either seminar report fails a gate, page count differs, source hash changes, or runtime exceeds five minutes. Diagnose the first seminar divergence with a new RED test.

---

### Task 7: Run one controlled Golden and checkpoint only measured improvement

**Files:**
- Read: newest `C:\WordReplica-Automation\diagnostics\<run>\golden_report.json`
- Read: newest retained `interactive\event_trace.jsonl`
- Commit only the reviewed fast-path files after all gates below succeed.

**Interfaces:**
- Baseline report: `C:\WordReplica-Automation\diagnostics\20260815T151655Z_5ed0c311\golden_report.json`.
- Baseline: G2 findings 157, table seconds 4,809.842, total profile seconds 7,909.862, pages 73/11.

- [ ] **Step 1: Run the only post-implementation Golden**

From the repository root:

```powershell
.\RUN_GOLDEN_CODEX.ps1
```

- [ ] **Step 2: Verify source and passing-gate invariants**

Require unchanged source SHA-256 and G0, G1, G5, and G7 still PASS. Stop immediately on an unexplained regression.

- [ ] **Step 3: Compare performance and fidelity**

Require table seconds below 2,404.921 (50% below the immediate baseline), no increase in G2-G9 mismatch counts attributable to batching, and total runtime materially closer to 45 minutes. Record `InsertTableBatch` phase metrics for all 24 eligible Golden tables.

- [ ] **Step 4: Create the production checkpoint only after evidence succeeds**

Stage only the table-fast-path implementation, tests, plan, and directly required support changes. Preserve unrelated working-tree changes. Commit and push only to `automation-dev`.

- [ ] **Step 5: Continue the mandatory first-divergence loop**

The expected next fidelity root cause before batching is `14/runs/0/font_cs`, expected `Cambria`, actual `Times New Roman`. After the performance checkpoint, return to that first G2 divergence with a separate RED/green cycle; do not begin downstream G3-G9 fixes first.
