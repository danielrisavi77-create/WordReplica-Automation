# Word Replica Remote Windows/Word Test Harness Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a portable one-command Windows harness that processes 10+ real DOCX files through Word Replica and desktop Microsoft Word in isolated child processes, then emits one diagnostic ZIP even when individual documents fail.

**Architecture:** The parent orchestrator owns prerequisite checks, process isolation, timeouts, artifact directories, aggregation, and ZIP creation. A child runner performs exactly one document/stage using production `DocxParser`, `BlueprintCompiler`, `RebuildService`, QA artifacts, and Word COM. Interactive execution gains observer-only `before/after/error` trace hooks without changing reconstruction semantics.

**Tech Stack:** Python 3.12+, stdlib (`argparse`, `csv`, `json`, `subprocess`, `zipfile`, `hashlib`, `platform`, `shutil`), existing Word Replica package, pywin32 on Windows, PowerShell 5.1-compatible ASCII bootstrap scripts.

## Global Constraints

- Source `.docx` files are read-only; SHA-256 before and after every document must match.
- One document/stage failure never stops later documents.
- Each stage runs in a fresh child Python process.
- Maximum Interactive correctly blocked by preflight is `EXPECTED_BLOCK`, not a failure.
- Harness does not fix renderer bugs or silently retry high-level document edit operations.
- Result ZIP is always attempted after corpus execution, including partial/all-fail runs.
- No source documents are included in result ZIP by default.
- `--logs-only` excludes generated DOCX/PDF artifacts.
- No automatic upload.
- No randomized typing, fake mistakes, fabricated edit history, or deceptive metadata.
- JSON artifact schema version is `1` and harness version is `0.1.0`.
- Default timeouts: static 120 s, Instant 600 s, Interactive 1800 s, L4 600 s.

---

### Task 1: Harness contracts, configuration, and status normalization

**Files:**
- Create: `scripts/remote_harness/__init__.py`
- Create: `scripts/remote_harness/contracts.py`
- Create: `scripts/remote_harness/config.py`
- Create: `harness_config.json`
- Test: `tests/unit/test_remote_harness_contracts.py`

**Interfaces:**
- Produces `HarnessConfig.load(path: Path) -> HarnessConfig`.
- Produces `normalize_failure(text: str) -> str` for root-cause grouping.
- Produces `HarnessStatus` values `PASS`, `WARN`, `EXPECTED_BLOCK`, `ENGINE_FAIL`, `COM_FAIL`, `TIMEOUT`, `PROCESS_CRASH`, `SOURCE_MUTATED`, `HARNESS_FAIL`.
- Produces JSON-serializable `StageResult` and `DocumentSummary` dataclasses.

- [ ] **Step 1: Write failing configuration/status tests**

```python
from scripts.remote_harness.config import HarnessConfig
from scripts.remote_harness.contracts import HarnessStatus, normalize_failure


def test_default_harness_timeouts_match_spec(tmp_path):
    cfg = HarnessConfig.load(tmp_path / "missing.json")
    assert cfg.static_timeout_seconds == 120
    assert cfg.instant_timeout_seconds == 600
    assert cfg.interactive_timeout_seconds == 1800
    assert cfg.l4_timeout_seconds == 600


def test_failure_normalizer_groups_rpc_call_rejected():
    assert normalize_failure("(-2147418111, 'Call was rejected by callee.', None, None)") == "RPC_E_CALL_REJECTED"
    assert HarnessStatus.COM_FAIL.value == "COM_FAIL"
```

- [ ] **Step 2: Run RED test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_contracts.py -q`

Expected: FAIL because `scripts.remote_harness` does not exist.

- [ ] **Step 3: Implement contracts and config**

Implement frozen/slots dataclasses with `to_dict()` using `dataclasses.asdict`, enum values above, exact default timeouts, `logs_only=False`, and failure normalization rules for `RPC_E_CALL_REJECTED`, `RPC_SERVER_UNAVAILABLE`, `INVALID_ACTIVE_RANGE_TYPE`, `IMAGE_RELATIONSHIP_ERROR`, falling back to `UNCLASSIFIED:<sha256-prefix>`.

- [ ] **Step 4: Run GREEN test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_contracts.py -q`

Expected: PASS.

---

### Task 2: Static DOCX/package/model/blueprint analyzer

**Files:**
- Create: `scripts/remote_harness/static_analysis.py`
- Test: `tests/unit/test_remote_harness_static_analysis.py`

**Interfaces:**
- Consumes `DocxParser`, `BlueprintCompiler`, `analyze_preflight`.
- Produces `analyze_document(source: Path) -> dict` containing package integrity, relationship coverage, model counts, capability classification, blueprint fingerprint and event counts.

- [ ] **Step 1: Write a failing fixture analysis test**

```python
from scripts.remote_harness.static_analysis import analyze_document


def test_static_analysis_reports_blueprint_and_structural_counts(corpus_dir):
    report = analyze_document(corpus_dir / "04_tables_merged.docx")
    assert report["parser"]["ok"] is True
    assert report["model"]["tables"] >= 1
    assert report["blueprint"]["total_events"] > 0
    assert report["source"]["sha256"]
```

- [ ] **Step 2: Run RED test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_static_analysis.py -q`

Expected: FAIL because analyzer is missing.

- [ ] **Step 3: Implement analyzer**

Use `zipfile.ZipFile.testzip()` plus parsed `[Content_Types].xml` and `.rels` targets for package checks. Parse with `DocxParser`, count paragraphs/runs/tables/merged cells/drawings/sections/stories/notes/fields/bookmarks/comments/revisions/preserved parts, compile with `BlueprintCompiler`, and run preflight using `word_probe=lambda: True` so static capability classification does not depend on the current host OS.

- [ ] **Step 4: Run GREEN test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_static_analysis.py -q`

Expected: PASS.

---

### Task 3: Interactive event trace instrumentation and one-stage child runner

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Create: `scripts/remote_harness/event_trace.py`
- Create: `scripts/remote_harness/child_runner.py`
- Test: `tests/unit/test_remote_harness_event_trace.py`
- Test: `tests/unit/test_remote_harness_child_runner.py`

**Interfaces:**
- Interactive executor emits optional observer callbacks `event_started(index, event, snapshot)` and `event_failed(index, event, snapshot, exc)` in addition to existing `event_completed`.
- `JsonlEventTraceObserver(path)` writes append-only records with `status=before|after|error`.
- Child CLI: `python scripts/remote_harness/child_runner.py --stage <static|instant|interactive_maximum|interactive_standard> --source <docx> --run-dir <dir> [--logs-only]`.
- Child always writes `<run-dir>/result.json` before normal exit whenever Python remains alive.

- [ ] **Step 1: Write failing trace callback tests**

```python

def test_event_trace_writes_before_after_and_error_records(tmp_path):
    observer = JsonlEventTraceObserver(tmp_path / "event_trace.jsonl")
    event = ReconstructionEvent("InsertCharacter", "r1", {"character": "A"})
    observer.event_started(3, event, {"story": "main", "range_start": 4, "range_end": 4})
    observer.event_completed(3, event)
    observer.event_failed(4, event, {"story": "main"}, RuntimeError("boom"))
    rows = [json.loads(x) for x in (tmp_path / "event_trace.jsonl").read_text().splitlines()]
    assert [r["status"] for r in rows] == ["before", "after", "error"]
```

- [ ] **Step 2: Run RED test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_event_trace.py tests/unit/test_remote_harness_child_runner.py -q`

Expected: FAIL because trace and child modules are missing and executor lacks hooks.

- [ ] **Step 3: Add observer-only executor hooks**

Immediately before an atomic `execute_event`, capture the existing renderer state snapshot and call `observer.event_started(...)` when available. Wrap only the event call in `try/except`; on exception call `observer.event_failed(...)` then re-raise unchanged. Existing edit ordering and error behavior must not change.

- [ ] **Step 4: Implement child runner**

`static` calls Task 2 analyzer. `instant` uses `RebuildOptions(renderer=RendererChoice.WORD, visibility=VisibilityMode.BACKGROUND, reconstruction_mode=INSTANT)`. Interactive Maximum/Standard use `ReconstructionMode.INTERACTIVE`, maximum speed, zero object delay, live verification enabled, and matching fidelity. Serialize `RunResult`, warnings/reasons, elapsed time, output/QA paths; copy produced output/QA to deterministic names in the stage directory unless `--logs-only`.

- [ ] **Step 5: Run GREEN tests**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_event_trace.py tests/unit/test_remote_harness_child_runner.py -q`

Expected: PASS.

---

### Task 4: Parent process isolation, timeout, and source mutation guard

**Files:**
- Create: `scripts/remote_harness/process_runner.py`
- Create: `scripts/remote_harness/runner.py`
- Test: `tests/unit/test_remote_harness_process_runner.py`
- Test: `tests/unit/test_remote_harness_runner.py`

**Interfaces:**
- `run_child(command: list[str], timeout_seconds: int, stdout_path: Path, stderr_path: Path) -> ChildProcessResult`.
- `RemoteHarnessRunner(...).run() -> Path` returns result root directory.
- Every stage is a new child process; timeout kills child tree on Windows with `taskkill /PID <pid> /T /F` and continues.

- [ ] **Step 1: Write failing isolation/timeout tests**

```python

def test_timeout_is_normalized_and_next_stage_can_run(tmp_path):
    first = run_child([sys.executable, "-c", "import time; time.sleep(5)"], 1, tmp_path/"a.out", tmp_path/"a.err")
    second = run_child([sys.executable, "-c", "print('ok')"], 5, tmp_path/"b.out", tmp_path/"b.err")
    assert first.timed_out is True
    assert second.exit_code == 0
```

- [ ] **Step 2: Run RED tests**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_process_runner.py tests/unit/test_remote_harness_runner.py -q`

Expected: FAIL because runner modules are missing.

- [ ] **Step 3: Implement parent runner**

Enumerate only direct `*.docx` files from `realworld_input`, sorted case-insensitively. Hash before stages; run `static`, `instant`, `interactive_maximum`, and only run `interactive_standard` when Maximum result is normalized `EXPECTED_BLOCK`. After all stages hash source again; if changed, overwrite document aggregate status with `SOURCE_MUTATED`. Never stop corpus for a document-stage failure.

- [ ] **Step 4: Run GREEN tests**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_process_runner.py tests/unit/test_remote_harness_runner.py -q`

Expected: PASS.

---

### Task 5: Environment/prerequisite capture and Word verification

**Files:**
- Create: `scripts/remote_harness/environment.py`
- Create: `scripts/remote_harness/prerequisites.py`
- Test: `tests/unit/test_remote_harness_environment.py`

**Interfaces:**
- `capture_environment(root: Path) -> dict` records only spec-approved fields.
- `check_prerequisites(input_dir: Path, source_root: Path) -> PrerequisiteReport` stops before corpus when Windows, Python, Word activation, free space, inputs, or source manifest are invalid.

- [ ] **Step 1: Write failing environment tests with injectable probes**

```python

def test_prerequisite_report_requires_docx_and_word(tmp_path):
    report = check_prerequisites(tmp_path/"input", tmp_path, os_name="nt", word_probe=lambda: False)
    assert report.ok is False
    assert any("Microsoft Word" in item for item in report.errors)
```

- [ ] **Step 2: Run RED test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_environment.py -q`

Expected: FAIL because modules are missing.

- [ ] **Step 3: Implement probes**

On Windows use `win32com.client.DispatchEx("Word.Application")`, read `Version`, `Build` when available, close/quit in `finally`, and capture exceptions verbatim. Use `platform`, `locale`, `shutil.disk_usage`, and PowerShell-independent printer probing via Word `ActivePrinter` when available. Do not enumerate unrelated hardware/user data.

- [ ] **Step 4: Run GREEN test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_environment.py -q`

Expected: PASS.

---

### Task 6: Aggregation, fingerprint grouping, CSV/JSON summary, and ZIP collector

**Files:**
- Create: `scripts/remote_harness/aggregate.py`
- Create: `scripts/remote_harness/package_results.py`
- Test: `tests/unit/test_remote_harness_aggregate.py`

**Interfaces:**
- `aggregate_results(result_root: Path) -> dict` writes `summary.json` and `summary.csv`.
- `package_results(result_root: Path, destination_zip: Path, logs_only: bool) -> Path` creates one ZIP even when all document stages failed.

- [ ] **Step 1: Write failing grouping/package tests**

```python

def test_aggregate_groups_same_rpc_failure_once(tmp_path):
    # two result.json files contain the same COM rejection string
    summary = aggregate_results(tmp_path)
    assert summary["failure_groups"]["RPC_E_CALL_REJECTED"]["runs"] == 2


def test_logs_only_zip_excludes_docx_and_pdf(tmp_path):
    package_results(tmp_path/"results", tmp_path/"out.zip", logs_only=True)
    with ZipFile(tmp_path/"out.zip") as zf:
        assert not any(name.lower().endswith((".docx", ".pdf")) for name in zf.namelist())
```

- [ ] **Step 2: Run RED test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_aggregate.py -q`

Expected: FAIL because aggregator is missing.

- [ ] **Step 3: Implement aggregate/package behavior**

Produce one summary row per document/mode and grouped fingerprints with run/document counts. ZIP uses paths relative to `remote_results/<run_id>/`; excludes any source input directory unconditionally. `logs_only` additionally excludes generated `.docx`, `.pdf`, `.png`, `.jpg`, `.jpeg`, `.webp` files while retaining JSON/CSV/JSONL/HTML/logs.

- [ ] **Step 4: Run GREEN test**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_aggregate.py -q`

Expected: PASS.

---

### Task 7: One-command CLI and Windows bootstrap

**Files:**
- Create: `scripts/remote_harness/main.py`
- Create: `RUN_REMOTE_WORD_HARNESS.ps1`
- Create: `START_REMOTE_HARNESS.cmd`
- Create: `realworld_input/PUT_DOCX_FILES_HERE.txt`
- Test: `tests/unit/test_remote_harness_cli.py`
- Test: `tests/unit/test_remote_harness_windows_scripts.py`

**Interfaces:**
- Python CLI accepts `--input-dir`, `--config`, `--logs-only`, `--result-parent`.
- PowerShell creates `.venv_harness`, installs `-e . --no-cache-dir`, checks prerequisites, runs Python harness, propagates exit code, and prints result ZIP path.
- Scripts contain only ASCII so Windows PowerShell 5.1 cannot corrupt source encoding.

- [ ] **Step 1: Write failing CLI/script tests**

```python

def test_windows_harness_scripts_are_ascii_and_fail_fast():
    ps1 = Path("RUN_REMOTE_WORD_HARNESS.ps1").read_bytes()
    assert all(byte < 128 for byte in ps1)
    text = ps1.decode("ascii")
    assert "$ErrorActionPreference = 'Stop'" in text
    assert "realworld_input" in text
```

- [ ] **Step 2: Run RED tests**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_cli.py tests/unit/test_remote_harness_windows_scripts.py -q`

Expected: FAIL because CLI/bootstrap files are missing.

- [ ] **Step 3: Implement CLI/bootstrap**

CLI performs prerequisite check, creates `remote_results/<timestamp>`, writes environment/harness log, executes runner, aggregates, packages ZIP in a `finally` path when result root exists, prints `REMOTE WORD HARNESS: COMPLETE` and exact ZIP. PowerShell refuses to run from Explorer's ZIP temp path, requires at least one DOCX, and never prints PASS if Python exits non-zero before ZIP creation.

- [ ] **Step 4: Run GREEN tests**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_cli.py tests/unit/test_remote_harness_windows_scripts.py -q`

Expected: PASS.

---

### Task 8: Full verification and portable Windows package

**Files:**
- Create: `docs/REMOTE_WORD_HARNESS.md`
- Create: `REMOTE_HARNESS_SOURCE_SHA256.txt`
- Package artifact: `/mnt/data/WordReplica-Remote-Word-Harness-v1.zip`

**Interfaces:**
- Final ZIP root directly contains `START_REMOTE_HARNESS.cmd`, `RUN_REMOTE_WORD_HARNESS.ps1`, `harness_config.json`, `realworld_input/`, `scripts/remote_harness/`, `src/word_replica/`, and required project metadata.

- [ ] **Step 1: Run targeted harness tests**

Run: `PYTHONPATH=src:. python -m pytest tests/unit/test_remote_harness_*.py -q`

Expected: all PASS.

- [ ] **Step 2: Run complete non-Word regression suite**

Run: `PYTHONPATH=src:. python -m pytest -q`

Expected: zero failures; Windows/Word-only tests may skip on non-Windows.

- [ ] **Step 3: Compile source**

Run: `PYTHONPATH=src:. python -m compileall -q src scripts/remote_harness`

Expected: exit code 0.

- [ ] **Step 4: Generate source manifest and package ZIP**

Manifest includes SHA-256 for production `src/word_replica/**`, harness scripts, bootstrap files, config, pyproject, and docs; exclude caches, virtual environments, test runtime results, and input DOCX files.

- [ ] **Step 5: Verify ZIP contents and checksum**

Open ZIP with `zipfile`, call `testzip()`, assert required files exist, assert no `.venv`, `__pycache__`, `.pytest_cache`, `.pyc`, or real input DOCX is packaged, then calculate final ZIP SHA-256.

