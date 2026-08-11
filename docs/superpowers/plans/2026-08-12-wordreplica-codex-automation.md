# WordReplica Codex Local Automation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Windows/Codex development loop that can run the real Microsoft Word Golden #1 reconstruction, produce deterministic G0–G9 diagnostics, let Codex make evidence-based TDD fixes, and iterate without manual ZIP/update/upload work.

**Architecture:** The private `WordReplica-Automation` repository contains source, tests, Codex instructions, and deterministic automation scripts. The Golden DOCX and heavy diagnostics remain outside Git under `C:\WordReplica-Automation\`; Codex works on `automation-dev`, calls one local Golden command, reads `golden_report.json`, and repeats. GitHub is version control/backup only during normal development; no routine Actions artifacts are created.

**Tech Stack:** Python 3.12+, pywin32, Microsoft Word desktop COM, PowerShell 5.1+, pytest, lxml/python-docx/pypdfium2/Pillow, Git, GitHub private repository, Codex desktop on Windows.

## Global Constraints

- Repository: `danielrisavi77-create/WordReplica-Automation` and it must remain private.
- `main` contains stable code only; autonomous work happens on `automation-dev`.
- Golden source is local-only and immutable: `C:\WordReplica-Automation\golden\Glavna verzija rektorova (grupno)(1).docx`.
- Heavy run artifacts stay local under `C:\WordReplica-Automation\diagnostics\` and `C:\WordReplica-Automation\work\`.
- Normal development uploads no DOCX/PDF/PNG/event-trace artifacts to GitHub.
- Every run writes exactly `golden_report.json`.
- One active Golden Word run at a time.
- Every automation run creates and owns a separate Word COM instance.
- Never run `taskkill /IM WINWORD.EXE /F`; only an automation-owned PID may be terminated, and only after ownership verification.
- Golden source SHA-256 must be identical before and after every run.
- Every production fix follows RED → minimal fix → GREEN → full regression suite → real Word Golden run.
- Stop autonomous repair after three consecutive production iterations without gate improvement, on unexplained regression of a previously passing gate, or when diagnostics do not identify a concrete root cause.
- Promote to `main` only after all local tests pass and the same commit produces two consecutive G0–G9 FULL PASS runs with unchanged Golden source hash.

---

## Target File Structure

```text
WordReplica-Automation/
├── AGENTS.md
├── .gitignore
├── pyproject.toml
├── src/word_replica/...
├── scripts/
│   ├── codex_automation/
│   │   ├── __init__.py
│   │   ├── config.py
│   │   ├── word_process.py
│   │   ├── run_workspace.py
│   │   ├── golden_audit.py
│   │   ├── report.py
│   │   ├── retention.py
│   │   ├── run_golden.py
│   │   └── bootstrap_windows.ps1
│   └── remote_harness/...
├── tests/
│   ├── unit/
│   │   ├── test_codex_automation_config.py
│   │   ├── test_codex_word_process.py
│   │   ├── test_codex_run_workspace.py
│   │   ├── test_golden_audit.py
│   │   ├── test_golden_report.py
│   │   └── test_codex_retention.py
│   ├── integration/
│   │   ├── test_golden_rektorova_proxy.py
│   │   └── test_codex_golden_runner.py
│   └── windows/
│       └── test_codex_real_word_golden.py
└── docs/
    ├── superpowers/specs/2026-08-12-wordreplica-codex-automation-design.md
    ├── superpowers/plans/2026-08-12-wordreplica-codex-automation.md
    └── CODEX_WINDOWS_SETUP.md
```

Local-only machine layout:

```text
C:\WordReplica-Automation\
├── repo\
├── golden\
│   └── Glavna verzija rektorova (grupno)(1).docx
├── work\
├── diagnostics\
├── archive\
└── state\
```

---

### Task 1: Seed the private repository with the validated WordReplica 2.0.4 source

**Files:**
- Copy into repository: `src/word_replica/**`
- Copy into repository: `scripts/remote_harness/**`
- Copy into repository: `scripts/persistent_harness/**`
- Copy into repository: `tests/**`
- Copy into repository: `pyproject.toml`
- Modify: `.gitignore`
- Preserve: `docs/superpowers/specs/2026-08-12-wordreplica-codex-automation-design.md`

**Interfaces:**
- Consumes: validated local WordReplica 2.0.4 source tree.
- Produces: a reproducible Git repository on `main` containing the exact code baseline that produced the `269 passed, 52 skipped` local gate.

- [ ] **Step 1: Add a repository hygiene test**

Create `tests/unit/test_repository_hygiene.py`:

```python
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]

FORBIDDEN_TRACKED_PATTERNS = (
    "*.docx",
    "*.pdf",
    "*.png",
    "*.jpg",
    "*.jpeg",
    "*.zip",
    "event_trace.jsonl",
    "golden_report.json",
)

def test_local_golden_and_heavy_diagnostics_are_not_inside_repo():
    forbidden_roots = {"golden", "work", "diagnostics", "archive", "state"}
    assert not any((ROOT / name).exists() for name in forbidden_roots)
```

- [ ] **Step 2: Verify the hygiene test fails against an intentionally created local-only directory**

Run:

```powershell
New-Item -ItemType Directory -Force golden | Out-Null
python -m pytest tests/unit/test_repository_hygiene.py -q
Remove-Item -Recurse -Force golden
```

Expected: FAIL while `golden\` exists.

- [ ] **Step 3: Copy the 2.0.4 source baseline and strengthen `.gitignore`**

`.gitignore` must include:

```gitignore
.worktrees/
__pycache__/
.pytest_cache/
*.py[cod]

# Local Codex/Golden data must never enter Git.
golden/
work/
diagnostics/
archive/
state/
*.docx
*.pdf
*.png
*.jpg
*.jpeg
*.zip
event_trace.jsonl
golden_report.json
```

Do not copy `.pytest_cache`, `__pycache__`, built EXEs, user result ZIPs, Golden DOCX, PDFs, or rendered pages.

- [ ] **Step 4: Run the full baseline suite**

Run:

```powershell
python -m pytest -q
```

Expected: no failures; Windows-only tests may skip when Microsoft Word is unavailable in the implementation environment.

- [ ] **Step 5: Commit the imported baseline**

```powershell
git add .
git commit -m "chore: seed WordReplica 2.0.4 automation baseline"
```

---

### Task 2: Add the persistent Codex operating contract and branch guard

**Files:**
- Create: `AGENTS.md`
- Create: `scripts/codex_automation/__init__.py`
- Create: `scripts/codex_automation/config.py`
- Create: `tests/unit/test_codex_automation_config.py`

**Interfaces:**
- Consumes: repository root and environment variables.
- Produces: `AutomationConfig.load()` returning validated local paths and policy values; Codex reads `AGENTS.md` automatically as project guidance.

- [ ] **Step 1: Write failing config tests**

Create `tests/unit/test_codex_automation_config.py`:

```python
from pathlib import Path
import pytest

from scripts.codex_automation.config import AutomationConfig


def test_config_requires_paths_outside_repo(tmp_path, monkeypatch):
    repo = tmp_path / "repo"
    repo.mkdir()
    golden = tmp_path / "golden" / "golden.docx"
    golden.parent.mkdir()
    golden.write_bytes(b"PK")
    monkeypatch.setenv("WORD_REPLICA_AUTOMATION_ROOT", str(tmp_path))
    monkeypatch.setenv("WORD_REPLICA_GOLDEN_DOCX", str(golden))
    cfg = AutomationConfig.load(repo_root=repo)
    assert cfg.root == tmp_path.resolve()
    assert cfg.golden_docx == golden.resolve()
    assert cfg.work_dir == (tmp_path / "work").resolve()
    assert cfg.diagnostics_dir == (tmp_path / "diagnostics").resolve()
    assert cfg.state_dir == (tmp_path / "state").resolve()


def test_config_rejects_golden_inside_repo(tmp_path, monkeypatch):
    repo = tmp_path / "repo"
    repo.mkdir()
    golden = repo / "golden.docx"
    golden.write_bytes(b"PK")
    monkeypatch.setenv("WORD_REPLICA_AUTOMATION_ROOT", str(tmp_path))
    monkeypatch.setenv("WORD_REPLICA_GOLDEN_DOCX", str(golden))
    with pytest.raises(ValueError, match="outside the Git repository"):
        AutomationConfig.load(repo_root=repo)
```

- [ ] **Step 2: Run tests to confirm RED**

```powershell
python -m pytest tests/unit/test_codex_automation_config.py -q
```

Expected: import/module failure because `config.py` does not exist.

- [ ] **Step 3: Implement `AutomationConfig`**

Create `scripts/codex_automation/config.py` with:

```python
from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path


@dataclass(frozen=True)
class AutomationConfig:
    repo_root: Path
    root: Path
    golden_docx: Path
    work_dir: Path
    diagnostics_dir: Path
    archive_dir: Path
    state_dir: Path
    max_no_improvement_iterations: int = 3

    @classmethod
    def load(cls, *, repo_root: Path) -> "AutomationConfig":
        repo_root = repo_root.resolve()
        root_raw = os.environ.get("WORD_REPLICA_AUTOMATION_ROOT", r"C:\WordReplica-Automation")
        golden_raw = os.environ.get(
            "WORD_REPLICA_GOLDEN_DOCX",
            str(Path(root_raw) / "golden" / "Glavna verzija rektorova (grupno)(1).docx"),
        )
        root = Path(root_raw).resolve()
        golden = Path(golden_raw).resolve()
        if repo_root == golden or repo_root in golden.parents:
            raise ValueError("Golden DOCX must remain outside the Git repository")
        return cls(
            repo_root=repo_root,
            root=root,
            golden_docx=golden,
            work_dir=(root / "work").resolve(),
            diagnostics_dir=(root / "diagnostics").resolve(),
            archive_dir=(root / "archive").resolve(),
            state_dir=(root / "state").resolve(),
        )
```

- [ ] **Step 4: Create `AGENTS.md`**

`AGENTS.md` must state, verbatim in operational terms:

```markdown
# WordReplica Automation Contract

Primary goal: bring GOLDEN #1 to G0-G9 FULL FIDELITY PASS.

Mandatory loop for every production fix:
1. Read the latest `golden_report.json` and local diagnostics.
2. Identify the first concrete divergence/root cause.
3. Write a regression test that fails for that root cause.
4. Run it and record RED.
5. Make the smallest production change.
6. Run the targeted test and record GREEN.
7. Run the full local pytest suite.
8. Only if the suite is green, run the real Microsoft Word Golden command.
9. Compare the new G0-G9 report with the previous run.
10. Continue only when the next change is supported by evidence.

Safety:
- Never modify the Golden source DOCX.
- Never globally terminate WINWORD.EXE.
- Never close unrelated user Word documents.
- Never work directly on `main`.
- Stop after three production iterations with no Golden gate improvement.
- Stop on an unexplained regression of a previously passing Golden gate.
- Stop when process ownership cannot be proven.
- Do not claim FULL PASS without two consecutive full passes on the same commit and unchanged source hash.

Git:
- Autonomous work branch: `automation-dev`.
- `main` is stable only.
- Do not commit failed hypotheses.
- Checkpoint only after tests are green and Golden evidence improves.

Data:
- Golden and diagnostics stay under `C:\WordReplica-Automation\`, outside Git.
- Do not add DOCX/PDF/PNG/ZIP/event traces to Git.
```

- [ ] **Step 5: Verify GREEN and commit**

```powershell
python -m pytest tests/unit/test_codex_automation_config.py -q
git add AGENTS.md scripts/codex_automation tests/unit/test_codex_automation_config.py
git commit -m "feat: add Codex automation contract and local config"
```

---

### Task 3: Add isolated per-run workspace and immutable Golden source guard

**Files:**
- Create: `scripts/codex_automation/run_workspace.py`
- Create: `tests/unit/test_codex_run_workspace.py`

**Interfaces:**
- Consumes: `AutomationConfig`, commit SHA, run id.
- Produces: `RunWorkspace.create(config, commit_sha)` with `source_copy`, `run_dir`, `diagnostics_dir`, `source_sha256_before`; `verify_source_unchanged()`.

- [ ] **Step 1: Write failing workspace tests**

```python
from pathlib import Path

from scripts.codex_automation.run_workspace import RunWorkspace


def test_workspace_copies_source_and_verifies_original_hash(tmp_path):
    root = tmp_path / "automation"
    golden = root / "golden" / "golden.docx"
    golden.parent.mkdir(parents=True)
    golden.write_bytes(b"golden-original")
    ws = RunWorkspace.create(
        root=root,
        golden_docx=golden,
        commit_sha="abc123",
        run_id="run-001",
    )
    assert ws.source_copy.read_bytes() == b"golden-original"
    assert ws.source_copy != golden
    ws.verify_source_unchanged()


def test_workspace_detects_original_source_mutation(tmp_path):
    root = tmp_path / "automation"
    golden = root / "golden" / "golden.docx"
    golden.parent.mkdir(parents=True)
    golden.write_bytes(b"before")
    ws = RunWorkspace.create(root=root, golden_docx=golden, commit_sha="abc123", run_id="run-002")
    golden.write_bytes(b"after")
    try:
        ws.verify_source_unchanged()
    except RuntimeError as exc:
        assert "Golden source SHA-256 changed" in str(exc)
    else:
        raise AssertionError("source mutation must fail")
```

- [ ] **Step 2: Run RED**

```powershell
python -m pytest tests/unit/test_codex_run_workspace.py -q
```

Expected: FAIL because module is missing.

- [ ] **Step 3: Implement workspace creation**

Use `hashlib.sha256`, `shutil.copy2`, and a unique run directory:

```text
C:\WordReplica-Automation\work\<run-id>\
├── source\golden.docx
├── interactive\
└── audit\
```

Store the original source hash at creation and compare the original file again at run end.

- [ ] **Step 4: Run GREEN and commit**

```powershell
python -m pytest tests/unit/test_codex_run_workspace.py -q
git add scripts/codex_automation/run_workspace.py tests/unit/test_codex_run_workspace.py
git commit -m "feat: isolate Golden runs and guard source immutability"
```

---

### Task 4: Track the automation-owned Word process safely

**Files:**
- Create: `scripts/codex_automation/word_process.py`
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `src/word_replica/qa/word_render.py`
- Create: `tests/unit/test_codex_word_process.py`
- Modify: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- Consumes: a COM `Word.Application` created with `DispatchEx`.
- Produces: `OwnedWordProcess(pid, hwnd, started_filetime)` and `OwnedWordProcess.verify_current_identity()`; controller exposes owned PID metadata to the automation runner.

- [ ] **Step 1: Write failing process identity tests**

Create pure unit tests around an injectable Windows API adapter:

```python
from scripts.codex_automation.word_process import OwnedWordProcess


class FakeWindowsApi:
    def __init__(self, identity):
        self.identity = identity
        self.terminated = []

    def process_identity(self, pid):
        return self.identity

    def terminate_pid(self, pid):
        self.terminated.append(pid)


def test_owned_process_terminates_only_when_identity_still_matches():
    api = FakeWindowsApi((4242, 100000))
    owned = OwnedWordProcess(pid=4242, started_filetime=100000, api=api)
    owned.terminate_if_still_owned()
    assert api.terminated == [4242]


def test_owned_process_refuses_pid_reuse():
    api = FakeWindowsApi((4242, 200000))
    owned = OwnedWordProcess(pid=4242, started_filetime=100000, api=api)
    assert owned.terminate_if_still_owned() is False
    assert api.terminated == []
```

- [ ] **Step 2: Run RED**

```powershell
python -m pytest tests/unit/test_codex_word_process.py -q
```

Expected: missing module failure.

- [ ] **Step 3: Implement Windows PID identity**

On Windows:
- derive PID from the `Application.Hwnd` window handle using `win32process.GetWindowThreadProcessId`;
- read process creation time using Win32 process APIs;
- record `(pid, creation_time)` immediately after `DispatchEx`;
- before forced termination, reopen the PID and verify its creation time still matches;
- terminate only that PID when identity matches.

On non-Windows platforms, process-ownership functions must raise a clear `RuntimeError("Windows Word process ownership is unavailable")`.

- [ ] **Step 4: Expose ownership from `InteractiveWordController`**

Immediately after each `DispatchEx("Word.Application")`, capture owned process metadata. Normal shutdown remains `Document.Close` + `Application.Quit`; force termination is a parent-runner recovery action, not the default path.

- [ ] **Step 5: Run targeted and full tests**

```powershell
python -m pytest tests/unit/test_codex_word_process.py tests/unit/test_interactive_word_executor.py -q
python -m pytest -q
```

Expected: all runnable tests pass.

- [ ] **Step 6: Commit**

```powershell
git add scripts/codex_automation/word_process.py src/word_replica/renderers/interactive_word.py src/word_replica/qa/word_render.py tests/unit/test_codex_word_process.py tests/unit/test_interactive_word_executor.py
git commit -m "feat: track and isolate automation-owned Word process"
```

---

### Task 5: Build deterministic G0–G9 Golden audit and `golden_report.json`

**Files:**
- Create: `scripts/codex_automation/golden_audit.py`
- Create: `scripts/codex_automation/report.py`
- Create: `tests/unit/test_golden_audit.py`
- Create: `tests/unit/test_golden_report.py`

**Interfaces:**
- Consumes: source DOCX, reconstructed DOCX, source PDF, reconstructed PDF, interactive result/trace metadata.
- Produces: `GoldenReport` serialized exactly to `golden_report.json`.

- [ ] **Step 1: Define the report contract with failing tests**

`tests/unit/test_golden_report.py`:

```python
import json

from scripts.codex_automation.report import GoldenReport, GateResult


def test_golden_report_serializes_required_contract(tmp_path):
    report = GoldenReport(
        run_id="run-001",
        commit_sha="abc123",
        source_sha256="deadbeef",
        word_version="14.0",
        word_build="14.0.7015",
        reconstruction_status="PASS",
        duration_seconds=12.5,
        source_pages=72,
        output_pages=72,
        gates={f"G{i}": GateResult(status="PASS", summary="ok") for i in range(10)},
        first_divergence=None,
        first_completed_event=0,
        last_completed_event=168368,
        warnings=[],
    )
    path = tmp_path / "golden_report.json"
    report.write(path)
    payload = json.loads(path.read_text(encoding="utf-8"))
    assert payload["schema_version"] == 1
    assert payload["gates"]["G9"]["status"] == "PASS"
    assert payload["full_pass"] is True
```

- [ ] **Step 2: Define Golden gate behavior with failing tests**

`tests/unit/test_golden_audit.py` must include fixtures where exactly one property differs and assert the corresponding gate fails:

```python
def test_g2_typography_fails_on_effective_font_mismatch(sample_audit_inputs):
    sample_audit_inputs.output.typography["body.font"] = "Calibri"
    report = audit_golden(sample_audit_inputs)
    assert report.gates["G2"].status == "FAIL"
    assert report.first_divergence["property"] == "body.font"


def test_g3_tables_fails_on_cell_style_mismatch(sample_audit_inputs):
    sample_audit_inputs.output.tables[0]["cells"][0]["style"] = "Heading 1"
    report = audit_golden(sample_audit_inputs)
    assert report.gates["G3"].status == "FAIL"


def test_g8_pagination_fails_on_page_count_or_break_drift(sample_audit_inputs):
    sample_audit_inputs.output.page_count += 1
    report = audit_golden(sample_audit_inputs)
    assert report.gates["G8"].status == "FAIL"
```

- [ ] **Step 3: Implement gate definitions**

Implement the gates with these exact meanings:

```text
G0 Content:
  normalized visible text and field display text match.

G1 Structure:
  paragraph/table/row/cell/section/image/bookmark/field/story counts and ordered structural signatures match.

G2 Typography:
  effective paragraph/run style, font family, size, bold, italic, underline, strike, color, highlight,
  language, alignment, spacing, indentation, keep rules, character spacing/position match.

G3 Tables:
  row/column counts, grid widths, cell spans/merges, cell margins, row heights/rules, borders, shading,
  vertical alignment, and cell paragraph/run effective formatting match.

G4 Images:
  image identity/hash, inline-vs-anchor type, dimensions, crop, wrap/position data available to the parser match.

G5 Page setup:
  page size, orientation, margins, header/footer distance, columns, section break type match.

G6 Header/footer:
  story text, fields, effective formatting, tab positions, linkage, and page-number restart properties match.

G7 Fields:
  field instruction, cached/display result, bookmark targets, and field order match.

G8 Pagination:
  source/output PDF page counts must match exactly; extract text per PDF page with pypdfium2, normalize line endings and whitespace-only runs, hash each page's text, and require the ordered page-text hashes to match exactly. Report the first page whose text partition differs.

G9 Visual fidelity:
  rasterize both PDFs at 144 DPI and reuse `compare_pdfs` thresholds exactly:
  `changed_pixel_tolerance=0.001` and `mae_tolerance=0.25`; report the first failing page.
```

Reuse existing parser and QA functions rather than duplicating their logic where possible.

- [ ] **Step 4: Run RED/GREEN cycle**

```powershell
python -m pytest tests/unit/test_golden_report.py tests/unit/test_golden_audit.py -q
```

Expected after implementation: PASS.

- [ ] **Step 5: Add the existing Golden source proxy as a contract dependency**

Run:

```powershell
$env:WORD_REPLICA_GOLDEN_DOCX='C:\WordReplica-Automation\golden\Glavna verzija rektorova (grupno)(1).docx'
python -m pytest tests/integration/test_golden_rektorova_proxy.py -q
```

Expected on the user's machine after setup: PASS.

- [ ] **Step 6: Commit**

```powershell
git add scripts/codex_automation/golden_audit.py scripts/codex_automation/report.py tests/unit/test_golden_audit.py tests/unit/test_golden_report.py
git commit -m "feat: add deterministic G0-G9 Golden audit contract"
```

---

### Task 6: Build the single-command real Word Golden runner

**Files:**
- Create: `scripts/codex_automation/run_golden.py`
- Create: `tests/integration/test_codex_golden_runner.py`
- Create: `tests/windows/test_codex_real_word_golden.py`
- Reuse: `scripts/remote_harness/child_runner.py`
- Reuse: `src/word_replica/services/rebuild.py`

**Interfaces:**
- CLI: `python -m scripts.codex_automation.run_golden`
- Produces: current run directory, reconstructed DOCX, local diagnostics, and exactly one `golden_report.json`.
- Exit codes: `0` = FULL PASS, `10` = completed with Golden gate failure, `20` = engine/COM failure, `30` = safety/source-integrity failure.

- [ ] **Step 1: Write failing CLI/integration tests**

```python
def test_runner_writes_report_even_when_reconstruction_fails(tmp_path, fake_config, fake_rebuilder):
    fake_rebuilder.raise_error = RuntimeError("synthetic failure")
    code = run_golden(config=fake_config, rebuilder=fake_rebuilder)
    assert code == 20
    report = next(fake_config.diagnostics_dir.rglob("golden_report.json"))
    assert report.exists()
    assert '"reconstruction_status": "ENGINE_FAIL"' in report.read_text(encoding="utf-8")
```

Add a second test asserting the runner calls source-hash verification in a `finally` path.

- [ ] **Step 2: Run RED**

```powershell
python -m pytest tests/integration/test_codex_golden_runner.py -q
```

- [ ] **Step 3: Implement orchestration**

The runner must:
1. load `AutomationConfig`;
2. refuse to run unless the current branch is `automation-dev` or an explicit `--verify-stable` flag is supplied;
3. create a unique `RunWorkspace`;
4. copy the Golden source into the run workspace;
5. capture environment/Word version;
6. run static preflight;
7. run Interactive Maximum against the copied source;
8. export source and output to PDF with the automation-owned Word path;
9. run G0–G9 audit;
10. write `golden_report.json` even on failure;
11. verify original Golden SHA in `finally`;
12. release normal COM resources;
13. on timeout, force-terminate only the tracked automation-owned PID after identity verification;
14. return the documented exit code.

- [ ] **Step 4: Add a Windows-only real Word smoke test**

`tests/windows/test_codex_real_word_golden.py`:

```python
import os
import pytest

pytestmark = pytest.mark.word


def test_real_word_golden_runner_produces_machine_readable_report():
    if os.name != "nt":
        pytest.skip("Windows only")
    if not os.environ.get("WORD_REPLICA_GOLDEN_DOCX"):
        pytest.skip("Golden path is not configured")
    from scripts.codex_automation.run_golden import main
    code = main(["--single-run"])
    assert code in {0, 10}
```

A COM/engine failure must not be accepted as success.

- [ ] **Step 5: Run targeted tests and full suite**

```powershell
python -m pytest tests/integration/test_codex_golden_runner.py -q
python -m pytest -q
```

- [ ] **Step 6: Commit**

```powershell
git add scripts/codex_automation/run_golden.py tests/integration/test_codex_golden_runner.py tests/windows/test_codex_real_word_golden.py
git commit -m "feat: add single-command real Word Golden runner"
```

---

### Task 7: Add local retention and fail-safe state tracking

**Files:**
- Create: `scripts/codex_automation/retention.py`
- Create: `scripts/codex_automation/state.py`
- Create: `tests/unit/test_codex_retention.py`
- Create: `tests/unit/test_codex_automation_state.py`

**Interfaces:**
- Produces: `prune_runs(...)`, `AutomationState.record(report)`, `AutomationState.stop_reason()`.
- Keeps latest successful run, latest two failed runs, and explicitly pinned run IDs.

- [ ] **Step 1: Write failing retention tests**

```python
def test_retention_keeps_latest_success_two_failures_and_pinned(tmp_path):
    # create run dirs with report metadata: success-old, success-new, fail-1, fail-2, fail-3, pinned
    kept = prune_runs(tmp_path, pinned={"pinned"})
    assert kept == {"success-new", "fail-2", "fail-3", "pinned"}
```

- [ ] **Step 2: Write failing fail-safe tests**

```python
def test_state_stops_after_three_non_improving_iterations():
    state = AutomationState.empty()
    state = state.record(fake_report(score=4))
    state = state.record(fake_report(score=4))
    state = state.record(fake_report(score=4))
    state = state.record(fake_report(score=4))
    assert "three consecutive" in state.stop_reason().lower()


def test_state_stops_on_regressed_gate():
    state = AutomationState.empty().record(fake_report(passing={"G0", "G1", "G2"}))
    state = state.record(fake_report(passing={"G0", "G1"}))
    assert "regressed" in state.stop_reason().lower()
```

- [ ] **Step 3: Implement retention and state**

Use report timestamps and gate transitions; never delete the current run. Write state atomically via a temporary file + `Path.replace()`.

- [ ] **Step 4: Run GREEN and commit**

```powershell
python -m pytest tests/unit/test_codex_retention.py tests/unit/test_codex_automation_state.py -q
git add scripts/codex_automation/retention.py scripts/codex_automation/state.py tests/unit/test_codex_retention.py tests/unit/test_codex_automation_state.py
git commit -m "feat: add local run retention and automation fail-safes"
```

---

### Task 8: Add one-time Windows bootstrap

**Files:**
- Create: `scripts/codex_automation/bootstrap_windows.ps1`
- Create: `docs/CODEX_WINDOWS_SETUP.md`
- Create: `tests/unit/test_codex_bootstrap_script.py`

**Interfaces:**
- CLI: `powershell -ExecutionPolicy Bypass -File scripts\codex_automation\bootstrap_windows.ps1`
- Optional CLI override: `-GoldenDocx` accepts an absolute existing DOCX path for non-interactive setup.
- Without `-GoldenDocx`, the script opens a Windows file picker so the user can select the existing Golden DOCX once.
- Produces local directories, Python virtualenv, environment configuration, `automation-dev` branch, and verifies Word/Codex prerequisites without uploading the Golden.

- [ ] **Step 1: Write a static failing test for required safety markers**

```python
from pathlib import Path

def test_bootstrap_script_has_required_safety_and_paths():
    text = Path("scripts/codex_automation/bootstrap_windows.ps1").read_text(encoding="utf-8")
    for required in (
        "C:\\WordReplica-Automation",
        "WORD_REPLICA_AUTOMATION_ROOT",
        "WORD_REPLICA_GOLDEN_DOCX",
        "automation-dev",
        "python -m pytest",
        "DispatchEx",
    ):
        assert required in text
    assert "taskkill /IM WINWORD.EXE" not in text
```

- [ ] **Step 2: Run RED**

```powershell
python -m pytest tests/unit/test_codex_bootstrap_script.py -q
```

- [ ] **Step 3: Implement bootstrap**

The script must:
- require Windows;
- accept an optional existing `.docx` via `-GoldenDocx`; when omitted, use `System.Windows.Forms.OpenFileDialog` filtered to `*.docx`, and abort cleanly if the user cancels;
- create `C:\WordReplica-Automation\{golden,work,diagnostics,archive,state}`;
- copy the supplied Golden once to the canonical local-only path;
- record its SHA-256 in `state\golden_source.sha256`;
- create/reuse `.venv` under the repository and `pip install -e ".[test]"`;
- verify Python >= 3.12;
- verify `win32com.client.DispatchEx("Word.Application")` can open and quit a dedicated Word instance;
- create/check out `automation-dev` from `main` if needed;
- configure the two environment variables at user scope;
- run `tests/integration/test_golden_rektorova_proxy.py`;
- print the exact next Codex instruction, not start an uncontrolled loop.

- [ ] **Step 4: Document the one-time user steps**

`docs/CODEX_WINDOWS_SETUP.md` must tell the user:
1. install/update the official ChatGPT desktop app and select Codex;
2. clone/open `danielrisavi77-create/WordReplica-Automation`;
3. run the bootstrap PowerShell command and choose the original Golden DOCX in the file picker;
4. restart Codex if environment variables were newly set;
5. open the repository in Codex and paste the canonical mission from Task 9.

Codex can work with local folders/repos/terminals, and repository-level `AGENTS.md` supplies persistent project guidance.

- [ ] **Step 5: Run tests and commit**

```powershell
python -m pytest tests/unit/test_codex_bootstrap_script.py -q
git add scripts/codex_automation/bootstrap_windows.ps1 docs/CODEX_WINDOWS_SETUP.md tests/unit/test_codex_bootstrap_script.py
git commit -m "feat: add one-time Windows Codex bootstrap"
```

---

### Task 9: Add the canonical Codex mission and operator command

**Files:**
- Create: `docs/CODEX_GOLDEN_MISSION.md`
- Create: `RUN_GOLDEN_CODEX.ps1`
- Create: `tests/unit/test_codex_operator_files.py`

**Interfaces:**
- User-facing test command: `.\RUN_GOLDEN_CODEX.ps1`
- Codex mission: a single stable prompt that tells Codex to continue until FULL PASS or a fail-safe triggers.

- [ ] **Step 1: Write failing operator-file tests**

```python
from pathlib import Path

def test_operator_script_invokes_only_deterministic_runner():
    text = Path("RUN_GOLDEN_CODEX.ps1").read_text(encoding="utf-8")
    assert "scripts.codex_automation.run_golden" in text
    assert "taskkill /IM WINWORD.EXE" not in text


def test_mission_requires_tdd_and_fail_safe():
    text = Path("docs/CODEX_GOLDEN_MISSION.md").read_text(encoding="utf-8")
    for required in ("RED", "GREEN", "G0", "G9", "three consecutive", "automation-dev"):
        assert required in text
```

- [ ] **Step 2: Run RED**

```powershell
python -m pytest tests/unit/test_codex_operator_files.py -q
```

- [ ] **Step 3: Implement operator script**

`RUN_GOLDEN_CODEX.ps1`:

```powershell
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repo
$python = Join-Path $repo ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "Virtual environment missing. Run scripts\codex_automation\bootstrap_windows.ps1 first."
}
& $python -m scripts.codex_automation.run_golden
exit $LASTEXITCODE
```

- [ ] **Step 4: Write the canonical mission**

`docs/CODEX_GOLDEN_MISSION.md` must instruct Codex:

```text
Continue WordReplica Golden #1 development on automation-dev.

Start by running .\RUN_GOLDEN_CODEX.ps1 and reading the newest golden_report.json.
If FULL PASS is false, select only the first evidence-backed divergence.
Use systematic debugging.
Before production code, write a regression test and demonstrate RED.
Implement the minimal fix, demonstrate GREEN, then run the full pytest suite.
Only after the suite is green, rerun .\RUN_GOLDEN_CODEX.ps1.
Compare G0-G9 against the previous run and record whether a gate improved/regressed.
Repeat while evidence supports progress.
Stop immediately when AGENTS.md fail-safe rules trigger.
When G0-G9 FULL PASS occurs, rerun the exact same commit once more.
Only after two consecutive FULL PASS runs on the same commit and unchanged Golden hash, prepare promotion to main; do not merge without explicit user approval.
```

- [ ] **Step 5: Run GREEN and commit**

```powershell
python -m pytest tests/unit/test_codex_operator_files.py -q
git add RUN_GOLDEN_CODEX.ps1 docs/CODEX_GOLDEN_MISSION.md tests/unit/test_codex_operator_files.py
git commit -m "feat: add Codex Golden mission and operator command"
```

---

### Task 10: End-to-end verification on the user's Windows + Word 2010 machine

**Files:**
- No new production file required unless a failure is reproduced by a RED test.
- Update only if necessary: `docs/CODEX_WINDOWS_SETUP.md`

**Interfaces:**
- Consumes: installed local Codex, private repo checkout, canonical Golden DOCX, Word 2010.
- Produces: verified local automation environment and first autonomous Golden report.

- [ ] **Step 1: Run bootstrap once**

From the repository root on the user's Windows PC:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\codex_automation\bootstrap_windows.ps1
```

The script opens a one-time DOCX picker. If the canonical Golden already exists, it verifies the recorded hash and reuses it rather than overwriting it.

- [ ] **Step 2: Run full local suite**

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

Expected: all non-Word tests PASS; Word-marked tests either PASS or execute only when explicitly selected according to project marker policy.

- [ ] **Step 3: Run the real Word Golden command**

```powershell
.\RUN_GOLDEN_CODEX.ps1
```

Expected:
- no mutation of the Golden source;
- a new local `golden_report.json`;
- exit `0`, `10`, or a documented failure exit;
- no unrelated Word document is closed.

- [ ] **Step 4: Verify report contract**

The report must include:
- `schema_version == 1`;
- commit SHA;
- source SHA;
- the Word version/build returned by COM (the current target machine is expected to report Word 14.0 / build 14.0.7015);
- G0–G9 statuses;
- source/output page counts where output exists;
- first divergence when not full pass;
- first/last completed reconstruction event;
- reconstruction status.

- [ ] **Step 5: Start Codex with the canonical mission**

Open the repository in Codex and use the content of `docs/CODEX_GOLDEN_MISSION.md` as the task. Verify that Codex:
- works on `automation-dev`;
- reads `AGENTS.md`;
- runs the deterministic Golden command;
- performs a RED/GREEN cycle before a production fix;
- reruns the Golden test without user ZIP/update/upload intervention.

- [ ] **Step 6: Verify fail-safe behavior**

Before enabling long autonomous work, use a controlled test fixture/state to prove:
- three no-improvement iterations cause stop;
- regressing a passing gate causes stop;
- process identity mismatch prevents forced termination.

- [ ] **Step 7: Verify Git/storage behavior**

Run:

```powershell
git status --short
git ls-files
```

Expected:
- no `golden/`, `work/`, `diagnostics/`, `archive/`, `state/`;
- no DOCX/PDF/PNG/ZIP/event-trace;
- normal Golden execution creates zero GitHub Actions artifacts because Actions are not part of the development loop.

- [ ] **Step 8: Commit setup-document corrections only if verification exposed an instruction mismatch**

If documentation was accurate, make no empty commit. If a concrete setup instruction required correction, test the corrected command and commit only that correction.

---

## Final Verification Gate

Before declaring the automation setup complete, run fresh:

```powershell
python -m pytest -q
$env:WORD_REPLICA_GOLDEN_DOCX='C:\WordReplica-Automation\golden\Glavna verzija rektorova (grupno)(1).docx'
python -m pytest tests/integration/test_golden_rektorova_proxy.py -q
.\RUN_GOLDEN_CODEX.ps1
git status --short
```

Required evidence:
- full test suite has zero failures;
- Golden proxy passes;
- real Word runner creates `golden_report.json`;
- original Golden SHA is unchanged;
- Git status contains no local Golden/heavy diagnostic artifacts;
- no unrelated Word process was terminated;
- Codex can begin the next evidence-backed RED/GREEN iteration without any manual ZIP transfer.

## Promotion Gate

Do **not** merge `automation-dev` to `main` merely because the automation infrastructure works. WordReplica product promotion remains separate:

```text
same commit
  -> Golden #1 FULL PASS
  -> repeat same commit
  -> Golden #1 FULL PASS again
  -> unchanged source hash
  -> full pytest suite PASS
  -> user explicitly approves promotion
  -> merge automation-dev to main
```
