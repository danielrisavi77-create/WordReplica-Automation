# Word Replica Persistent Remote Harness + Safe Updater Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convert the version-per-folder Remote Word Harness into one persistent Windows harness directory that preserves the real-world DOCX corpus and historical results while accepting rollback-safe code-only update ZIPs.

**Architecture:** Keep persistent user data (`realworld_input`, `results`, `harness_config.json`, `.venv_harness`) outside the replaceable `current` tree. A stable top-level launcher invokes `current/RUN_REMOTE_WORD_HARNESS.ps1` with explicit persistent paths and version metadata. A Python updater validates update archives and payload hashes, stages code, backs up the active tree, atomically swaps `current`, validates the new version, and rolls back on failure; thin ASCII-safe CMD/PowerShell launchers expose update and rollback to Windows users.

**Tech Stack:** Python 3.12+, pathlib/json/hashlib/zipfile/shutil/subprocess, PowerShell 5.1-compatible ASCII scripts, existing Word Replica pytest suite and Remote Harness runner.

## Global Constraints

- `realworld_input/**`, `results/**`, top-level `harness_config.json`, and user files outside `current`, `backup`, and `updates` are protected persistent data.
- Replaceable harness code lives only under `current/`.
- Top-level `.venv_harness/` is reusable and is never part of an update payload.
- Update archives may contain only `update_manifest.json` and `payload/**`; traversal, absolute, drive-qualified, and link-like escape paths are rejected.
- Update payload hashes are verified before any payload code executes.
- A failed preflight leaves `current` unchanged; a failed post-install validation restores the previous `current` and version metadata.
- `START_HERE.cmd`, `UPDATE_HARNESS.cmd`, and `ROLLBACK_HARNESS.cmd` must be Windows PowerShell 5.1/ASCII safe.
- The updater never downloads or uploads anything automatically.
- Existing Remote Harness document-processing semantics remain unchanged except for explicit persistent path/version arguments.
- Existing Instant/Interactive renderer behavior is out of scope for this feature.

---

### Task 1: Persistent Layout and Version Contracts

**Files:**
- Create: `scripts/persistent_harness/__init__.py`
- Create: `scripts/persistent_harness/layout.py`
- Create: `scripts/persistent_harness/versioning.py`
- Test: `tests/unit/test_persistent_harness_layout.py`

**Interfaces:**
- Produces: `PersistentLayout.from_root(root: Path) -> PersistentLayout`
- Produces: `VersionMetadata.load(path: Path) -> VersionMetadata`
- Produces: `VersionMetadata.write_atomic(path: Path) -> None`

- [ ] **Step 1: Write failing layout/version tests**

```python
from pathlib import Path
from scripts.persistent_harness.layout import PersistentLayout
from scripts.persistent_harness.versioning import VersionMetadata


def test_layout_keeps_user_data_outside_current(tmp_path: Path):
    layout = PersistentLayout.from_root(tmp_path)
    assert layout.input_dir == tmp_path / "realworld_input"
    assert layout.results_root == tmp_path / "results"
    assert layout.current == tmp_path / "current"
    assert layout.venv == tmp_path / ".venv_harness"


def test_version_metadata_round_trips_atomically(tmp_path: Path):
    path = tmp_path / "version.json"
    value = VersionMetadata(1, "2.0.0", "2026-08-11T19:00:00+02:00", "a" * 64, "1.0.0")
    value.write_atomic(path)
    assert VersionMetadata.load(path) == value
    assert not (tmp_path / "version.json.tmp").exists()
```

- [ ] **Step 2: Run RED**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_layout.py -q`
Expected: FAIL because `scripts.persistent_harness` does not exist.

- [ ] **Step 3: Implement immutable layout/version contracts**

```python
@dataclass(frozen=True, slots=True)
class PersistentLayout:
    root: Path
    input_dir: Path
    results_root: Path
    current: Path
    backup: Path
    updates: Path
    venv: Path
    config: Path
    version_file: Path

    @classmethod
    def from_root(cls, root: Path) -> "PersistentLayout":
        root = Path(root).resolve()
        return cls(root, root / "realworld_input", root / "results", root / "current",
                   root / "backup", root / "updates", root / ".venv_harness",
                   root / "harness_config.json", root / "version.json")
```

`VersionMetadata.write_atomic` writes JSON to a sibling temporary file and uses `Path.replace()`.

- [ ] **Step 4: Run GREEN**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_layout.py -q`
Expected: PASS.

---

### Task 2: Versioned Run Directories and Explicit Harness Paths

**Files:**
- Modify: `scripts/remote_harness/__init__.py`
- Modify: `scripts/remote_harness/main.py`
- Modify: `scripts/remote_harness/environment.py`
- Test: `tests/unit/test_remote_harness_cli.py`
- Test: `tests/unit/test_persistent_harness_results.py`

**Interfaces:**
- Extends CLI with `--harness-version`.
- Result directory format: `<YYYY-MM-DD_HHMMSS>_v<safe-version>`.
- Result ZIP format: `WordReplica-Remote-Results-<YYYYMMDD-HHMMSS>-v<safe-version>.zip` inside that run directory.

- [ ] **Step 1: Write failing CLI/result-name tests**

```python
def test_cli_accepts_explicit_persistent_paths_and_version():
    args = build_parser().parse_args([
        "--input-dir", "input", "--result-parent", "results",
        "--config", "cfg.json", "--harness-version", "2.0.0",
    ])
    assert args.harness_version == "2.0.0"


def test_run_names_include_version(monkeypatch, tmp_path):
    name = build_run_name("2.0.0", "20260811-190500")
    assert name == "2026-08-11_190500_v2.0.0"
```

- [ ] **Step 2: Run RED**

Run: `PYTHONPATH=src:. pytest tests/unit/test_remote_harness_cli.py tests/unit/test_persistent_harness_results.py -q`
Expected: FAIL because version-aware naming does not exist.

- [ ] **Step 3: Implement explicit version propagation**

`main.py` must use the explicit `--input-dir`, `--result-parent`, and `--harness-version` values, write `harness_version` into `environment.json` and `summary.json`, and package the ZIP inside the run directory rather than beside source code.

- [ ] **Step 4: Run GREEN**

Run: `PYTHONPATH=src:. pytest tests/unit/test_remote_harness_cli.py tests/unit/test_persistent_harness_results.py -q`
Expected: PASS.

---

### Task 3: Stable Persistent Launcher and Reusable Environment

**Files:**
- Create: `persistent_root/START_HERE.cmd`
- Create: `persistent_root/RUN_PERSISTENT_HARNESS.ps1`
- Modify: `RUN_REMOTE_WORD_HARNESS.ps1`
- Test: `tests/unit/test_persistent_harness_windows_scripts.py`

**Interfaces:**
- `START_HERE.cmd` invokes top-level `RUN_PERSISTENT_HARNESS.ps1`.
- `RUN_PERSISTENT_HARNESS.ps1` invokes `current/RUN_REMOTE_WORD_HARNESS.ps1` with `-InputDir`, `-ResultsRoot`, `-HarnessVersion`, `-ConfigPath`, and `-VenvPath`.
- `current/RUN_REMOTE_WORD_HARNESS.ps1` accepts those explicit parameters and does not derive input/results from its own directory.

- [ ] **Step 1: Write failing script contract test**

```python
def test_persistent_launcher_passes_protected_paths_and_version():
    text = Path("persistent_root/RUN_PERSISTENT_HARNESS.ps1").read_text("ascii")
    assert "realworld_input" in text
    assert "results" in text
    assert "version.json" in text
    assert "-InputDir" in text and "-ResultsRoot" in text and "-HarnessVersion" in text
    assert ".venv_harness" in text
```

- [ ] **Step 2: Run RED**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_windows_scripts.py -q`
Expected: FAIL because persistent launchers do not exist.

- [ ] **Step 3: Implement ASCII-safe stable launcher**

The current PowerShell runner gets a `param(...)` block with explicit paths. It installs dependencies from its code root into the provided reusable venv only when needed and writes results exclusively under the provided results root.

- [ ] **Step 4: Run GREEN**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_windows_scripts.py -q`
Expected: PASS.

---

### Task 4: Update Manifest Validation and Archive Safety

**Files:**
- Create: `scripts/persistent_harness/update_manifest.py`
- Create: `scripts/persistent_harness/archive.py`
- Test: `tests/unit/test_persistent_harness_update_validation.py`

**Interfaces:**
- Produces: `UpdateManifest.load(path: Path) -> UpdateManifest`
- Produces: `validate_update_zip(zip_path: Path, installed_version: str) -> ValidatedUpdate`
- `ValidatedUpdate` contains target version, package SHA-256, manifest, and safe member list.

- [ ] **Step 1: Write RED tests for valid hash and traversal rejection**

```python
def test_update_rejects_traversal_member(tmp_path):
    archive = make_update_zip(tmp_path, extra_member="payload/../../realworld_input/x.docx")
    with pytest.raises(UpdateValidationError, match="unsafe archive path"):
        validate_update_zip(archive, "2.0.0")


def test_update_rejects_payload_hash_mismatch(tmp_path):
    archive = make_update_zip(tmp_path, manifest_hash="0" * 64)
    with pytest.raises(UpdateValidationError, match="SHA-256"):
        validate_update_zip(archive, "2.0.0")
```

- [ ] **Step 2: Run RED**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_update_validation.py -q`
Expected: FAIL because validator is missing.

- [ ] **Step 3: Implement strict archive/manifest validation**

Reject absolute paths, `..`, drive-qualified paths, non-`payload/` members other than `update_manifest.json`, duplicate normalized members, and link-like external attributes. Verify every payload file listed in the manifest and reject unlisted payload files.

- [ ] **Step 4: Run GREEN**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_update_validation.py -q`
Expected: PASS.

---

### Task 5: Transactional Update, Backup, Rollback, and Recovery

**Files:**
- Create: `scripts/persistent_harness/updater.py`
- Create: `scripts/persistent_harness/rollback.py`
- Test: `tests/unit/test_persistent_harness_update_transaction.py`

**Interfaces:**
- Produces: `apply_update(layout: PersistentLayout, zip_path: Path, smoke_test: Callable[[Path], None]) -> VersionMetadata`
- Produces: `rollback_latest(layout: PersistentLayout, smoke_test: Callable[[Path], None]) -> VersionMetadata`
- Uses transaction marker: `<root>/.update_transaction.json`.

- [ ] **Step 1: Write RED tests for protected data, failed preflight, failed post-install, and rollback**

```python
def test_successful_update_never_changes_input_or_results(tmp_path):
    layout = seeded_persistent_layout(tmp_path)
    before_input = tree_hash(layout.input_dir)
    before_results = tree_hash(layout.results_root)
    apply_update(layout, valid_update_zip(tmp_path), smoke_test=lambda current: None)
    assert tree_hash(layout.input_dir) == before_input
    assert tree_hash(layout.results_root) == before_results


def test_post_install_failure_restores_previous_current(tmp_path):
    layout = seeded_persistent_layout(tmp_path)
    old_hash = tree_hash(layout.current)
    with pytest.raises(UpdateTransactionError):
        apply_update(layout, valid_update_zip(tmp_path), smoke_test=lambda current: (_ for _ in ()).throw(RuntimeError("bad")))
    assert tree_hash(layout.current) == old_hash
    assert VersionMetadata.load(layout.version_file).harness_version == "2.0.0"
```

- [ ] **Step 2: Run RED**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_update_transaction.py -q`
Expected: FAIL because transaction logic is missing.

- [ ] **Step 3: Implement staging and rollback-safe swap**

Workflow: validate ZIP → extract validated payload to staging → run preflight smoke test against staging → create timestamped backup containing old `current` and old `version.json` → write transaction marker → rename live `current` to temporary old-live name → rename staging to `current` → atomically write new version → run post-install smoke test → on failure restore old-live/backup and old version → delete transaction marker only after a known-good final state.

- [ ] **Step 4: Add interrupted transaction recovery test**

Seed a marker representing a swapped-but-unvalidated state and assert the next updater invocation restores the referenced known-good backup before accepting another update.

- [ ] **Step 5: Run GREEN**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_update_transaction.py -q`
Expected: PASS.

---

### Task 6: Dependency Fingerprint and Windows Update/Rollback Entry Points

**Files:**
- Create: `scripts/persistent_harness/dependencies.py`
- Create: `scripts/persistent_harness/update_cli.py`
- Create: `scripts/persistent_harness/rollback_cli.py`
- Create: `persistent_root/UPDATE_HARNESS.cmd`
- Create: `persistent_root/UPDATE_HARNESS.ps1`
- Create: `persistent_root/ROLLBACK_HARNESS.cmd`
- Create: `persistent_root/ROLLBACK_HARNESS.ps1`
- Test: `tests/unit/test_persistent_harness_dependency_reuse.py`
- Test: `tests/unit/test_persistent_harness_windows_scripts.py`

**Interfaces:**
- Produces: `dependency_fingerprint(pyproject_path: Path) -> str`.
- Updater synchronizes dependencies only when old/new fingerprints differ.
- Windows updater selects one ZIP from `updates/` or fails clearly if zero/multiple packages exist.

- [ ] **Step 1: Write RED dependency reuse test**

```python
def test_unchanged_dependency_fingerprint_reuses_venv(tmp_path):
    old = dependency_fingerprint(write_pyproject(tmp_path / "old", "lxml>=5"))
    new = dependency_fingerprint(write_pyproject(tmp_path / "new", "lxml>=5"))
    assert old == new
```

- [ ] **Step 2: Write RED Windows script tests**

Assert all four new scripts are ASCII, use fail-fast error handling, reference `updates`, `backup`, and `version.json`, and never contain deletion commands targeting `realworld_input` or `results`.

- [ ] **Step 3: Implement dependency fingerprint and CLI wrappers**

The fingerprint hashes normalized dependency declarations from `pyproject.toml`. `UPDATE_HARNESS.ps1` uses the top-level `.venv_harness` Python interpreter when available, bootstraps it if missing, and invokes only trusted installed updater code; payload code is not invoked until Python validation passes.

- [ ] **Step 4: Run GREEN**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_dependency_reuse.py tests/unit/test_persistent_harness_windows_scripts.py -q`
Expected: PASS.

---

### Task 7: Persistent Bootstrap Builder and v2 Migration

**Files:**
- Create: `scripts/persistent_harness/build_bootstrap.py`
- Create: `scripts/persistent_harness/build_update.py`
- Create: `persistent_root/IMPORT_EXISTING_INPUT.cmd`
- Create: `persistent_root/IMPORT_EXISTING_INPUT.ps1`
- Test: `tests/unit/test_persistent_harness_bootstrap.py`
- Test: `tests/unit/test_persistent_harness_migration.py`

**Interfaces:**
- Produces bootstrap tree with top-level persistent scripts/data folders and replaceable code under `current/`.
- Produces future update ZIPs with `update_manifest.json` + `payload/**` only.
- Migration copies DOCX files and refuses overwrite when an existing same-name file has a different SHA-256.

- [ ] **Step 1: Write RED bootstrap structure test**

```python
def test_bootstrap_places_code_only_under_current(tmp_path):
    root = build_bootstrap(source_root=PROJECT_ROOT, destination=tmp_path / "persistent")
    assert (root / "current" / "scripts" / "remote_harness" / "main.py").exists()
    assert (root / "realworld_input").is_dir()
    assert (root / "results").is_dir()
    assert not (root / "current" / "realworld_input").exists()
```

- [ ] **Step 2: Write RED migration collision test**

Copy one input successfully, then place a different same-name DOCX in destination and assert migration refuses to overwrite it.

- [ ] **Step 3: Implement bootstrap/update builders and migration**

Bootstrap version is `2.0.0-persistent.1`. Its `current/` payload is copied from the verified v2 source while excluding caches, virtual environments, result folders, input documents, build artifacts, and prior ZIPs. The initial `version.json` records the source package SHA-256 and no previous persistent version.

- [ ] **Step 4: Run GREEN**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_bootstrap.py tests/unit/test_persistent_harness_migration.py -q`
Expected: PASS.

---

### Task 8: Full Regression Gate, Package Integrity, and User Artifact

**Files:**
- Modify: `README.md`
- Create: `docs/PERSISTENT_REMOTE_HARNESS.md`
- Test: full existing suite plus all new persistent tests.

**Interfaces:**
- Final artifact: `WordReplica-Remote-Harness-Persistent-v1.zip`.
- Future update artifact builder can produce `WordReplica-Remote-Harness-Update-<version>.zip` without persistent user data.

- [ ] **Step 1: Run targeted persistent suite**

Run: `PYTHONPATH=src:. pytest tests/unit/test_persistent_harness_*.py tests/unit/test_remote_harness_*.py -q`
Expected: all PASS.

- [ ] **Step 2: Run full non-Word regression suite**

Run: `PYTHONPATH=src:. pytest -q`
Expected: zero failures; Windows/Microsoft Word-only tests may skip in this sandbox.

- [ ] **Step 3: Compile Python tree**

Run: `python -m compileall -q src scripts`
Expected: exit code 0.

- [ ] **Step 4: Build bootstrap artifact and verify protected layout**

Build `WordReplica-Remote-Harness-Persistent-v1.zip`; verify ZIP integrity, manifest hashes, absence of `.venv`, `__pycache__`, `.pytest_cache`, `realworld_input/*.docx`, and any previous result ZIPs.

- [ ] **Step 5: Verify update package builder on a no-op synthetic version bump**

Generate a test update ZIP from a copied source tree, validate it with the production validator, and assert it contains only `update_manifest.json` and `payload/**`.

- [ ] **Step 6: Document Windows workflow**

Document: extract persistent ZIP once → copy/import DOCX once → use one `START_HERE.cmd` forever → place future update ZIP in `updates/` → run `UPDATE_HARNESS.cmd` → retain all versioned result folders → use `ROLLBACK_HARNESS.cmd` only when needed.
