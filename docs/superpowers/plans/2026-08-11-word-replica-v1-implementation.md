# Word Replica v1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Windows-first application that reconstructs an input `.docx` into a new `.docx` with maximum practical fidelity, truthful save/audit history, dual renderers, and automated L0-L4 QA.

**Architecture:** Parse the source DOCX/OPC package into a deterministic renderer-neutral canonical model, then render that model through either Microsoft Word COM or a pure-DOCX fallback. A shared application service owns source protection, checkpoints, metadata policy, project storage, audit logging, QA, CLI, and GUI so no interface duplicates reconstruction logic.

**Tech Stack:** Python 3.12+, `lxml`, `python-docx`, `pywin32` on Windows, `PySide6`, SQLite, JSONL, `pypdfium2`, Pillow, Jinja2, pytest.

## Global Constraints

- Input is `.docx` only in v1.
- Source document is read-only by default and must be hash-verified before and after reconstruction.
- Full Fidelity is enhanced when desktop Microsoft Word is available; fallback must work without Word.
- Rendering profiles: `clean` and `full`.
- Metadata profiles: `fresh` and `preserve`.
- Preserve mode may copy only legitimate descriptive properties; it must never copy or fabricate source creation/modification timestamps, revision count, total editing time, or historical save/revision values.
- Every logged save must correspond to a real file save operation.
- Exact save count is exposed through the truthful custom document property `WordReplicaActualSaveCount`; built-in Word revision count/total editing time are never repurposed as a fake save counter.
- No simulated human typing, fake keystroke timing, deliberate typos, fake corrections, forged revision history, backdated timestamps, inflated editing time, or provenance/AI-detection evasion.
- Logs must not copy full document text by default.
- GUI and CLI must call the same application/service layer.
- Unit/core tests must run without Microsoft Word; Word COM tests are explicitly gated Windows integration tests.
- Critical reconstruction failures stop the run; non-critical fidelity gaps continue with warnings.

---

## File Map

The implementation should converge on this structure:

```text
word-replica-v1/
├─ pyproject.toml
├─ README.md
├─ src/word_replica/
│  ├─ __init__.py
│  ├─ cli.py
│  ├─ config.py
│  ├─ domain/
│  │  ├─ enums.py
│  │  ├─ errors.py
│  │  ├─ model.py
│  │  └─ results.py
│  ├─ opc/
│  │  ├─ package_reader.py
│  │  └─ properties.py
│  ├─ parser/
│  │  ├─ parser.py
│  │  ├─ text.py
│  │  ├─ tables.py
│  │  ├─ relationships.py
│  │  └─ fidelity.py
│  ├─ renderers/
│  │  ├─ base.py
│  │  ├─ pure_docx.py
│  │  └─ word_com.py
│  ├─ services/
│  │  ├─ audit.py
│  │  ├─ checkpoints.py
│  │  ├─ project_store.py
│  │  ├─ source_guard.py
│  │  └─ rebuild.py
│  ├─ qa/
│  │  ├─ content.py
│  │  ├─ structure.py
│  │  ├─ formatting.py
│  │  ├─ layout.py
│  │  ├─ render.py
│  │  ├─ policy.py
│  │  └─ report.py
│  └─ gui/
│     ├─ app.py
│     ├─ main_window.py
│     └─ controller.py
├─ tests/
│  ├─ conftest.py
│  ├─ fixtures/
│  │  ├─ build_fixtures.py
│  │  └─ corpus/
│  ├─ unit/
│  ├─ integration/
│  │  ├─ test_pure_docx_e2e.py
│  │  └─ word/
│  └─ acceptance/
│     └─ test_acceptance_matrix.py
└─ scripts/
   ├─ build_windows.ps1
   └─ run_word_integration.ps1
```

---

### Task 1: Project Skeleton, Profiles, and Typed Run Configuration

**Files:**
- Create: `pyproject.toml`
- Create: `src/word_replica/__init__.py`
- Create: `src/word_replica/config.py`
- Create: `src/word_replica/domain/enums.py`
- Create: `src/word_replica/domain/errors.py`
- Create: `src/word_replica/domain/results.py`
- Create: `tests/unit/test_config.py`

**Interfaces:**
- Produces: `RebuildOptions`, `RendererChoice`, `VisibilityMode`, `FidelityMode`, `MetadataMode`, `RunStatus`, and shared exception classes.
- Later tasks must import these types rather than using raw strings.

- [ ] **Step 1: Add packaging metadata and test dependencies**

```toml
[build-system]
requires = ["setuptools>=75", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "word-replica"
version = "0.1.0"
requires-python = ">=3.12"
dependencies = [
  "lxml>=5,<7",
  "python-docx>=1.1,<2",
  "platformdirs>=4,<5",
  "PySide6>=6.8,<7",
  "pypdfium2>=4,<6",
  "Pillow>=10,<13",
  "Jinja2>=3.1,<4",
  "pywin32>=306; sys_platform == 'win32'",
]

[project.optional-dependencies]
test = ["pytest>=8,<10", "pytest-cov>=5,<8"]
build = ["pyinstaller>=6,<8"]

[project.scripts]
word-replica = "word_replica.cli:main"
word-replica-gui = "word_replica.gui.app:main"

[tool.pytest.ini_options]
testpaths = ["tests"]
markers = ["word: requires Windows + Microsoft Word desktop"]
```

- [ ] **Step 2: Write the failing profile-validation tests**

```python
# tests/unit/test_config.py
import pytest
from word_replica.config import RebuildOptions
from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, VisibilityMode


def test_defaults_are_safe_and_truthful():
    options = RebuildOptions()
    assert options.renderer is RendererChoice.AUTO
    assert options.visibility is VisibilityMode.BACKGROUND
    assert options.fidelity is FidelityMode.CLEAN
    assert options.metadata is MetadataMode.FRESH
    assert options.allow_source_overwrite is False
    assert options.preserve_author_fields is False
    assert options.custom_metadata_allowlist == ()
    assert options.periodic_save_seconds == 60


def test_periodic_save_must_be_positive():
    with pytest.raises(ValueError, match="periodic_save_seconds"):
        RebuildOptions(periodic_save_seconds=0)
```

- [ ] **Step 3: Run tests and verify failure**

Run: `python -m pytest tests/unit/test_config.py -v`

Expected: FAIL because `word_replica.config` and enums do not exist.

- [ ] **Step 4: Implement enums, configuration, statuses, and errors**

```python
# src/word_replica/domain/enums.py
from enum import StrEnum


class RendererChoice(StrEnum):
    AUTO = "auto"
    WORD = "word"
    DOCX = "docx"


class VisibilityMode(StrEnum):
    VISIBLE = "visible"
    BACKGROUND = "background"


class FidelityMode(StrEnum):
    CLEAN = "clean"
    FULL = "full"


class MetadataMode(StrEnum):
    FRESH = "fresh"
    PRESERVE = "preserve"


class RunStatus(StrEnum):
    PASS = "PASS"
    WARN = "WARN"
    FAIL = "FAIL"
```

```python
# src/word_replica/config.py
from dataclasses import dataclass
from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, VisibilityMode


@dataclass(frozen=True, slots=True)
class RebuildOptions:
    renderer: RendererChoice = RendererChoice.AUTO
    visibility: VisibilityMode = VisibilityMode.BACKGROUND
    fidelity: FidelityMode = FidelityMode.CLEAN
    metadata: MetadataMode = MetadataMode.FRESH
    allow_source_overwrite: bool = False
    preserve_author_fields: bool = False
    custom_metadata_allowlist: tuple[str, ...] = ()
    periodic_save_seconds: int = 60

    def __post_init__(self) -> None:
        if self.periodic_save_seconds <= 0:
            raise ValueError("periodic_save_seconds must be > 0")
```

```python
# src/word_replica/domain/errors.py
class WordReplicaError(Exception):
    pass


class CriticalRebuildError(WordReplicaError):
    pass


class SourceIntegrityError(CriticalRebuildError):
    pass


class RendererUnavailableError(CriticalRebuildError):
    pass


class PackageReadError(CriticalRebuildError):
    pass
```

```python
# src/word_replica/domain/results.py
from dataclasses import dataclass, field
from pathlib import Path
from word_replica.domain.enums import RunStatus


@dataclass(slots=True)
class WarningItem:
    code: str
    message: str
    element_id: str | None = None
    affects_status: bool = True


@dataclass(slots=True)
class RunResult:
    status: RunStatus
    output_path: Path | None
    qa_report_path: Path | None
    project_id: str | None = None
    save_count: int = 0
    warnings: list[WarningItem] = field(default_factory=list)
    reasons: list[str] = field(default_factory=list)
```

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_config.py -v`

Expected: PASS.

```bash
git add pyproject.toml src/word_replica tests/unit/test_config.py
git commit -m "feat: scaffold word replica core configuration"
```

---

### Task 2: Project Store, Source Protection, and Central SQLite Index

**Files:**
- Create: `src/word_replica/services/source_guard.py`
- Create: `src/word_replica/services/project_store.py`
- Create: `tests/unit/test_project_store.py`
- Create: `tests/unit/test_source_guard.py`

**Interfaces:**
- Consumes: `RebuildOptions`.
- Produces: `sha256_file(path: Path) -> str`, `SourceSnapshot`, `ProjectPaths`, `ProjectStore.create_project(...)`, `ProjectStore.set_status(...)`, and `ProjectStore.list_projects()`.

- [ ] **Step 1: Write source-integrity tests**

```python
# tests/unit/test_source_guard.py
from pathlib import Path
import pytest
from word_replica.services.source_guard import capture_source, assert_source_unchanged
from word_replica.domain.errors import SourceIntegrityError


def test_source_hash_detects_mutation(tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"before")
    snapshot = capture_source(source)
    source.write_bytes(b"after")
    with pytest.raises(SourceIntegrityError):
        assert_source_unchanged(snapshot)
```

- [ ] **Step 2: Write project-layout/index tests**

```python
# tests/unit/test_project_store.py
from pathlib import Path
from word_replica.config import RebuildOptions
from word_replica.services.project_store import ProjectStore


def test_create_project_builds_required_directories(tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"docx")
    store = ProjectStore(app_root=tmp_path / "app")
    paths = store.create_project(source, RebuildOptions())
    assert paths.source_snapshot_dir.is_dir()
    assert paths.working_dir.is_dir()
    assert paths.output_dir.is_dir()
    assert paths.backups_dir.is_dir()
    assert paths.logs_dir.is_dir()
    assert paths.qa_dir.is_dir()
    assert paths.project_json.exists()
    assert store.get_project(paths.project_id)["source_path"] == str(source.resolve())
```

- [ ] **Step 3: Run tests and verify failure**

Run: `python -m pytest tests/unit/test_source_guard.py tests/unit/test_project_store.py -v`

Expected: FAIL because services do not exist.

- [ ] **Step 4: Implement source hashing and project store**

```python
# src/word_replica/services/source_guard.py
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from word_replica.domain.errors import SourceIntegrityError


@dataclass(frozen=True, slots=True)
class SourceSnapshot:
    path: Path
    sha256: str
    size: int


def sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def capture_source(path: Path) -> SourceSnapshot:
    resolved = path.resolve()
    return SourceSnapshot(resolved, sha256_file(resolved), resolved.stat().st_size)


def assert_source_unchanged(snapshot: SourceSnapshot) -> None:
    if not snapshot.path.exists() or sha256_file(snapshot.path) != snapshot.sha256:
        raise SourceIntegrityError(f"Source changed during reconstruction: {snapshot.path}")
```

```python
# src/word_replica/services/project_store.py
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
import shutil
import sqlite3
import uuid
from platformdirs import user_documents_dir
from word_replica.config import RebuildOptions
from word_replica.services.source_guard import capture_source


@dataclass(frozen=True, slots=True)
class ProjectPaths:
    project_id: str
    root: Path
    source_snapshot_dir: Path
    working_dir: Path
    output_dir: Path
    backups_dir: Path
    logs_dir: Path
    qa_dir: Path
    project_json: Path


class ProjectStore:
    def __init__(self, app_root: Path | None = None) -> None:
        self.app_root = app_root or Path(user_documents_dir()) / "WordReplica" / "Projects"
        self.app_root.mkdir(parents=True, exist_ok=True)
        self.db_path = self.app_root / "projects.sqlite3"
        self._init_db()

    def _init_db(self) -> None:
        with sqlite3.connect(self.db_path) as db:
            db.execute("CREATE TABLE IF NOT EXISTS projects (id TEXT PRIMARY KEY, source_path TEXT NOT NULL, root_path TEXT NOT NULL, created_utc TEXT NOT NULL, status TEXT NOT NULL)")

    def create_project(self, source: Path, options: RebuildOptions) -> ProjectPaths:
        project_id = uuid.uuid4().hex
        root = source.parent / f"{source.stem}_rebuild" / project_id
        dirs = {name: root / name for name in ("source_snapshot", "working", "output", "backups", "logs", "qa")}
        for path in dirs.values():
            path.mkdir(parents=True, exist_ok=True)
        snapshot = capture_source(source)
        shutil.copy2(source, dirs["source_snapshot"] / source.name)
        project_json = root / "project.json"
        payload = {"project_id": project_id, "source_path": str(source.resolve()), "source_sha256": snapshot.sha256, "options": asdict(options)}
        project_json.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
        with sqlite3.connect(self.db_path) as db:
            db.execute("INSERT INTO projects VALUES (?, ?, ?, ?, ?)", (project_id, str(source.resolve()), str(root), datetime.now(timezone.utc).isoformat(), "CREATED"))
        return ProjectPaths(project_id, root, dirs["source_snapshot"], dirs["working"], dirs["output"], dirs["backups"], dirs["logs"], dirs["qa"], project_json)

    def get_project(self, project_id: str) -> dict[str, str]:
        with sqlite3.connect(self.db_path) as db:
            row = db.execute("SELECT id, source_path, root_path, created_utc, status FROM projects WHERE id = ?", (project_id,)).fetchone()
        if row is None:
            raise KeyError(project_id)
        return dict(zip(("id", "source_path", "root_path", "created_utc", "status"), row, strict=True))

    def set_status(self, project_id: str, status: str) -> None:
        with sqlite3.connect(self.db_path) as db:
            db.execute("UPDATE projects SET status = ? WHERE id = ?", (status, project_id))

    def list_projects(self) -> list[dict[str, str]]:
        with sqlite3.connect(self.db_path) as db:
            rows = db.execute("SELECT id, source_path, root_path, created_utc, status FROM projects ORDER BY created_utc DESC").fetchall()
        keys = ("id", "source_path", "root_path", "created_utc", "status")
        return [dict(zip(keys, row, strict=True)) for row in rows]
```

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_source_guard.py tests/unit/test_project_store.py -v`

Expected: PASS.

```bash
git add src/word_replica/services tests/unit/test_source_guard.py tests/unit/test_project_store.py
git commit -m "feat: protect sources and create project store"
```

---

### Task 3: Append-Only Audit Log and Truthful Checkpoint Save History

**Files:**
- Create: `src/word_replica/services/audit.py`
- Create: `src/word_replica/services/checkpoints.py`
- Create: `tests/unit/test_audit.py`
- Create: `tests/unit/test_checkpoints.py`

**Interfaces:**
- Produces: `AuditLog.append(event_type, payload)`, `CheckpointManager.next_sequence`, and `CheckpointManager.save(reason, stage, save_callable, document_path)`.
- `save_callable` must perform the real renderer/file save; the manager logs only after success. The final metadata-seal save uses `next_sequence` before saving so `WordReplicaActualSaveCount` equals the sequence that is about to be committed.

- [ ] **Step 1: Write audit/checkpoint tests**

```python
# tests/unit/test_checkpoints.py
import json
from pathlib import Path
from word_replica.services.audit import AuditLog
from word_replica.services.checkpoints import CheckpointManager


def test_checkpoint_logs_only_after_real_save(tmp_path: Path):
    document = tmp_path / "out.docx"
    history = tmp_path / "save_history.jsonl"
    audit = AuditLog(tmp_path / "audit.jsonl")
    manager = CheckpointManager(history, audit)

    def real_save() -> None:
        document.write_bytes(b"saved")

    event = manager.save("chapter complete", "chapter_1", real_save, document)
    row = json.loads(history.read_text(encoding="utf-8").splitlines()[0])
    assert row["sequence"] == 1
    assert row["event_id"]
    assert row["timestamp_utc"]
    assert row["timestamp_local"]
    assert row["reason"] == "chapter complete"
    assert row["document_sha256"] == event.document_sha256
```

- [ ] **Step 2: Run test and verify failure**

Run: `python -m pytest tests/unit/test_audit.py tests/unit/test_checkpoints.py -v`

Expected: FAIL because logging services do not exist.

- [ ] **Step 3: Implement append-only audit and checkpoint manager**

```python
# src/word_replica/services/audit.py
from datetime import datetime, timezone
import json
from pathlib import Path
import uuid


class AuditLog:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def append(self, event_type: str, payload: dict) -> dict:
        event = {"event_id": uuid.uuid4().hex, "event_type": event_type, "timestamp_utc": datetime.now(timezone.utc).isoformat(), "payload": payload}
        with self.path.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(event, ensure_ascii=False, default=str) + "\n")
        return event
```

```python
# src/word_replica/services/checkpoints.py
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Callable
import uuid
from word_replica.services.audit import AuditLog
from word_replica.services.source_guard import sha256_file


@dataclass(frozen=True, slots=True)
class SaveEvent:
    event_id: str
    sequence: int
    reason: str
    stage: str
    timestamp_utc: str
    timestamp_local: str
    document_path: str
    document_sha256: str


class CheckpointManager:
    def __init__(self, history_path: Path, audit: AuditLog) -> None:
        self.history_path = history_path
        self.audit = audit
        self.sequence = 0

    @property
    def next_sequence(self) -> int:
        return self.sequence + 1

    def save(self, reason: str, stage: str, save_callable: Callable[[], None], document_path: Path) -> SaveEvent:
        save_callable()
        self.sequence += 1
        now_utc = datetime.now(timezone.utc)
        event = SaveEvent(uuid.uuid4().hex, self.sequence, reason, stage, now_utc.isoformat(), now_utc.astimezone().isoformat(), str(document_path), sha256_file(document_path))
        with self.history_path.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(asdict(event)) + "\n")
        self.audit.append("DOCUMENT_SAVED", asdict(event))
        return event
```

- [ ] **Step 4: Add failure-path assertion**

```python
def test_failed_save_does_not_increment_or_log(tmp_path: Path):
    history = tmp_path / "save_history.jsonl"
    manager = CheckpointManager(history, AuditLog(tmp_path / "audit.jsonl"))
    try:
        manager.save("x", "x", lambda: (_ for _ in ()).throw(OSError("disk full")), tmp_path / "missing.docx")
    except OSError:
        pass
    assert manager.sequence == 0
    assert not history.exists()
```

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_audit.py tests/unit/test_checkpoints.py -v`

Expected: PASS.

```bash
git add src/word_replica/services/audit.py src/word_replica/services/checkpoints.py tests/unit/test_audit.py tests/unit/test_checkpoints.py
git commit -m "feat: add truthful audit and checkpoint history"
```

---

### Task 4: DOCX/OPC Package Reader and Metadata Policy

**Files:**
- Create: `src/word_replica/opc/package_reader.py`
- Create: `src/word_replica/opc/properties.py`
- Create: `tests/unit/test_package_reader.py`
- Create: `tests/unit/test_properties.py`

**Interfaces:**
- Produces: `DocxPackage.open(path)`, `read_xml(part_name)`, `read_bytes(part_name)`, `iter_parts(prefix)`, `relationships(part_name)`.
- Produces: `read_properties(package) -> DocumentProperties` and `select_preservable_properties(properties, explicit_author=False)`.

- [ ] **Step 1: Write a minimal in-test DOCX package fixture**

```python
# tests/unit/test_package_reader.py
from pathlib import Path
from zipfile import ZipFile
from word_replica.opc.package_reader import DocxPackage


def make_package(path: Path) -> None:
    with ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
        z.writestr("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>")
        z.writestr("word/_rels/document.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='image' Target='media/image1.png'/></Relationships>")
        z.writestr("word/media/image1.png", b"png")


def test_reader_lists_parts_and_relationships(tmp_path: Path):
    path = tmp_path / "a.docx"
    make_package(path)
    with DocxPackage.open(path) as package:
        assert "word/document.xml" in package.parts
        assert package.read_bytes("word/media/image1.png") == b"png"
        assert package.relationships("word/document.xml")["rId1"].target == "media/image1.png"
```

- [ ] **Step 2: Write metadata allowlist test**

```python
# tests/unit/test_properties.py
from word_replica.opc.properties import DocumentProperties, select_preservable_properties


def test_preserve_policy_excludes_lifecycle_history():
    props = DocumentProperties(title="Paper", subject="AI", creator="Alice", created="2025-01-01", modified="2025-02-01", revision="42", total_editing_time="999")
    selected = select_preservable_properties(props, explicit_author=False)
    assert selected == {"title": "Paper", "subject": "AI"}
```

- [ ] **Step 3: Run tests and verify failure**

Run: `python -m pytest tests/unit/test_package_reader.py tests/unit/test_properties.py -v`

Expected: FAIL because OPC modules do not exist.

- [ ] **Step 4: Implement reader and metadata policy**

```python
# src/word_replica/opc/package_reader.py
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile
from lxml import etree
from word_replica.domain.errors import PackageReadError

REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}


@dataclass(frozen=True, slots=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    target_mode: str | None


class DocxPackage:
    def __init__(self, path: Path, archive: ZipFile) -> None:
        self.path = path
        self.archive = archive
        self.parts = frozenset(archive.namelist())

    @classmethod
    def open(cls, path: Path) -> "DocxPackage":
        try:
            return cls(path, ZipFile(path, "r"))
        except (BadZipFile, OSError) as exc:
            raise PackageReadError(str(exc)) from exc

    def __enter__(self): return self
    def __exit__(self, *args): self.archive.close()
    def read_bytes(self, part_name: str) -> bytes: return self.archive.read(part_name)
    def read_xml(self, part_name: str): return etree.fromstring(self.read_bytes(part_name))
    def iter_parts(self, prefix: str): return sorted(p for p in self.parts if p.startswith(prefix))

    def relationships(self, source_part: str) -> dict[str, Relationship]:
        source = Path(source_part)
        rel_part = str(source.parent / "_rels" / f"{source.name}.rels").replace("\\", "/")
        if rel_part not in self.parts:
            return {}
        root = self.read_xml(rel_part)
        result = {}
        for node in root.xpath("//pr:Relationship", namespaces=REL_NS):
            rel = Relationship(node.get("Id"), node.get("Type"), node.get("Target"), node.get("TargetMode"))
            result[rel.rel_id] = rel
        return result
```

```python
# src/word_replica/opc/properties.py
from dataclasses import dataclass


@dataclass(slots=True)
class DocumentProperties:
    title: str | None = None
    subject: str | None = None
    keywords: str | None = None
    category: str | None = None
    language: str | None = None
    creator: str | None = None
    company: str | None = None
    created: str | None = None
    modified: str | None = None
    revision: str | None = None
    total_editing_time: str | None = None
    custom: dict[str, str] | None = None


def select_preservable_properties(
    props: DocumentProperties,
    explicit_author: bool = False,
    custom_allowlist: tuple[str, ...] = (),
) -> dict[str, str]:
    names = ("title", "subject", "keywords", "category", "language")
    selected = {name: value for name in names if (value := getattr(props, name))}
    if explicit_author and props.creator:
        selected["creator"] = props.creator
    if explicit_author and props.company:
        selected["company"] = props.company
    for key in custom_allowlist:
        if props.custom and key in props.custom:
            selected[f"custom:{key}"] = props.custom[key]
    return selected

def build_output_metadata(
    props: DocumentProperties,
    mode: MetadataMode,
    preserve_author_fields: bool,
    custom_allowlist: tuple[str, ...],
) -> dict[str, str]:
    if mode is MetadataMode.FRESH:
        return {}
    return select_preservable_properties(props, preserve_author_fields, custom_allowlist)
```

- [ ] **Step 5: Extend `read_properties` for core/app/custom property parts and commit**

```python
CP = {"cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "dc": "http://purl.org/dc/elements/1.1/", "dcterms": "http://purl.org/dc/terms/"}
EP = {"ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}
CUST = {"c": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"}

def _text(root, xpath: str, ns: dict[str, str]) -> str | None:
    nodes = root.xpath(xpath, namespaces=ns)
    return nodes[0].text if nodes and nodes[0].text is not None else None

def read_properties(package: DocxPackage) -> DocumentProperties:
    props = DocumentProperties(custom={})
    if "docProps/core.xml" in package.parts:
        root = package.read_xml("docProps/core.xml")
        props.title = _text(root, "//dc:title", CP)
        props.subject = _text(root, "//dc:subject", CP)
        props.keywords = _text(root, "//cp:keywords", CP)
        props.category = _text(root, "//cp:category", CP)
        props.language = _text(root, "//dc:language", CP)
        props.creator = _text(root, "//dc:creator", CP)
        props.created = _text(root, "//dcterms:created", CP)
        props.modified = _text(root, "//dcterms:modified", CP)
        props.revision = _text(root, "//cp:revision", CP)
    if "docProps/app.xml" in package.parts:
        root = package.read_xml("docProps/app.xml")
        props.company = _text(root, "//ep:Company", EP)
        props.total_editing_time = _text(root, "//ep:TotalTime", EP)
    if "docProps/custom.xml" in package.parts:
        root = package.read_xml("docProps/custom.xml")
        for node in root.xpath("//c:property", namespaces=CUST):
            child = next(iter(node), None)
            if child is not None and child.text is not None:
                props.custom[node.get("name")] = child.text
    return props
```

Add a test with `custom={"DatasetVersion": "v13"}` proving it is copied only when `custom_allowlist=("DatasetVersion",)` is supplied.

Run: `python -m pytest tests/unit/test_package_reader.py tests/unit/test_properties.py -v`

Expected: PASS.

```bash
git add src/word_replica/opc tests/unit/test_package_reader.py tests/unit/test_properties.py
git commit -m "feat: inspect docx packages and enforce metadata allowlist"
```

---

### Task 5: Canonical Document Model and Deterministic Element IDs

**Files:**
- Create: `src/word_replica/domain/model.py`
- Create: `tests/unit/test_model.py`

**Interfaces:**
- Produces immutable-ish canonical element dataclasses and `ElementIdFactory`.
- Renderers and QA consume `DocumentModel` only; they must not depend on parser internals.

- [ ] **Step 1: Write deterministic-ID/model tests**

```python
# tests/unit/test_model.py
from word_replica.domain.model import DocumentModel, ElementIdFactory, Paragraph, Run


def test_element_ids_are_deterministic_for_source_positions():
    ids = ElementIdFactory("abc123")
    assert ids.make("paragraph", "body/0") == ids.make("paragraph", "body/0")
    assert ids.make("paragraph", "body/0") != ids.make("paragraph", "body/1")


def test_document_plain_text_preserves_order():
    doc = DocumentModel(source_sha256="abc123", body=[Paragraph("p1", runs=[Run("r1", text="Hello")]), Paragraph("p2", runs=[Run("r2", text=" world")])])
    assert doc.plain_text() == "Hello\n world"
```

- [ ] **Step 2: Run test and verify failure**

Run: `python -m pytest tests/unit/test_model.py -v`

Expected: FAIL because model types do not exist.

- [ ] **Step 3: Implement canonical model**

```python
# src/word_replica/domain/model.py
from dataclasses import dataclass, field
from hashlib import sha256
from typing import Any


class ElementIdFactory:
    def __init__(self, source_sha256: str) -> None:
        self.source_sha256 = source_sha256

    def make(self, kind: str, source_position: str) -> str:
        raw = f"{self.source_sha256}|{kind}|{source_position}".encode()
        return f"{kind[:3]}_{sha256(raw).hexdigest()[:16]}"


@dataclass(slots=True)
class Run:
    element_id: str
    text: str = ""
    properties: dict[str, Any] = field(default_factory=dict)
    hidden: bool = False


@dataclass(slots=True)
class Paragraph:
    element_id: str
    runs: list[Run] = field(default_factory=list)
    style_id: str | None = None
    properties: dict[str, Any] = field(default_factory=dict)

    def text(self) -> str:
        return "".join(run.text for run in self.runs)


@dataclass(slots=True)
class TableCell:
    element_id: str
    blocks: list[Any] = field(default_factory=list)
    properties: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class TableRow:
    element_id: str
    cells: list[TableCell] = field(default_factory=list)


@dataclass(slots=True)
class Table:
    element_id: str
    rows: list[TableRow] = field(default_factory=list)
    properties: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class BinaryAsset:
    asset_id: str
    part_name: str
    content_type: str | None
    sha256: str
    bytes_data: bytes


@dataclass(slots=True)
class RelationshipRef:
    rel_id: str
    rel_type: str
    target: str
    external: bool = False


@dataclass(slots=True)
class Section:
    element_id: str
    properties: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class DocumentModel:
    source_sha256: str
    body: list[Any] = field(default_factory=list)
    sections: list[Section] = field(default_factory=list)
    styles_xml: bytes | None = None
    numbering_xml: bytes | None = None
    settings_xml: bytes | None = None
    theme_parts: dict[str, bytes] = field(default_factory=dict)
    relationships: dict[str, RelationshipRef] = field(default_factory=dict)
    assets: dict[str, BinaryAsset] = field(default_factory=dict)
    extras: dict[str, Any] = field(default_factory=dict)

    def plain_text(self) -> str:
        return "\n".join(block.text() for block in self.body if isinstance(block, Paragraph))
```

- [ ] **Step 4: Add serialization fingerprint test**

Add an explicit canonical scalar serializer and fingerprint; binary payloads are represented by their stored SHA-256 rather than raw bytes:

```python
# add to src/word_replica/domain/model.py
import json
from dataclasses import asdict, is_dataclass

def _stable(value):
    if is_dataclass(value):
        return {k: _stable(v) for k, v in asdict(value).items() if k != "bytes_data"}
    if isinstance(value, dict):
        return {str(k): _stable(value[k]) for k in sorted(value, key=str)}
    if isinstance(value, list):
        return [_stable(v) for v in value]
    if isinstance(value, bytes):
        return sha256(value).hexdigest()
    return value

def fingerprint(self) -> str:
    payload = json.dumps(_stable(self), sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return sha256(payload).hexdigest()

def iter_paragraphs(self):
    def walk(blocks):
        for block in blocks:
            if isinstance(block, Paragraph):
                yield block
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        yield from walk(cell.blocks)
    yield from walk(self.body)
```

```python
def test_equivalent_models_have_same_fingerprint():
    a = DocumentModel(source_sha256="x", body=[Paragraph("p", runs=[Run("r", text="A")])])
    b = DocumentModel(source_sha256="x", body=[Paragraph("p", runs=[Run("r", text="A")])])
    assert a.fingerprint() == b.fingerprint()
```

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_model.py -v`

Expected: PASS.

```bash
git add src/word_replica/domain/model.py tests/unit/test_model.py
git commit -m "feat: define canonical document model"
```

---

### Task 6: Core Parser — Paragraphs, Runs, Styles, Sections, Breaks

**Files:**
- Create: `src/word_replica/parser/text.py`
- Create: `src/word_replica/parser/parser.py`
- Create: `tests/fixtures/build_fixtures.py`
- Create: `tests/unit/test_parser_core.py`

**Interfaces:**
- Consumes: `DocxPackage`, `DocumentModel`, `ElementIdFactory`.
- Produces: `DocxParser.parse(path: Path) -> DocumentModel`.

- [ ] **Step 1: Create deterministic core fixture builder**

```python
# tests/fixtures/build_fixtures.py
from pathlib import Path
from docx import Document
from docx.enum.section import WD_ORIENT


def build_core_fixture(path: Path) -> Path:
    doc = Document()
    doc.add_heading("Heading One", level=1)
    p = doc.add_paragraph()
    p.add_run("Bold").bold = True
    p.add_run(" normal")
    section = doc.add_section()
    section.orientation = WD_ORIENT.LANDSCAPE
    doc.save(path)
    return path
```

- [ ] **Step 2: Write parser test**

```python
# tests/unit/test_parser_core.py
from word_replica.parser.parser import DocxParser
from tests.fixtures.build_fixtures import build_core_fixture


def test_parser_extracts_text_styles_and_sections(tmp_path):
    source = build_core_fixture(tmp_path / "core.docx")
    model = DocxParser().parse(source)
    assert model.body[0].style_id == "Heading1"
    assert model.body[0].text() == "Heading One"
    assert model.body[1].runs[0].properties["bold"] is True
    assert len(model.sections) == 2
```

- [ ] **Step 3: Run test and verify failure**

Run: `python -m pytest tests/unit/test_parser_core.py -v`

Expected: FAIL because parser does not exist.

- [ ] **Step 4: Implement text extraction helpers and core parser**

```python
# src/word_replica/parser/text.py
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def run_text(run_node) -> str:
    pieces: list[str] = []
    for child in run_node:
        local = child.tag.rsplit("}", 1)[-1]
        if local == "t": pieces.append(child.text or "")
        elif local == "tab": pieces.append("\t")
        elif local in {"br", "cr"}: pieces.append("\n")
    return "".join(pieces)


def bool_prop(parent, name: str) -> bool | None:
    node = parent.find(f"w:{name}", namespaces=NS) if parent is not None else None
    if node is None: return None
    value = node.get(f"{{{W_NS}}}val")
    return value not in {"0", "false", "off"}
```

Implement `DocxParser.parse()` to:
- validate `word/document.xml` exists;
- hash source;
- read styles/numbering/settings raw bytes where present;
- call `read_properties(package)` and store the typed result as `model.extras["source_properties"]`;
- walk body children in order;
- create `Paragraph` blocks from `w:p` and preserve run order;
- capture paragraph style, bold/italic/underline/hidden, spacing/indent/alignment properties used by L2 QA;
- capture every `w:sectPr` into ordered `Section` objects, including page size, margins, orientation, columns, header/footer distances, and numbering start where present.

- [ ] **Step 5: Add break/tab round-trip assertions and commit**

Extend fixture with a tab and page break. Assert parser emits `\t` and `\n` tokens plus paragraph-level `pageBreakBefore/keepNext` properties when present.

Run: `python -m pytest tests/unit/test_parser_core.py -v`

Expected: PASS.

```bash
git add src/word_replica/parser tests/fixtures/build_fixtures.py tests/unit/test_parser_core.py
git commit -m "feat: parse core word document structures"
```

---

### Task 7: Extended Parser — Tables, Numbering, Media, Headers/Footers, Footnotes/Endnotes

**Files:**
- Create: `src/word_replica/parser/tables.py`
- Create: `src/word_replica/parser/relationships.py`
- Modify: `src/word_replica/parser/parser.py`
- Modify: `src/word_replica/domain/model.py`
- Modify: `tests/fixtures/build_fixtures.py`
- Create: `tests/unit/test_parser_extended.py`

**Interfaces:**
- Extends `DocumentModel` with `headers`, `footers`, `footnotes`, `endnotes`, image/object references, and numbering refs.
- All binary assets are keyed by content hash and retain original package part names.

- [ ] **Step 1: Add table/media/header/footnote fixture**

Use a deterministic self-contained fixture builder. The low-level `patch_notes()` helper must add the note part, content type, and document relationship together so the package stays valid:

```python
from PIL import Image
from docx import Document
from docx.shared import Inches

def build_extended_fixture(path: Path) -> Path:
    png = path.with_suffix(".png")
    Image.new("RGB", (8, 8), "white").save(png)
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).merge(table.cell(0, 1)).text = "merged"
    doc.add_paragraph().add_run().add_picture(str(png), width=Inches(0.25))
    doc.sections[0].header.paragraphs[0].text = "Header"
    doc.sections[0].footer.paragraphs[0].text = "Footer"
    doc.save(path)
    patch_notes(path, footnote_text="Footnote text", endnote_text="Endnote text")
    png.unlink()
    return path

def patch_notes(path: Path, footnote_text: str, endnote_text: str) -> None:
    # Rewrite the ZIP into a sibling temp file, adding word/footnotes.xml,
    # word/endnotes.xml, matching [Content_Types] overrides, relationships,
    # and noteReference/endnoteReference runs in word/document.xml.
    rewrite_docx_with_notes(path, footnote_text, endnote_text)
```

`rewrite_docx_with_notes()` is implemented in the same fixture module with `zipfile` + `lxml`; it must never depend on Word or network resources.

- [ ] **Step 2: Write parser assertions**

```python
# tests/unit/test_parser_extended.py
from word_replica.domain.model import Table
from word_replica.parser.parser import DocxParser
from tests.fixtures.build_fixtures import build_extended_fixture


def test_parser_extracts_tables_assets_headers_and_notes(tmp_path):
    source = build_extended_fixture(tmp_path / "extended.docx")
    model = DocxParser().parse(source)
    table = next(block for block in model.body if isinstance(block, Table))
    assert len(table.rows) == 2
    assert table.rows[0].cells[0].properties["grid_span"] == 2
    assert len(model.assets) == 1
    assert model.extras["headers"]
    assert model.extras["footnotes"]
```

- [ ] **Step 3: Run test and verify failure**

Run: `python -m pytest tests/unit/test_parser_extended.py -v`

Expected: FAIL on missing extended structures.

- [ ] **Step 4: Implement extended parsers**

Use one recursive block parser and explicit helpers with these signatures:

```python
def parse_blocks(parent, ids: ElementIdFactory, source_path: str, package: DocxPackage) -> list[object]:
    blocks = []
    for index, child in enumerate(parent):
        local = child.tag.rsplit("}", 1)[-1]
        path = f"{source_path}/{index}"
        if local == "p":
            blocks.append(parse_paragraph(child, ids, path, package))
        elif local == "tbl":
            blocks.append(parse_table(child, ids, path, package))
    return blocks

def parse_table(node, ids, path, package) -> Table:
    rows: list[TableRow] = []
    for r_index, tr in enumerate(node.findall("w:tr", namespaces=NS)):
        cells: list[TableCell] = []
        for c_index, tc in enumerate(tr.findall("w:tc", namespaces=NS)):
            tc_pr = tc.find("w:tcPr", namespaces=NS)
            span = tc_pr.find("w:gridSpan", namespaces=NS) if tc_pr is not None else None
            merge = tc_pr.find("w:vMerge", namespaces=NS) if tc_pr is not None else None
            props = {
                "grid_span": int(span.get(f"{{{W_NS}}}val", "1")) if span is not None else 1,
                "v_merge": merge.get(f"{{{W_NS}}}val", "continue") if merge is not None else None,
            }
            cell_path = f"{path}/row/{r_index}/cell/{c_index}"
            cells.append(TableCell(ids.make("cell", cell_path), parse_blocks(tc, ids, cell_path, package), props))
        rows.append(TableRow(ids.make("row", f"{path}/row/{r_index}"), cells))
    return Table(ids.make("table", path), rows=rows, properties=parse_table_properties(node))

def extract_asset(package: DocxPackage, part_name: str, content_type: str | None) -> BinaryAsset:
    data = package.read_bytes(part_name)
    digest = sha256(data).hexdigest()
    return BinaryAsset(f"asset_{digest[:16]}", part_name, content_type, digest, data)
```

`parse_table_properties()` is a pure helper that extracts table width/layout/borders/shading. Header/footer and notes call `parse_blocks()` on their XML roots, and paragraph numbering stores both `numId` and `ilvl` in paragraph properties while retaining raw `numbering.xml`.

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_parser_core.py tests/unit/test_parser_extended.py -v`

Expected: PASS.

```bash
git add src/word_replica/domain/model.py src/word_replica/parser tests/fixtures/build_fixtures.py tests/unit/test_parser_extended.py
git commit -m "feat: parse tables media headers and notes"
```

---

### Task 8: Full-Fidelity Structures and Clean-vs-Full Projection

**Files:**
- Create: `src/word_replica/parser/fidelity.py`
- Modify: `src/word_replica/parser/parser.py`
- Modify: `src/word_replica/domain/model.py`
- Create: `tests/unit/test_fidelity_projection.py`

**Interfaces:**
- Produces: `project_fidelity(model: DocumentModel, mode: FidelityMode) -> DocumentModel`.
- Canonical source model retains full available evidence; projection determines what a renderer receives.

- [ ] **Step 1: Write clean/full behavior tests**

```python
# tests/unit/test_fidelity_projection.py
from word_replica.domain.enums import FidelityMode
from word_replica.parser.fidelity import project_fidelity
from tests.fixtures.build_fixtures import build_review_fixture
from word_replica.parser.parser import DocxParser


def test_clean_projection_omits_comments_and_deleted_revision_text(tmp_path):
    source = build_review_fixture(tmp_path / "review.docx")
    model = DocxParser().parse(source)
    clean = project_fidelity(model, FidelityMode.CLEAN)
    assert clean.extras["comments"] == {}
    assert "deleted text" not in clean.plain_text()
    assert "inserted text" in clean.plain_text()


def test_full_projection_retains_review_structures(tmp_path):
    source = build_review_fixture(tmp_path / "review.docx")
    full = project_fidelity(DocxParser().parse(source), FidelityMode.FULL)
    assert full.extras["comments"]
    assert full.extras["revisions"]
    assert full.extras["bookmarks"]
    assert full.extras["fields"]
```

- [ ] **Step 2: Run test and verify failure**

Run: `python -m pytest tests/unit/test_fidelity_projection.py -v`

Expected: FAIL because review structures are not modeled.

- [ ] **Step 3: Extend model/parser for review structures**

Add concrete model types and store them under typed `DocumentModel` collections rather than unstructured strings:

```python
@dataclass(slots=True)
class Comment:
    comment_id: str
    author: str | None
    date: str | None
    blocks: list[object]

@dataclass(slots=True)
class Bookmark:
    bookmark_id: str
    name: str
    start_path: str
    end_path: str | None

@dataclass(slots=True)
class Field:
    field_id: str
    instruction: str
    result_text: str
    locked: bool = False

@dataclass(slots=True)
class RevisionSpan:
    revision_id: str
    kind: str  # insert|delete|move_from|move_to
    author: str | None
    date: str | None
    path: str
    text: str

@dataclass(slots=True)
class PreservedPart:
    part_name: str
    content_type: str | None
    relationship_type: str | None
    sha256: str
    data: bytes
```

Parser rules are exact: `w:ins` contributes visible text and a revision record; `w:del` contributes only revision/deleted-text evidence, not `Run.text`; comments are read from `word/comments.xml`; bookmarks pair start/end IDs; field instruction text is concatenated across `w:instrText`; chart/OLE/embedded package parts become `PreservedPart` entries keyed by part name.

- [ ] **Step 4: Implement profile projection rules**

```python
# src/word_replica/parser/fidelity.py
from copy import deepcopy
from word_replica.domain.enums import FidelityMode
from word_replica.domain.model import DocumentModel


def project_fidelity(model: DocumentModel, mode: FidelityMode) -> DocumentModel:
    projected = deepcopy(model)
    if mode is FidelityMode.FULL:
        return projected
    projected.extras["comments"] = {}
    projected.extras["revisions"] = [r for r in projected.extras.get("revisions", []) if r.kind == "insert"]
    projected.extras["tracked_changes_enabled"] = False
    for paragraph in projected.iter_paragraphs():
        paragraph.runs[:] = [run for run in paragraph.runs if not run.hidden]
    # Parser stores visible-final text separately from deleted revision text, so clean mode never re-inserts deletion text.
    return projected
```

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_fidelity_projection.py -v`

Expected: PASS.

```bash
git add src/word_replica/domain/model.py src/word_replica/parser tests/unit/test_fidelity_projection.py
git commit -m "feat: model full fidelity review structures"
```

---

### Task 9: Renderer Contract and Pure-DOCX Reconstruction Engine

**Files:**
- Create: `src/word_replica/renderers/base.py`
- Create: `src/word_replica/renderers/pure_docx.py`
- Create: `tests/unit/test_pure_docx_renderer.py`
- Create: `tests/integration/test_pure_docx_e2e.py`

**Interfaces:**
- Produces `Renderer` protocol with `render(model, output_path, context) -> RenderResult`, `save(output_path)`, and `set_custom_property(name, value)` semantics.
- When a `RenderContext` is supplied, every renderer must call `context.final_seal(self, output_path)` exactly once as its last real save **before** disposing its mutable Word/package state.
- Produces `PureDocxRenderer` that can rebuild core supported structures without Word.

- [ ] **Step 1: Define renderer contract and failing smoke test**

```python
# src/word_replica/renderers/base.py
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol
from word_replica.domain.results import WarningItem


@dataclass(slots=True)
class RenderResult:
    output_path: Path
    warnings: list[WarningItem] = field(default_factory=list)
    stages_completed: list[str] = field(default_factory=list)


class Renderer(Protocol):
    def render(self, model, output_path: Path, context) -> RenderResult: ...
    def save(self, output_path: Path) -> None: ...
    def set_custom_property(self, name: str, value: str) -> None: ...
```

```python
# tests/unit/test_pure_docx_renderer.py
from zipfile import ZipFile
from word_replica.domain.model import DocumentModel, Paragraph, Run
from word_replica.renderers.pure_docx import PureDocxRenderer


def test_renderer_creates_new_docx_with_document_xml(tmp_path):
    model = DocumentModel(source_sha256="abc", body=[Paragraph("p1", runs=[Run("r1", text="Hello")])])
    output = tmp_path / "out.docx"
    PureDocxRenderer().render(model, output, context=None)
    with ZipFile(output) as z:
        assert b"Hello" in z.read("word/document.xml")
```

- [ ] **Step 2: Run test and verify failure**

Run: `python -m pytest tests/unit/test_pure_docx_renderer.py -v`

Expected: FAIL because renderer does not exist.

- [ ] **Step 3: Implement pure renderer using a fresh DOCX shell**

Implement the renderer around a temporary fresh package; all text nodes are created by `lxml` so escaping is automatic:

```python
class PureDocxRenderer:
    def __init__(self):
        self._package = None

    def render(self, model: DocumentModel, output_path: Path, context) -> RenderResult:
        shell = output_path.with_suffix(".shell.docx")
        Document().save(shell)
        self._package = MutableDocxPackage.from_file(shell)
        stages = [
            ("body", lambda: self._package.set_document_body(build_body_xml(model.body, model.sections))),
            ("styles", lambda: self._package.set_styles(model.styles_xml)),
            ("numbering", lambda: self._package.set_numbering(model.numbering_xml)),
            ("settings", lambda: self._package.set_settings(project_settings(model.settings_xml))),
            ("assets", lambda: install_assets_and_relationships(self._package, model)),
            ("notes_headers", lambda: install_headers_footers_notes(self._package, model)),
            ("metadata", lambda: apply_output_properties(self._package, model.extras["metadata_policy"])),
        ]
        for index, (stage, mutate) in enumerate(stages, start=1):
            mutate()
            if context is not None:
                context.mark_content_changed(f"{model.fingerprint()}:{index}:{stage}")
                context.checkpoint(f"{stage} complete", stage, lambda: self.save(output_path), output_path)
        if context is None:
            self.save(output_path)
        else:
            context.final_seal(self, output_path)
        shell.unlink(missing_ok=True)
        return RenderResult(output_path=output_path, warnings=self._package.warnings, stages_completed=[s for s, _ in stages])

    def save(self, output_path: Path) -> None:
        if self._package is None:
            raise RuntimeError("PureDocxRenderer has no active package")
        self._package.write_atomic(output_path)

    def set_custom_property(self, name: str, value: str) -> None:
        if self._package is None:
            raise RuntimeError("PureDocxRenderer has no active package")
        self._package.set_custom_property(name, value)
```

`MutableDocxPackage`, `build_body_xml`, `install_assets_and_relationships`, and property writers are implemented in this task as focused private helpers; they must generate new relationship IDs and must not copy core/app lifecycle values from the source.

- [ ] **Step 4: Add end-to-end parse → render → parse test**

```python
# tests/integration/test_pure_docx_e2e.py
from word_replica.parser.parser import DocxParser
from word_replica.renderers.pure_docx import PureDocxRenderer
from tests.fixtures.build_fixtures import build_extended_fixture


def test_pure_docx_round_trip_preserves_core_content(tmp_path):
    source = build_extended_fixture(tmp_path / "source.docx")
    source_model = DocxParser().parse(source)
    output = tmp_path / "out.docx"
    PureDocxRenderer().render(source_model, output, context=None)
    rebuilt = DocxParser().parse(output)
    assert rebuilt.plain_text() == source_model.plain_text()
    assert len(rebuilt.sections) == len(source_model.sections)
```

- [ ] **Step 5: Add warning behavior for unsupported transfer parts and commit**

If a `PreservedPart` cannot be safely related into the output package, return `WarningItem(code="UNSUPPORTED_TRANSFER_PART", message=f"Could not transfer {part.part_name}")` rather than silently dropping it.

Run: `python -m pytest tests/unit/test_pure_docx_renderer.py tests/integration/test_pure_docx_e2e.py -v`

Expected: PASS.

```bash
git add src/word_replica/renderers tests/unit/test_pure_docx_renderer.py tests/integration/test_pure_docx_e2e.py
git commit -m "feat: add pure docx reconstruction renderer"
```

---

### Task 10: Microsoft Word Capability Detection and COM Session Lifecycle

**Files:**
- Create: `src/word_replica/renderers/word_com.py`
- Create: `tests/unit/test_word_com_selection.py`
- Create: `tests/integration/word/test_word_session.py`

**Interfaces:**
- Produces: `word_available() -> bool`, `WordSession`, `WordComRenderer`.
- `WordSession` owns COM initialization, Word process lifetime, visibility, alerts, and cleanup.

- [ ] **Step 1: Write platform-independent renderer selection test**

```python
# tests/unit/test_word_com_selection.py
from word_replica.domain.enums import RendererChoice
from word_replica.renderers.word_com import choose_renderer_name


def test_auto_falls_back_to_docx_when_word_unavailable():
    assert choose_renderer_name(RendererChoice.AUTO, word_is_available=False) == "docx"


def test_explicit_word_requires_word():
    import pytest
    from word_replica.domain.errors import RendererUnavailableError
    with pytest.raises(RendererUnavailableError):
        choose_renderer_name(RendererChoice.WORD, word_is_available=False)
```

- [ ] **Step 2: Implement capability selection without importing COM on non-Windows**

```python
# src/word_replica/renderers/word_com.py
import sys
from word_replica.domain.enums import RendererChoice
from word_replica.domain.errors import RendererUnavailableError


def choose_renderer_name(choice: RendererChoice, word_is_available: bool) -> str:
    if choice is RendererChoice.DOCX: return "docx"
    if choice is RendererChoice.WORD and not word_is_available:
        raise RendererUnavailableError("Microsoft Word desktop automation is unavailable")
    if choice is RendererChoice.WORD: return "word"
    return "word" if word_is_available else "docx"


def word_available() -> bool:
    if sys.platform != "win32": return False
    try:
        import win32com.client
        app = win32com.client.DispatchEx("Word.Application")
        app.Quit()
        return True
    except Exception:
        return False
```

- [ ] **Step 3: Add gated Word integration test**

```python
# tests/integration/word/test_word_session.py
import os
import pytest
pytestmark = pytest.mark.word

@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="set WORD_REPLICA_WORD_TESTS=1 on Windows with Word")
def test_word_session_creates_and_closes_document(tmp_path):
    from word_replica.renderers.word_com import WordSession
    with WordSession(visible=False) as session:
        doc = session.new_document()
        path = tmp_path / "smoke.docx"
        doc.SaveAs2(str(path))
        doc.Close(False)
    assert path.exists()
```

- [ ] **Step 4: Implement `WordSession` lifecycle**

```python
class WordSession:
    def __init__(self, visible: bool) -> None:
        self.visible = visible
        self.app = None
        self._owned_docs = []

    def __enter__(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = self.visible
        self.app.DisplayAlerts = 0
        return self

    def new_document(self):
        doc = self.app.Documents.Add()
        self._owned_docs.append(doc)
        return doc

    def release_document(self, doc) -> None:
        if doc in self._owned_docs:
            self._owned_docs.remove(doc)

    def __exit__(self, exc_type, exc, tb):
        try:
            for doc in reversed(self._owned_docs):
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if self.app is not None:
                self.app.Quit()
        finally:
            self._owned_docs.clear()
            self.app = None
            self._pythoncom.CoUninitialize()
```

The production version must remove a document from `_owned_docs` when the renderer explicitly closes it, preventing a second close attempt.

- [ ] **Step 5: Run unit tests, optionally integration test, and commit**

Run: `python -m pytest tests/unit/test_word_com_selection.py -v`

On Windows + Word: `set WORD_REPLICA_WORD_TESTS=1 && python -m pytest tests/integration/word/test_word_session.py -v`

```bash
git add src/word_replica/renderers/word_com.py tests/unit/test_word_com_selection.py tests/integration/word/test_word_session.py
git commit -m "feat: add safe microsoft word com session"
```

---

### Task 11: Word COM Renderer — Core and Full-Fidelity Reconstruction

**Files:**
- Modify: `src/word_replica/renderers/word_com.py`
- Create: `tests/integration/word/test_word_com_renderer.py`
- Create: `scripts/run_word_integration.ps1`

**Interfaces:**
- Consumes canonical `DocumentModel` projected for clean/full mode.
- Produces `.docx` through real Microsoft Word saves and supports visible/background mode.

- [ ] **Step 1: Write gated core reconstruction integration test**

```python
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word integration disabled")
def test_word_renderer_rebuilds_core_fixture(tmp_path):
    source = build_core_fixture(tmp_path / "source.docx")
    model = DocxParser().parse(source)
    output = tmp_path / "rebuilt.docx"
    renderer = WordComRenderer(visible=False)
    renderer.render(model, output, context=None)
    rebuilt = DocxParser().parse(output)
    assert rebuilt.plain_text() == model.plain_text()
    assert len(rebuilt.sections) == len(model.sections)
```

- [ ] **Step 2: Implement deterministic Word reconstruction order**

The renderer must create a **new blank document** and rebuild in this order:
1. base page/section settings;
2. styles needed by body blocks;
3. ordered body paragraphs and tables;
4. images/shapes and their anchor/inline mode where supported;
5. headers/footers and page numbering;
6. footnotes/endnotes;
7. bookmarks and fields;
8. comments/revisions only in Full Fidelity when reconstructable;
9. TOC/field update;
10. final pagination/repagination before save.

Use these exact method names and responsibilities:

```python
class WordComRenderer:
    def __init__(self, visible: bool):
        self.visible = visible
        self._warnings: list[WarningItem] = []
        self._stages: list[str] = []

    def _insert_block(self, doc, block) -> None:
        if isinstance(block, Paragraph):
            self._insert_paragraph(doc, block)
        elif isinstance(block, Table):
            self._insert_table(doc, block)
        else:
            self._warnings.append(WarningItem("UNSUPPORTED_BLOCK", f"Unsupported block type: {type(block).__name__}"))

    def save(self, output_path: Path) -> None:
        if self._document is None:
            raise RuntimeError("No active Word document")
        self._document.SaveAs2(str(output_path))

    def set_custom_property(self, name: str, value: str) -> None:
        props = self._document.CustomDocumentProperties
        try:
            props(name).Value = value
        except Exception:
            # msoPropertyTypeString = 4
            props.Add(name, False, 4, value)
```

The same class also defines `_apply_sections`, `_insert_paragraph`, `_insert_table`, `_insert_assets`, `_restore_headers_footers`, `_restore_notes`, and `_restore_fields_bookmarks`; each maps only canonical-model data into Word COM objects and appends a typed warning instead of silently dropping unsupported properties.

- [ ] **Step 3: Implement visible/background behavior**

Use a single code path; only the Word application visibility flag differs:

```python
def render(self, model: DocumentModel, output_path: Path, context) -> RenderResult:
    with WordSession(visible=self.visible) as session:
        doc = session.new_document()
        self._document = doc
        self._apply_sections(doc, model)
        self._apply_metadata(doc, model.extras["metadata_policy"])
        for block_index, block in enumerate(model.body, start=1):
            self._insert_block(doc, block)
            if context is not None:
                context.mark_content_changed(f"{model.fingerprint()}:body:{block_index}")
                context.maybe_periodic_save("body", lambda: doc.SaveAs2(str(output_path)), output_path)
        self._restore_headers_footers(doc, model)
        self._restore_notes(doc, model)
        self._restore_fields_bookmarks(doc, model)
        doc.Fields.Update()
        doc.Repaginate()
        if context is None:
            doc.SaveAs2(str(output_path))
        else:
            context.mark_content_changed(f"{model.fingerprint()}:finalize")
            context.checkpoint("fields and pagination complete", "finalize", lambda: doc.SaveAs2(str(output_path)), output_path)
            context.final_seal(self, output_path)
        session.release_document(doc)
        doc.Close(False)
        self._document = None
    return RenderResult(output_path, warnings=self._warnings, stages_completed=self._stages)
```

No sleep/random-delay/keyboard API is imported or called anywhere in the renderer.

- [ ] **Step 4: Add Full Fidelity warnings and embedded-object transfer rule**

Make transfer decisions explicit and testable:

```python
def _handle_preserved_part(self, part: PreservedPart) -> None:
    if part.relationship_type in SAFE_PACKAGE_TRANSFER_REL_TYPES:
        self._post_save_transfers.append(part)
        return
    code = "OLE_TRANSFER_UNSUPPORTED" if "oleObject" in (part.relationship_type or "") else "EMBEDDED_PART_UNSUPPORTED"
    self._warnings.append(WarningItem(code=code, message=f"Could not safely transfer {part.part_name}"))
```

After Word's final `SaveAs2`, post-save transfer must occur only through a validator that reopens the DOCX ZIP, verifies all relationship targets exist, and then opens the package through `DocxPackage`; invalid transfer is rolled back from a backup and remains a warning.

- [ ] **Step 5: Add PowerShell runner and commit**

```powershell
# scripts/run_word_integration.ps1
$env:WORD_REPLICA_WORD_TESTS = "1"
python -m pytest tests/integration/word -m word -v
exit $LASTEXITCODE
```

Run on Windows + Word: `powershell -ExecutionPolicy Bypass -File scripts/run_word_integration.ps1`

```bash
git add src/word_replica/renderers/word_com.py tests/integration/word/test_word_com_renderer.py scripts/run_word_integration.ps1
git commit -m "feat: reconstruct documents through word com"
```

---

### Task 12: Reconstruction Orchestrator, Error Model, and Real Save Scheduling

**Files:**
- Create: `src/word_replica/services/rebuild.py`
- Modify: `src/word_replica/services/project_store.py`
- Modify: `src/word_replica/services/checkpoints.py`
- Create: `tests/unit/test_rebuild_service.py`

**Interfaces:**
- Produces: `RebuildService.rebuild(source: Path, options: RebuildOptions) -> RunResult`.
- This is the single entry point used by CLI and GUI.

- [ ] **Step 1: Write orchestration test with fake renderer**

```python
# tests/unit/test_rebuild_service.py
from pathlib import Path
from word_replica.config import RebuildOptions
from word_replica.domain.enums import RunStatus
from word_replica.renderers.base import RenderResult
from word_replica.services.rebuild import RebuildService


class FakeRenderer:
    def __init__(self):
        self.properties = {}
        self.payload = b"fake-docx"

    def set_custom_property(self, name: str, value: str) -> None:
        self.properties[name] = value

    def save(self, output_path: Path) -> None:
        output_path.write_bytes(self.payload + repr(sorted(self.properties.items())).encode())

    def render(self, model, output_path: Path, context):
        context.mark_content_changed("fake:1")
        context.checkpoint("document shell created", "shell", lambda: self.save(output_path), output_path)
        context.final_seal(self, output_path)
        return RenderResult(output_path)


def test_rebuild_protects_source_and_returns_status(tmp_path, monkeypatch):
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")
    service = RebuildService.for_testing(tmp_path / "app", parser=lambda _: object(), renderer=FakeRenderer(), qa=lambda *_: RunStatus.PASS)
    result = service.rebuild(source, RebuildOptions())
    assert result.status is RunStatus.PASS
    assert source.read_bytes() == b"fixture"


def test_overwrite_mode_requires_verified_backup(tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")
    service = RebuildService.for_testing(tmp_path / "app", parser=lambda _: object(), renderer=FakeRenderer(), qa=lambda *_: RunStatus.PASS)
    result = service.rebuild(source, RebuildOptions(allow_source_overwrite=True))
    backups = list((source.parent / "source_rebuild").rglob("source_before_overwrite.docx"))
    assert result.status is RunStatus.PASS
    assert len(backups) == 1
    assert backups[0].read_bytes() == b"fixture"
    assert source.read_bytes().startswith(b"fake-docx")
```

- [ ] **Step 2: Run test and verify failure**

Run: `python -m pytest tests/unit/test_rebuild_service.py -v`

Expected: FAIL because service does not exist.

- [ ] **Step 3: Implement orchestration phases**

`RebuildService.rebuild()` must execute and audit these exact phases:

```text
PROJECT_CREATED
SOURCE_HASHED
PARSE_STARTED
PARSE_COMPLETED
FIDELITY_PROJECTED
RENDERER_SELECTED
REBUILD_STARTED
<real checkpoint events>
REBUILD_COMPLETED
SOURCE_RECHECKED | SOURCE_INTENTIONAL_REPLACEMENT_VERIFIED
QA_STARTED
QA_COMPLETED
RUN_COMPLETED
```

Any `CriticalRebuildError` produces `RunStatus.FAIL`, records the reason, and prevents QA claims of success. Non-critical `WarningItem` values flow into `warnings.json` and can result in `WARN`. Before renderer selection, build `model.extras["metadata_policy"]` from the source properties and `RebuildOptions`; Fresh mode returns no source descriptive fields, while Preserve mode uses only the approved allowlist.

```python
def _prepare_metadata_policy(model: DocumentModel, options: RebuildOptions) -> None:
    source_props = model.extras.get("source_properties", DocumentProperties())
    model.extras["metadata_policy"] = build_output_metadata(
        source_props,
        options.metadata,
        options.preserve_author_fields,
        options.custom_metadata_allowlist,
    )

def _resolve_output_path(source: Path, paths: ProjectPaths, options: RebuildOptions) -> Path:
    if not options.allow_source_overwrite:
        return paths.output_dir / f"{source.stem}_reconstructed.docx"
    backup = paths.backups_dir / f"{source.stem}_before_overwrite.docx"
    shutil.copy2(source, backup)
    if sha256_file(backup) != sha256_file(source):
        raise SourceIntegrityError("Verified backup could not be created before source overwrite")
    return source

def _final_source_check(snapshot: SourceSnapshot, output_path: Path, options: RebuildOptions, paths: ProjectPaths) -> None:
    if not options.allow_source_overwrite:
        assert_source_unchanged(snapshot)
        return
    backup = paths.backups_dir / f"{snapshot.path.stem}_before_overwrite.docx"
    if sha256_file(backup) != snapshot.sha256:
        raise SourceIntegrityError("Original backup no longer matches the pre-run source hash")
```

The final output save is a truthful metadata seal owned by the shared render context:

```python
class RenderContext:
    def final_seal(self, renderer: Renderer, output_path: Path) -> SaveEvent:
        if self.final_sealed:
            raise CriticalRebuildError("Renderer attempted more than one final metadata seal")
        next_sequence = self.checkpoints.next_sequence
        renderer.set_custom_property("WordReplicaProjectId", self.project_id)
        renderer.set_custom_property("WordReplicaReconstructed", "true")
        renderer.set_custom_property("WordReplicaActualSaveCount", str(next_sequence))
        event = self.checkpoint("final metadata/save-count seal", "final", lambda: renderer.save(output_path), output_path)
        self.final_sealed = True
        return event
```

No renderer may save after `context.final_seal()`. The seal must happen before Word/package state is closed. The acceptance test compares this custom property with the number of rows in `save_history.jsonl`. Word's built-in revision number and Total Editing Time are left to Word/default package semantics and are never presented as the authoritative save count.

`warnings.json` is written on every completed/failed run as a JSON array, including `[]` when there are no warnings. When `allow_source_overwrite=True`, audit `SOURCE_OVERWRITE_ENABLED` and `SOURCE_BACKUP_VERIFIED`; do not emit `SOURCE_RECHECKED` as if the original were unchanged—instead emit `SOURCE_INTENTIONAL_REPLACEMENT_VERIFIED`.

- [ ] **Step 4: Add periodic save scheduler without meaningless save loops**

Implement a monotonic-clock guard in the render context:

```python
class RenderContext:
    def __init__(self, project_id: str, options: RebuildOptions, checkpoints: CheckpointManager) -> None:
        self.project_id = project_id
        self.options = options
        self.checkpoints = checkpoints
        self.last_save_monotonic = time.monotonic()
        self.current_content_fingerprint: str | None = None
        self.last_saved_content_fingerprint: str | None = None
        self.final_sealed = False

    def mark_content_changed(self, fingerprint: str) -> None:
        self.current_content_fingerprint = fingerprint

    def checkpoint(self, reason: str, stage: str, save_callable, path: Path) -> SaveEvent:
        if self.final_sealed:
            raise CriticalRebuildError("No save is permitted after final metadata seal")
        event = self.checkpoints.save(reason, stage, save_callable, path)
        self.last_save_monotonic = time.monotonic()
        self.last_saved_content_fingerprint = self.current_content_fingerprint
        return event

    def maybe_periodic_save(self, stage: str, save_callable, path: Path) -> None:
        now = time.monotonic()
        if now - self.last_save_monotonic < self.options.periodic_save_seconds:
            return
        if self.current_content_fingerprint == self.last_saved_content_fingerprint:
            return
        self.checkpoint("periodic save", stage, save_callable, path)
```

Add `final_seal()` from Step 3 to this class; `RebuildService` fails the run if a renderer returns without setting `context.final_sealed=True`.

The renderer must call `mark_content_changed(fingerprint)` after actual document mutations. This prevents meaningless saves that inflate counts.

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_rebuild_service.py tests/unit/test_checkpoints.py tests/unit/test_source_guard.py -v`

Expected: PASS.

```bash
git add src/word_replica/services tests/unit/test_rebuild_service.py
git commit -m "feat: orchestrate protected rebuild runs"
```

---

### Task 13: QA L0-L3 and PASS/WARN/FAIL Policy

**Files:**
- Create: `src/word_replica/qa/content.py`
- Create: `src/word_replica/qa/structure.py`
- Create: `src/word_replica/qa/formatting.py`
- Create: `src/word_replica/qa/layout.py`
- Create: `src/word_replica/qa/policy.py`
- Create: `tests/unit/test_qa_levels.py`
- Create: `tests/unit/test_qa_policy.py`

**Interfaces:**
- Produces `QaFinding`, `QaLevelResult`, `QaBundle`, `run_l0_l3(source_model, rebuilt_model)`, `classify_run(bundle, warnings) -> RunStatus`.

- [ ] **Step 1: Write exact L0 and fail-policy tests**

Define the shared QA result types first:

```python
@dataclass(slots=True)
class QaFinding:
    code: str
    path: str
    expected: object
    actual: object
    severity: str

@dataclass(slots=True)
class QaLevelResult:
    level: str
    passed: bool
    findings: list[QaFinding | str]
    has_required_loss: bool = False

@dataclass(slots=True)
class QaBundle:
    levels: dict[str, QaLevelResult]
    render: RenderQaResult | None = None
```

```python
# tests/unit/test_qa_policy.py
from word_replica.domain.enums import RunStatus
from word_replica.domain.results import WarningItem
from word_replica.qa.policy import QaBundle, QaLevelResult, classify_run


def test_content_loss_is_fail():
    bundle = QaBundle(levels={"L0": QaLevelResult("L0", passed=False, findings=["text mismatch"])})
    assert classify_run(bundle, warnings=[]) is RunStatus.FAIL


def test_noncritical_visual_warning_is_warn():
    bundle = QaBundle(levels={"L0": QaLevelResult("L0", True, []), "L1": QaLevelResult("L1", True, []), "L2": QaLevelResult("L2", True, []), "L3": QaLevelResult("L3", True, [])})
    assert classify_run(bundle, warnings=[WarningItem("VISUAL_DIFF", "minor mismatch")]) is RunStatus.WARN
```

- [ ] **Step 2: Implement canonical comparison projections**

```python
def l0_projection(model: DocumentModel) -> list[tuple[str, str]]:
    rows = [(f"body/{i}", block.text()) for i, block in enumerate(model.body) if isinstance(block, Paragraph)]
    rows.extend((f"footnote/{k}", note.text()) for k, note in sorted(model.footnotes.items()))
    return rows

def l1_projection(model: DocumentModel) -> dict:
    return {
        "body_kinds": [type(b).__name__ for b in model.body],
        "tables": [table_shape(b) for b in model.body if isinstance(b, Table)],
        "asset_hashes": sorted(a.sha256 for a in model.assets.values()),
        "footnotes": len(model.footnotes),
        "endnotes": len(model.endnotes),
        "bookmarks": len(model.bookmarks),
        "fields": len(model.fields),
    }

def l2_projection(model: DocumentModel) -> list[dict]:
    return [normalize_formatting(p) for p in model.iter_paragraphs()]

def l3_projection(model: DocumentModel) -> list[dict]:
    return [normalize_section_properties(section.properties) for section in model.sections]
```

Do not compare element IDs; `compare_projection(level, expected, actual)` recursively emits path-based `QaFinding` objects.

- [ ] **Step 3: Implement human-readable findings**

Every mismatch must have `code`, `path`, `expected`, `actual`, and `severity`. For text mismatches, `expected`/`actual` in logs/reports are capped to a 200-character excerpt plus a SHA-256 of the full value so operational artifacts do not duplicate whole document text. Example:

```python
QaFinding(code="RUN_BOLD_MISMATCH", path="body/3/run/1", expected=True, actual=False, severity="error")
```

- [ ] **Step 4: Implement classification policy**

```python
def classify_run(bundle: QaBundle, warnings: list[WarningItem], critical_reasons: list[str] | None = None) -> RunStatus:
    if critical_reasons:
        return RunStatus.FAIL
    if not bundle.levels["L0"].passed:
        return RunStatus.FAIL
    if not bundle.levels["L1"].passed and bundle.levels["L1"].has_required_loss:
        return RunStatus.FAIL
    if any(w.affects_status for w in warnings):
        return RunStatus.WARN
    for name in ("L2", "L3"):
        if name in bundle.levels and not bundle.levels[name].passed:
            return RunStatus.WARN
    if bundle.render is not None and not bundle.render.within_tolerance:
        return RunStatus.WARN
    return RunStatus.PASS
```

`L4_UNAVAILABLE` is recorded with `affects_status=False` when the selected environment legitimately lacks Word/PDF rendering; it is not used to fake a visual PASS.

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_qa_levels.py tests/unit/test_qa_policy.py -v`

Expected: PASS.

```bash
git add src/word_replica/qa tests/unit/test_qa_levels.py tests/unit/test_qa_policy.py
git commit -m "feat: compare reconstruction quality through l0 l3"
```

---

### Task 14: L4 PDF Render Diff and Final HTML QA Report

**Files:**
- Create: `src/word_replica/qa/render.py`
- Create: `src/word_replica/qa/report.py`
- Create: `tests/unit/test_render_diff.py`
- Create: `tests/unit/test_report.py`

**Interfaces:**
- Produces: `compare_page_images(source_png: Path, rebuilt_png: Path) -> VisualMetric`, `compare_pdfs(source_pdf: Path, rebuilt_pdf: Path, qa_dir: Path, dpi: int = 144) -> RenderQaResult`, and `write_qa_report(path: Path, **context) -> Path`.
- Word COM export is an optional capability; report must state when L4 was not verifiable.

- [ ] **Step 1: Write image-diff metric test**

```python
# tests/unit/test_render_diff.py
from PIL import Image
from word_replica.qa.render import compare_page_images


def test_identical_pages_have_zero_difference(tmp_path):
    a = tmp_path / "a.png"; b = tmp_path / "b.png"
    Image.new("RGB", (100, 100), "white").save(a)
    Image.new("RGB", (100, 100), "white").save(b)
    metric = compare_page_images(a, b)
    assert metric.changed_pixel_ratio == 0.0
    assert metric.mean_absolute_error == 0.0
```

- [ ] **Step 2: Implement PDF rasterization and page comparison**

```python
def rasterize_pdf(pdf_path: Path, out_dir: Path, dpi: int = 144) -> list[Path]:
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(pdf_path)
    scale = dpi / 72.0
    pages = []
    for index in range(len(pdf)):
        image = pdf[index].render(scale=scale).to_pil().convert("RGB")
        path = out_dir / f"page_{index + 1:04d}.png"
        image.save(path)
        pages.append(path)
    return pages

def compare_page_images(a: Path, b: Path) -> VisualMetric:
    ia, ib = Image.open(a).convert("RGB"), Image.open(b).convert("RGB")
    if ia.size != ib.size:
        return VisualMetric(False, 1.0, 255.0, ia.size, ib.size)
    diff = ImageChops.difference(ia, ib)
    hist = diff.histogram()
    pixels = ia.width * ia.height
    changed = sum(hist[i] for i in range(1, 256)) + sum(hist[i] for i in range(257, 512)) + sum(hist[i] for i in range(513, 768))
    changed_ratio = min(1.0, changed / max(1, pixels * 3))
    mae = sum((i % 256) * count for i, count in enumerate(hist)) / max(1, pixels * 3)
    return VisualMetric(True, changed_ratio, mae, ia.size, ib.size)
```

`compare_pdfs()` first checks page count, then compares matching page dimensions/metrics and writes optional diff PNGs.

- [ ] **Step 3: Write report content test**

```python
# tests/unit/test_report.py
from word_replica.qa.report import write_qa_report


def test_report_explains_status_and_l4_limitations(tmp_path):
    path = write_qa_report(tmp_path / "qa_report.html", status="WARN", levels={}, warnings=[{"code": "L4_UNAVAILABLE", "message": "Word not installed"}], render_result=None)
    html = path.read_text(encoding="utf-8")
    assert "WARN" in html
    assert "L4_UNAVAILABLE" in html
    assert "Word not installed" in html
```

- [ ] **Step 4: Implement a self-contained HTML report**

Use a literal embedded Jinja template, with autoescape enabled and inline CSS only:

```python
_TEMPLATE = """<!doctype html><html><head><meta charset='utf-8'><style>body{font-family:system-ui;margin:2rem} .PASS{font-weight:700}.finding{margin:.4rem 0}</style></head><body>
<h1>Word Replica QA — {{ status }}</h1>
<p>Source SHA-256: <code>{{ source_sha256 }}</code></p>
<p>Output SHA-256: <code>{{ output_sha256 }}</code></p>
<p>Renderer: {{ renderer }} | Fidelity: {{ fidelity }} | Metadata: {{ metadata }}</p>
<h2>Actual saves ({{ saves|length }})</h2><ol>{% for save in saves %}<li>#{{ save.sequence }} — {{ save.reason }} — {{ save.timestamp_local }}</li>{% endfor %}</ol>
<h2>QA</h2>{% for level, result in levels.items() %}<section><h3>{{ level }} — {{ 'PASS' if result.passed else 'DIFF' }}</h3>{% for f in result.findings %}<div class='finding'>{{ f.code }} @ {{ f.path }}</div>{% endfor %}</section>{% endfor %}
<h2>Warnings / limitations</h2>{% for w in warnings %}<div>{{ w.code }} — {{ w.message }}</div>{% endfor %}
<p><strong>Integrity:</strong> This document was reconstructed automatically. Word Replica does not simulate manual typing or fabricate timestamps, editing time, save history, or provenance.</p>
</body></html>"""

def write_qa_report(path: Path, **context) -> Path:
    env = Environment(autoescape=select_autoescape(default_for_string=True))
    path.write_text(env.from_string(_TEMPLATE).render(**context), encoding="utf-8")
    return path
```

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_render_diff.py tests/unit/test_report.py -v`

Expected: PASS.

```bash
git add src/word_replica/qa/render.py src/word_replica/qa/report.py tests/unit/test_render_diff.py tests/unit/test_report.py
git commit -m "feat: add l4 visual diff and qa report"
```

---

### Task 15: CLI on Top of the Shared Rebuild Service

**Files:**
- Create: `src/word_replica/cli.py`
- Create: `tests/unit/test_cli.py`

**Interfaces:**
- Commands: `rebuild`, `inspect`, `qa`, `projects list`, `project show`.
- CLI only parses inputs and delegates to services.

- [ ] **Step 1: Write CLI parser tests**

```python
# tests/unit/test_cli.py
from word_replica.cli import build_parser


def test_rebuild_command_maps_modes():
    args = build_parser().parse_args(["rebuild", "paper.docx", "--renderer", "auto", "--visibility", "visible", "--fidelity", "full", "--metadata", "fresh"])
    assert args.command == "rebuild"
    assert args.visibility == "visible"
    assert args.fidelity == "full"
```

- [ ] **Step 2: Implement argparse command tree**

```python
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="word-replica")
    sub = parser.add_subparsers(dest="command", required=True)
    rebuild = sub.add_parser("rebuild")
    rebuild.add_argument("source", type=Path)
    rebuild.add_argument("--renderer", choices=[e.value for e in RendererChoice], default="auto")
    rebuild.add_argument("--visibility", choices=[e.value for e in VisibilityMode], default="background")
    rebuild.add_argument("--fidelity", choices=[e.value for e in FidelityMode], default="clean")
    rebuild.add_argument("--metadata", choices=[e.value for e in MetadataMode], default="fresh")
    rebuild.add_argument("--preserve-author-fields", action="store_true")
    rebuild.add_argument("--preserve-custom-property", action="append", default=[])
    rebuild.add_argument("--allow-source-overwrite", action="store_true")
    inspect = sub.add_parser("inspect"); inspect.add_argument("source", type=Path)
    qa = sub.add_parser("qa"); qa.add_argument("source", type=Path); qa.add_argument("rebuilt", type=Path)
    projects = sub.add_parser("projects"); projects.add_argument("action", choices=["list"])
    project = sub.add_parser("project"); project.add_argument("action", choices=["show"]); project.add_argument("id")
    return parser

def require_docx(path: Path) -> Path:
    if path.suffix.lower() != ".docx":
        raise SystemExit(f"Only .docx input is supported: {path}")
    return path
```

- [ ] **Step 3: Implement rebuild output contract**

CLI output must print:

```text
Status: PASS|WARN|FAIL
Output: <absolute path or ->
QA report: <absolute path or ->
Saves: <actual count>
Warnings: <count>
```

Return process codes `0=PASS`, `1=WARN`, `2=FAIL/usage error`.

```python
def run_rebuild(args) -> int:
    options = RebuildOptions(
        renderer=RendererChoice(args.renderer),
        visibility=VisibilityMode(args.visibility),
        fidelity=FidelityMode(args.fidelity),
        metadata=MetadataMode(args.metadata),
        allow_source_overwrite=args.allow_source_overwrite,
        preserve_author_fields=args.preserve_author_fields,
        custom_metadata_allowlist=tuple(args.preserve_custom_property),
    )
    result = RebuildService.default().rebuild(require_docx(args.source), options)
    print(f"Status: {result.status.value}")
    print(f"Output: {result.output_path or '-'}")
    print(f"QA report: {result.qa_report_path or '-'}")
    print(f"Saves: {result.save_count}")
    print(f"Warnings: {len(result.warnings)}")
    return {RunStatus.PASS: 0, RunStatus.WARN: 1, RunStatus.FAIL: 2}[result.status]

def main() -> int:
    args = build_parser().parse_args()
    if args.command == "rebuild": return run_rebuild(args)
    if args.command == "inspect": return run_inspect(args.source)
    store = ProjectStore()
    if args.command == "projects": return run_projects(store)
    if args.command == "project": return run_project_show(store, args.id)
    if args.command == "qa": return run_qa(args.source, args.rebuilt)
    return 2
```

- [ ] **Step 4: Add inspect/project commands**

```python
def run_inspect(path: Path) -> int:
    model = DocxParser().parse(require_docx(path))
    print(json.dumps({"paragraphs": sum(1 for _ in model.iter_paragraphs()), "sections": len(model.sections), "assets": len(model.assets)}, indent=2))
    return 0

def run_projects(store: ProjectStore) -> int:
    print(json.dumps(store.list_projects(), indent=2))
    return 0

def run_project_show(store: ProjectStore, project_id: str) -> int:
    print(json.dumps(store.get_project(project_id), indent=2))
    return 0

def run_qa(source: Path, rebuilt: Path) -> int:
    bundle = run_l0_l3(DocxParser().parse(require_docx(source)), DocxParser().parse(require_docx(rebuilt)))
    status = classify_run(bundle, warnings=[])
    print(status.value)
    return {RunStatus.PASS: 0, RunStatus.WARN: 1, RunStatus.FAIL: 2}[status]
```

- [ ] **Step 5: Run tests and commit**

Run: `python -m pytest tests/unit/test_cli.py -v`

Expected: PASS.

```bash
git add src/word_replica/cli.py tests/unit/test_cli.py
git commit -m "feat: expose reconstruction through cli"
```

---

### Task 16: PySide6 Desktop GUI and Progress Controller

**Files:**
- Create: `src/word_replica/gui/app.py`
- Create: `src/word_replica/gui/main_window.py`
- Create: `src/word_replica/gui/controller.py`
- Create: `tests/unit/test_gui_controller.py`

**Interfaces:**
- GUI controller delegates to `RebuildService` on a worker thread and emits phase/progress/status signals.
- Window never manipulates DOCX/COM directly.

- [ ] **Step 1: Write controller-state test without opening a window**

```python
# tests/unit/test_gui_controller.py
from word_replica.gui.controller import UiState


def test_initial_ui_state_is_safe():
    state = UiState()
    assert state.source_path is None
    assert state.running is False
    assert state.renderer == "auto"
    assert state.fidelity == "clean"
    assert state.metadata == "fresh"
```

- [ ] **Step 2: Implement `UiState` and worker contract**

```python
@dataclass(slots=True)
class UiState:
    source_path: Path | None = None
    renderer: str = "auto"
    visibility: str = "background"
    fidelity: str = "clean"
    metadata: str = "fresh"
    allow_source_overwrite: bool = False
    preserve_author_fields: bool = False
    custom_metadata_allowlist: tuple[str, ...] = ()
    running: bool = False
    current_phase: str = "Idle"
    save_count: int = 0
    warnings: list[str] = field(default_factory=list)
```

Use `QThread` + worker `QObject`; signal payloads are simple strings/ints/`RunResult`. No COM object may cross threads; the Word renderer session is created and destroyed entirely inside the worker thread that executes the rebuild.

```python
def options_from_state(state: UiState) -> RebuildOptions:
    return RebuildOptions(
        renderer=RendererChoice(state.renderer),
        visibility=VisibilityMode(state.visibility),
        fidelity=FidelityMode(state.fidelity),
        metadata=MetadataMode(state.metadata),
        allow_source_overwrite=state.allow_source_overwrite,
        preserve_author_fields=state.preserve_author_fields,
        custom_metadata_allowlist=state.custom_metadata_allowlist,
    )

class RebuildWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)
    def __init__(self, service: RebuildService, source: Path, options: RebuildOptions):
        super().__init__(); self.service = service; self.source = source; self.options = options
    @Slot()
    def run(self):
        try: self.finished.emit(self.service.rebuild(self.source, self.options))
        except Exception as exc: self.failed.emit(str(exc))
```

- [ ] **Step 3: Build minimal main window**

```python
class MainWindow(QMainWindow):
    def __init__(self, controller: RebuildController):
        super().__init__()
        self.controller = controller
        self.source_edit = QLineEdit(); self.browse = QPushButton("Select DOCX")
        self.renderer = QComboBox(); self.renderer.addItems(["Auto", "Microsoft Word", "Pure DOCX"])
        self.visibility = QComboBox(); self.visibility.addItems(["Background", "Visible"])
        self.fidelity = QComboBox(); self.fidelity.addItems(["Clean Replica", "Full Fidelity"])
        self.metadata = QComboBox(); self.metadata.addItems(["Fresh", "Preserve Legitimate"])
        self.start = QPushButton("Start reconstruction")
        self.phase = QLabel("Idle"); self.saves = QLabel("Saves: 0")
        self.messages = QPlainTextEdit(); self.messages.setReadOnly(True)
        self.open_output = QPushButton("Open output folder"); self.open_output.setEnabled(False)
        self.open_report = QPushButton("Open QA report"); self.open_report.setEnabled(False)
        self.allow_overwrite = QCheckBox("Advanced: allow replacing the selected source after backup")
        self.allow_overwrite.setToolTip("Creates a verified backup first; default is always read-only source.")
        self.preserve_author = QCheckBox("Preserve truthful author/company descriptive fields")
        self.custom_properties = QLineEdit()
        self.custom_properties.setPlaceholderText("Custom property names to preserve, comma-separated")

        form = QFormLayout()
        source_row = QHBoxLayout(); source_row.addWidget(self.source_edit); source_row.addWidget(self.browse)
        source_widget = QWidget(); source_widget.setLayout(source_row)
        form.addRow("Source", source_widget)
        form.addRow("Renderer", self.renderer); form.addRow("Visibility", self.visibility)
        form.addRow("Fidelity", self.fidelity); form.addRow("Metadata", self.metadata)
        form.addRow(self.preserve_author); form.addRow("Custom properties", self.custom_properties)
        form.addRow(self.allow_overwrite)
        actions = QHBoxLayout(); actions.addWidget(self.start); actions.addWidget(self.open_output); actions.addWidget(self.open_report)
        layout = QVBoxLayout(); layout.addLayout(form); layout.addWidget(self.phase); layout.addWidget(self.saves); layout.addWidget(self.messages); layout.addLayout(actions)
        central = QWidget(); central.setLayout(layout); self.setCentralWidget(central)
        self.browse.clicked.connect(self.controller.choose_source)
        self.start.clicked.connect(self.controller.start)
        self.renderer.currentTextChanged.connect(lambda _: self.apply_word_capability(self.controller.word_available))
```

- [ ] **Step 4: Add non-Word capability UX**

```python
def apply_word_capability(self, available: bool) -> None:
    word_index = self.renderer.findText("Microsoft Word")
    item = self.renderer.model().item(word_index)
    item.setEnabled(available)
    if not available:
        self.messages.appendPlainText("Microsoft Word desktop not detected — Pure DOCX fallback will be used")
        if self.renderer.currentText() == "Microsoft Word":
            self.renderer.setCurrentText("Auto")
    self.visibility.setEnabled(available and self.renderer.currentText() != "Pure DOCX")
```

Connect renderer changes so selecting `Pure DOCX` disables visibility regardless of Word availability.

- [ ] **Step 5: Run controller tests and manual smoke test, then commit**

Run: `python -m pytest tests/unit/test_gui_controller.py -v`

Manual: `python -m word_replica.gui.app`; select a fixture; verify buttons/mode disabling and that the window remains responsive during a rebuild.

```bash
git add src/word_replica/gui tests/unit/test_gui_controller.py
git commit -m "feat: add desktop reconstruction gui"
```

---

### Task 17: Full Fixture Corpus, Acceptance Matrix, Documentation, and Windows Build

**Files:**
- Modify: `tests/fixtures/build_fixtures.py`
- Create fixtures under: `tests/fixtures/corpus/`
- Create: `tests/acceptance/test_acceptance_matrix.py`
- Create: `README.md`
- Create: `scripts/build_windows.ps1`

**Interfaces:**
- Produces the final v1 acceptance gate and reproducible Windows build instructions.

- [ ] **Step 1: Build the 13 required deterministic fixtures**

`build_fixtures.py` must generate exactly these names:

```text
01_plain_text.docx
02_headings_styles.docx
03_lists.docx
04_tables_merged.docx
05_images_inline_floating.docx
06_sections_orientations.docx
07_headers_footers_numbers.docx
08_footnotes_endnotes.docx
09_toc_fields.docx
10_comments_tracked_changes.docx
11_bookmarks_crossrefs.docx
12_charts_embedded.docx
13_academic_complex.docx
```

```python
BUILDERS = [
    ("01_plain_text.docx", build_plain_text),
    ("02_headings_styles.docx", build_headings_styles),
    ("03_lists.docx", build_lists),
    ("04_tables_merged.docx", build_tables_merged),
    ("05_images_inline_floating.docx", build_images_inline_floating),
    ("06_sections_orientations.docx", build_sections_orientations),
    ("07_headers_footers_numbers.docx", build_headers_footers_numbers),
    ("08_footnotes_endnotes.docx", build_footnotes_endnotes),
    ("09_toc_fields.docx", build_toc_fields),
    ("10_comments_tracked_changes.docx", build_comments_tracked_changes),
    ("11_bookmarks_crossrefs.docx", build_bookmarks_crossrefs),
    ("12_charts_embedded.docx", build_charts_embedded),
    ("13_academic_complex.docx", build_academic_complex),
]

def build_all_corpus(out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    built = []
    for filename, builder in BUILDERS:
        target = out_dir / filename
        builder(target)
        with ZipFile(target) as z:
            assert z.testzip() is None
        built.append(target)
    return built
```

Each builder uses `python-docx` where supported and deterministic low-level XML patches for unsupported features. Each low-level builder rewrites to a temporary sibling ZIP and `Path.replace()`s only after `ZipFile.testzip()` passes, making the generator idempotent and network-independent.

- [ ] **Step 2: Write pure-DOCX acceptance matrix**

```python
# tests/acceptance/test_acceptance_matrix.py
import pytest
from word_replica.config import RebuildOptions
from word_replica.domain.enums import RendererChoice, RunStatus
from word_replica.services.rebuild import RebuildService

PURE_REQUIRED = [
    "01_plain_text.docx", "02_headings_styles.docx", "03_lists.docx",
    "04_tables_merged.docx", "06_sections_orientations.docx",
    "07_headers_footers_numbers.docx", "08_footnotes_endnotes.docx",
]

@pytest.mark.parametrize("fixture_name", PURE_REQUIRED)
def test_required_fallback_fixtures_do_not_fail(corpus_dir, fixture_name):
    result = RebuildService.default_for_tests().rebuild(corpus_dir / fixture_name, RebuildOptions(renderer=RendererChoice.DOCX))
    assert result.status in {RunStatus.PASS, RunStatus.WARN}
```

- [ ] **Step 3: Add Word Full-Fidelity acceptance matrix**

```python
WORD_REQUIRED = [f"{i:02d}_{name}.docx" for i, name in [
    (1,"plain_text"),(2,"headings_styles"),(3,"lists"),(4,"tables_merged"),(5,"images_inline_floating"),
    (6,"sections_orientations"),(7,"headers_footers_numbers"),(8,"footnotes_endnotes"),(9,"toc_fields"),
    (10,"comments_tracked_changes"),(11,"bookmarks_crossrefs"),(12,"charts_embedded"),(13,"academic_complex")]]

@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word acceptance disabled")
@pytest.mark.parametrize("fidelity", [FidelityMode.CLEAN, FidelityMode.FULL])
@pytest.mark.parametrize("fixture_name", WORD_REQUIRED)
def test_word_acceptance_matrix(corpus_dir, fixture_name, fidelity):
    source = corpus_dir / fixture_name
    before = sha256_file(source)
    result = RebuildService.default().rebuild(source, RebuildOptions(renderer=RendererChoice.WORD, fidelity=fidelity))
    assert sha256_file(source) == before
    assert result.status in {RunStatus.PASS, RunStatus.WARN}
    assert result.qa_report_path and result.qa_report_path.exists()
```

For `13_academic_complex.docx`, add an assertion that `save_history.jsonl` contains at least one logical checkpoint plus final save, then read output custom properties and assert `int(WordReplicaActualSaveCount) == len(save_history_rows)`. Also assert source lifecycle properties (`created`, `modified`, `revision`, `TotalTime`) were not copied from the source in Fresh mode.

- [ ] **Step 4: Write README and explicit integrity boundary**

Write `README.md` with this concrete top-level structure and commands:

```markdown
# Word Replica

Local `.docx` reconstruction and QA for Windows.

## Install
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[test]"

## GUI
word-replica-gui

## CLI
word-replica rebuild "C:\Docs\paper.docx" --renderer auto --visibility visible --fidelity full --metadata fresh

## Modes
Explain Visible/Background, Clean/Full Fidelity, Fresh/Preserve Legitimate Metadata exactly as the design spec defines them.

## Output and QA
Document project.json, audit.jsonl, save_history.jsonl, warnings.json, qa_report.html and PASS/WARN/FAIL.

## Tests
python -m pytest tests/unit tests/integration/test_pure_docx_e2e.py -m "not word" -v
$env:WORD_REPLICA_WORD_TESTS="1"
python -m pytest tests/integration/word -m word -v

## Integrity boundary
Word Replica performs automated document reconstruction. It does not simulate manual typing and does not fabricate timestamps, editing time, revision counts, save history, or provenance. Preserve Metadata copies only explicitly allowed descriptive fields.

## Known limitations
Document pixel layout can vary with Word version, installed fonts, printer/layout environment, and unsupported embedded objects; the QA report must state those limitations instead of claiming unverifiable pixel identity.
```

- [ ] **Step 5: Add Windows build script**

```powershell
# scripts/build_windows.ps1
$ErrorActionPreference = "Stop"
python -m pytest tests/unit tests/integration/test_pure_docx_e2e.py tests/acceptance -m "not word" -v
python -m PyInstaller --noconfirm --clean --name WordReplica --windowed --collect-all PySide6 src/word_replica/gui/app.py
Write-Host "Built: dist\WordReplica\WordReplica.exe"
```

- [ ] **Step 6: Run the full non-Word gate**

Run:

```bash
python -m pytest tests/unit tests/integration/test_pure_docx_e2e.py tests/acceptance -m "not word" -v
```

Expected: all tests PASS; no source fixture hash changes.

- [ ] **Step 7: Run Windows + Word gate before declaring v1 complete**

On the target Windows machine with desktop Word:

```powershell
$env:WORD_REPLICA_WORD_TESTS="1"
python -m pytest tests/integration/word tests/acceptance -m word -v
```

Expected: mandatory Full Fidelity fixtures reconstruct without critical errors, all source hashes remain unchanged, real save history exists, and each run yields a readable QA report.

- [ ] **Step 8: Commit the acceptance gate**

```bash
git add tests/fixtures tests/acceptance README.md scripts/build_windows.ps1
git commit -m "test: lock word replica v1 acceptance matrix"
```

---

## Cross-Task Review Gates

After Tasks 1-5:
- run all unit tests;
- confirm no renderer imports COM on non-Windows;
- confirm no source file mutation is possible through default APIs.

After Tasks 6-9:
- parse → pure-render → parse the core fixture corpus;
- verify deterministic canonical fingerprints for repeated parses of the same source;
- inspect generated packages with `zipfile.testzip()` and Word/Open XML validation where available.

After Tasks 10-12:
- run Word integration suite on a real Windows + Word machine;
- verify Visible and Background modes use the same reconstruction logic;
- verify save history increments only after actual saves and content changes.

After Tasks 13-17:
- verify every run has `project.json`, `audit.jsonl`, `save_history.jsonl`, `warnings.json`, and `qa_report.html`;
- inspect one PASS, one WARN, and one forced FAIL report manually;
- verify metadata tests prove lifecycle/history fields are neither copied nor fabricated;
- run the entire acceptance matrix before tagging `v0.1.0`.

## Definition of Done

Word Replica v1 is complete only when:

1. GUI and CLI accept a valid `.docx` and call the same `RebuildService`.
2. Default operation leaves source bytes unchanged and proves that with pre/post SHA-256.
3. Canonical model parsing is deterministic for all supported fixtures.
4. Pure-DOCX renderer handles the documented fallback subset without Word.
5. Word COM renderer handles the mandatory v1 fixture corpus on Windows with Word.
6. Clean/Full and Fresh/Preserve modes have automated tests proving their differences.
7. Every recorded save corresponds to a real file save after a real content change or logical checkpoint, and final `WordReplicaActualSaveCount` equals the number of `save_history.jsonl` rows.
8. L0-L3 run on every completed output; L4 runs only when a controlled PDF-render path is available.
9. Critical loss yields FAIL; non-critical fidelity gaps yield WARN; exact supported reconstruction yields PASS.
10. The HTML report explains all differences and limitations without claiming unverifiable pixel identity.
11. No code path simulates human typing or falsifies timestamps, revision counts, editing time, save history, or provenance.
12. Full unit/non-Word acceptance suite passes; Windows Word integration gate passes before release.
