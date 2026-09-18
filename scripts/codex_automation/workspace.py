from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from hashlib import sha256
import json
import os
from pathlib import Path
import shutil
from typing import Callable

from scripts.codex_automation.config import CodexAutomationConfig
from word_replica.renderers.word_ownership import process_creation_filetime


def sha256_file(path: Path) -> str:
    digest = sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


# OpenProcess sets ERROR_INVALID_PARAMETER when the pid does not exist. Every
# other failure -- ERROR_ACCESS_DENIED above all -- means the process may well
# be alive and we simply cannot see it.
_ERROR_INVALID_PARAMETER = 87


def _process_identity(pid: int) -> int | None:
    """Creation FILETIME of a live process, or None when no such process exists.

    Raises RuntimeError when the answer cannot be established -- a platform that
    cannot report process identity, or a process we are not allowed to query.
    Callers must treat that as "unknown", never as "gone".
    """
    try:
        return process_creation_filetime(pid)
    except OSError as exc:
        if exc.errno == _ERROR_INVALID_PARAMETER or getattr(exc, "winerror", None) == _ERROR_INVALID_PARAMETER:
            return None
        raise RuntimeError(f"cannot determine whether PID {pid} is alive: {exc}") from exc


def _self_creation_filetime() -> int | None:
    try:
        return _process_identity(os.getpid())
    except RuntimeError:
        return None


def _read_lock_record(path: Path) -> dict | None:
    try:
        record = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return record if isinstance(record, dict) else None


def lock_holder_state(record: dict | None, *, identity_probe: Callable[[int], int | None]) -> str:
    """One of "held", "stale", "unknown"."""
    if not record:
        return "unknown"
    try:
        pid = int(record["pid"])
    except (KeyError, TypeError, ValueError):
        return "unknown"
    try:
        live_identity = identity_probe(pid)
    except RuntimeError:
        return "unknown"
    if live_identity is None:
        return "stale"
    recorded = record.get("creation_filetime")
    if isinstance(recorded, int) and int(live_identity) != recorded:
        return "stale"
    return "held"


def _lock_busy_message(record: dict | None) -> str:
    pid = (record or {}).get("pid")
    if pid is None:
        return "a Golden Word run is already active"
    return f"a Golden Word run is already active (holder pid {pid})"


@dataclass(slots=True)
class GoldenRunPaths:
    run_id: str
    run_dir: Path
    source_copy: Path
    source_sha256_before: str
    ownership_file: Path
    golden_id: str = "golden_1"


class GoldenWorkspace:
    def __init__(self, config: CodexAutomationConfig) -> None:
        self.config = config
        self._lock_fd: int | None = None
        self._lock_path = self.config.state_dir / "golden_run.lock"

    def ensure_layout(self) -> None:
        for path in (
            self.config.golden_dir,
            self.config.work_dir,
            self.config.diagnostics_dir,
            self.config.archive_dir,
            self.config.state_dir,
        ):
            path.mkdir(parents=True, exist_ok=True)

    def acquire_lock(self, *, identity_probe: Callable[[int], int | None] | None = None) -> None:
        """Take the machine-wide "one Golden Word run at a time" lock.

        The lock is reclaimed only when the recorded holder can be *proved*
        gone: its pid no longer exists, or it exists with a different process
        creation FILETIME (i.e. the pid was reused). Anything we cannot prove
        -- an unreadable record, a platform that cannot answer -- counts as
        held, because wrongly clearing this lock lets two Word runs overlap.
        This mirrors the pid + creation-FILETIME ownership proof AGENTS.md
        already requires before terminating a Word process.
        """
        self.ensure_layout()
        probe = identity_probe or _process_identity
        for attempt in range(2):
            try:
                self._lock_fd = os.open(self._lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                break
            except FileExistsError as exc:
                holder = _read_lock_record(self._lock_path)
                if attempt == 1 or lock_holder_state(holder, identity_probe=probe) != "stale":
                    raise RuntimeError(_lock_busy_message(holder)) from exc
                try:
                    self._lock_path.unlink()
                except FileNotFoundError:
                    pass
        os.write(self._lock_fd, json.dumps({
            "pid": os.getpid(),
            "creation_filetime": _self_creation_filetime(),
            "created_utc": datetime.now(timezone.utc).isoformat(),
        }).encode("utf-8"))

    def release_lock(self) -> None:
        if self._lock_fd is not None:
            os.close(self._lock_fd)
            self._lock_fd = None
        try:
            self._lock_path.unlink()
        except FileNotFoundError:
            pass

    def create_run(self, commit_sha: str, golden_id: str = "golden_1") -> GoldenRunPaths:
        self.ensure_layout()
        golden = self.config.golden_path_for(golden_id)
        if not golden.is_file():
            raise FileNotFoundError(f"Golden source not found: {golden}")
        source_hash = sha256_file(golden)
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        run_id = f"{stamp}_{commit_sha[:8]}"
        run_dir = self.config.work_dir / run_id
        suffix = 1
        while run_dir.exists():
            suffix += 1
            run_dir = self.config.work_dir / f"{run_id}_{suffix}"
        run_dir.mkdir(parents=True)
        source_copy = run_dir / "source.docx"
        shutil.copy2(golden, source_copy)
        return GoldenRunPaths(
            run_id=run_dir.name,
            run_dir=run_dir,
            source_copy=source_copy,
            source_sha256_before=source_hash,
            ownership_file=run_dir / "owned_word.json",
            golden_id=golden_id,
        )

    def verify_golden_unchanged(self, run: GoldenRunPaths) -> bool:
        golden = self.config.golden_path_for(run.golden_id)
        return golden.is_file() and sha256_file(golden) == run.source_sha256_before
