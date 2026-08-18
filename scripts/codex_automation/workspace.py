from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from hashlib import sha256
import json
import os
from pathlib import Path
import shutil

from scripts.codex_automation.config import CodexAutomationConfig


def sha256_file(path: Path) -> str:
    digest = sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


@dataclass(slots=True)
class GoldenRunPaths:
    run_id: str
    run_dir: Path
    source_copy: Path
    source_sha256_before: str
    ownership_file: Path


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

    def acquire_lock(self) -> None:
        self.ensure_layout()
        try:
            self._lock_fd = os.open(self._lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError as exc:
            raise RuntimeError("a Golden Word run is already active") from exc
        os.write(self._lock_fd, json.dumps({
            "pid": os.getpid(),
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

    def create_run(self, commit_sha: str) -> GoldenRunPaths:
        self.ensure_layout()
        golden = self.config.golden_path
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
        )

    def verify_golden_unchanged(self, run: GoldenRunPaths) -> bool:
        return self.config.golden_path.is_file() and sha256_file(self.config.golden_path) == run.source_sha256_before
