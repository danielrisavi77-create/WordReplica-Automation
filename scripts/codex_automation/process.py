from __future__ import annotations

from dataclasses import dataclass, field
import os
from pathlib import Path
import subprocess
import time
from typing import Mapping, Sequence

from scripts.codex_automation.word_process import (
    list_owned_word_processes,
    terminate_owned_word_processes,
)


@dataclass(slots=True)
class ChildResult:
    exit_code: int | None
    timed_out: bool
    elapsed_seconds: float
    terminated_word_pids: list[int] = field(default_factory=list)


def run_owned_child(
    command: Sequence[str],
    *,
    timeout_seconds: int,
    stdout_path: Path,
    stderr_path: Path,
    ownership_file: Path,
    cwd: Path | None = None,
    env: Mapping[str, str] | None = None,
) -> ChildResult:
    stdout_path = Path(stdout_path)
    stderr_path = Path(stderr_path)
    ownership_file = Path(ownership_file)
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    stderr_path.parent.mkdir(parents=True, exist_ok=True)
    ownership_file.parent.mkdir(parents=True, exist_ok=True)
    try:
        ownership_file.unlink()
    except FileNotFoundError:
        pass

    child_env = os.environ.copy()
    if env:
        child_env.update({str(key): str(value) for key, value in env.items()})
    child_env["WORD_REPLICA_WORD_OWNERSHIP_FILE"] = str(ownership_file.resolve())

    started = time.perf_counter()
    with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open("w", encoding="utf-8") as stderr:
        process = subprocess.Popen(
            list(command),
            cwd=str(Path(cwd).resolve()) if cwd else None,
            env=child_env,
            stdout=stdout,
            stderr=stderr,
            text=True,
        )
        timed_out = False
        try:
            exit_code = process.wait(timeout=timeout_seconds)
        except subprocess.TimeoutExpired:
            timed_out = True
            process.kill()
            exit_code = getattr(process, "returncode", None)
            try:
                exit_code = process.wait(timeout=15)
            except Exception:
                exit_code = getattr(process, "returncode", exit_code)

        records = list_owned_word_processes(ownership_file)
        terminated = terminate_owned_word_processes(records, expected_owner_pid=process.pid)

    return ChildResult(
        exit_code=exit_code,
        timed_out=timed_out,
        elapsed_seconds=time.perf_counter() - started,
        terminated_word_pids=terminated,
    )
