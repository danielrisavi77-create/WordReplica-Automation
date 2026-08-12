from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
import subprocess
import sys
import time


@dataclass(slots=True)
class ChildProcessResult:
    exit_code: int | None
    timed_out: bool
    elapsed_seconds: float


def _kill_process_tree(process: subprocess.Popen) -> None:
    if process.poll() is not None:
        return
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    else:
        process.kill()
    try:
        process.wait(timeout=10)
    except Exception:
        pass


def run_child(command: list[str], timeout_seconds: int, stdout_path: Path, stderr_path: Path) -> ChildProcessResult:
    stdout_path = Path(stdout_path); stderr_path = Path(stderr_path)
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    stderr_path.parent.mkdir(parents=True, exist_ok=True)
    started = time.perf_counter()
    with stdout_path.open("w", encoding="utf-8", errors="replace") as out, stderr_path.open("w", encoding="utf-8", errors="replace") as err:
        process = subprocess.Popen(command, stdout=out, stderr=err, text=True)
        try:
            exit_code = process.wait(timeout=timeout_seconds)
            return ChildProcessResult(exit_code=exit_code, timed_out=False, elapsed_seconds=time.perf_counter() - started)
        except subprocess.TimeoutExpired:
            _kill_process_tree(process)
            return ChildProcessResult(exit_code=process.poll(), timed_out=True, elapsed_seconds=time.perf_counter() - started)
