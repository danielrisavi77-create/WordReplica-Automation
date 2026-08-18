from __future__ import annotations

from dataclasses import dataclass, field
import ctypes
import os
from pathlib import Path
import signal
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


def descendant_process_pids(root_pid: int, parent_by_pid: Mapping[int, int]) -> set[int]:
    descendants = {int(root_pid)}
    while True:
        found = {
            int(pid) for pid, parent_pid in parent_by_pid.items()
            if int(parent_pid) in descendants
        }
        expanded = descendants | found
        if expanded == descendants:
            return descendants
        descendants = expanded


def _windows_parent_process_ids() -> dict[int, int]:
    from ctypes import wintypes

    TH32CS_SNAPPROCESS = 0x00000002
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32FirstW.restype = wintypes.BOOL
    kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32NextW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL

    snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snapshot == INVALID_HANDLE_VALUE:
        raise OSError(ctypes.get_last_error(), "CreateToolhelp32Snapshot failed")
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(entry)
        parents: dict[int, int] = {}
        if not kernel32.Process32FirstW(snapshot, ctypes.byref(entry)):
            raise OSError(ctypes.get_last_error(), "Process32FirstW failed")
        while True:
            parents[int(entry.th32ProcessID)] = int(entry.th32ParentProcessID)
            if not kernel32.Process32NextW(snapshot, ctypes.byref(entry)):
                break
        return parents
    finally:
        kernel32.CloseHandle(snapshot)


def process_tree_pids(root_pid: int) -> set[int]:
    if os.name != "nt":
        return {int(root_pid)}
    try:
        return descendant_process_pids(root_pid, _windows_parent_process_ids())
    except Exception:
        return {int(root_pid)}


def kill_process_tree(root_pid: int) -> None:
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(int(root_pid)), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=30,
            check=False,
        )
        return
    os.kill(int(root_pid), signal.SIGKILL)


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
            expected_owner_pids = process_tree_pids(process.pid)
            records = list_owned_word_processes(ownership_file)
            kill_process_tree(process.pid)
            exit_code = getattr(process, "returncode", None)
            try:
                exit_code = process.wait(timeout=15)
            except Exception:
                exit_code = getattr(process, "returncode", exit_code)
        else:
            expected_owner_pids = {int(process.pid)}
            records = list_owned_word_processes(ownership_file)
        terminated = terminate_owned_word_processes(
            records, expected_owner_pids=expected_owner_pids
        )

    return ChildResult(
        exit_code=exit_code,
        timed_out=timed_out,
        elapsed_seconds=time.perf_counter() - started,
        terminated_word_pids=terminated,
    )
