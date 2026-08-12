from __future__ import annotations

from datetime import datetime, timezone
import json
import os
from pathlib import Path
from typing import Callable


def process_creation_filetime(pid: int) -> int:
    """Return the Windows process creation FILETIME as a stable 64-bit integer."""
    if os.name != "nt":
        raise RuntimeError("Windows process identity is unavailable")
    import ctypes
    from ctypes import wintypes

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

    class FILETIME(ctypes.Structure):
        _fields_ = [("dwLowDateTime", wintypes.DWORD), ("dwHighDateTime", wintypes.DWORD)]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.GetProcessTimes.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
    ]
    kernel32.GetProcessTimes.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL

    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        raise OSError(ctypes.get_last_error(), f"OpenProcess failed for PID {pid}")
    try:
        created = FILETIME()
        exited = FILETIME()
        kernel = FILETIME()
        user = FILETIME()
        if not kernel32.GetProcessTimes(handle, ctypes.byref(created), ctypes.byref(exited), ctypes.byref(kernel), ctypes.byref(user)):
            raise OSError(ctypes.get_last_error(), f"GetProcessTimes failed for PID {pid}")
        return (int(created.dwHighDateTime) << 32) | int(created.dwLowDateTime)
    finally:
        kernel32.CloseHandle(handle)

ENV_OWNERSHIP_FILE = "WORD_REPLICA_WORD_OWNERSHIP_FILE"


def _default_pid_resolver(hwnd: int) -> int:
    import win32process
    _thread_id, pid = win32process.GetWindowThreadProcessId(int(hwnd))
    return int(pid)


def _record_path() -> Path | None:
    raw = os.environ.get(ENV_OWNERSHIP_FILE)
    return Path(raw).resolve() if raw else None


def _read_registry(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []
    if isinstance(payload.get("processes"), list):
        return [dict(item) for item in payload["processes"] if isinstance(item, dict)]
    if "pid" in payload:
        return [dict(payload)]
    return []


def _write_registry(path: Path, processes: list[dict]) -> None:
    if not processes:
        try:
            path.unlink()
        except FileNotFoundError:
            pass
        return
    latest = dict(processes[-1])
    payload = {**latest, "schema_version": 2, "processes": processes}
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    os.replace(temp, path)


def list_owned_words() -> list[dict]:
    path = _record_path()
    return _read_registry(path) if path is not None else []


def record_owned_word(application, *, role: str, pid_resolver: Callable[[int], int] | None = None,
                      process_identity_resolver: Callable[[int], int] | None = None,
                      owner_pid: int | None = None) -> int | None:
    path = _record_path()
    if path is None:
        return None
    hwnd = int(getattr(application, "Hwnd"))
    resolver = pid_resolver or _default_pid_resolver
    pid = int(resolver(hwnd))
    identity_resolver = process_identity_resolver or process_creation_filetime
    started_filetime = int(identity_resolver(pid))
    entry = {
        "pid": pid,
        "hwnd": hwnd,
        "owner_process_pid": int(os.getpid() if owner_pid is None else owner_pid),
        "role": str(role),
        "started_filetime": started_filetime,
        "recorded_utc": datetime.now(timezone.utc).isoformat(),
    }
    processes = [item for item in _read_registry(path) if int(item.get("pid", -1)) != pid]
    processes.append(entry)
    _write_registry(path, processes)
    return pid


def clear_owned_word(pid: int | None) -> None:
    if pid is None:
        return
    path = _record_path()
    if path is None or not path.exists():
        return
    processes = [item for item in _read_registry(path) if int(item.get("pid", -1)) != int(pid)]
    _write_registry(path, processes)
