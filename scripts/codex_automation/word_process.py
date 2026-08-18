from __future__ import annotations

from dataclasses import dataclass
import csv
import io
import json
from pathlib import Path
import subprocess
from typing import Callable, Iterable


@dataclass(frozen=True, slots=True)
class OwnedWordProcess:
    pid: int
    hwnd: int
    owner_process_pid: int
    role: str
    started_filetime: int


def _from_payload(payload: dict) -> OwnedWordProcess:
    return OwnedWordProcess(
        pid=int(payload["pid"]),
        hwnd=int(payload.get("hwnd", 0)),
        owner_process_pid=int(payload["owner_process_pid"]),
        role=str(payload.get("role", "word")),
        started_filetime=int(payload.get("started_filetime", 0)),
    )


def list_owned_word_processes(path: Path) -> list[OwnedWordProcess]:
    path = Path(path)
    if not path.exists():
        return []
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload.get("processes"), list):
        return [_from_payload(item) for item in payload["processes"] if isinstance(item, dict) and "pid" in item]
    return [_from_payload(payload)] if "pid" in payload else []


def read_owned_word_process(path: Path) -> OwnedWordProcess | None:
    records = list_owned_word_processes(path)
    return records[-1] if records else None


def may_terminate_owned_word(
    record: OwnedWordProcess | None, *, expected_owner_pid: int | None = None,
    expected_owner_pids: Iterable[int] | None = None, image_name: str, live_started_filetime: int
) -> bool:
    if record is None:
        return False
    allowed_owners = (
        {int(pid) for pid in expected_owner_pids}
        if expected_owner_pids is not None
        else ({int(expected_owner_pid)} if expected_owner_pid is not None else set())
    )
    return (
        record.owner_process_pid in allowed_owners
        and image_name.upper() == "WINWORD.EXE"
        and record.started_filetime > 0
        and int(live_started_filetime) == record.started_filetime
    )


def _windows_image_name(pid: int) -> str:
    proc = subprocess.run(
        ["tasklist", "/FI", f"PID eq {int(pid)}", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )
    text = proc.stdout.strip()
    if not text or text.upper().startswith("INFO:"):
        return ""
    rows = list(csv.reader(io.StringIO(text)))
    return rows[0][0] if rows and rows[0] else ""



def _windows_process_creation_filetime(pid: int) -> int:
    from word_replica.renderers.word_ownership import process_creation_filetime

    return int(process_creation_filetime(int(pid)))


def _windows_kill_pid(pid: int) -> None:
    subprocess.run(
        ["taskkill", "/PID", str(int(pid)), "/T", "/F"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        timeout=30,
        check=False,
    )


def terminate_owned_word_processes(
    records: Iterable[OwnedWordProcess],
    *,
    expected_owner_pid: int | None = None,
    expected_owner_pids: Iterable[int] | None = None,
    image_resolver: Callable[[int], str] | None = None,
    identity_resolver: Callable[[int], int] | None = None,
    killer: Callable[[int], None] | None = None,
) -> list[int]:
    resolve = image_resolver or _windows_image_name
    identify = identity_resolver or _windows_process_creation_filetime
    kill = killer or _windows_kill_pid
    terminated: list[int] = []
    for record in records:
        try:
            image = resolve(record.pid)
            live_started_filetime = int(identify(record.pid))
        except Exception:
            continue
        if not may_terminate_owned_word(
            record,
            expected_owner_pid=expected_owner_pid,
            expected_owner_pids=expected_owner_pids,
            image_name=image,
            live_started_filetime=live_started_filetime,
        ):
            continue
        kill(record.pid)
        terminated.append(record.pid)
    return terminated
