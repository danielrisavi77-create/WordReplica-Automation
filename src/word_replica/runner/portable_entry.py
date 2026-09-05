"""Double-click entry point for one paid Lekta local Word repair.

The executable is one stable Authenticode-signed artifact. Lekta changes only
its downloaded filename, which carries the already single-use job id and claim
token. The token is parsed in memory and is never written to a launch file.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path

from word_replica.runner.lekta_claim import ClaimProtocolError, LaunchTicket


_PORTABLE_NAME = re.compile(
    r"^LektaRepair-"
    r"(?P<job>[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})-"
    r"(?P<token>[A-Za-z0-9_-]{43})\.exe$",
    re.IGNORECASE,
)


_SELF_DELETE_SCRIPT = r"""
$target = [Environment]::GetEnvironmentVariable('WORDREPLICA_DELETE_TARGET', 'Process')
$parentText = [Environment]::GetEnvironmentVariable('WORDREPLICA_PARENT_PID', 'Process')
if ([string]::IsNullOrWhiteSpace($target) -or [string]::IsNullOrWhiteSpace($parentText)) { exit 2 }
$parentId = 0
if (-not [int]::TryParse($parentText, [ref]$parentId)) { exit 2 }
for ($wait = 0; $wait -lt 240; $wait++) {
    if (-not (Get-Process -Id $parentId -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Milliseconds 500
}
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $target)) { exit 0 }
    Start-Sleep -Milliseconds 500
}
exit 1
""".strip()


def schedule_portable_cleanup(
    executable_path: Path,
    *,
    spawn: Callable[..., object] = subprocess.Popen,
    parent_pid: int | None = None,
) -> bool:
    """Delete only this frozen, validly named runner after its process exits."""
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return False
    target = Path(executable_path).resolve()
    if target != Path(sys.executable).resolve() or not target.is_file():
        return False
    try:
        parse_portable_executable_name(target)
    except ClaimProtocolError:
        return False

    environment = {
        key: value for key, value in os.environ.items()
        if key.upper() in {"SYSTEMROOT", "WINDIR", "TEMP", "TMP"}
    }
    environment["WORDREPLICA_DELETE_TARGET"] = str(target)
    environment["WORDREPLICA_PARENT_PID"] = str(parent_pid or os.getpid())
    creationflags = subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS
    try:
        spawn(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", _SELF_DELETE_SCRIPT],
            env=environment,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
            creationflags=creationflags,
        )
    except OSError:
        return False
    return True


def parse_portable_executable_name(executable_path: Path) -> LaunchTicket:
    name = Path(executable_path).name
    match = _PORTABLE_NAME.fullmatch(name)
    if match is None:
        raise ClaimProtocolError("invalid portable repair executable name")
    return LaunchTicket.parse({
        "version": 1,
        "jobId": match.group("job"),
        "claimToken": match.group("token"),
    })


def _choose_output() -> Path | None:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        selected = filedialog.askdirectory(
            parent=root,
            title="Odaberi mapu za popravljeni Word dokument",
            mustexist=True,
        )
        return Path(selected) if selected else None
    finally:
        root.destroy()


def _execute(ticket: LaunchTicket, output_dir: Path) -> int:
    from word_replica.cli import execute_repair_runner_ticket

    return execute_repair_runner_ticket(ticket, output_dir)


def _show_error(message: str) -> None:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    try:
        messagebox.showerror("Lekta Word popravak", message, parent=root)
    finally:
        root.destroy()


def run_portable_self_test(
    expected_key_id: str, *, trust_path: Path | None = None
) -> int:
    """Fail closed unless the frozen bundle owns exactly the expected public key."""
    from word_replica.runner.trust_store import load_trust_keys

    if not re.fullmatch(r"[A-Za-z0-9._-]{1,80}", expected_key_id):
        return 2
    packaged_trust_path = (
        Path(trust_path)
        if trust_path is not None
        else Path(__file__).resolve().parent / "trusted_keys.json"
    )
    try:
        trusted = load_trust_keys(packaged_trust_path)
    except Exception:
        return 2
    return 0 if set(trusted) == {expected_key_id} else 2


def run_portable_entry(
    executable_path: Path,
    *,
    choose_output: Callable[[], Path | None] = _choose_output,
    execute: Callable[[LaunchTicket, Path], int] = _execute,
    show_error: Callable[[str], None] = _show_error,
    schedule_cleanup: Callable[[Path], object] = schedule_portable_cleanup,
) -> int:
    try:
        ticket = parse_portable_executable_name(executable_path)
    except ClaimProtocolError:
        show_error("Program nije povezan s valjanim plaćenim Lekta popravkom.")
        return 2

    output_dir = choose_output()
    if output_dir is None:
        return 1
    try:
        result = execute(ticket, Path(output_dir))
        if result == 0:
            schedule_cleanup(Path(executable_path))
        return result
    except Exception:
        show_error(
            "Lokalni Word popravak nije dovršen. Serverska verzija ostaje dostupna u Lekti. "
            "Ponovno pokreni isti EXE za siguran nastavak."
        )
        return 2


def main() -> int:
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        return run_portable_self_test(sys.argv[2])
    executable_path = Path(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
    return run_portable_entry(executable_path)


if __name__ == "__main__":
    raise SystemExit(main())
