from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
import os
from pathlib import Path
import shutil
from typing import Callable


@dataclass(slots=True)
class PrerequisiteReport:
    ok: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    word: dict = field(default_factory=dict)
    free_space_mb: int = 0


def _default_word_probe() -> tuple[bool, dict]:
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = None
        try:
            app = win32com.client.DispatchEx("Word.Application")
            info = {
                "version": str(getattr(app, "Version", "unknown")),
                "build": str(getattr(app, "Build", "unknown")),
                "name": str(getattr(app, "Name", "Microsoft Word")),
                "active_printer": str(getattr(app, "ActivePrinter", "unknown")),
            }
            return True, info
        finally:
            if app is not None:
                try: app.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()
    except Exception as exc:
        return False, {"error": str(exc), "exception_type": type(exc).__name__}


def _default_free_space_probe(root: Path) -> int:
    return int(shutil.disk_usage(Path(root)).free // (1024 * 1024))


def _sha256_file(path: Path) -> str:
    h=sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024*1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _default_manifest_probe(root: Path) -> tuple[bool, list[str]]:
    root=Path(root).resolve()
    manifest=root/"REMOTE_HARNESS_SOURCE_SHA256.txt"
    if not manifest.exists():
        return False, ["REMOTE_HARNESS_SOURCE_SHA256.txt is missing"]
    errors=[]
    for line in manifest.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            expected, rel = line.split(None, 1)
        except ValueError:
            errors.append(f"malformed manifest line: {line}")
            continue
        rel=rel.strip()
        if rel.startswith("./") or rel.startswith(".\\"):
            rel=rel[2:]
        relative_path=Path(rel)
        if relative_path.is_absolute() or ".." in relative_path.parts:
            errors.append(f"unsafe manifest path: {rel}")
            continue
        path=(root/relative_path).resolve()
        try:
            path.relative_to(root)
        except ValueError:
            errors.append(f"unsafe manifest path: {rel}")
            continue
        if not path.exists():
            errors.append(f"manifest file missing: {rel}")
            continue
        actual=_sha256_file(path)
        if actual.lower()!=expected.lower():
            errors.append(f"manifest hash mismatch: {rel}")
    return not errors, errors


def check_prerequisites(
    input_dir: Path,
    source_root: Path,
    *,
    minimum_free_space_mb: int = 750,
    os_name: str | None = None,
    word_probe: Callable[[], tuple[bool, dict]] | None = None,
    manifest_probe: Callable[[Path], tuple[bool, list[str]]] | None = None,
    free_space_probe: Callable[[Path], int] | None = None,
) -> PrerequisiteReport:
    errors=[]; warnings=[]
    os_name = os_name if os_name is not None else os.name
    if os_name != "nt":
        errors.append("Windows 10/11 is required for Microsoft Word COM diagnostics")
    input_dir=Path(input_dir)
    if not input_dir.exists() or not any(p.is_file() and p.suffix.lower()==".docx" for p in input_dir.iterdir()):
        errors.append("realworld_input must contain at least one .docx file")
    free_probe=free_space_probe or _default_free_space_probe
    try:
        free_mb=int(free_probe(source_root))
        if free_mb < minimum_free_space_mb:
            errors.append(f"Insufficient free disk space: {free_mb} MB available; {minimum_free_space_mb} MB required")
    except Exception as exc:
        free_mb=0
        errors.append(f"Free disk space check failed: {exc}")
    manifest_check=manifest_probe or _default_manifest_probe
    try:
        manifest_ok, manifest_errors = manifest_check(Path(source_root))
        if not manifest_ok:
            errors.extend(manifest_errors or ["Source manifest integrity check failed"])
    except Exception as exc:
        errors.append(f"Source manifest integrity check failed: {exc}")
    probe=word_probe or _default_word_probe
    try:
        word_ok, word_info=probe()
    except Exception as exc:
        word_ok=False; word_info={"error":str(exc),"exception_type":type(exc).__name__}
    if not word_ok:
        errors.append("Microsoft Word desktop COM activation failed")
    return PrerequisiteReport(ok=not errors, errors=errors, warnings=warnings, word=word_info, free_space_mb=free_mb)
