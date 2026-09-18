from __future__ import annotations

import json
from pathlib import Path
import re
import shutil


_GOLDEN_RUN_DIR = re.compile(r"^\d{8}T\d{6}Z_[0-9a-fA-F]{8}(?:_\d+)?$")


def _report(path: Path) -> dict:
    report_path = path / "golden_report.json"
    if not report_path.exists():
        return {}
    try:
        return json.loads(report_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def prune_diagnostics(root: Path, *, keep_success: int = 1, keep_failures: int = 2) -> list[Path]:
    root = Path(root)
    if not root.exists():
        return []
    runs = sorted(
        (path for path in root.iterdir() if path.is_dir() and (path / "golden_report.json").is_file()),
        key=lambda p: p.name,
    )
    reports = {path: _report(path) for path in runs}
    pinned = {path for path in runs if bool(reports[path].get("pinned"))}
    successes = [path for path in runs if reports[path].get("full_pass") and path not in pinned]
    failures = [path for path in runs if not reports[path].get("full_pass") and path not in pinned]
    keep = pinned | set(successes[-keep_success:] if keep_success else []) | set(failures[-keep_failures:] if keep_failures else [])
    for path in runs:
        if path not in keep:
            shutil.rmtree(path)
    return sorted(keep, key=lambda p: p.name)


def prune_stale_work(root: Path) -> list[Path]:
    """Remove abandoned Golden run directories while preserving all other work."""
    root = Path(root)
    if not root.exists():
        return []
    removed = []
    for path in sorted(root.iterdir(), key=lambda item: item.name):
        if not path.is_dir() or not _GOLDEN_RUN_DIR.fullmatch(path.name):
            continue
        shutil.rmtree(path)
        removed.append(path)
    return removed
