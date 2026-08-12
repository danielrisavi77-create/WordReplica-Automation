from __future__ import annotations

from pathlib import Path

from scripts.persistent_harness.layout import PersistentLayout
from scripts.persistent_harness.updater import UpdateTransactionError, _restore_backup
from scripts.persistent_harness.versioning import VersionMetadata


def available_backups(layout: PersistentLayout) -> list[Path]:
    if not layout.backup.exists():
        return []
    return sorted(
        (p for p in layout.backup.iterdir() if p.is_dir() and (p / "current").is_dir() and (p / "version.json").is_file()),
        key=lambda p: p.name,
        reverse=True,
    )


def rollback_latest(layout: PersistentLayout, smoke_test) -> VersionMetadata:
    backups = available_backups(layout)
    if not backups:
        raise UpdateTransactionError("no valid harness backup is available")
    return _restore_backup(layout, backups[0], smoke_test)
