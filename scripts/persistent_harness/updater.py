from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
import shutil
import uuid
from zipfile import ZipFile

from scripts.persistent_harness.archive import UpdateValidationError, validate_update_zip
from scripts.persistent_harness.layout import PersistentLayout
from scripts.persistent_harness.versioning import VersionMetadata


class UpdateTransactionError(RuntimeError):
    pass


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")


def _marker_path(layout: PersistentLayout) -> Path:
    return layout.root / ".update_transaction.json"


def _write_marker(layout: PersistentLayout, payload: dict) -> None:
    path = _marker_path(layout)
    temp = path.with_name(path.name + ".tmp")
    temp.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(path)


def _safe_rmtree(path: Path) -> None:
    if Path(path).exists():
        shutil.rmtree(path)


def _extract_payload(validated, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=False)
    with ZipFile(validated.zip_path) as zf:
        for member in validated.payload_members:
            relative = Path(*member.split("/")[1:])
            target = destination / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member) as src, target.open("wb") as dst:
                shutil.copyfileobj(src, dst)


def _restore_backup(layout: PersistentLayout, backup_dir: Path, smoke_test) -> VersionMetadata:
    backup_dir = Path(backup_dir)
    backup_current = backup_dir / "current"
    backup_version = backup_dir / "version.json"
    if not backup_current.is_dir() or not backup_version.is_file():
        raise UpdateTransactionError(f"backup is incomplete: {backup_dir}")
    _safe_rmtree(layout.current)
    shutil.copytree(backup_current, layout.current)
    restored = VersionMetadata.load(backup_version)
    restored.write_atomic(layout.version_file)
    try:
        smoke_test(layout.current)
    except Exception as exc:
        raise UpdateTransactionError(f"restored backup failed validation: {exc}") from exc
    return restored


def recover_incomplete_transaction(layout: PersistentLayout, smoke_test) -> bool:
    marker_path = _marker_path(layout)
    if not marker_path.exists():
        return False
    try:
        marker = json.loads(marker_path.read_text(encoding="utf-8"))
        backup_dir = Path(marker["backup_dir"])
    except Exception as exc:
        raise UpdateTransactionError(f"transaction marker is unreadable: {exc}") from exc
    _restore_backup(layout, backup_dir, smoke_test)
    marker_path.unlink(missing_ok=True)
    old_live = marker.get("old_live")
    if old_live:
        _safe_rmtree(Path(old_live))
    staging = marker.get("staging")
    if staging:
        _safe_rmtree(Path(staging))
    return True


def apply_update(layout: PersistentLayout, zip_path: Path, smoke_test) -> VersionMetadata:
    layout = PersistentLayout.from_root(layout.root)
    if not layout.current.is_dir() or not layout.version_file.is_file():
        raise UpdateTransactionError("persistent harness is not initialized")
    for path in (layout.backup, layout.updates):
        path.mkdir(parents=True, exist_ok=True)
    recover_incomplete_transaction(layout, smoke_test)
    installed = VersionMetadata.load(layout.version_file)
    try:
        validated = validate_update_zip(zip_path, installed.harness_version)
    except UpdateValidationError:
        raise

    token = uuid.uuid4().hex[:12]
    staging = layout.root / f".update_staging_{token}"
    old_live = layout.root / f".update_old_current_{token}"
    backup_dir = layout.backup / f"{_stamp()}-{installed.harness_version}"
    _safe_rmtree(staging)
    _safe_rmtree(old_live)

    try:
        _extract_payload(validated, staging)
        try:
            smoke_test(staging)
        except Exception as exc:
            raise UpdateTransactionError(f"pre-install validation failed: {exc}") from exc

        backup_dir.mkdir(parents=True, exist_ok=False)
        shutil.copytree(layout.current, backup_dir / "current")
        shutil.copy2(layout.version_file, backup_dir / "version.json")

        marker = {
            "schema_version": 1,
            "phase": "prepared",
            "backup_dir": str(backup_dir),
            "staging": str(staging),
            "old_live": str(old_live),
            "old_version": installed.harness_version,
            "target_version": validated.target_version,
        }
        _write_marker(layout, marker)

        layout.current.rename(old_live)
        staging.rename(layout.current)
        marker["phase"] = "swapped"
        _write_marker(layout, marker)

        new_meta = VersionMetadata(
            schema_version=1,
            harness_version=validated.target_version,
            installed_at=_utc_now(),
            package_sha256=validated.package_sha256,
            previous_version=installed.harness_version,
        )
        new_meta.write_atomic(layout.version_file)
        marker["phase"] = "version_written"
        _write_marker(layout, marker)

        try:
            smoke_test(layout.current)
        except Exception as exc:
            try:
                _restore_backup(layout, backup_dir, smoke_test)
            finally:
                _safe_rmtree(old_live)
                _marker_path(layout).unlink(missing_ok=True)
            raise UpdateTransactionError(f"post-install validation failed: {exc}") from exc

        _safe_rmtree(old_live)
        _marker_path(layout).unlink(missing_ok=True)
        return new_meta
    except UpdateTransactionError:
        _safe_rmtree(staging)
        raise
    except Exception as exc:
        # If the live tree was already moved/swapped, restore from backup.
        if backup_dir.exists() and (backup_dir / "current").exists():
            try:
                _restore_backup(layout, backup_dir, smoke_test)
            except Exception:
                pass
        _safe_rmtree(staging)
        _safe_rmtree(old_live)
        _marker_path(layout).unlink(missing_ok=True)
        raise UpdateTransactionError(f"update transaction failed: {exc}") from exc
