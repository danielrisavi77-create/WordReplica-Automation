from __future__ import annotations

from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path
import shutil

from scripts.persistent_harness.build_update import iter_payload_files
from scripts.persistent_harness.source_manifest import SOURCE_MANIFEST_NAME, source_manifest_bytes
from scripts.persistent_harness.versioning import VersionMetadata


BOOTSTRAP_VERSION = "2.0.0-persistent.1"


def _tree_hash(root: Path) -> str:
    h = sha256()
    for path in sorted(p for p in Path(root).rglob("*") if p.is_file()):
        rel = path.relative_to(root).as_posix()
        h.update(rel.encode("utf-8"))
        h.update(path.read_bytes())
    return h.hexdigest()


def build_bootstrap(source_root: Path, destination: Path) -> Path:
    source_root = Path(source_root).resolve()
    destination = Path(destination).resolve()
    if destination.exists():
        shutil.rmtree(destination)
    destination.mkdir(parents=True)
    current = destination / "current"
    current.mkdir()
    manifest_entries: list[tuple[str, bytes]] = []
    for source, rel in iter_payload_files(source_root):
        target = current / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        manifest_entries.append((rel.as_posix(), source.read_bytes()))
    (current / SOURCE_MANIFEST_NAME).write_bytes(source_manifest_bytes(manifest_entries))
    persistent_root = source_root / "persistent_root"
    for name in (
        "START_HERE.cmd", "RUN_PERSISTENT_HARNESS.ps1",
        "UPDATE_HARNESS.cmd", "UPDATE_HARNESS.ps1",
        "ROLLBACK_HARNESS.cmd", "ROLLBACK_HARNESS.ps1",
        "IMPORT_EXISTING_INPUT.cmd", "IMPORT_EXISTING_INPUT.ps1",
    ):
        source = persistent_root / name
        if source.exists():
            shutil.copy2(source, destination / name)
    for dirname in ("realworld_input", "results", "backup", "updates"):
        (destination / dirname).mkdir()
    config_source = source_root / "harness_config.json"
    shutil.copy2(config_source, destination / "harness_config.json")
    (destination / "realworld_input" / "PUT_DOCX_FILES_HERE.txt").write_text(
        "Copy your real-world .docx files into this folder once. Future harness updates do not touch this folder.\n",
        encoding="ascii",
    )
    package_sha = _tree_hash(current)
    VersionMetadata(
        schema_version=1,
        harness_version=BOOTSTRAP_VERSION,
        installed_at=datetime.now(timezone.utc).isoformat(),
        package_sha256=package_sha,
        previous_version=None,
    ).write_atomic(destination / "version.json")
    return destination
