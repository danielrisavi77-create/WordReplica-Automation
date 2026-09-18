from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from scripts.persistent_harness.source_manifest import SOURCE_MANIFEST_NAME, source_manifest_bytes


EXCLUDED_DIR_NAMES = {".git", ".pytest_cache", "__pycache__", ".venv", ".venv_harness", "realworld_input", "results", "remote_results", "backup", "updates", "dist", "build", "persistent_root"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo", ".zip"}


def iter_payload_files(source_root: Path):
    source_root = Path(source_root).resolve()
    for path in sorted(p for p in source_root.rglob("*") if p.is_file()):
        rel = path.relative_to(source_root)
        if any(
            part in EXCLUDED_DIR_NAMES or part.startswith(".venv")
            for part in rel.parts[:-1]
        ):
            continue
        if path.suffix.lower() in EXCLUDED_SUFFIXES:
            continue
        if path.name in {"SOURCE_TREE_SHA256.txt", SOURCE_MANIFEST_NAME}:
            continue
        yield path, rel


def build_update_zip(source_root: Path, destination: Path, *, target_version: str, minimum_installed_version: str) -> Path:
    source_root = Path(source_root).resolve()
    destination = Path(destination).resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    entries = []
    files: dict[str, str] = {}
    for path, rel in iter_payload_files(source_root):
        rel_posix = rel.as_posix()
        data = path.read_bytes()
        files[rel_posix] = sha256(data).hexdigest()
        entries.append((rel_posix, data))
    if not files:
        raise ValueError("update payload is empty")
    source_manifest = source_manifest_bytes(entries)
    files[SOURCE_MANIFEST_NAME] = sha256(source_manifest).hexdigest()
    entries.append((SOURCE_MANIFEST_NAME, source_manifest))
    manifest = {
        "schema_version": 1,
        "package_id": f"word-replica-remote-harness-update-{target_version}",
        "target_version": target_version,
        "minimum_installed_version": minimum_installed_version,
        "files": files,
    }
    with ZipFile(destination, "w", ZIP_DEFLATED) as zf:
        zf.writestr("update_manifest.json", json.dumps(manifest, indent=2, sort_keys=True))
        for rel, data in entries:
            zf.writestr(f"payload/{rel}", data)
    return destination
