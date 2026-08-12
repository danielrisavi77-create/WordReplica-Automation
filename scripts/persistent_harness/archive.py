from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
import json
from pathlib import Path, PurePosixPath
import re
from zipfile import ZipFile, ZipInfo

from scripts.persistent_harness.update_manifest import UpdateManifest, UpdateValidationError, version_key


@dataclass(frozen=True, slots=True)
class ValidatedUpdate:
    zip_path: Path
    package_sha256: str
    target_version: str
    manifest: UpdateManifest
    payload_members: tuple[str, ...]


def _file_sha256(path: Path) -> str:
    h = sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _safe_member(info: ZipInfo) -> str:
    name = info.filename.replace("\\", "/")
    if not name or name.startswith("/") or re.match(r"^[A-Za-z]:", name):
        raise UpdateValidationError(f"unsafe archive path: {name}")
    parts = PurePosixPath(name).parts
    if any(part in {"..", ""} for part in parts):
        raise UpdateValidationError(f"unsafe archive path: {name}")
    mode = (info.external_attr >> 16) & 0o170000
    if mode == 0o120000:
        raise UpdateValidationError(f"unsafe symlink-like archive entry: {name}")
    return "/".join(parts)


def validate_update_zip(zip_path: Path, installed_version: str) -> ValidatedUpdate:
    zip_path = Path(zip_path).resolve()
    if not zip_path.is_file():
        raise UpdateValidationError(f"update ZIP not found: {zip_path}")
    package_sha = _file_sha256(zip_path)
    with ZipFile(zip_path) as zf:
        safe_names: list[str] = []
        seen: set[str] = set()
        for info in zf.infolist():
            name = _safe_member(info)
            if name in seen:
                raise UpdateValidationError(f"duplicate archive path: {name}")
            seen.add(name)
            safe_names.append(name)
        if "update_manifest.json" not in seen:
            raise UpdateValidationError("update_manifest.json is missing")
        try:
            manifest = UpdateManifest.from_dict(json.loads(zf.read("update_manifest.json").decode("utf-8")))
        except UpdateValidationError:
            raise
        except Exception as exc:
            raise UpdateValidationError(f"manifest is unreadable: {exc}") from exc
        if version_key(installed_version) < version_key(manifest.minimum_installed_version):
            raise UpdateValidationError(
                f"installed version {installed_version} is below minimum installed version {manifest.minimum_installed_version}"
            )
        if version_key(manifest.target_version) <= version_key(installed_version):
            raise UpdateValidationError("target version must be newer than installed version")
        payload_members = tuple(sorted(name for name in safe_names if name.startswith("payload/") and not name.endswith("/")))
        unexpected = sorted(name for name in safe_names if name != "update_manifest.json" and not name.startswith("payload/"))
        if unexpected:
            raise UpdateValidationError(f"archive contains files outside payload: {unexpected[0]}")
        listed = {f"payload/{name}" for name in manifest.files}
        actual = set(payload_members)
        missing = sorted(listed - actual)
        unlisted = sorted(actual - listed)
        if missing:
            raise UpdateValidationError(f"manifest payload file missing: {missing[0]}")
        if unlisted:
            raise UpdateValidationError(f"unlisted payload file: {unlisted[0]}")
        for relative, expected in manifest.files.items():
            actual_digest = sha256(zf.read(f"payload/{relative}")).hexdigest()
            if actual_digest != expected:
                raise UpdateValidationError(f"SHA-256 mismatch for payload/{relative}")
    return ValidatedUpdate(zip_path, package_sha, manifest.target_version, manifest, payload_members)


__all__ = ["UpdateValidationError", "ValidatedUpdate", "validate_update_zip"]
