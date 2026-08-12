from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import re


class UpdateValidationError(ValueError):
    pass


_VERSION_RE = re.compile(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?")


def version_key(value: str) -> tuple[int, int, int]:
    match = _VERSION_RE.match(str(value).strip())
    if not match:
        raise UpdateValidationError(f"invalid version: {value}")
    return tuple(int(part or 0) for part in match.groups())


@dataclass(frozen=True, slots=True)
class UpdateManifest:
    schema_version: int
    package_id: str
    target_version: str
    minimum_installed_version: str
    files: dict[str, str]

    @classmethod
    def from_dict(cls, raw: dict) -> "UpdateManifest":
        required = {"schema_version", "package_id", "target_version", "minimum_installed_version", "files"}
        missing = sorted(required - set(raw))
        if missing:
            raise UpdateValidationError(f"manifest missing fields: {', '.join(missing)}")
        if int(raw["schema_version"]) != 1:
            raise UpdateValidationError("unsupported update manifest schema version")
        files = raw["files"]
        if not isinstance(files, dict) or not files:
            raise UpdateValidationError("manifest files must be a non-empty object")
        normalized: dict[str, str] = {}
        for name, digest in files.items():
            name = str(name).replace("\\", "/")
            digest = str(digest).lower()
            if not re.fullmatch(r"[0-9a-f]{64}", digest):
                raise UpdateValidationError(f"invalid SHA-256 for payload file: {name}")
            normalized[name] = digest
        return cls(
            schema_version=1,
            package_id=str(raw["package_id"]),
            target_version=str(raw["target_version"]),
            minimum_installed_version=str(raw["minimum_installed_version"]),
            files=normalized,
        )

    @classmethod
    def load(cls, path: Path) -> "UpdateManifest":
        return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))
