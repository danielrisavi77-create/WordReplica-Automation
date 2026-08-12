from __future__ import annotations

from hashlib import sha256
from pathlib import Path
import json
import tomllib


def dependency_fingerprint(pyproject_path: Path) -> str:
    raw = tomllib.loads(Path(pyproject_path).read_text(encoding="utf-8"))
    project = raw.get("project", {})
    optional = project.get("optional-dependencies", {})
    payload = {
        "requires-python": project.get("requires-python"),
        "dependencies": sorted(str(item) for item in project.get("dependencies", [])),
        "test": sorted(str(item) for item in optional.get("test", [])),
    }
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return sha256(encoded).hexdigest()


def dependencies_changed(old_fingerprint: str, new_fingerprint: str) -> bool:
    return str(old_fingerprint) != str(new_fingerprint)
