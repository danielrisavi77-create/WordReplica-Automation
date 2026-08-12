from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path


@dataclass(frozen=True, slots=True)
class VersionMetadata:
    schema_version: int
    harness_version: str
    installed_at: str
    package_sha256: str
    previous_version: str | None

    @classmethod
    def load(cls, path: Path) -> "VersionMetadata":
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls(
            schema_version=int(raw["schema_version"]),
            harness_version=str(raw["harness_version"]),
            installed_at=str(raw["installed_at"]),
            package_sha256=str(raw["package_sha256"]),
            previous_version=raw.get("previous_version"),
        )

    def write_atomic(self, path: Path) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = path.with_name(path.name + ".tmp")
        temp.write_text(json.dumps(asdict(self), indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        temp.replace(path)
