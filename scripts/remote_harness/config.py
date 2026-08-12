from __future__ import annotations

from dataclasses import dataclass, fields
import json
from pathlib import Path


@dataclass(frozen=True, slots=True)
class HarnessConfig:
    setup_timeout_seconds: int = 300
    static_timeout_seconds: int = 120
    instant_timeout_seconds: int = 600
    interactive_timeout_seconds: int = 1800
    l4_timeout_seconds: int = 600
    minimum_free_space_mb: int = 750
    logs_only: bool = False

    @classmethod
    def load(cls, path: Path) -> "HarnessConfig":
        path = Path(path)
        if not path.exists():
            return cls()
        raw = json.loads(path.read_text(encoding="utf-8"))
        allowed = {item.name for item in fields(cls)}
        kwargs = {key: value for key, value in raw.items() if key in allowed}
        return cls(**kwargs)

    def timeout_for_stage(self, stage: str) -> int:
        if stage == "static":
            return self.static_timeout_seconds
        if stage == "instant":
            return self.instant_timeout_seconds
        if stage in {"interactive_maximum", "interactive_standard"}:
            return self.interactive_timeout_seconds
        if stage == "l4":
            return self.l4_timeout_seconds
        return self.setup_timeout_seconds
