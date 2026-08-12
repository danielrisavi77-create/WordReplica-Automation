from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class PersistentLayout:
    root: Path
    input_dir: Path
    results_root: Path
    current: Path
    backup: Path
    updates: Path
    venv: Path
    config: Path
    version_file: Path

    @classmethod
    def from_root(cls, root: Path) -> "PersistentLayout":
        root = Path(root).resolve()
        return cls(
            root=root,
            input_dir=root / "realworld_input",
            results_root=root / "results",
            current=root / "current",
            backup=root / "backup",
            updates=root / "updates",
            venv=root / ".venv_harness",
            config=root / "harness_config.json",
            version_file=root / "version.json",
        )
