from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path


@dataclass(slots=True)
class CodexAutomationConfig:
    local_root: Path = Path(r"C:\WordReplica-Automation")
    golden_filename: str = "Glavna verzija rektorova (grupno)(1).docx"
    reconstruction_timeout_seconds: int = 7200
    audit_timeout_seconds: int = 1800
    keep_success_runs: int = 1
    keep_failure_runs: int = 2
    visual_dpi: int = 144
    changed_pixel_tolerance: float = 0.001
    mae_tolerance: float = 0.25

    def __post_init__(self) -> None:
        self.local_root = Path(self.local_root)

    @property
    def golden_dir(self) -> Path:
        return self.local_root / "golden"

    @property
    def golden_path(self) -> Path:
        return self.golden_dir / self.golden_filename

    @property
    def work_dir(self) -> Path:
        return self.local_root / "work"

    @property
    def diagnostics_dir(self) -> Path:
        return self.local_root / "diagnostics"

    @property
    def archive_dir(self) -> Path:
        return self.local_root / "archive"

    @property
    def state_dir(self) -> Path:
        return self.local_root / "state"


def load_config(path: Path, *, local_root_override: Path | None = None) -> CodexAutomationConfig:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    allowed = set(CodexAutomationConfig.__dataclass_fields__)
    values = {key: value for key, value in payload.items() if key in allowed}
    if local_root_override is not None:
        values["local_root"] = Path(local_root_override)
    return CodexAutomationConfig(**values)
