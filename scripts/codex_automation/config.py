from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path


@dataclass(slots=True)
class CodexAutomationConfig:
    local_root: Path = Path(r"C:\WordReplica-Automation")
    golden_filename: str = "Glavna verzija rektorova (grupno)(1).docx"
    # Golden #1 (golden_filename) is the sole promotion gate for `main`, per
    # AGENTS.md. Additional golden documents are a regression-only layer:
    # a regression on one still stops the autonomous loop, but reaching
    # FULL PASS on them does not by itself make a commit promotable.
    additional_golden_filenames: list[str] = field(default_factory=list)
    reconstruction_timeout_seconds: int = 14_400
    audit_timeout_seconds: int = 1800
    keep_success_runs: int = 1
    minimum_free_space_mb: int = 0
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
    def golden_documents(self) -> list[tuple[str, str]]:
        """(golden_id, filename) pairs, Golden #1 always first."""
        documents = [("golden_1", self.golden_filename)]
        for index, filename in enumerate(self.additional_golden_filenames, start=2):
            documents.append((f"golden_{index}", filename))
        return documents

    def golden_filename_for(self, golden_id: str) -> str:
        for doc_id, filename in self.golden_documents:
            if doc_id == golden_id:
                return filename
        raise KeyError(f"unknown golden document id: {golden_id}")

    def golden_path_for(self, golden_id: str) -> Path:
        return self.golden_dir / self.golden_filename_for(golden_id)

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
