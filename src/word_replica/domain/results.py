from dataclasses import dataclass, field
from pathlib import Path
from word_replica.domain.enums import RunStatus


@dataclass(slots=True)
class WarningItem:
    code: str
    message: str
    element_id: str | None = None
    affects_status: bool = True


@dataclass(slots=True)
class RunResult:
    status: RunStatus
    output_path: Path | None
    qa_report_path: Path | None
    project_id: str | None = None
    save_count: int = 0
    warnings: list[WarningItem] = field(default_factory=list)
    reasons: list[str] = field(default_factory=list)
