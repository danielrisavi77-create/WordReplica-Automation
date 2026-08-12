from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

from word_replica.domain.results import WarningItem


@dataclass(slots=True)
class RenderResult:
    output_path: Path
    warnings: list[WarningItem] = field(default_factory=list)
    stages_completed: list[str] = field(default_factory=list)


class Renderer(Protocol):
    def render(self, model, output_path: Path, context) -> RenderResult: ...
    def save(self, output_path: Path) -> None: ...
    def set_custom_property(self, name: str, value: str) -> None: ...
