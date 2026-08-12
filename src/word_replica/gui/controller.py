from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from word_replica.config import RebuildOptions
from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, VisibilityMode
from word_replica.renderers.word_com import word_available
from word_replica.services.rebuild import RebuildService


@dataclass(slots=True)
class UiState:
    source_path: Path | None = None
    renderer: str = "auto"
    visibility: str = "background"
    fidelity: str = "clean"
    metadata: str = "fresh"
    allow_source_overwrite: bool = False
    preserve_author_fields: bool = False
    custom_metadata_allowlist: tuple[str, ...] = ()
    running: bool = False
    current_phase: str = "Idle"
    save_count: int = 0
    warnings: list[str] = field(default_factory=list)


def options_from_state(state: UiState) -> RebuildOptions:
    return RebuildOptions(
        renderer=RendererChoice(state.renderer),
        visibility=VisibilityMode(state.visibility),
        fidelity=FidelityMode(state.fidelity),
        metadata=MetadataMode(state.metadata),
        allow_source_overwrite=state.allow_source_overwrite,
        preserve_author_fields=state.preserve_author_fields,
        custom_metadata_allowlist=state.custom_metadata_allowlist,
    )


try:
    from PySide6.QtCore import QObject, QThread, Signal, Slot
    from PySide6.QtWidgets import QFileDialog
except ImportError:  # Core library remains importable/testable without optional GUI runtime.
    QObject = QThread = Signal = Slot = QFileDialog = None  # type: ignore[assignment]


if QObject is not None:
    class RebuildWorker(QObject):
        finished = Signal(object)
        failed = Signal(str)

        def __init__(self, service: RebuildService, source: Path, options: RebuildOptions):
            super().__init__()
            self.service = service
            self.source = source
            self.options = options

        @Slot()
        def run(self) -> None:
            try:
                self.finished.emit(self.service.rebuild(self.source, self.options))
            except Exception as exc:  # Worker boundary converts unexpected errors into UI status.
                self.failed.emit(str(exc))


    class RebuildController(QObject):
        state_changed = Signal(object)
        run_finished = Signal(object)
        run_failed = Signal(str)
        source_selected = Signal(str)

        def __init__(self, service: RebuildService | None = None):
            super().__init__()
            self.service = service or RebuildService.default()
            self.state = UiState()
            self.word_available = word_available()
            self._thread: QThread | None = None
            self._worker: RebuildWorker | None = None

        @Slot()
        def choose_source(self) -> None:
            selected, _ = QFileDialog.getOpenFileName(None, "Select DOCX", "", "Word documents (*.docx)")
            if selected:
                self.state.source_path = Path(selected)
                self.source_selected.emit(selected)
                self.state_changed.emit(self.state)

        @Slot()
        def start(self) -> None:
            if self.state.running:
                return
            if self.state.source_path is None:
                self.run_failed.emit("Select a .docx source first")
                return
            self.state.running = True
            self.state.current_phase = "Reconstructing"
            self.state_changed.emit(self.state)
            thread = QThread(self)
            worker = RebuildWorker(self.service, self.state.source_path, options_from_state(self.state))
            worker.moveToThread(thread)
            thread.started.connect(worker.run)
            worker.finished.connect(self._finish)
            worker.failed.connect(self._fail)
            worker.finished.connect(thread.quit)
            worker.failed.connect(thread.quit)
            thread.finished.connect(worker.deleteLater)
            thread.finished.connect(thread.deleteLater)
            self._thread, self._worker = thread, worker
            thread.start()

        @Slot(object)
        def _finish(self, result: Any) -> None:
            self.state.running = False
            self.state.current_phase = result.status.value
            self.state.save_count = result.save_count
            self.state.warnings = [warning.message for warning in result.warnings]
            self.state_changed.emit(self.state)
            self.run_finished.emit(result)
            self._worker = None
            self._thread = None

        @Slot(str)
        def _fail(self, message: str) -> None:
            self.state.running = False
            self.state.current_phase = "FAIL"
            self.state.warnings = [message]
            self.state_changed.emit(self.state)
            self.run_failed.emit(message)
            self._worker = None
            self._thread = None
else:
    class RebuildWorker:  # pragma: no cover - runtime guard only
        def __init__(self, *args, **kwargs):
            raise RuntimeError("PySide6 is required for the desktop GUI")


    class RebuildController:  # pragma: no cover - runtime guard only
        def __init__(self, *args, **kwargs):
            raise RuntimeError("PySide6 is required for the desktop GUI")
