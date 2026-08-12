from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit,
    QMainWindow, QPlainTextEdit, QPushButton, QVBoxLayout, QWidget,
)

from word_replica.gui.controller import RebuildController


class MainWindow(QMainWindow):
    def __init__(self, controller: RebuildController):
        super().__init__()
        self.controller = controller
        self._last_output: Path | None = None
        self._last_report: Path | None = None
        self.setWindowTitle("Word Replica")
        self.resize(820, 600)

        self.source_edit = QLineEdit(); self.source_edit.setReadOnly(True)
        self.browse = QPushButton("Select DOCX")
        self.renderer = QComboBox(); self.renderer.addItems(["Auto", "Microsoft Word", "Pure DOCX"])
        self.visibility = QComboBox(); self.visibility.addItems(["Background", "Visible"])
        self.fidelity = QComboBox(); self.fidelity.addItems(["Clean Replica", "Full Fidelity"])
        self.metadata = QComboBox(); self.metadata.addItems(["Fresh", "Preserve Legitimate"])
        self.start = QPushButton("Start reconstruction")
        self.phase = QLabel("Idle"); self.saves = QLabel("Saves: 0")
        self.messages = QPlainTextEdit(); self.messages.setReadOnly(True)
        self.open_output = QPushButton("Open output folder"); self.open_output.setEnabled(False)
        self.open_report = QPushButton("Open QA report"); self.open_report.setEnabled(False)
        self.allow_overwrite = QCheckBox("Advanced: allow replacing the selected source after backup")
        self.allow_overwrite.setToolTip("Creates a verified backup first; default is always read-only source.")
        self.preserve_author = QCheckBox("Preserve truthful author/company descriptive fields")
        self.custom_properties = QLineEdit(); self.custom_properties.setPlaceholderText("Custom property names to preserve, comma-separated")

        form = QFormLayout()
        source_row = QHBoxLayout(); source_row.addWidget(self.source_edit); source_row.addWidget(self.browse)
        source_widget = QWidget(); source_widget.setLayout(source_row)
        form.addRow("Source", source_widget); form.addRow("Renderer", self.renderer)
        form.addRow("Visibility", self.visibility); form.addRow("Fidelity", self.fidelity)
        form.addRow("Metadata", self.metadata); form.addRow(self.preserve_author)
        form.addRow("Custom properties", self.custom_properties); form.addRow(self.allow_overwrite)
        actions = QHBoxLayout(); actions.addWidget(self.start); actions.addWidget(self.open_output); actions.addWidget(self.open_report)
        layout = QVBoxLayout(); layout.addLayout(form); layout.addWidget(self.phase); layout.addWidget(self.saves); layout.addWidget(self.messages); layout.addLayout(actions)
        central = QWidget(); central.setLayout(layout); self.setCentralWidget(central)

        self.browse.clicked.connect(controller.choose_source)
        self.start.clicked.connect(self._start)
        self.renderer.currentTextChanged.connect(self._renderer_changed)
        controller.source_selected.connect(self.source_edit.setText)
        controller.state_changed.connect(self._state_changed)
        controller.run_finished.connect(self._run_finished)
        controller.run_failed.connect(self.messages.appendPlainText)
        self.open_output.clicked.connect(self._open_output)
        self.open_report.clicked.connect(self._open_report)
        self.apply_word_capability(controller.word_available)

    def apply_word_capability(self, available: bool) -> None:
        word_index = self.renderer.findText("Microsoft Word")
        item = self.renderer.model().item(word_index)
        item.setEnabled(available)
        if not available:
            self.messages.appendPlainText("Microsoft Word desktop not detected — Pure DOCX fallback will be used")
            if self.renderer.currentText() == "Microsoft Word":
                self.renderer.setCurrentText("Auto")
        self._renderer_changed(self.renderer.currentText())

    def _renderer_changed(self, text: str) -> None:
        self.visibility.setEnabled(self.controller.word_available and text != "Pure DOCX")

    def _start(self) -> None:
        state = self.controller.state
        state.renderer = {"Auto":"auto", "Microsoft Word":"word", "Pure DOCX":"docx"}[self.renderer.currentText()]
        state.visibility = self.visibility.currentText().lower()
        state.fidelity = {"Clean Replica":"clean", "Full Fidelity":"full"}[self.fidelity.currentText()]
        state.metadata = {"Fresh":"fresh", "Preserve Legitimate":"preserve"}[self.metadata.currentText()]
        state.allow_source_overwrite = self.allow_overwrite.isChecked()
        state.preserve_author_fields = self.preserve_author.isChecked()
        state.custom_metadata_allowlist = tuple(v.strip() for v in self.custom_properties.text().split(",") if v.strip())
        self.controller.start()

    def _state_changed(self, state) -> None:
        self.phase.setText(state.current_phase); self.saves.setText(f"Saves: {state.save_count}")
        self.start.setEnabled(not state.running); self.browse.setEnabled(not state.running)
        for warning in state.warnings:
            self.messages.appendPlainText(warning)

    def _run_finished(self, result) -> None:
        self._last_output = result.output_path; self._last_report = result.qa_report_path
        self.open_output.setEnabled(result.output_path is not None)
        self.open_report.setEnabled(result.qa_report_path is not None)
        self.messages.appendPlainText(f"Status: {result.status.value}")

    def _open_output(self) -> None:
        if self._last_output:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._last_output.parent)))

    def _open_report(self) -> None:
        if self._last_report:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._last_report)))
