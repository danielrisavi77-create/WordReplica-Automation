import os

import pytest

from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import InteractiveSpeedMode, ReconstructionMode, RunStatus
from word_replica.interactive.control import InteractiveRunControl
from word_replica.services.rebuild import RebuildService

pytestmark = [
    pytest.mark.word,
    pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word integration disabled"),
]


def _options():
    return RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(
            speed_mode=InteractiveSpeedMode.MAXIMUM,
            object_step_delay_ms=0,
            checkpoint_event_interval=100000,
            verify_during_run=False,
        ),
    )


def test_real_word_stop_mid_table_then_restart_resume_finishes_same_project(corpus_dir, tmp_path):
    source = corpus_dir / "04_tables_merged.docx"
    control = InteractiveRunControl(); control.start()

    class StopAfterFirstCellCharacter:
        def __init__(self): self.inside_cell = False; self.stopped = False
        def event_completed(self, index, event):
            if event.event_type == "EnterCell": self.inside_cell = True
            elif event.event_type == "LeaveCell": self.inside_cell = False
            elif event.event_type == "InsertCharacter" and self.inside_cell and not self.stopped:
                self.stopped = True
                control.stop()

    service = RebuildService(app_root=tmp_path / "projects")
    first = service.rebuild(source, _options(), interactive_control=control, interactive_observer=StopAfterFirstCellCharacter())
    assert first.status is RunStatus.WARN
    assert first.output_path is not None and first.output_path.exists()
    assert first.project_id
    assert any("resume" in reason.lower() for reason in first.reasons)

    resume_control = InteractiveRunControl(); resume_control.start()
    resumed = service.resume_interactive(first.project_id, interactive_control=resume_control)
    assert resumed.status in {RunStatus.PASS, RunStatus.WARN}, resumed.reasons
    assert resumed.output_path == first.output_path
    assert resumed.qa_report_path is not None and resumed.qa_report_path.exists()
    html = resumed.qa_report_path.read_text(encoding="utf-8")
    assert "L0 — PASS" in html
    assert "L1 — PASS" in html
