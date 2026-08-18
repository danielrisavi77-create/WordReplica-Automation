from pathlib import Path

from word_replica.domain.enums import InteractiveRunState
from word_replica.domain.reconstruction import ReconstructionBlueprint, ReconstructionEvent, WordStateSnapshot
from word_replica.interactive.control import InteractiveRunControl
from word_replica.interactive.verification import InteractiveProgressTracker, state_snapshots_match
from word_replica.renderers.interactive_word import InteractiveWordRenderer


def make_snapshot(**overrides):
    values=dict(document_identity="doc", story="body", range_start=1, range_end=1,
        section_index=0, table_element_id=None, cell_element_id=None, last_completed_event_index=0)
    values.update(overrides)
    return WordStateSnapshot(**values)


def test_state_snapshot_mismatch_pauses_before_next_character():
    class Controller:
        def __init__(self): self.events=[]; self.snapshot=make_snapshot(last_completed_event_index=-1, range_start=0, range_end=0)
        def execute_event(self,event):
            self.events.append(event.payload.get("character"))
            self.snapshot=make_snapshot(last_completed_event_index=len(self.events)-1, range_start=len(self.events), range_end=len(self.events))
        def current_state_snapshot(self): return self.snapshot
    controller=Controller()
    renderer=InteractiveWordRenderer(controller=controller)
    renderer.speed._sleep=lambda seconds:None
    bp=ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=(
        ReconstructionEvent("InsertCharacter","r",{"character":"A"}),
        ReconstructionEvent("InsertCharacter","r",{"character":"B"}),
    ))
    control=InteractiveRunControl(); control.start()
    class Observer:
        def __init__(self): self.mismatches=[]
        def event_completed(self,index,event):
            if index == 0:
                controller.snapshot=make_snapshot(last_completed_event_index=0, range_start=999, range_end=999)
        def state_mismatch(self,index,event,expected,actual): self.mismatches.append((index,event,expected,actual))
    observer=Observer()
    outcome=renderer.execute_blueprint(bp, control, observer)
    assert controller.events == ["A"]
    assert outcome.status == "PAUSED_STATE_MISMATCH"
    assert control.state is InteractiveRunState.PAUSED
    assert observer.mismatches and observer.mismatches[0][0] == 1


def test_snapshot_comparison_ignores_last_completed_index_but_not_word_context():
    assert state_snapshots_match(make_snapshot(last_completed_event_index=1), make_snapshot(last_completed_event_index=9))
    assert not state_snapshots_match(make_snapshot(), make_snapshot(story="header:default"))
    assert not state_snapshots_match(make_snapshot(), make_snapshot(range_start=2, range_end=2))


def test_progress_tracker_reports_semantic_counters_and_current_source():
    bp=ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=(
        ReconstructionEvent("BeginSection","s",{"section_index":0}),
        ReconstructionEvent("InsertCharacter","r",{"character":"A"}),
        ReconstructionEvent("EndTable","t",{}),
        ReconstructionEvent("SetImageZOrder","img",{}),
    ), semantic_counts={"sections":1,"tables":1,"images":1})
    tracker=InteractiveProgressTracker(bp)
    for i,event in enumerate(bp.events): tracker.event_completed(i,event)
    progress=tracker.snapshot(state=InteractiveRunState.RUNNING)
    assert progress.completed_events == 4 and progress.total_events == 4
    assert progress.completed_characters == 1 and progress.total_characters == 1
    assert progress.completed_tables == 1
    assert progress.completed_images == 1
    assert progress.completed_sections == 1
    assert progress.source_element_id == "img"


def test_live_verifier_reports_prefix_mismatch_without_mutating_source(tmp_path):
    from word_replica.domain.model import DocumentModel, Paragraph, Run
    from word_replica.interactive.verification import LiveVerifier
    source=DocumentModel(source_sha256="a"*64, body=[Paragraph("p",[Run("r","ABC")])])
    output=DocumentModel(source_sha256="b"*64, body=[Paragraph("p2",[Run("r2","AX")])])
    before=source.fingerprint()
    class Parser:
        def parse(self,path): return output
    result=LiveVerifier(Parser()).verify_boundary(source,tmp_path/"partial.docx",{"name":"p1","expected_text_prefix":"AB"})
    assert result.passed is False
    assert result.status == "PAUSED_FIDELITY_MISMATCH"
    assert "text prefix mismatch" in result.reasons
    assert source.fingerprint() == before


def test_observer_can_halt_after_safe_boundary_before_next_event():
    class Controller:
        def __init__(self): self.events=[]
        def execute_event(self,event): self.events.append(event.event_type)
        def current_state_snapshot(self): return make_snapshot()
    controller=Controller(); renderer=InteractiveWordRenderer(controller=controller); renderer.speed._sleep=lambda _:None
    bp=ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=(
        ReconstructionEvent("EndTable","t",{}),
        ReconstructionEvent("InsertCharacter","r",{"character":"X"}),
    ))
    control=InteractiveRunControl(); control.start()
    class Observer:
        halt_status=None
        def event_completed(self,index,event):
            if event.event_type == "EndTable": self.halt_status="PAUSED_FIDELITY_MISMATCH"
    observer=Observer()
    outcome=renderer.execute_blueprint(bp,control,observer)
    assert controller.events == ["EndTable"]
    assert outcome.status == "PAUSED_FIDELITY_MISMATCH"
    assert outcome.last_completed_index == 0


def test_stop_requested_inside_non_resumable_structure_waits_until_safe_boundary():
    class Controller:
        def __init__(self): self.events=[]; self.safe=True
        def execute_event(self,event):
            self.events.append(event.event_type)
            if event.event_type == "BeginTable": self.safe=False
            if event.event_type == "EnterCell": self.safe=True
        def current_state_snapshot(self): return make_snapshot()
        def is_restart_safe(self): return self.safe
    controller=Controller(); renderer=InteractiveWordRenderer(controller=controller); renderer.speed._sleep=lambda _:None
    bp=ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=(
        ReconstructionEvent("BeginTable","t",{"rows":1,"columns":1}),
        ReconstructionEvent("MergeCells","t",{"row_start":1,"column_start":1,"row_end":1,"column_end":1}),
        ReconstructionEvent("EnterCell","c",{"row":1,"column":1}),
        ReconstructionEvent("InsertCharacter","r",{"character":"X"}),
    ))
    control=InteractiveRunControl(); control.start()
    class Observer:
        halt_status=None
        def event_completed(self,index,event):
            if event.event_type == "BeginTable": control.stop()
    outcome=renderer.execute_blueprint(bp,control,Observer())
    assert controller.events == ["BeginTable","MergeCells","EnterCell"]
    assert outcome.status == "STOPPED"
    assert outcome.last_completed_index == 2


def test_progress_tracker_can_prime_resume_prefix_without_reexecuting_word_events():
    bp=ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=(
        ReconstructionEvent("InsertCharacter","r",{"character":"A"}),
        ReconstructionEvent("InsertCharacter","r",{"character":"B"}),
        ReconstructionEvent("EndTable","t",{}),
        ReconstructionEvent("InsertCharacter","r",{"character":"C"}),
    ))
    tracker=InteractiveProgressTracker(bp)
    tracker.prime_through(2)
    snap=tracker.snapshot(state=InteractiveRunState.RUNNING)
    assert snap.completed_events == 3
    assert snap.completed_characters == 2
    assert snap.completed_tables == 1


def test_service_observer_expected_body_text_includes_semantic_field_cached_result():
    from types import SimpleNamespace
    from word_replica.services.interactive_rebuild import _InteractiveServiceObserver
    bp = ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=())
    control = InteractiveRunControl(); control.start()
    observer = _InteractiveServiceObserver(bp, control, SimpleNamespace(append=lambda *args, **kwargs: None))
    for event in (
        ReconstructionEvent("InsertCharacter", "r", {"character": "T"}),
        ReconstructionEvent("InsertCharacter", "r", {"character": "a"}),
        ReconstructionEvent("CreateField", "f", {"instruction": "REF ref_tab_1 \\h", "cached_result": "1"}),
        ReconstructionEvent("InsertCharacter", "r", {"character": "."}),
    ):
        observer._track_expected_body_text(event)
    assert "".join(observer._expected_body_text) == "Ta1."


def test_service_observer_does_not_add_header_field_result_to_body_prefix():
    from types import SimpleNamespace
    from word_replica.services.interactive_rebuild import _InteractiveServiceObserver
    bp = ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=())
    control = InteractiveRunControl(); control.start()
    observer = _InteractiveServiceObserver(bp, control, SimpleNamespace(append=lambda *args, **kwargs: None))
    observer._track_expected_body_text(ReconstructionEvent("InsertCharacter", "r", {"character": "A"}))
    observer._track_expected_body_text(ReconstructionEvent("BeginHeader", "h", {}))
    observer._track_expected_body_text(ReconstructionEvent("CreateField", "f", {"instruction": "STYLEREF 1", "cached_result": "Header"}))
    observer._track_expected_body_text(ReconstructionEvent("EndHeader", "h", {}))
    observer._track_expected_body_text(ReconstructionEvent("InsertCharacter", "r", {"character": "B"}))
    assert "".join(observer._expected_body_text) == "AB"


def test_table_batch_updates_progress_text_prefix_and_safe_live_verification(tmp_path):
    from types import SimpleNamespace
    from word_replica.services.interactive_rebuild import _InteractiveServiceObserver

    event = ReconstructionEvent("InsertTableBatch", "t", {
        "rows": 2,
        "columns": 2,
        "cells": ({}, {}, {}, {}),
        "paragraph_count": 4,
        "text_projection": "AB",
    })
    bp = ReconstructionBlueprint.build(
        source_sha256="a" * 64,
        source_model_fingerprint="m",
        events=(event,),
        semantic_counts={"tables": 1, "paragraphs": 4},
    )

    class Control:
        state = "RUNNING"
        def pause(self): raise AssertionError("passing verification must not pause")

    class Verifier:
        def __init__(self): self.boundaries = []
        def verify_boundary(self, source_model, output_path, boundary):
            self.boundaries.append(boundary)
            return SimpleNamespace(status="PASS", passed=True, reasons=())

    verifier = Verifier()
    observer = _InteractiveServiceObserver(
        bp,
        Control(),
        SimpleNamespace(append=lambda *args, **kwargs: None),
        live_verifier=verifier,
        source_model=object(),
        output_path=tmp_path / "output.docx",
        settings={"verify_during_run": True},
    )

    observer.event_completed(0, event)

    assert observer.tracker.completed_characters == 2
    assert observer.tracker.completed_tables == 1
    assert observer.tracker.table_index == 0
    assert "".join(observer._expected_body_text) == "AB"
    assert verifier.boundaries == [{
        "name": "InsertTableBatch:t",
        "expected_text_prefix": "AB",
        "completed_tables": 1,
    }]
