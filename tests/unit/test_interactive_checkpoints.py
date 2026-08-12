from pathlib import Path
import pytest

from word_replica.domain.reconstruction import ReconstructionBlueprint, ReconstructionEvent
from word_replica.interactive.checkpoints import InteractiveCheckpoint, InteractiveCheckpointStore
from word_replica.services.source_guard import sha256_file


def blueprint():
    return ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="model", events=(ReconstructionEvent("InsertCharacter","r",{"character":"A"}), ReconstructionEvent("InsertCharacter","r",{"character":"B"})))


def make_checkpoint(source, output, bp, **overrides):
    values=dict(project_id="p1", source_sha256=sha256_file(source), source_model_fingerprint=bp.source_model_fingerprint,
        blueprint_fingerprint=bp.fingerprint, blueprint_schema_version=bp.schema_version,
        output_path=str(output), output_sha256=sha256_file(output), last_completed_event_index=0,
        section_element_id="s1", block_element_id="p1", table_element_id=None, cell_element_id=None,
        save_sequence=3, settings={"speed_mode":"fast","fidelity":"maximum"}, timestamp_utc="2026-08-11T12:00:00+00:00", status="STOPPED")
    values.update(overrides); return InteractiveCheckpoint(**values)


def test_checkpoint_round_trip_preserves_every_field(tmp_path):
    source=tmp_path/"source.docx"; source.write_bytes(b"source")
    output=tmp_path/"partial.docx"; output.write_bytes(b"output")
    bp=blueprint(); checkpoint=make_checkpoint(source,output,bp)
    store=InteractiveCheckpointStore(tmp_path/"checkpoint.json")
    store.write(checkpoint)
    assert store.load_latest() == checkpoint


@pytest.mark.parametrize("dimension,mutator", [
    ("source hash", lambda source,output,bp,cp: source.write_bytes(b"changed")),
    ("blueprint fingerprint", lambda source,output,bp,cp: object.__setattr__(cp,"blueprint_fingerprint","bad")),
    ("schema", lambda source,output,bp,cp: object.__setattr__(cp,"blueprint_schema_version",999)),
    ("output hash", lambda source,output,bp,cp: output.write_bytes(b"tampered")),
    ("output missing", lambda source,output,bp,cp: output.unlink()),
    ("event index", lambda source,output,bp,cp: object.__setattr__(cp,"last_completed_event_index",99)),
])
def test_resume_validation_names_failed_dimension(tmp_path, dimension, mutator):
    source=tmp_path/"source.docx"; source.write_bytes(b"source")
    output=tmp_path/"partial.docx"; output.write_bytes(b"output")
    bp=blueprint(); cp=make_checkpoint(source,output,bp)
    mutator(source,output,bp,cp)
    store=InteractiveCheckpointStore(tmp_path/"checkpoint.json")
    result=store.validate_resume(cp, source_path=source, blueprint=bp)
    assert result.valid is False
    assert dimension in result.reason.lower()


def test_checkpoint_round_trip_preserves_word_insertion_state(tmp_path):
    source=tmp_path/"source.docx"; source.write_bytes(b"source")
    output=tmp_path/"partial.docx"; output.write_bytes(b"output")
    bp=blueprint()
    cp=make_checkpoint(source, output, bp,
        story="body", range_start=42, range_end=42, paragraph_started=True,
    )
    store=InteractiveCheckpointStore(tmp_path/"checkpoint.json")
    store.write(cp)
    loaded=store.load_latest()
    assert loaded.story == "body"
    assert loaded.range_start == 42
    assert loaded.range_end == 42
    assert loaded.paragraph_started is True


def test_checkpoint_coordinator_saves_word_before_persisting_metadata(tmp_path):
    from word_replica.interactive.checkpoints import InteractiveCheckpointCoordinator
    from word_replica.services.audit import AuditLog
    from word_replica.services.checkpoints import CheckpointManager

    source=tmp_path/"source.docx"; source.write_bytes(b"source")
    output=tmp_path/"partial.docx"
    bp=blueprint()
    audit=AuditLog(tmp_path/"audit.jsonl")
    saves=CheckpointManager(tmp_path/"save_history.jsonl", audit)
    calls=[]

    class Renderer:
        def save(self, path):
            calls.append("save")
            Path(path).write_bytes(b"actual-word-save")
        def current_state_snapshot(self):
            return {"story":"body", "range_start":12, "range_end":12, "paragraph_started":True}

    store=InteractiveCheckpointStore(tmp_path/"interactive_checkpoint.json")
    coordinator=InteractiveCheckpointCoordinator(
        project_id="p1", source_path=source, blueprint=bp, output_path=output,
        settings={"checkpoint_event_interval":1}, renderer=Renderer(),
        save_manager=saves, checkpoint_store=store,
    )
    coordinator.event_completed(0, bp.events[0])
    cp=store.load_latest()

    assert calls == ["save"]
    assert cp.output_sha256 == sha256_file(output)
    assert cp.save_sequence == 1
    assert cp.last_completed_event_index == 0
    assert cp.range_start == 12 and cp.range_end == 12
    assert cp.paragraph_started is True
    assert saves.sequence == 1


def test_checkpoint_coordinator_boundary_and_stop_triggers_are_truthful(tmp_path):
    from word_replica.interactive.checkpoints import InteractiveCheckpointCoordinator
    from word_replica.services.audit import AuditLog
    from word_replica.services.checkpoints import CheckpointManager

    source=tmp_path/"source.docx"; source.write_bytes(b"source")
    output=tmp_path/"partial.docx"
    bp=ReconstructionBlueprint.build(source_sha256=sha256_file(source), source_model_fingerprint="m", events=(
        ReconstructionEvent("InsertCharacter","p",{"character":"A"}),
        ReconstructionEvent("EndTable","t",{}),
    ))
    class Renderer:
        def save(self,path): Path(path).write_bytes(f"save-{manager.next_sequence}".encode())
        def current_state_snapshot(self): return {"story":"body","range_start":1,"range_end":1,"paragraph_started":True}
    manager=CheckpointManager(tmp_path/"history.jsonl", AuditLog(tmp_path/"audit.jsonl"))
    store=InteractiveCheckpointStore(tmp_path/"cp.json")
    coordinator=InteractiveCheckpointCoordinator(project_id="p", source_path=source, blueprint=bp, output_path=output,
        settings={"checkpoint_event_interval":999, "checkpoint_after_tables":True, "checkpoint_after_images":True, "checkpoint_after_sections":True},
        renderer=Renderer(), save_manager=manager, checkpoint_store=store)
    coordinator.event_completed(0,bp.events[0])
    assert not store.path.exists()
    coordinator.event_completed(1,bp.events[1])
    assert store.load_latest().status == "RUNNING"
    coordinator.stop(1)
    assert store.load_latest().status == "STOPPED"
    assert manager.sequence == 2


def test_checkpoint_coordinator_accepts_typed_word_state_snapshot(tmp_path):
    from word_replica.domain.reconstruction import WordStateSnapshot
    from word_replica.interactive.checkpoints import InteractiveCheckpointCoordinator
    from word_replica.services.audit import AuditLog
    from word_replica.services.checkpoints import CheckpointManager
    source=tmp_path/"source.docx"; source.write_bytes(b"source")
    output=tmp_path/"partial.docx"; bp=blueprint()
    class Renderer:
        def save(self,path): Path(path).write_bytes(b"saved")
        def current_state_snapshot(self): return WordStateSnapshot("doc","body",5,5,0,None,None,0)
    coordinator=InteractiveCheckpointCoordinator(project_id="p", source_path=source, blueprint=bp, output_path=output,
        settings={"checkpoint_event_interval":1}, renderer=Renderer(),
        save_manager=CheckpointManager(tmp_path/"h.jsonl",AuditLog(tmp_path/"a.jsonl")),
        checkpoint_store=InteractiveCheckpointStore(tmp_path/"cp.json"))
    coordinator.event_completed(0,bp.events[0])
    assert coordinator.checkpoint_store.load_latest().range_start == 5


def test_interval_checkpoint_skips_non_resumable_atomic_structure(tmp_path):
    from word_replica.interactive.checkpoints import InteractiveCheckpointCoordinator
    from word_replica.services.audit import AuditLog
    from word_replica.services.checkpoints import CheckpointManager
    source=tmp_path/"source.docx"; source.write_bytes(b"source")
    output=tmp_path/"partial.docx"
    bp=ReconstructionBlueprint.build(source_sha256=sha256_file(source), source_model_fingerprint="m", events=(
        ReconstructionEvent("MergeCells","t",{"row_start":1,"column_start":1,"row_end":1,"column_end":2}),
        ReconstructionEvent("EnterCell","c",{"row":1,"column":1}),
    ))
    class Renderer:
        safe=False
        def is_restart_safe(self): return self.safe
        def save(self,path): Path(path).write_bytes(b"saved")
        def current_state_snapshot(self): return {"story":"body","range_start":1,"range_end":1}
    renderer=Renderer(); manager=CheckpointManager(tmp_path/"h.jsonl",AuditLog(tmp_path/"a.jsonl")); store=InteractiveCheckpointStore(tmp_path/"cp.json")
    coordinator=InteractiveCheckpointCoordinator(project_id="p",source_path=source,blueprint=bp,output_path=output,
        settings={"checkpoint_event_interval":1},renderer=renderer,save_manager=manager,checkpoint_store=store)
    coordinator.event_completed(0,bp.events[0])
    assert not store.path.exists()
    renderer.safe=True
    coordinator.event_completed(1,bp.events[1])
    assert store.load_latest().last_completed_event_index == 1


def test_nested_table_context_is_not_restart_safe_until_inner_table_exits():
    from types import SimpleNamespace
    from word_replica.renderers.interactive_word import InteractiveWordController

    controller = InteractiveWordController.for_testing(active_range=SimpleNamespace())
    controller._table_stack = [
        {"structure_complete": True, "element_id": "outer"},
        {"structure_complete": True, "element_id": "inner"},
    ]
    assert controller.is_restart_safe() is False
    controller._table_stack.pop()
    assert controller.is_restart_safe() is True
