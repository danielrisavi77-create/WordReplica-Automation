import json
import pytest

from scripts.remote_harness.event_trace import JsonlEventTraceObserver
from word_replica.domain.reconstruction import ExecutionOutcome, ReconstructionBlueprint, ReconstructionEvent, WordStateSnapshot
from word_replica.renderers.interactive_word import InteractiveWordRenderer


class FakeControl:
    def before_next_event(self, last_completed):
        class D: stop_requested = False
        return D()
    def pause(self): pass


class FakeController:
    def __init__(self):
        self.calls=[]; self.fail=False
    def current_state_snapshot(self):
        return WordStateSnapshot("doc", "body", 4, 4, 0, None, None, None)
    def execute_event(self, event):
        self.calls.append(event.event_type)
        if self.fail: raise RuntimeError("boom")
    def is_restart_safe(self): return True
    def close(self): pass


def _blueprint():
    e = ReconstructionEvent("InsertCharacter", "r1", {"character":"A"})
    return ReconstructionBlueprint.build(source_sha256="s", source_model_fingerprint="m", events=(e,))


def test_event_trace_writes_before_after_and_error_records(tmp_path):
    observer = JsonlEventTraceObserver(tmp_path / "event_trace.jsonl")
    event = ReconstructionEvent("InsertCharacter", "r1", {"character": "A"})
    snapshot = WordStateSnapshot("doc", "body", 4, 4, 0, None, None, None)
    observer.event_started(3, event, snapshot)
    observer.event_completed(3, event)
    observer.event_finished(3, event, snapshot)
    observer.event_failed(4, event, snapshot, RuntimeError("boom"))
    rows = [json.loads(x) for x in (tmp_path / "event_trace.jsonl").read_text().splitlines()]
    assert [r["status"] for r in rows] == ["before", "after", "error"]
    assert rows[0]["story"] == "body"
    assert rows[0]["active_range_type"] is not None
    assert "character" not in rows[0]


def test_renderer_notifies_started_and_failed_without_swallowing_exception(tmp_path):
    controller=FakeController(); controller.fail=True
    renderer=InteractiveWordRenderer(controller=controller)
    observer=JsonlEventTraceObserver(tmp_path/"trace.jsonl")
    try:
        renderer.execute_blueprint(_blueprint(), FakeControl(), observer)
    except RuntimeError as exc:
        assert str(exc)=="boom"
    else:
        raise AssertionError("expected RuntimeError")
    rows=[json.loads(x) for x in (tmp_path/"trace.jsonl").read_text().splitlines()]
    assert [r["status"] for r in rows] == ["before", "error"]


def test_service_observer_forwards_started_and_failed_to_harness_trace(tmp_path):
    from word_replica.services.interactive_rebuild import _InteractiveServiceObserver

    class Control:
        state = "RUNNING"
        def pause(self): pass

    class Audit:
        def append(self, *args, **kwargs): pass

    trace = JsonlEventTraceObserver(tmp_path / "service_trace.jsonl")
    observer = _InteractiveServiceObserver(_blueprint(), Control(), Audit(), downstream=trace)
    event = _blueprint().events[0]
    snapshot = WordStateSnapshot("doc", "body", 1, 1, 0, None, None, None)
    observer.event_started(0, event, snapshot)
    observer.event_failed(0, event, snapshot, RuntimeError("boom"))
    rows = [json.loads(x) for x in (tmp_path / "service_trace.jsonl").read_text().splitlines()]
    assert [row["status"] for row in rows] == ["before", "error"]

class CountingSnapshotController(FakeController):
    def __init__(self):
        super().__init__(); self.snapshot_calls = 0
    def current_state_snapshot(self):
        self.snapshot_calls += 1
        return super().current_state_snapshot()


def test_renderer_reuses_verified_pre_event_snapshot_for_before_trace(tmp_path):
    controller = CountingSnapshotController()
    renderer = InteractiveWordRenderer(controller=controller)
    observer = JsonlEventTraceObserver(tmp_path / "trace.jsonl")
    outcome = renderer.execute_blueprint(_blueprint(), FakeControl(), observer)
    assert outcome.status == "COMPLETED"
    # initial expected state + verified pre-event state + post-event state
    assert controller.snapshot_calls == 3


def test_renderer_reuses_post_snapshot_before_text_after_run_properties(tmp_path):
    controller = CountingSnapshotController()
    blueprint = ReconstructionBlueprint.build(
        source_sha256="s",
        source_model_fingerprint="m",
        events=(
            ReconstructionEvent("ApplyRunProperties", "r1", {"bold": True}),
            ReconstructionEvent("InsertText", "r1", {"text": "A"}),
        ),
    )
    renderer = InteractiveWordRenderer(controller=controller)
    observer = JsonlEventTraceObserver(tmp_path / "trace.jsonl")

    outcome = renderer.execute_blueprint(blueprint, FakeControl(), observer)

    assert outcome.status == "COMPLETED"
    assert controller.snapshot_calls == 4


class SnapshotRejectingController(FakeController):
    def __init__(self):
        super().__init__(); self.snapshot_calls = 0
    def current_state_snapshot(self):
        self.snapshot_calls += 1
        if self.snapshot_calls == 2:
            raise RuntimeError("snapshot exploded")
        return super().current_state_snapshot()


def test_renderer_records_snapshot_failure_with_event_context(tmp_path):
    controller = SnapshotRejectingController()
    renderer = InteractiveWordRenderer(controller=controller)
    observer = JsonlEventTraceObserver(tmp_path / "trace.jsonl")
    try:
        renderer.execute_blueprint(_blueprint(), FakeControl(), observer)
    except RuntimeError as exc:
        assert str(exc) == "snapshot exploded"
    else:
        raise AssertionError("expected snapshot failure")
    rows = [json.loads(x) for x in (tmp_path / "trace.jsonl").read_text().splitlines()]
    assert rows[-1]["status"] == "snapshot_error"
    assert rows[-1]["event_type"] == "InsertCharacter"
    assert rows[-1]["phase"] == "pre_state_check"


def _table_batch_metrics(*, success=True, failed_phase=None):
    return {
        "table_id": "t1",
        "row_count": 2,
        "cell_count": 4,
        "run_count": 5,
        "insert_seconds": 0.1,
        "convert_seconds": 0.2,
        "geometry_seconds": 0.3,
        "formatting_seconds": 0.4,
        "verification_seconds": 0.01,
        "total_seconds": 1.01,
        "success": success,
        "failed_phase": failed_phase,
    }


def test_table_batch_trace_writes_phase_profile_without_cell_text(tmp_path):
    observer = JsonlEventTraceObserver(tmp_path / "trace.jsonl")
    event = ReconstructionEvent("InsertTableBatch", "t1", {
        "text_projection": "SECRET CELL TEXT",
        "cells": ({"text": "SECRET CELL TEXT"},),
    })

    observer.table_batch_profile(7, event, _table_batch_metrics())

    row = json.loads((tmp_path / "trace.jsonl").read_text(encoding="utf-8"))
    assert row["status"] == "table_batch_profile"
    assert row["event_index"] == 7
    assert row["table_id"] == "t1"
    assert row["formatting_seconds"] == 0.4
    assert row["cell_count"] == 4
    assert "text_projection" not in row
    assert "cells" not in row
    assert "SECRET" not in json.dumps(row)


def test_renderer_emits_failed_table_profile_before_event_failed():
    event = ReconstructionEvent("InsertTableBatch", "t1", {"text_projection": "AB"})
    bp = ReconstructionBlueprint.build(
        source_sha256="s", source_model_fingerprint="m", events=(event,),
    )

    class Controller(FakeController):
        def __init__(self):
            super().__init__()
            self.fail = True
            self.metrics = _table_batch_metrics(success=False, failed_phase="convert")
        def consume_table_batch_metrics(self):
            result, self.metrics = self.metrics, None
            return result

    class Observer:
        def __init__(self): self.calls = []
        def event_started(self, *args): pass
        def table_batch_profile(self, index, observed_event, metrics):
            self.calls.append(("profile", metrics["failed_phase"]))
        def event_failed(self, *args): self.calls.append(("failed", None))

    observer = Observer()
    with pytest.raises(RuntimeError, match="boom"):
        InteractiveWordRenderer(controller=Controller()).execute_blueprint(bp, FakeControl(), observer)

    assert observer.calls == [("profile", "convert"), ("failed", None)]


def test_composite_and_service_observers_forward_table_batch_profile(tmp_path):
    from word_replica.services.interactive_rebuild import _CompositeObserver, _InteractiveServiceObserver

    class Control:
        state = "RUNNING"
        def pause(self): pass

    trace = JsonlEventTraceObserver(tmp_path / "trace.jsonl")
    service = _InteractiveServiceObserver(
        _blueprint(),
        Control(),
        type("Audit", (), {"append": lambda self, *args, **kwargs: None})(),
        downstream=trace,
    )
    event = ReconstructionEvent("InsertTableBatch", "t1", {})

    _CompositeObserver(service).table_batch_profile(4, event, _table_batch_metrics())

    row = json.loads((tmp_path / "trace.jsonl").read_text(encoding="utf-8"))
    assert row["status"] == "table_batch_profile"
    assert row["event_index"] == 4
