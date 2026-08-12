import json

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
