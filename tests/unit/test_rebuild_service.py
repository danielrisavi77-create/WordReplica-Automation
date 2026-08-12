import json
from pathlib import Path

from word_replica.config import RebuildOptions
from word_replica.domain.enums import RunStatus
from word_replica.renderers.base import RenderResult
from word_replica.services.rebuild import RebuildService


class FakeRenderer:
    def __init__(self):
        self.properties = {}
        self.payload = b"fake-docx"

    def set_custom_property(self, name: str, value: str) -> None:
        self.properties[name] = value

    def save(self, output_path: Path) -> None:
        output_path.write_bytes(self.payload + repr(sorted(self.properties.items())).encode())

    def render(self, model, output_path: Path, context):
        context.mark_content_changed("fake:1")
        context.checkpoint(
            "document shell created",
            "shell",
            lambda: self.save(output_path),
            output_path,
        )
        context.final_seal(self, output_path)
        return RenderResult(output_path)


def _service(tmp_path: Path, renderer: FakeRenderer | None = None) -> RebuildService:
    return RebuildService.for_testing(
        tmp_path / "app",
        parser=lambda _: object(),
        renderer=renderer or FakeRenderer(),
        qa=lambda *_: RunStatus.PASS,
    )


def test_rebuild_protects_source_and_returns_status(tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")
    renderer = FakeRenderer()
    result = _service(tmp_path, renderer).rebuild(source, RebuildOptions())
    assert result.status is RunStatus.PASS
    assert source.read_bytes() == b"fixture"
    assert result.save_count == 2
    assert renderer.properties["WordReplicaActualSaveCount"] == "2"
    assert renderer.properties["WordReplicaReconstructed"] == "true"
    assert result.output_path is not None and result.output_path.exists()


def test_overwrite_mode_requires_verified_backup(tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")
    result = _service(tmp_path).rebuild(source, RebuildOptions(allow_source_overwrite=True))
    backups = list((source.parent / "source_rebuild").rglob("source_before_overwrite.docx"))
    assert result.status is RunStatus.PASS
    assert len(backups) == 1
    assert backups[0].read_bytes() == b"fixture"
    assert source.read_bytes().startswith(b"fake-docx")


def test_save_history_matches_truthful_custom_count(tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")
    result = _service(tmp_path).rebuild(source, RebuildOptions())
    histories = list((source.parent / "source_rebuild").rglob("save_history.jsonl"))
    assert len(histories) == 1
    rows = histories[0].read_text(encoding="utf-8").splitlines()
    assert len(rows) == result.save_count == 2
    warnings_files = list((source.parent / "source_rebuild").rglob("warnings.json"))
    assert json.loads(warnings_files[0].read_text(encoding="utf-8")) == []


def test_periodic_save_skips_unchanged_content_and_saves_changed_content(tmp_path, monkeypatch):
    from word_replica.services.audit import AuditLog
    from word_replica.services.checkpoints import CheckpointManager
    from word_replica.services.rebuild import RenderContext

    clock = {"now": 0.0}
    monkeypatch.setattr("word_replica.services.rebuild.time.monotonic", lambda: clock["now"])
    history = tmp_path / "history.jsonl"
    checkpoints = CheckpointManager(history, AuditLog(tmp_path / "audit.jsonl"))
    context = RenderContext("project", RebuildOptions(periodic_save_seconds=10), checkpoints)
    output = tmp_path / "out.docx"
    saves = []

    def real_save():
        saves.append(len(saves) + 1)
        output.write_bytes(f"save-{len(saves)}".encode())

    context.mark_content_changed("v1")
    context.checkpoint("first", "body", real_save, output)
    clock["now"] = 20.0
    context.maybe_periodic_save("body", real_save, output)
    assert len(saves) == 1

    context.mark_content_changed("v2")
    context.maybe_periodic_save("body", real_save, output)
    assert len(saves) == 2


def test_default_qa_writes_report_and_compares_rebuilt_output(tmp_path):
    from tests.fixtures.build_fixtures import build_core_fixture
    from word_replica.domain.enums import RendererChoice

    source = build_core_fixture(tmp_path / 'core.docx')
    result = RebuildService(app_root=tmp_path / 'app').rebuild(
        source, RebuildOptions(renderer=RendererChoice.DOCX)
    )
    assert result.status in {RunStatus.PASS, RunStatus.WARN}
    assert result.qa_report_path is not None and result.qa_report_path.exists()
    html = result.qa_report_path.read_text(encoding='utf-8')
    assert 'L0' in html and 'L1' in html and 'WordReplicaActualSaveCount' not in html


def test_qa_fail_exposes_l0_l1_findings_in_run_reasons(tmp_path):
    from tests.fixtures.build_fixtures import build_plain_text
    from word_replica.config import RebuildOptions
    from word_replica.domain.enums import RendererChoice, RunStatus
    from word_replica.services.rebuild import RebuildService

    source = build_plain_text(tmp_path / 'plain.docx')
    service = RebuildService.default_for_tests()

    # Inject a QA function returning FAIL: the service must not return an opaque reasons=[].
    service._qa = lambda source_path, output_path, model: RunStatus.FAIL
    result = service.rebuild(source, RebuildOptions(renderer=RendererChoice.DOCX))
    assert result.status is RunStatus.FAIL
    assert result.reasons


def test_explicit_word_renderer_does_not_spawn_a_probe_word_process(monkeypatch):
    import word_replica.services.rebuild as module
    from word_replica.config import RebuildOptions
    from word_replica.domain.enums import RendererChoice

    monkeypatch.setattr(module.sys, "platform", "win32")
    monkeypatch.setattr(module, "word_available", lambda: (_ for _ in ()).throw(AssertionError("probe must not run")))

    renderer, name = module.RebuildService()._select_renderer(RebuildOptions(renderer=RendererChoice.WORD))
    assert name == "word"
    assert renderer.__class__.__name__ == "WordComRenderer"


def test_interactive_mode_routes_without_selecting_instant_renderer(tmp_path, monkeypatch):
    from types import SimpleNamespace
    from word_replica.domain.enums import ReconstructionMode
    source=tmp_path/"source.docx"; source.write_bytes(b"fixture")
    calls=[]
    class InteractiveService:
        project_store=None
        def prepare(self, source, options, paths, audit, observer=None):
            calls.append(("prepare", source, options.reconstruction_mode)); return SimpleNamespace(paths=paths)
        def start(self, prepared, controller_factory=None, control=None, observer=None):
            calls.append(("start", prepared.paths.project_id));
            return __import__('word_replica.domain.results',fromlist=['RunResult']).RunResult(RunStatus.WARN,None,None,project_id=prepared.paths.project_id)
    service=RebuildService(app_root=tmp_path/"app", interactive_service=InteractiveService())
    monkeypatch.setattr(service,"_select_renderer",lambda options: (_ for _ in ()).throw(AssertionError("Instant renderer must not be selected")))
    result=service.rebuild(source,RebuildOptions(reconstruction_mode=ReconstructionMode.INTERACTIVE))
    assert result.status is RunStatus.WARN
    assert [x[0] for x in calls] == ["prepare","start"]


def test_interactive_word_unavailable_never_falls_back_to_pure_docx(tmp_path):
    from tests.fixtures.build_fixtures import build_plain_text
    from word_replica.domain.enums import ReconstructionMode
    from word_replica.services.interactive_rebuild import InteractiveRebuildService
    source=build_plain_text(tmp_path/"source.docx")
    interactive=InteractiveRebuildService(word_probe=lambda:False)
    service=RebuildService(app_root=tmp_path/"app", interactive_service=interactive)
    result=service.rebuild(source,RebuildOptions(reconstruction_mode=ReconstructionMode.INTERACTIVE))
    assert result.status is RunStatus.FAIL
    assert any("Word" in reason for reason in result.reasons)


def test_resume_interactive_routes_to_interactive_service(tmp_path):
    calls=[]
    class InteractiveService:
        project_store=None
        def resume(self,project_id,control,observer=None):
            calls.append((project_id,control,observer)); return __import__('word_replica.domain.results',fromlist=['RunResult']).RunResult(RunStatus.WARN,None,None,project_id=project_id)
    service=RebuildService(app_root=tmp_path/"app",interactive_service=InteractiveService())
    marker=object(); observer=object()
    result=service.resume_interactive("p1",interactive_control=marker,interactive_observer=observer)
    assert result.project_id == "p1"
    assert calls == [("p1",marker,observer)]
