from pathlib import Path

from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import ReconstructionMode, RunStatus, VisibilityMode
from word_replica.services.audit import AuditLog
from word_replica.services.interactive_rebuild import InteractiveRebuildService
from word_replica.services.project_store import ProjectStore
from tests.fixtures.build_fixtures import build_plain_text


def test_blocked_preflight_never_creates_word_controller(tmp_path):
    source=build_plain_text(tmp_path/"source.docx")
    store=ProjectStore(tmp_path/"projects")
    options=RebuildOptions(reconstruction_mode=ReconstructionMode.INTERACTIVE)
    paths=store.create_project(source, options)
    audit=AuditLog(paths.logs_dir/"audit.jsonl")
    service=InteractiveRebuildService(word_probe=lambda:False)
    prepared=service.prepare(source, options, paths, audit)
    called=[]
    result=service.start(prepared, controller_factory=lambda: called.append(True))
    assert called == []
    assert result.status is RunStatus.FAIL
    assert any("Word" in reason for reason in result.reasons)
    assert (paths.working_dir/"blueprint.json").exists()
    assert (paths.logs_dir/"preflight.json").exists()


def test_ready_start_executes_blueprint_and_writes_truthful_checkpoint(tmp_path):
    from word_replica.interactive.control import InteractiveRunControl
    from word_replica.qa.render import RenderQaResult
    source=build_plain_text(tmp_path/"source.docx")
    store=ProjectStore(tmp_path/"projects")
    options=RebuildOptions(reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(checkpoint_event_interval=1))
    paths=store.create_project(source, options)
    audit=AuditLog(paths.logs_dir/"audit.jsonl")
    def exporter(docx, pdf, visible=False):
        Path(pdf).write_bytes(b"pdf")
        return Path(pdf)
    def comparer(source_pdf, rebuilt_pdf, qa_dir):
        return RenderQaResult(True, True, True, 1, 1, [], [])
    service=InteractiveRebuildService(
        word_probe=lambda:True, project_store=store,
        pdf_exporter=exporter, pdf_comparer=comparer,
    )
    prepared=service.prepare(source, options, paths, audit)

    class FakeController:
        def __init__(self): self.opened=False; self.events=[]; self.range=0
        def open_blank(self): self.opened=True
        def set_asset_resolver(self,resolver): self.resolver=resolver
        def execute_event(self,event):
            self.events.append(event.event_type)
            if event.event_type == "InsertCharacter": self.range += 1
            if event.event_type == "InsertText": self.range += len(event.payload["text"])
        def save(self,path):
            import shutil
            shutil.copy2(source, path)
        def current_state_snapshot(self): return {"story":"body","range_start":self.range,"range_end":self.range,"paragraph_started":True}
        def set_custom_property(self,name,value): pass
        def close(self): pass
    controller=FakeController()
    control=InteractiveRunControl(); control.start()
    result=service.start(prepared, controller_factory=lambda:controller, control=control)
    assert controller.opened is True
    assert result.output_path and result.output_path.exists()
    assert result.save_count > 0
    assert (paths.logs_dir/"interactive_checkpoint.json").exists()
    assert store.get_project(paths.project_id)["status"] in {"VERIFYING","COMPLETED"}


def test_default_controller_receives_background_visibility(tmp_path, monkeypatch):
    import shutil
    from word_replica.qa.render import RenderQaResult
    from word_replica.services import interactive_rebuild as service_module

    source = build_plain_text(tmp_path / "source.docx")
    store = ProjectStore(tmp_path / "projects")
    options = RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        visibility=VisibilityMode.BACKGROUND,
        interactive=InteractiveOptions(checkpoint_event_interval=9999),
    )
    paths = store.create_project(source, options)
    audit = AuditLog(paths.logs_dir / "audit.jsonl")
    seen_visibility = []

    class FakeController:
        def __init__(self, *, visible):
            seen_visibility.append(visible)
            self.position = 0
        def open_blank(self): pass
        def set_asset_resolver(self, resolver): self.resolver = resolver
        def execute_event(self, event):
            if event.event_type == "InsertText":
                self.position += len(event.payload["text"])
        def save(self, path): shutil.copy2(source, path)
        def current_state_snapshot(self):
            return {"story": "body", "range_start": self.position, "range_end": self.position,
                    "paragraph_started": True}
        def set_custom_property(self, name, value): pass
        def close(self): pass

    monkeypatch.setattr(service_module, "InteractiveWordController", FakeController)
    def exporter(docx, pdf, visible=False):
        Path(pdf).write_bytes(b"pdf")
        return Path(pdf)

    service = InteractiveRebuildService(
        word_probe=lambda: True,
        project_store=store,
        pdf_exporter=exporter,
        pdf_comparer=lambda source_pdf, rebuilt_pdf, qa_dir: RenderQaResult(
            True, True, True, 1, 1, [], []
        ),
    )
    prepared = service.prepare(source, options, paths, audit)

    service.start(prepared)

    assert seen_visibility == [False]


def test_resume_rejects_changed_source_before_opening_word(tmp_path):
    from word_replica.interactive.checkpoints import InteractiveCheckpoint, InteractiveCheckpointStore
    from word_replica.interactive.control import InteractiveRunControl
    from word_replica.interactive.blueprint import BlueprintCompiler
    from word_replica.parser.parser import DocxParser
    from word_replica.services.source_guard import sha256_file

    source=build_plain_text(tmp_path/"source.docx")
    store=ProjectStore(tmp_path/"projects")
    options=RebuildOptions(reconstruction_mode=ReconstructionMode.INTERACTIVE)
    paths=store.create_project(source, options)
    model=DocxParser().parse(source); bp=BlueprintCompiler().compile(model)
    output=paths.output_dir/"partial.docx"; output.write_bytes(b"partial")
    cp=InteractiveCheckpoint.now(project_id=paths.project_id, source_sha256=sha256_file(source),
        source_model_fingerprint=bp.source_model_fingerprint, blueprint_fingerprint=bp.fingerprint,
        blueprint_schema_version=bp.schema_version, output_path=str(output), output_sha256=sha256_file(output),
        last_completed_event_index=0, section_element_id=None, block_element_id=None, table_element_id=None,
        cell_element_id=None, save_sequence=1, settings={"speed_mode":"fast","characters_per_second":25,
        "object_step_delay_ms":0,"fidelity":"maximum","checkpoint_after_tables":True,
        "checkpoint_after_images":True,"checkpoint_after_sections":True,"checkpoint_event_interval":500,
        "verify_during_run":True,"block_on_unsupported":True,"allow_preserved_objects":False}, status="STOPPED",
        story="body", range_start=1, range_end=1, paragraph_started=True)
    InteractiveCheckpointStore(paths.logs_dir/"interactive_checkpoint.json").write(cp)
    store.set_status(paths.project_id,"STOPPED")
    source.write_bytes(b"changed")
    called=[]
    service=InteractiveRebuildService(project_store=store, word_probe=lambda:True)
    control=InteractiveRunControl(); control.start()
    result=service.resume(paths.project_id, control, controller_factory=lambda:called.append(True))
    assert result.status is RunStatus.FAIL
    assert called == []
    assert any("source hash" in r.lower() for r in result.reasons)


def test_resume_opens_partial_restores_range_and_continues_after_checkpoint(tmp_path):
    from dataclasses import asdict
    from word_replica.interactive.blueprint import BlueprintCompiler
    from word_replica.interactive.checkpoints import InteractiveCheckpoint, InteractiveCheckpointStore
    from word_replica.interactive.control import InteractiveRunControl
    from word_replica.parser.parser import DocxParser
    from word_replica.services.source_guard import sha256_file

    source=build_plain_text(tmp_path/"source.docx")
    store=ProjectStore(tmp_path/"projects")
    options=RebuildOptions(reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(checkpoint_event_interval=999))
    paths=store.create_project(source, options)
    model=DocxParser().parse(source); bp=BlueprintCompiler().compile(model)
    first_char = next(i for i,e in enumerate(bp.events) if e.event_type == "InsertText")
    output=paths.output_dir/"partial.docx"; output.write_bytes(b"partial")
    settings=asdict(options.interactive)
    settings["speed_mode"]=options.interactive.speed_mode.value; settings["fidelity"]=options.interactive.fidelity.value
    cp=InteractiveCheckpoint.now(project_id=paths.project_id, source_sha256=sha256_file(source),
        source_model_fingerprint=bp.source_model_fingerprint, blueprint_fingerprint=bp.fingerprint,
        blueprint_schema_version=bp.schema_version, output_path=str(output), output_sha256=sha256_file(output),
        last_completed_event_index=first_char, section_element_id=None, block_element_id=None, table_element_id=None,
        cell_element_id=None, save_sequence=1, settings=settings, status="STOPPED", story="body",
        range_start=7, range_end=7, paragraph_started=True)
    InteractiveCheckpointStore(paths.logs_dir/"interactive_checkpoint.json").write(cp)
    store.set_status(paths.project_id,"STOPPED")

    class FakeController:
        def __init__(self): self.opened=None; self.restored=None; self.events=[]; self.range=7
        def open_existing(self,path): self.opened=Path(path)
        def restore_checkpoint_state(self,checkpoint): self.restored=(checkpoint.range_start, checkpoint.range_end)
        def set_asset_resolver(self,resolver): self.resolver=resolver
        def execute_event(self,event): self.events.append(event)
        def save(self,path): Path(path).write_bytes(b"resumed")
        def current_state_snapshot(self): return {"story":"body","range_start":8,"range_end":8,"paragraph_started":True}
        def set_custom_property(self,name,value): pass
        def close(self): pass
    controller=FakeController()
    control=InteractiveRunControl(); control.start()
    service=InteractiveRebuildService(project_store=store, word_probe=lambda:True)
    result=service.resume(paths.project_id, control, controller_factory=lambda:controller)
    assert controller.opened == output
    assert controller.restored == (7,7)
    assert controller.events[0] == bp.events[first_char+1]
    assert result.output_path == output
    assert result.save_count >= 2


def test_completed_interactive_run_requires_final_l0_l3_and_l4_before_pass(tmp_path):
    import shutil
    from word_replica.interactive.control import InteractiveRunControl
    from word_replica.qa.render import RenderQaResult
    source=build_plain_text(tmp_path/"source.docx")
    store=ProjectStore(tmp_path/"projects")
    options=RebuildOptions(reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(checkpoint_event_interval=9999))
    paths=store.create_project(source,options); audit=AuditLog(paths.logs_dir/"audit.jsonl")
    exports=[]
    def exporter(docx,pdf,visible=False):
        exports.append(Path(docx)); Path(pdf).write_bytes(b"pdf"); return Path(pdf)
    def comparer(source_pdf,rebuilt_pdf,qa_dir):
        return RenderQaResult(True,True,True,1,1,[],[])
    service=InteractiveRebuildService(project_store=store, word_probe=lambda:True,
        pdf_exporter=exporter, pdf_comparer=comparer)
    prepared=service.prepare(source,options,paths,audit)
    class Controller:
        def __init__(self): self.props={}; self.pos=0
        def open_blank(self): pass
        def set_asset_resolver(self,resolver): pass
        def execute_event(self,event):
            if event.event_type=="InsertCharacter": self.pos+=1
            if event.event_type=="InsertText": self.pos+=len(event.payload["text"])
        def current_state_snapshot(self): return {"document_identity":"d","story":"body","range_start":self.pos,"range_end":self.pos,"section_index":0,"table_element_id":None,"cell_element_id":None}
        def save(self,path): shutil.copy2(source,path)
        def set_custom_property(self,name,value): self.props[name]=value
        def close(self): pass
    controller=Controller(); control=InteractiveRunControl(); control.start()
    result=service.start(prepared,controller_factory=lambda:controller,control=control)
    assert result.status is RunStatus.PASS
    assert result.qa_report_path and result.qa_report_path.exists()
    assert exports == [source.resolve(), result.output_path]
    assert store.get_project(paths.project_id)["status"] == "COMPLETED"
    assert int(controller.props["WordReplicaActualSaveCount"]) == result.save_count
    audit_rows = [__import__("json").loads(line) for line in (paths.logs_dir/"audit.jsonl").read_text(encoding="utf-8").splitlines() if line]
    metrics = [row for row in audit_rows if row.get("event_type") == "INTERACTIVE_METRICS"]
    assert metrics
    assert metrics[-1]["payload"]["completed_characters"] == prepared.blueprint.total_visible_characters
    assert metrics[-1]["payload"]["completed_events"] == prepared.blueprint.total_events


def test_live_verification_mismatch_pauses_at_saved_table_boundary(tmp_path):
    import shutil
    from word_replica.domain.reconstruction import LiveVerificationResult
    from word_replica.interactive.control import InteractiveRunControl
    from tests.fixtures.build_fixtures import build_tables_merged

    source = build_tables_merged(tmp_path / "table.docx")
    store = ProjectStore(tmp_path / "projects")
    options = RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(
            verify_during_run=True,
            checkpoint_after_tables=True,
            checkpoint_event_interval=99999,
        ),
    )
    paths = store.create_project(source, options)
    audit = AuditLog(paths.logs_dir / "audit.jsonl")

    class FailingVerifier:
        def __init__(self): self.calls = []
        def verify_boundary(self, source_model, output_path, boundary):
            self.calls.append((Path(output_path), dict(boundary)))
            return LiveVerificationResult(
                "PAUSED_FIDELITY_MISMATCH", str(boundary["name"]), False, ("table count mismatch",)
            )

    verifier = FailingVerifier()
    service = InteractiveRebuildService(
        word_probe=lambda: True,
        project_store=store,
        live_verifier=verifier,
    )
    prepared = service.prepare(source, options, paths, audit)

    class Controller:
        def __init__(self): self.events=[]; self.pos=0
        def open_blank(self): pass
        def set_asset_resolver(self, resolver): pass
        def execute_event(self, event):
            self.events.append(event.event_type)
            if event.event_type == "InsertCharacter": self.pos += 1
            if event.event_type == "InsertText": self.pos += len(event.payload["text"])
        def current_state_snapshot(self):
            return {"document_identity":"d","story":"body","range_start":self.pos,"range_end":self.pos,
                    "section_index":0,"table_element_id":None,"cell_element_id":None}
        def save(self, path): shutil.copy2(source, path)
        def set_custom_property(self, name, value): pass
        def close(self): pass

    controller=Controller(); control=InteractiveRunControl(); control.start()
    result=service.start(prepared, controller_factory=lambda:controller, control=control)
    assert result.status is RunStatus.WARN
    assert verifier.calls
    assert control.state.value == "PAUSED"
    assert any("fidelity" in reason.lower() or "table count mismatch" in reason.lower() for reason in result.reasons)
    # The executor must return at the failing safe boundary, before the next event is emitted.
    end_table_index = next(i for i,e in enumerate(prepared.blueprint.events) if e.event_type == "EndTable")
    assert len(controller.events) == end_table_index + 1


def test_prepare_builds_interactive_metadata_policy_from_source(tmp_path):
    from docx import Document
    from word_replica.domain.enums import MetadataMode

    source = tmp_path / "metadata.docx"
    doc = Document(); doc.add_paragraph("Body")
    doc.core_properties.title = "Research title"
    doc.core_properties.subject = "Policy analysis"
    doc.save(source)
    store = ProjectStore(tmp_path / "projects")
    options = RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        metadata=MetadataMode.PRESERVE,
    )
    paths = store.create_project(source, options)
    prepared = InteractiveRebuildService(word_probe=lambda: True).prepare(
        source, options, paths, AuditLog(paths.logs_dir / "audit.jsonl")
    )
    assert prepared.model.extras["metadata_policy"]["title"] == "Research title"
    assert prepared.model.extras["metadata_policy"]["subject"] == "Policy analysis"


def test_resume_checkpoint_preserves_metadata_mode_settings(tmp_path):
    from word_replica.domain.enums import MetadataMode
    service = InteractiveRebuildService()
    options = RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        metadata=MetadataMode.PRESERVE,
        preserve_author_fields=True,
        custom_metadata_allowlist=("StudyId",),
    )
    settings = service._run_settings_dict(options)
    restored = service._rebuild_options_from_settings(settings)
    assert restored.metadata is MetadataMode.PRESERVE
    assert restored.preserve_author_fields is True
    assert restored.custom_metadata_allowlist == ("StudyId",)


def test_prepare_and_resume_recompute_identical_model_fingerprint_with_metadata_policy(tmp_path):
    from docx import Document
    from word_replica.domain.enums import MetadataMode
    from word_replica.interactive.blueprint import BlueprintCompiler

    source = tmp_path / "resume_metadata.docx"
    doc = Document(); doc.add_paragraph("Body"); doc.core_properties.title = "Stable title"; doc.save(source)
    store = ProjectStore(tmp_path / "projects")
    options = RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        metadata=MetadataMode.PRESERVE,
    )
    paths = store.create_project(source, options)
    service = InteractiveRebuildService(word_probe=lambda: True)
    prepared = service.prepare(source, options, paths, AuditLog(paths.logs_dir / "audit.jsonl"))
    restored_options = service._rebuild_options_from_settings(service._run_settings_dict(options))
    restored_model = service._parse_model_for_options(source, restored_options)
    restored_blueprint = BlueprintCompiler().compile(restored_model)
    assert restored_blueprint.source_model_fingerprint == prepared.blueprint.source_model_fingerprint
    assert restored_blueprint.fingerprint == prepared.blueprint.fingerprint
