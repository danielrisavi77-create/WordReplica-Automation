from __future__ import annotations

from dataclasses import asdict
import json
from pathlib import Path
from typing import Callable

from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import InteractiveFidelity, InteractiveSpeedMode, MetadataMode, ReconstructionMode, RunStatus
from word_replica.domain.reconstruction import PreparedInteractiveRun
from word_replica.domain.results import RunResult, WarningItem
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.interactive.capabilities import detect_image_extension
from word_replica.interactive.checkpoints import (
    InteractiveCheckpointCoordinator,
    InteractiveCheckpointStore,
)
from word_replica.interactive.control import InteractiveRunControl
from word_replica.interactive.preflight import analyze_preflight
from word_replica.interactive.verification import InteractiveProgressTracker, LiveVerifier
from word_replica.parser.parser import DocxParser
from word_replica.opc.properties import DocumentProperties, build_output_metadata
from word_replica.qa.policy import classify_run, run_l0_l3
from word_replica.qa.render import compare_pdfs
from word_replica.qa.report import write_qa_report
from word_replica.qa.word_render import export_docx_to_pdf_with_word
from word_replica.renderers.interactive_word import InteractiveWordController, InteractiveWordRenderer
from word_replica.services.audit import AuditLog
from word_replica.services.checkpoints import CheckpointManager
from word_replica.services.project_store import ProjectStore
from word_replica.services.source_guard import sha256_file


class _CompositeObserver:
    def __init__(self, *observers) -> None:
        self.observers = tuple(observer for observer in observers if observer is not None)

    def event_started(self, index, event, snapshot) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_started", None)
            if callback is not None:
                callback(index, event, snapshot)

    def event_completed(self, index, event) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_completed", None)
            if callback is not None:
                callback(index, event)

    def event_finished(self, index, event, snapshot) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_finished", None)
            if callback is not None:
                callback(index, event, snapshot)

    def event_failed(self, index, event, snapshot, exc) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_failed", None)
            if callback is not None:
                callback(index, event, snapshot, exc)

    @property
    def halt_status(self):
        for observer in self.observers:
            status = getattr(observer, "halt_status", None)
            if status:
                return status
        return None

    def state_mismatch(self, index, event, expected, actual) -> None:
        for observer in self.observers:
            callback = getattr(observer, "state_mismatch", None)
            if callback is not None:
                callback(index, event, expected, actual)


class _InteractiveServiceObserver:
    SAFE_VERIFY_EVENTS = {"EndTable", "SetImageZOrder", "EndSection"}

    def __init__(self, blueprint, control, audit, downstream=None, *, live_verifier=None,
                 source_model=None, output_path=None, coordinator=None, settings=None) -> None:
        self.blueprint = blueprint
        self.tracker = InteractiveProgressTracker(blueprint)
        self.control = control
        self.audit = audit
        self.downstream = downstream
        self.live_verifier = live_verifier
        self.source_model = source_model
        self.output_path = Path(output_path) if output_path is not None else None
        self.coordinator = coordinator
        self.settings = dict(settings or {})
        self.halt_status = None
        self.halt_reasons: tuple[str, ...] = ()
        self._current_story = "body"
        self._expected_body_text: list[str] = []

    def _track_expected_body_text(self, event) -> None:
        et = event.event_type
        if et == "BeginHeader": self._current_story = "header"
        elif et == "BeginFooter": self._current_story = "footer"
        elif et == "BeginFootnoteStory": self._current_story = "footnote"
        elif et == "BeginEndnoteStory": self._current_story = "endnote"
        elif et in {"EndHeader", "EndFooter", "EndFootnoteStory", "EndEndnoteStory"}: self._current_story = "body"
        elif et == "InsertCharacter" and self._current_story == "body":
            self._expected_body_text.append(str(event.payload.get("character", "")))
        elif et == "InsertTab" and self._current_story == "body": self._expected_body_text.append("\t")
        elif et in {"InsertLineBreak", "InsertPageBreak"} and self._current_story == "body": self._expected_body_text.append("\n")
        elif et == "CreateField" and self._current_story == "body":
            self._expected_body_text.append(str(event.payload.get("cached_result", "")))

    def _verify_safe_boundary(self, index, event) -> None:
        if not self.settings.get("verify_during_run", False): return
        if event.event_type not in self.SAFE_VERIFY_EVENTS: return
        if self.live_verifier is None or self.output_path is None or self.source_model is None: return
        # Verification reads the saved DOCX. Force a truthful save when the configured
        # checkpoint policy did not already save this exact boundary.
        if self.coordinator is not None and self.coordinator.last_checkpoint_event_index != index:
            checkpoint = self.coordinator.milestone(index, "live_verification_boundary")
            self.tracker.checkpoint_saved(index, checkpoint.timestamp_utc)
        boundary = {
            "name": f"{event.event_type}:{event.source_element_id}",
            "expected_text_prefix": "".join(self._expected_body_text),
            "completed_tables": self.tracker.completed_tables,
        }
        result = self.live_verifier.verify_boundary(self.source_model, self.output_path, boundary)
        self.tracker.verification_updated(result.status)
        self.audit.append("INTERACTIVE_LIVE_VERIFICATION", {
            "event_index": index, "boundary": boundary["name"], "status": result.status,
            "passed": result.passed, "reasons": list(result.reasons),
        })
        if not result.passed:
            self.halt_status = "PAUSED_FIDELITY_MISMATCH"
            self.halt_reasons = tuple(result.reasons)
            self.control.pause()

    def prime_through(self, last_completed_index: int) -> None:
        if last_completed_index < 0:
            return
        end = min(last_completed_index, self.blueprint.total_events - 1)
        for index in range(0, end + 1):
            event = self.blueprint.events[index]
            self._track_expected_body_text(event)
            self.tracker.event_completed(index, event)

    def event_started(self, index, event, snapshot) -> None:
        callback = getattr(self.downstream, "event_started", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, snapshot)

    def event_failed(self, index, event, snapshot, exc) -> None:
        callback = getattr(self.downstream, "event_failed", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, snapshot, exc)

    def event_finished(self, index, event, snapshot) -> None:
        callback = getattr(self.downstream, "event_finished", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, snapshot)

    def event_completed(self, index, event) -> None:
        self._track_expected_body_text(event)
        self.tracker.event_completed(index, event)
        self._verify_safe_boundary(index, event)
        callback = getattr(self.downstream, "event_completed", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event)
        progress = getattr(self.downstream, "progress", None) if self.downstream is not None else None
        if progress is not None:
            progress(self.tracker.snapshot(state=self.control.state))

    def state_mismatch(self, index, event, expected, actual) -> None:
        from dataclasses import asdict, is_dataclass
        left = asdict(expected) if is_dataclass(expected) else dict(expected)
        right = asdict(actual) if is_dataclass(actual) else dict(actual)
        self.tracker.verification_updated("WORD_STATE_MISMATCH")
        self.audit.append("WORD_STATE_MISMATCH", {
            "event_index": index, "event_type": event.event_type,
            "source_element_id": event.source_element_id,
            "expected": left, "actual": right,
        })
        callback = getattr(self.downstream, "state_mismatch", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, expected, actual)
        progress = getattr(self.downstream, "progress", None) if self.downstream is not None else None
        if progress is not None:
            progress(self.tracker.snapshot(state=self.control.state))


class InteractiveRebuildService:
    def __init__(
        self,
        *,
        parser=None,
        word_probe: Callable[[], bool] | None = None,
        font_probe=None,
        project_store: ProjectStore | None = None,
        pdf_exporter=None,
        pdf_comparer=None,
        live_verifier=None,
    ) -> None:
        self.parser = parser or DocxParser()
        self.word_probe = word_probe
        self.font_probe = font_probe
        self.project_store = project_store
        self.pdf_exporter = pdf_exporter or export_docx_to_pdf_with_word
        self.pdf_comparer = pdf_comparer or compare_pdfs
        self.live_verifier = live_verifier or LiveVerifier(self.parser)

    def _parse_model_for_options(self, source: Path, options: RebuildOptions):
        # Output metadata policy is deliberately not part of the canonical source
        # fingerprint or reconstruction blueprint. It is applied after compilation.
        return self.parser.parse(Path(source)) if hasattr(self.parser, "parse") else self.parser(Path(source))

    @staticmethod
    def _apply_metadata_policy(model, options: RebuildOptions) -> None:
        source_props = model.extras.get("source_properties", DocumentProperties())
        model.extras["metadata_policy"] = build_output_metadata(
            source_props, options.metadata, options.preserve_author_fields, options.custom_metadata_allowlist
        )

    def prepare(self, source: Path, options: RebuildOptions, paths, audit, observer=None) -> PreparedInteractiveRun:
        model = self._parse_model_for_options(Path(source), options)
        blueprint = BlueprintCompiler().compile(model)
        self._apply_metadata_policy(model, options)
        kwargs = {"font_probe": self.font_probe}
        if self.word_probe is not None:
            kwargs["word_probe"] = self.word_probe
        preflight = analyze_preflight(model, blueprint, options.interactive, **kwargs)
        blueprint_payload = {
            "schema_version": blueprint.schema_version,
            "source_sha256": blueprint.source_sha256,
            "source_model_fingerprint": blueprint.source_model_fingerprint,
            "fingerprint": blueprint.fingerprint,
            "total_events": blueprint.total_events,
            "total_visible_characters": blueprint.total_visible_characters,
            "semantic_counts": blueprint.semantic_counts,
            "events": [asdict(event) for event in blueprint.events],
        }
        (paths.working_dir / "blueprint.json").write_text(
            json.dumps(blueprint_payload, indent=2, ensure_ascii=False, default=str), encoding="utf-8"
        )
        (paths.logs_dir / "preflight.json").write_text(
            json.dumps(asdict(preflight), indent=2, ensure_ascii=False, default=str), encoding="utf-8"
        )
        audit.append("BLUEPRINT_COMPILED", {"fingerprint": blueprint.fingerprint, "events": blueprint.total_events})
        audit.append("PREFLIGHT_COMPLETED", {"can_proceed": preflight.can_proceed, "blocking_reasons": preflight.blocking_reasons})
        callback = getattr(observer, "preflight", None) if observer is not None else None
        if callback is not None:
            callback(preflight)
        return PreparedInteractiveRun(str(Path(source).resolve()), model, blueprint, preflight, paths, options)

    @staticmethod
    def _settings_dict(options: InteractiveOptions) -> dict:
        data = asdict(options)
        data["speed_mode"] = options.speed_mode.value
        data["fidelity"] = options.fidelity.value
        return data

    @classmethod
    def _run_settings_dict(cls, options: RebuildOptions) -> dict:
        data = cls._settings_dict(options.interactive)
        data["metadata_mode"] = options.metadata.value
        data["preserve_author_fields"] = bool(options.preserve_author_fields)
        data["custom_metadata_allowlist"] = list(options.custom_metadata_allowlist)
        return data

    @staticmethod
    def _options_from_settings(settings: dict) -> InteractiveOptions:
        allowed = set(InteractiveOptions.__dataclass_fields__)
        values = {key: value for key, value in dict(settings).items() if key in allowed}
        values["speed_mode"] = InteractiveSpeedMode(values.get("speed_mode", InteractiveSpeedMode.FAST.value))
        values["fidelity"] = InteractiveFidelity(values.get("fidelity", InteractiveFidelity.MAXIMUM.value))
        return InteractiveOptions(**values)

    @classmethod
    def _rebuild_options_from_settings(cls, settings: dict) -> RebuildOptions:
        interactive = cls._options_from_settings(settings)
        return RebuildOptions(
            reconstruction_mode=ReconstructionMode.INTERACTIVE,
            interactive=interactive,
            metadata=MetadataMode(settings.get("metadata_mode", MetadataMode.FRESH.value)),
            preserve_author_fields=bool(settings.get("preserve_author_fields", False)),
            custom_metadata_allowlist=tuple(settings.get("custom_metadata_allowlist", ()) or ()),
        )

    @staticmethod
    def _materialize_assets(model, working_dir: Path) -> dict[str, Path]:
        assets_dir = working_dir / "interactive_assets"
        assets_dir.mkdir(parents=True, exist_ok=True)
        result: dict[str, Path] = {}
        for asset_id, asset in model.assets.items():
            suffix = detect_image_extension(asset.part_name, asset) or Path(asset.part_name).suffix or ".bin"
            target = assets_dir / f"{asset_id}{suffix}"
            if not target.exists() or target.read_bytes() != asset.bytes_data:
                target.write_bytes(asset.bytes_data)
            result[asset_id] = target
        return result

    @staticmethod
    def _load_save_history(path: Path) -> list[dict]:
        if not path.exists():
            return []
        rows: list[dict] = []
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.strip():
                rows.append(json.loads(line))
        return rows

    def _finalize_qa(self, prepared, output_path: Path, save_manager, audit, renderer) -> RunResult:
        paths = prepared.paths
        # The custom property states the number of real saves after this final save completes.
        renderer.set_custom_property("WordReplicaActualSaveCount", save_manager.next_sequence)
        checkpoint_store = InteractiveCheckpointStore(paths.logs_dir / "interactive_checkpoint.json")
        coordinator = InteractiveCheckpointCoordinator(
            project_id=paths.project_id, source_path=Path(prepared.source_path), blueprint=prepared.blueprint,
            output_path=output_path, settings=self._run_settings_dict(prepared.options),
            renderer=renderer, save_manager=save_manager, checkpoint_store=checkpoint_store,
        )
        final_checkpoint = coordinator.milestone(prepared.blueprint.total_events - 1, "final_output")
        warnings: list[WarningItem] = []
        for item in prepared.preflight.capability_items:
            if item.classification.value != "RECONSTRUCTED":
                warnings.append(WarningItem(
                    code=f"CAPABILITY_{item.classification.value}", message=item.reason,
                    element_id=item.source_element_id, affects_status=True,
                ))
        try:
            rebuilt_model = self.parser.parse(output_path) if hasattr(self.parser, "parse") else self.parser(output_path)
            bundle = run_l0_l3(prepared.model, rebuilt_model)
        except Exception as exc:
            if self.project_store is not None:
                self.project_store.set_status(paths.project_id, "FAILED")
            audit.append("INTERACTIVE_FINAL_QA_FAILED", {"reason": f"output parse failed: {exc}"})
            return RunResult(RunStatus.FAIL, output_path, None, project_id=paths.project_id,
                save_count=save_manager.sequence, warnings=warnings, reasons=[f"output parse failed: {exc}"])

        source_pdf = paths.qa_dir / "l4_source.pdf"
        rebuilt_pdf = paths.qa_dir / "l4_rebuilt.pdf"
        try:
            self.pdf_exporter(Path(prepared.source_path), source_pdf, visible=False)
            self.pdf_exporter(output_path, rebuilt_pdf, visible=False)
            bundle.render = self.pdf_comparer(source_pdf, rebuilt_pdf, paths.qa_dir)
        except Exception as exc:
            warnings.append(WarningItem("L4_UNAVAILABLE", f"Controlled Word PDF comparison failed: {exc}", affects_status=True))
            bundle.render = None

        status = classify_run(bundle, warnings)
        qa_report = paths.qa_dir / "qa_report.html"
        saves = self._load_save_history(paths.logs_dir / "save_history.jsonl")
        write_qa_report(
            qa_report, status=status.value, source_sha256=prepared.blueprint.source_sha256,
            output_sha256=final_checkpoint.output_sha256, renderer="interactive-word",
            fidelity=prepared.options.interactive.fidelity.value, metadata=prepared.options.metadata.value,
            saves=saves, levels=bundle.levels, warnings=warnings, render_result=bundle.render,
        )
        reasons: list[str] = []
        for level_name in ("L0", "L1"):
            level = bundle.levels.get(level_name)
            if level is not None and not level.passed:
                for finding in level.findings[:10]:
                    reasons.append(f"{getattr(finding, 'code', level_name)} at {getattr(finding, 'path', '/')}")
        project_status = "FAILED" if status is RunStatus.FAIL else "COMPLETED"
        if self.project_store is not None:
            self.project_store.set_status(paths.project_id, project_status)
        audit.append("INTERACTIVE_FINAL_QA_COMPLETED", {
            "status": status.value, "qa_report": str(qa_report),
            "l4_available": bundle.render is not None,
            "l4_within_tolerance": getattr(bundle.render, "within_tolerance", None),
        })
        return RunResult(status, output_path, qa_report, project_id=paths.project_id,
            save_count=save_manager.sequence, warnings=warnings, reasons=reasons)

    def _execute(
        self,
        prepared: PreparedInteractiveRun,
        *,
        controller_factory,
        control,
        observer=None,
        resume_checkpoint=None,
    ) -> RunResult:
        paths = prepared.paths
        audit = AuditLog(paths.logs_dir / "audit.jsonl")
        settings = self._run_settings_dict(prepared.options)
        output_path = Path(resume_checkpoint.output_path) if resume_checkpoint is not None else (
            paths.output_dir / f"{Path(prepared.source_path).stem}_reconstructed.docx"
        )
        renderer = InteractiveWordRenderer(
            controller=controller_factory(),
            options=prepared.options.interactive,
        )
        assets = self._materialize_assets(prepared.model, paths.working_dir)
        renderer.controller.set_asset_resolver(lambda asset_id: assets[asset_id])
        save_manager = CheckpointManager(paths.logs_dir / "save_history.jsonl", audit)
        if resume_checkpoint is not None:
            save_manager.sequence = int(resume_checkpoint.save_sequence)
        checkpoint_store = InteractiveCheckpointStore(paths.logs_dir / "interactive_checkpoint.json")
        coordinator = InteractiveCheckpointCoordinator(
            project_id=paths.project_id,
            source_path=Path(prepared.source_path),
            blueprint=prepared.blueprint,
            output_path=output_path,
            settings=settings,
            renderer=renderer,
            save_manager=save_manager,
            checkpoint_store=checkpoint_store,
        )
        service_observer = _InteractiveServiceObserver(
            prepared.blueprint, control, audit, observer,
            live_verifier=self.live_verifier, source_model=prepared.model,
            output_path=output_path, coordinator=coordinator, settings=settings,
        )
        composite = _CompositeObserver(coordinator, service_observer)
        if self.project_store is not None:
            self.project_store.set_status(paths.project_id, "RUNNING")
        try:
            if resume_checkpoint is None:
                renderer.open_blank()
                start_index = 0
            else:
                renderer.open_existing(output_path)
                renderer.restore_checkpoint_state(resume_checkpoint)
                start_index = resume_checkpoint.last_completed_event_index + 1
            renderer.apply_metadata(prepared.model.extras.get("metadata_policy", {}))
            if start_index > 0:
                service_observer.prime_through(start_index - 1)
            audit.append("INTERACTIVE_EXECUTION_STARTED", {"start_index": start_index})
            outcome = renderer.execute_blueprint(
                prepared.blueprint, control, composite, start_index=start_index
            )
            audit.append("INTERACTIVE_METRICS", {
                "completed_events": service_observer.tracker.completed_events,
                "completed_characters": service_observer.tracker.completed_characters,
                "completed_tables": service_observer.tracker.completed_tables,
                "completed_images": service_observer.tracker.completed_images,
                "completed_sections": service_observer.tracker.completed_sections,
                "total_events": prepared.blueprint.total_events,
                "total_characters": prepared.blueprint.total_visible_characters,
                "outcome": outcome.status,
            })
            if outcome.status == "PAUSED_FIDELITY_MISMATCH":
                if coordinator.last_checkpoint_event_index != outcome.last_completed_index:
                    coordinator.milestone(outcome.last_completed_index, "fidelity_mismatch")
                if self.project_store is not None:
                    self.project_store.set_status(paths.project_id, "PAUSED")
                reasons = list(service_observer.halt_reasons) or ["Live fidelity verification mismatch"]
                audit.append("INTERACTIVE_EXECUTION_PAUSED_FIDELITY", {
                    "last_completed_event_index": outcome.last_completed_index, "reasons": reasons,
                })
                return RunResult(
                    status=RunStatus.WARN, output_path=output_path, qa_report_path=None,
                    project_id=paths.project_id, save_count=save_manager.sequence, reasons=reasons,
                )
            if outcome.status == "PAUSED_STATE_MISMATCH":
                coordinator.milestone(outcome.last_completed_index, "word_state_mismatch")
                if self.project_store is not None:
                    self.project_store.set_status(paths.project_id, "PAUSED")
                return RunResult(
                    status=RunStatus.WARN, output_path=output_path, qa_report_path=None,
                    project_id=paths.project_id, save_count=save_manager.sequence,
                    reasons=["Word state mismatch; reconstruction paused before the next event"],
                )
            if outcome.status == "STOPPED":
                coordinator.stop(outcome.last_completed_index)
                if self.project_store is not None:
                    self.project_store.set_status(paths.project_id, "STOPPED")
                audit.append("INTERACTIVE_EXECUTION_STOPPED", {"last_completed_event_index": outcome.last_completed_index})
                return RunResult(
                    status=RunStatus.WARN,
                    output_path=output_path,
                    qa_report_path=None,
                    project_id=paths.project_id,
                    save_count=save_manager.sequence,
                    reasons=["Interactive reconstruction stopped safely; resume checkpoint saved"],
                )
            if self.project_store is not None:
                self.project_store.set_status(paths.project_id, "VERIFYING")
            audit.append("INTERACTIVE_EXECUTION_COMPLETED", {"last_completed_event_index": outcome.last_completed_index})
            return self._finalize_qa(prepared, output_path, save_manager, audit, renderer)
        finally:
            renderer.close()

    def start(self, prepared: PreparedInteractiveRun, controller_factory=None, *, control=None, observer=None) -> RunResult:
        if not prepared.preflight.can_proceed:
            return RunResult(
                status=RunStatus.FAIL,
                output_path=None,
                qa_report_path=None,
                project_id=prepared.paths.project_id,
                reasons=list(prepared.preflight.blocking_reasons),
            )
        controller_factory = controller_factory or InteractiveWordController
        control = control or InteractiveRunControl()
        if control.state.value == "CREATED":
            control.start()
        return self._execute(
            prepared,
            controller_factory=controller_factory,
            control=control,
            observer=observer,
        )

    def resume(self, project_id: str, control, observer=None, *, controller_factory=None) -> RunResult:
        if self.project_store is None:
            raise RuntimeError("resume requires a ProjectStore")
        paths = self.project_store.paths_for_project(project_id)
        project = self.project_store.get_project(project_id)
        source = Path(project["source_path"])
        checkpoint_store = InteractiveCheckpointStore(paths.logs_dir / "interactive_checkpoint.json")
        try:
            checkpoint = checkpoint_store.load_latest()
        except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=[f"checkpoint unreadable: {exc}"])
        if not source.exists() or sha256_file(source) != checkpoint.source_sha256:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=["source hash mismatch"])
        options = self._rebuild_options_from_settings(checkpoint.settings)
        interactive = options.interactive
        model = self._parse_model_for_options(source, options)
        blueprint = BlueprintCompiler().compile(model)
        validation = checkpoint_store.validate_resume(checkpoint, source_path=source, blueprint=blueprint)
        if not validation.valid:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=[validation.reason])
        self._apply_metadata_policy(model, options)
        # Resume has already validated the saved preflight contract; re-run environment checks before Word opens.
        kwargs = {"font_probe": self.font_probe}
        if self.word_probe is not None:
            kwargs["word_probe"] = self.word_probe
        preflight = analyze_preflight(model, blueprint, interactive, **kwargs)
        if not preflight.can_proceed:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=list(preflight.blocking_reasons))
        prepared = PreparedInteractiveRun(str(source.resolve()), model, blueprint, preflight, paths, options)
        controller_factory = controller_factory or InteractiveWordController
        return self._execute(
            prepared,
            controller_factory=controller_factory,
            control=control,
            observer=observer,
            resume_checkpoint=checkpoint,
        )
