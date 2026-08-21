from __future__ import annotations

from dataclasses import asdict
import json
from pathlib import Path
import shutil
import tempfile
import time
import sys
from typing import Callable

from word_replica.config import RebuildOptions
from word_replica.domain.enums import ReconstructionMode, RendererChoice, RunStatus, VisibilityMode
from word_replica.domain.errors import CriticalRebuildError, RendererUnavailableError, SourceIntegrityError
from word_replica.domain.model import DocumentModel
from word_replica.domain.results import RunResult, WarningItem
from word_replica.opc.properties import DocumentProperties, build_output_metadata
from word_replica.parser.fidelity import project_fidelity
from word_replica.parser.parser import DocxParser
from word_replica.qa.environment import capture_environment_fingerprint
from word_replica.qa.policy import classify_run, run_l0_l3
from word_replica.qa.report import write_qa_report
from word_replica.renderers.base import Renderer
from word_replica.renderers.pure_docx import PureDocxRenderer
from word_replica.renderers.word_com import WordComRenderer, choose_renderer_name, word_available
from word_replica.services.audit import AuditLog
from word_replica.services.checkpoints import CheckpointManager, SaveEvent
from word_replica.services.project_store import ProjectPaths, ProjectStore
from word_replica.services.source_guard import (
    SourceSnapshot,
    assert_source_unchanged,
    capture_source,
    sha256_file,
)


class RenderContext:
    def __init__(
        self,
        project_id: str,
        options: RebuildOptions,
        checkpoints: CheckpointManager,
    ) -> None:
        self.project_id = project_id
        self.options = options
        self.checkpoints = checkpoints
        self.last_save_monotonic = time.monotonic()
        self.current_content_fingerprint: str | None = None
        self.last_saved_content_fingerprint: str | None = None
        self.final_sealed = False

    def mark_content_changed(self, fingerprint: str) -> None:
        self.current_content_fingerprint = fingerprint

    def checkpoint(
        self,
        reason: str,
        stage: str,
        save_callable,
        path: Path,
    ) -> SaveEvent:
        if self.final_sealed:
            raise CriticalRebuildError("No save is permitted after final metadata seal")
        event = self.checkpoints.save(reason, stage, save_callable, path)
        self.last_save_monotonic = time.monotonic()
        self.last_saved_content_fingerprint = self.current_content_fingerprint
        return event

    def maybe_periodic_save(self, stage: str, save_callable, path: Path) -> None:
        now = time.monotonic()
        if now - self.last_save_monotonic < self.options.periodic_save_seconds:
            return
        if self.current_content_fingerprint == self.last_saved_content_fingerprint:
            return
        self.checkpoint("periodic save", stage, save_callable, path)

    def final_seal(self, renderer: Renderer, output_path: Path) -> SaveEvent:
        if self.final_sealed:
            raise CriticalRebuildError("Renderer attempted more than one final metadata seal")
        next_sequence = self.checkpoints.next_sequence
        renderer.set_custom_property("WordReplicaProjectId", self.project_id)
        renderer.set_custom_property("WordReplicaReconstructed", "true")
        renderer.set_custom_property("WordReplicaActualSaveCount", str(next_sequence))
        event = self.checkpoint(
            "final metadata/save-count seal",
            "final",
            lambda: renderer.save(output_path),
            output_path,
        )
        self.final_sealed = True
        return event


def _prepare_metadata_policy(model: DocumentModel, options: RebuildOptions) -> None:
    source_props = model.extras.get("source_properties", DocumentProperties())
    model.extras["metadata_policy"] = build_output_metadata(
        source_props,
        options.metadata,
        options.preserve_author_fields,
        options.custom_metadata_allowlist,
    )


def _resolve_output_path(
    source: Path,
    paths: ProjectPaths,
    options: RebuildOptions,
) -> Path:
    if not options.allow_source_overwrite:
        return paths.output_dir / f"{source.stem}_reconstructed.docx"
    backup = paths.backups_dir / f"{source.stem}_before_overwrite.docx"
    shutil.copy2(source, backup)
    if sha256_file(backup) != sha256_file(source):
        raise SourceIntegrityError(
            "Verified backup could not be created before source overwrite"
        )
    return source


def _final_source_check(
    snapshot: SourceSnapshot,
    output_path: Path,
    options: RebuildOptions,
    paths: ProjectPaths,
) -> None:
    if not options.allow_source_overwrite:
        assert_source_unchanged(snapshot)
        return
    backup = paths.backups_dir / f"{snapshot.path.stem}_before_overwrite.docx"
    if not backup.exists() or sha256_file(backup) != snapshot.sha256:
        raise SourceIntegrityError(
            "Original backup no longer matches the pre-run source hash"
        )
    if output_path.resolve() != snapshot.path.resolve():
        raise SourceIntegrityError("Overwrite mode did not target the original source path")


class RebuildService:
    @classmethod
    def default(cls) -> "RebuildService":
        return cls()

    @classmethod
    def default_for_tests(cls) -> "RebuildService":
        return cls(app_root=Path(tempfile.mkdtemp(prefix="word-replica-tests-")))

    def __init__(
        self,
        app_root: Path | None = None,
        parser=None,
        renderer=None,
        qa: Callable | None = None,
        interactive_service=None,
        projects_under_app_root: bool = False,
    ) -> None:
        self.app_root = app_root
        self._parser = parser
        self._renderer = renderer
        self._qa = qa
        self._interactive_service = interactive_service
        # Passed through to ProjectStore: a caller reading a corpus it does not
        # own needs its projects kept out of that corpus's tree.
        self.projects_under_app_root = projects_under_app_root

    @classmethod
    def for_testing(
        cls,
        app_root: Path,
        parser,
        renderer,
        qa: Callable,
    ) -> "RebuildService":
        return cls(app_root=app_root, parser=parser, renderer=renderer, qa=qa)

    @staticmethod
    def _parse(parser, source: Path):
        if callable(parser) and not hasattr(parser, "parse"):
            return parser(source)
        return parser.parse(source)

    @staticmethod
    def _write_warnings(path: Path, warnings: list[WarningItem]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps([asdict(warning) for warning in warnings], indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def _select_renderer(self, options: RebuildOptions):
        if self._renderer is not None:
            return self._renderer, "injected"
        if options.renderer is RendererChoice.WORD:
            # Explicit Word mode should not start and quit a separate probe Word
            # process before the real renderer. The renderer's open step is the
            # definitive availability/compatibility gate.
            if sys.platform != "win32":
                raise RendererUnavailableError("Microsoft Word desktop automation is unavailable")
            return WordComRenderer(visible=options.visibility is VisibilityMode.VISIBLE), "word"
        if options.renderer is RendererChoice.DOCX:
            return PureDocxRenderer(), "docx"
        name = choose_renderer_name(options.renderer, word_available())
        if name == "word":
            return WordComRenderer(visible=options.visibility is VisibilityMode.VISIBLE), "word"
        return PureDocxRenderer(), "docx"

    def resume_interactive(self, project_id: str, *, interactive_control=None, interactive_observer=None) -> RunResult:
        from word_replica.interactive.control import InteractiveRunControl
        from word_replica.services.interactive_rebuild import InteractiveRebuildService

        store = ProjectStore(
            app_root=self.app_root,
            projects_under_app_root=self.projects_under_app_root,
        )
        interactive = self._interactive_service or InteractiveRebuildService(
            parser=self._parser or DocxParser(), project_store=store
        )
        if getattr(interactive, "project_store", None) is None:
            try:
                interactive.project_store = store
            except Exception:
                pass
        control = interactive_control or InteractiveRunControl()
        if interactive_control is None:
            control.start()
        return interactive.resume(project_id, control, observer=interactive_observer)

    def rebuild(self, source: Path, options: RebuildOptions, *, interactive_control=None, interactive_observer=None) -> RunResult:
        source = Path(source).resolve()
        if source.suffix.lower() != ".docx":
            return RunResult(
                status=RunStatus.FAIL,
                output_path=None,
                qa_report_path=None,
                reasons=["Word Replica v1 accepts .docx input only"],
            )

        store = ProjectStore(
            app_root=self.app_root,
            projects_under_app_root=self.projects_under_app_root,
        )
        snapshot = capture_source(source)
        paths = store.create_project(source, options)
        audit = AuditLog(paths.logs_dir / "audit.jsonl")
        warnings_path = paths.logs_dir / "warnings.json"
        warnings: list[WarningItem] = []
        output_path: Path | None = None
        checkpoints = CheckpointManager(paths.logs_dir / "save_history.jsonl", audit)

        audit.append("PROJECT_CREATED", {"project_id": paths.project_id})
        audit.append(
            "SOURCE_HASHED",
            {"path": str(snapshot.path), "sha256": snapshot.sha256, "size": snapshot.size},
        )
        if options.allow_source_overwrite:
            audit.append("SOURCE_OVERWRITE_ENABLED", {"source_path": str(source)})

        try:
            if options.reconstruction_mode is ReconstructionMode.INTERACTIVE:
                from word_replica.services.interactive_rebuild import InteractiveRebuildService
                if options.allow_source_overwrite:
                    raise CriticalRebuildError("Interactive reconstruction always preserves the source document; source overwrite is unavailable")
                interactive = self._interactive_service or InteractiveRebuildService(
                    parser=self._parser or DocxParser(), project_store=store
                )
                if getattr(interactive, "project_store", None) is None:
                    try:
                        interactive.project_store = store
                    except Exception:
                        pass
                prepared = interactive.prepare(source, options, paths, audit, observer=interactive_observer)
                result = interactive.start(
                    prepared, control=interactive_control, observer=interactive_observer
                )
                assert_source_unchanged(snapshot)
                audit.append("SOURCE_RECHECKED", {"sha256": snapshot.sha256})
                return result

            store.set_status(paths.project_id, "RUNNING")
            audit.append("PARSE_STARTED", {"source_path": str(source)})
            parser = self._parser or DocxParser()
            model = self._parse(parser, source)
            audit.append("PARSE_COMPLETED", {})

            if isinstance(model, DocumentModel):
                _prepare_metadata_policy(model, options)
                model = project_fidelity(model, options.fidelity)
            audit.append("FIDELITY_PROJECTED", {"mode": options.fidelity.value})

            renderer, renderer_name = self._select_renderer(options)
            audit.append("RENDERER_SELECTED", {"renderer": renderer_name})
            output_path = _resolve_output_path(source, paths, options)
            if options.allow_source_overwrite:
                backup = paths.backups_dir / f"{source.stem}_before_overwrite.docx"
                audit.append(
                    "SOURCE_BACKUP_VERIFIED",
                    {"backup_path": str(backup), "sha256": sha256_file(backup)},
                )

            context = RenderContext(paths.project_id, options, checkpoints)
            audit.append("REBUILD_STARTED", {"output_path": str(output_path)})
            render_result = renderer.render(model, output_path, context)
            warnings.extend(render_result.warnings)
            if not context.final_sealed:
                raise CriticalRebuildError(
                    "Renderer returned without performing the required final metadata seal"
                )
            audit.append(
                "REBUILD_COMPLETED",
                {
                    "output_path": str(output_path),
                    "stages_completed": render_result.stages_completed,
                    "save_count": checkpoints.sequence,
                },
            )

            _final_source_check(snapshot, output_path, options, paths)
            if options.allow_source_overwrite:
                audit.append(
                    "SOURCE_INTENTIONAL_REPLACEMENT_VERIFIED",
                    {"source_path": str(source), "backup_sha256": snapshot.sha256},
                )
            else:
                audit.append("SOURCE_RECHECKED", {"sha256": snapshot.sha256})

            audit.append("QA_STARTED", {})
            qa_bundle = None
            qa_reasons: list[str] = []
            if self._qa is not None:
                qa_status = self._qa(source, output_path, model)
                if not isinstance(qa_status, RunStatus):
                    qa_status = RunStatus(str(qa_status))
                status = qa_status
                if status is RunStatus.PASS and any(w.affects_status for w in warnings):
                    status = RunStatus.WARN
                if status is RunStatus.FAIL:
                    qa_reasons.append("QA classified the reconstructed document as FAIL")
            else:
                rebuilt_model = DocxParser().parse(output_path)
                qa_bundle = run_l0_l3(model, rebuilt_model)
                warnings.append(WarningItem(
                    code="L4_UNAVAILABLE",
                    message="Controlled PDF render comparison was not available in this run",
                    affects_status=False,
                ))
                status = classify_run(qa_bundle, warnings)
                if status is RunStatus.FAIL:
                    for level_name in ("L0", "L1"):
                        level = qa_bundle.levels.get(level_name)
                        if level is None:
                            continue
                        for finding in level.findings[:10]:
                            if hasattr(finding, "code"):
                                qa_reasons.append(
                                    f"{finding.code} at {finding.path}: expected={finding.expected!r}, actual={finding.actual!r}"
                                )
                            else:
                                qa_reasons.append(f"{level_name}: {finding}")
            audit.append("QA_COMPLETED", {"status": status.value})

            self._write_warnings(warnings_path, warnings)
            save_rows = []
            history_path = paths.logs_dir / "save_history.jsonl"
            if history_path.exists():
                save_rows = [json.loads(line) for line in history_path.read_text(encoding="utf-8").splitlines() if line]
            qa_report_path = write_qa_report(
                paths.qa_dir / "qa_report.html",
                status=status.value,
                source_sha256=snapshot.sha256,
                output_sha256=sha256_file(output_path),
                renderer=renderer_name,
                fidelity=options.fidelity.value,
                metadata=options.metadata.value,
                saves=save_rows,
                levels=qa_bundle.levels if qa_bundle is not None else {},
                warnings=warnings,
                render_result=qa_bundle.render if qa_bundle is not None else None,
                environment=capture_environment_fingerprint(),
            )
            store.set_status(paths.project_id, status.value)
            audit.append(
                "RUN_COMPLETED",
                {"status": status.value, "save_count": checkpoints.sequence, "qa_report_path": str(qa_report_path)},
            )
            return RunResult(
                status=status,
                output_path=output_path,
                qa_report_path=qa_report_path,
                project_id=paths.project_id,
                save_count=checkpoints.sequence,
                warnings=warnings,
                reasons=qa_reasons,
            )
        except Exception as exc:
            reason = str(exc) or type(exc).__name__
            self._write_warnings(warnings_path, warnings)
            store.set_status(paths.project_id, RunStatus.FAIL.value)
            audit.append(
                "RUN_COMPLETED",
                {
                    "status": RunStatus.FAIL.value,
                    "reason": reason,
                    "error_type": type(exc).__name__,
                    "save_count": checkpoints.sequence,
                },
            )
            return RunResult(
                status=RunStatus.FAIL,
                output_path=output_path if output_path and output_path.exists() else None,
                qa_report_path=None,
                project_id=paths.project_id,
                save_count=checkpoints.sequence,
                warnings=warnings,
                reasons=[reason],
            )
