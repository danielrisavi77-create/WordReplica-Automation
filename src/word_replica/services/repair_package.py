"""Orchestrates one Lekta repair-package run: preflight, visible rebuild of
the signed target, output copy/audit, and an original-untouched recheck.

Lekta fixer semantics never enter this module — by the time a package
reaches here, Lekta has already produced the corrected target DOCX; this
service only reconstructs *that* target visibly through the existing
interactive engine and proves the reconstruction and the untouched original
both hold up. `rebuild_service` and `gate_auditor` are injected so unit
tests never touch Microsoft Word or COM; production wiring passes the real
RebuildService and word_replica.qa.golden_audit.audit_docx_pair.
"""
from __future__ import annotations

import shutil
import time
from collections.abc import Callable
from dataclasses import replace
from datetime import datetime
from pathlib import Path

from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    InteractiveFidelity,
    InteractiveSpeedMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
    RunStatus,
    VisibilityMode,
)
from word_replica.domain.errors import RepairPackageError
from word_replica.repair_contract.binding import RepairRunBinding, RepairRunBindingStore
from word_replica.repair_contract.contract import GOLDEN_GATES
from word_replica.repair_contract.package import RepairPackageRequest, ValidatedRepairPackage, load_and_validate_package
from word_replica.repair_contract.report import RepairCompletionReport, build_completion_report
from word_replica.services.source_guard import assert_source_unchanged, sha256_file


def repair_contract_required_gates(*, signed_target: bool) -> tuple[str, ...]:
    if not signed_target:
        return GOLDEN_GATES
    return (*GOLDEN_GATES, "G10")


def default_rebuild_options() -> RebuildOptions:
    """Visible Word, maximum fidelity and speed, with the table fast path."""
    return RebuildOptions(
        renderer=RendererChoice.WORD,
        visibility=VisibilityMode.VISIBLE,
        fidelity=FidelityMode.FULL,
        metadata=MetadataMode.FRESH,
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(
            speed_mode=InteractiveSpeedMode.MAXIMUM,
            fidelity=InteractiveFidelity.MAXIMUM,
            enable_table_fast_path=True,
        ),
    )


def default_pure_docx_rebuild_options() -> RebuildOptions:
    """No Microsoft Word dependency at all: the same Pure DOCX path the
    desktop app's licence-free live preview uses. "Visible" here means the
    caller's preview_sink, not a real Word window — pair with a preview_sink
    on RepairPackageService to get the same watch-it-type experience without
    requiring an activated Word install.
    """
    return RebuildOptions(
        renderer=RendererChoice.DOCX,
        visibility=VisibilityMode.BACKGROUND,
        fidelity=FidelityMode.FULL,
        metadata=MetadataMode.FRESH,
        reconstruction_mode=ReconstructionMode.INSTANT,
    )


def _status_value(status: object) -> str:
    return str(getattr(status, "value", status))


def _default_open_and_repair_checker(output_path: Path) -> bool | None:
    from word_replica.qa.word_render import detect_open_and_repair

    return detect_open_and_repair(output_path)


def _default_fields_update_checker(output_path: Path) -> bool | None:
    from word_replica.qa.word_render import check_fields_update_equality

    return check_fields_update_equality(output_path)


def _copy_output_exclusive(source: Path, destination: Path) -> None:
    created = False
    try:
        with Path(source).open("rb") as input_stream, Path(destination).open("xb") as output_stream:
            created = True
            shutil.copyfileobj(input_stream, output_stream)
    except FileExistsError as exc:
        raise RepairPackageError("output-path-collision", str(destination)) from exc
    except Exception:
        if created:
            Path(destination).unlink(missing_ok=True)
        raise


def _is_reserved_output_name(candidate: Path, suggested_file_name: str) -> bool:
    if candidate.name == suggested_file_name:
        return True
    suggested = Path(suggested_file_name)
    if candidate.suffix != suggested.suffix:
        return False
    prefix = f"{suggested.stem} ("
    candidate_stem = candidate.stem
    if not candidate_stem.startswith(prefix) or not candidate_stem.endswith(")"):
        return False
    counter = candidate_stem[len(prefix):-1]
    return counter.isdigit() and int(counter) >= 2


def _same_file_identity_or_fail_closed(candidate: Path, protected: Path) -> bool:
    try:
        return candidate.samefile(protected)
    except OSError as exc:
        raise RepairPackageError(
            "binding-mismatch", f"destination identity check failed: {exc}"
        ) from exc


def _assert_destination_not_protected_alias(
    destination: Path, validated: ValidatedRepairPackage
) -> None:
    for protected_path in (
        validated.original_snapshot.path,
        validated.target_snapshot.path,
    ):
        if _same_file_identity_or_fail_closed(destination, protected_path):
            raise RepairPackageError("binding-mismatch", "destination_path")


def _validated_bound_destination(
    request: RepairPackageRequest,
    validated: ValidatedRepairPackage,
    binding: RepairRunBinding,
    working_output_path: Path,
) -> Path:
    destination = Path(binding.destination_path).resolve()
    output_dir = request.output_dir.resolve()
    if destination.parent != output_dir:
        raise RepairPackageError("binding-mismatch", "destination_path")
    if destination in {
        validated.original_snapshot.path.resolve(),
        validated.target_snapshot.path.resolve(),
    }:
        raise RepairPackageError("binding-mismatch", "destination_path")
    if not _is_reserved_output_name(
        destination, validated.contract.output_policy.suggested_file_name
    ):
        raise RepairPackageError("binding-mismatch", "destination_path")

    if destination.exists():
        _assert_destination_not_protected_alias(destination, validated)
        working_output_path = Path(working_output_path).resolve()
        if not destination.is_file() or not working_output_path.is_file():
            raise RepairPackageError("binding-mismatch", "working_output_path")
        if sha256_file(destination) != sha256_file(working_output_path):
            raise RepairPackageError("output-path-collision", str(destination))
    return destination


class RepairPackageService:
    def __init__(
        self,
        *,
        rebuild_service,
        gate_auditor: Callable[..., dict],
        binding_store: RepairRunBindingStore,
        clock: Callable[[], datetime],
        engine_version: str,
        rebuild_options: RebuildOptions | None = None,
        preview_sink: Callable[[str, dict], None] | None = None,
        preview_interactive_options: InteractiveOptions | None = None,
        open_and_repair_checker: Callable[[Path], bool | None] | None = None,
        fields_update_checker: Callable[[Path], bool | None] | None = None,
    ) -> None:
        self.rebuild_service = rebuild_service
        self.gate_auditor = gate_auditor
        self.binding_store = binding_store
        self.clock = clock
        self.engine_version = engine_version
        # preview_sink opts into the licence-free path: no Word COM at all,
        # "visible" comes from replaying the blueprint into preview_sink
        # instead of driving a real Word window. Without it, behavior is
        # unchanged from before this option existed (real Word, unaffected).
        self.preview_sink = preview_sink
        self.preview_interactive_options = preview_interactive_options or InteractiveOptions()
        if rebuild_options is not None:
            self.rebuild_options = rebuild_options
        elif preview_sink is not None:
            self.rebuild_options = default_pure_docx_rebuild_options()
        else:
            self.rebuild_options = default_rebuild_options()
        # Real best-effort Word probes (word_render.detect_open_and_repair /
        # check_fields_update_equality). Both only need Word to be
        # installed, not activated — opening/inspecting a document works in
        # Word's reduced-functionality mode even though live editing does
        # not, verified live against this machine's unactivated Word. They
        # return None (not verified) if Word is unavailable at all or the
        # probe otherwise fails; None still fails full_pass rather than
        # being treated as a pass.
        self.open_and_repair_checker = open_and_repair_checker or _default_open_and_repair_checker
        self.fields_update_checker = fields_update_checker or _default_fields_update_checker

    def _play_preview(self, validated: ValidatedRepairPackage, control) -> None:
        from word_replica.interactive.blueprint import compile_blueprint_from_source
        from word_replica.interactive.control import InteractiveRunControl
        from word_replica.domain.enums import InteractiveRunState
        from word_replica.interactive.text_preview import TextPreviewPlayer

        blueprint = compile_blueprint_from_source(
            validated.target_snapshot.path,
            enable_table_fast_path=self.preview_interactive_options.enable_table_fast_path,
        )
        control = control or InteractiveRunControl()
        if control.state is InteractiveRunState.CREATED:
            control.start()
        player = TextPreviewPlayer(self.preview_interactive_options, self.preview_sink)
        player.play(blueprint, control)

    def run(
        self, request: RepairPackageRequest, *, interactive_control=None, interactive_observer=None
    ) -> RepairCompletionReport:
        now = self.clock()
        validated = load_and_validate_package(request, now=now, engine_version=self.engine_version)
        started = time.monotonic()
        if self.preview_sink is not None:
            self._play_preview(validated, interactive_control)
            result = self.rebuild_service.rebuild(
                validated.target_snapshot.path,
                self.rebuild_options,
                expected_source_snapshot=validated.target_snapshot,
            )
            if result.project_id is not None and result.output_path is not None:
                self._persist_binding(
                    validated,
                    project_id=result.project_id,
                    working_output_path=result.output_path,
                )
        else:
            def bind_before_word(prepared) -> None:
                source_stem = Path(prepared.source_path).stem
                self._persist_binding(
                    validated,
                    project_id=prepared.paths.project_id,
                    working_output_path=prepared.paths.output_dir / f"{source_stem}_reconstructed.docx",
                )

            result = self.rebuild_service.rebuild(
                validated.target_snapshot.path,
                self.rebuild_options,
                interactive_control=interactive_control,
                interactive_observer=interactive_observer,
                interactive_pre_start=bind_before_word,
                expected_source_snapshot=validated.target_snapshot,
            )
        elapsed = time.monotonic() - started
        return self._finish(validated, result, timing_seconds={"rebuild": elapsed})

    def retry_checkpoint_sha256(self, job_id: str) -> str | None:
        binding = self.binding_store.load(job_id)
        project = self.rebuild_service.describe_interactive_project(binding.project_id)
        if project.project_id != binding.project_id:
            raise RepairPackageError("binding-mismatch", "project_id")
        return self.rebuild_service.interactive_checkpoint_sha256(binding.project_id)

    def resume(
        self, request: RepairPackageRequest, job_id: str, *, interactive_control=None, interactive_observer=None
    ) -> RepairCompletionReport:
        now = self.clock()
        validated = load_and_validate_package(request, now=now, engine_version=self.engine_version)
        candidate_binding = self.binding_store.load(job_id)
        try:
            project = self.rebuild_service.describe_interactive_project(candidate_binding.project_id)
        except Exception as exc:
            raise RepairPackageError("binding-mismatch", f"project_id: {exc}") from exc
        if Path(project.source_path).resolve() != validated.target_snapshot.path.resolve():
            raise RepairPackageError("binding-mismatch", "project source does not match the signed target")
        binding = self.binding_store.validate(
            job_id,
            contract_sha256=validated.contract_sha256,
            source_sha256=validated.original_snapshot.sha256,
            target_sha256=validated.target_snapshot.sha256,
            project_id=project.project_id,
            working_output_path=project.working_output_path,
            destination_path=candidate_binding.destination_path,
            engine_version=self.engine_version,
        )
        validated = replace(
            validated,
            output_path=_validated_bound_destination(
                request, validated, binding, Path(project.working_output_path)
            ),
        )
        started = time.monotonic()
        if self.preview_sink is not None:
            # The Pure DOCX path has no interactive checkpoint of its own to
            # resume (Instant reconstruction is a single fast pass), so a
            # retry just replays the preview and reconstructs again.
            self._play_preview(validated, interactive_control)
            result = self.rebuild_service.rebuild(validated.target_snapshot.path, self.rebuild_options)
        else:
            result = self.rebuild_service.resume_interactive(
                binding.project_id,
                interactive_control=interactive_control,
                interactive_observer=interactive_observer,
            )
        if result.project_id != binding.project_id:
            raise RepairPackageError("binding-mismatch", "result.project_id")
        if result.output_path is not None:
            if Path(result.output_path).resolve() != Path(binding.working_output_path).resolve():
                raise RepairPackageError("binding-mismatch", "result.output_path")
        elif result.status is RunStatus.PASS:
            raise RepairPackageError("binding-mismatch", "result.output_path")
        elapsed = time.monotonic() - started
        return self._finish(
            validated,
            result,
            timing_seconds={"resume": elapsed},
            reuse_existing_bound_output=True,
        )

    def _original_unchanged(self, validated: ValidatedRepairPackage) -> bool:
        try:
            assert_source_unchanged(validated.original_snapshot)
            return True
        except Exception:
            return False

    def _persist_binding(
        self,
        validated: ValidatedRepairPackage,
        *,
        project_id: str,
        working_output_path: Path,
    ) -> None:
        binding = RepairRunBinding.build(
            job_id=validated.contract.job_id,
            contract_sha256=validated.contract_sha256,
            source_sha256=validated.original_snapshot.sha256,
            target_sha256=validated.target_snapshot.sha256,
            project_id=project_id,
            working_output_path=working_output_path,
            destination_path=validated.output_path,
            engine_version=self.engine_version,
        )
        self.binding_store.create(binding)

    def _finish(
        self,
        validated: ValidatedRepairPackage,
        result,
        *,
        timing_seconds: dict[str, float],
        reuse_existing_bound_output: bool = False,
    ) -> RepairCompletionReport:
        required_gates = repair_contract_required_gates(signed_target=True)
        reasons = tuple(result.reasons)
        original_unchanged = self._original_unchanged(validated)
        reconstruction_status = _status_value(result.status)

        if result.status is not RunStatus.PASS or result.output_path is None:
            # WARN/FAIL: never copy a "successful" output, per the module contract.
            return build_completion_report(
                job_id=validated.contract.job_id,
                contract_sha256=validated.contract_sha256,
                source_sha256=validated.original_snapshot.sha256,
                source_size=validated.original_snapshot.size,
                target_sha256=validated.target_snapshot.sha256,
                target_size=validated.target_snapshot.size,
                required_gates=required_gates,
                project_id=result.project_id,
                reconstruction_status=reconstruction_status,
                original_unchanged=original_unchanged,
                reasons=reasons,
                timing_seconds=timing_seconds,
            )

        copy_started = time.monotonic()
        source_output_path = Path(result.output_path)
        source_output_sha256 = sha256_file(source_output_path)
        if reuse_existing_bound_output and validated.output_path.exists():
            if not validated.output_path.is_file() or sha256_file(validated.output_path) != source_output_sha256:
                raise RepairPackageError("output-path-collision", str(validated.output_path))
            _assert_destination_not_protected_alias(validated.output_path, validated)
        else:
            _copy_output_exclusive(source_output_path, validated.output_path)
        output_sha256 = sha256_file(validated.output_path)
        output_size = validated.output_path.stat().st_size
        if output_sha256 != source_output_sha256:
            reasons = (*reasons, "output-copy-hash-mismatch")
        timing = {**timing_seconds, "copy": time.monotonic() - copy_started}

        audit_started = time.monotonic()
        qa_dir = validated.output_path.parent / "qa" / validated.contract.job_id
        audit_report = self.gate_auditor(
            str(validated.target_snapshot.path),
            str(validated.output_path),
            qa_dir,
            run_id=validated.contract.job_id,
            source_sha256=validated.target_snapshot.sha256,
            commit_sha=self.engine_version,
            reconstruction_status=reconstruction_status,
            gate_names=required_gates,
            custom_properties_dropped_by_policy=True,
            application_properties_rewritten_by_policy=True,
        )
        timing["audit"] = time.monotonic() - audit_started
        gates: dict[str, bool] = dict(audit_report.get("gates", {}))

        return build_completion_report(
            job_id=validated.contract.job_id,
            contract_sha256=validated.contract_sha256,
            source_sha256=validated.original_snapshot.sha256,
            source_size=validated.original_snapshot.size,
            target_sha256=validated.target_snapshot.sha256,
            target_size=validated.target_snapshot.size,
            required_gates=required_gates,
            output_sha256=output_sha256,
            output_size=output_size,
            output_path=str(validated.output_path),
            project_id=result.project_id,
            blueprint_fingerprint=audit_report.get("blueprint_fingerprint"),
            reconstruction_status=reconstruction_status,
            original_unchanged=original_unchanged,
            open_and_repair=self.open_and_repair_checker(validated.output_path),
            visible_text_equal=gates.get("G0"),
            fields_update_equal=self.fields_update_checker(validated.output_path),
            gates=gates,
            reasons=reasons,
            timing_seconds=timing,
        )
