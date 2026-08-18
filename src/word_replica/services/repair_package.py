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
from word_replica.repair_contract.binding import RepairRunBinding, RepairRunBindingStore
from word_replica.repair_contract.contract import GOLDEN_GATES
from word_replica.repair_contract.package import RepairPackageRequest, ValidatedRepairPackage, load_and_validate_package
from word_replica.repair_contract.report import RepairCompletionReport, build_completion_report
from word_replica.services.source_guard import assert_source_unchanged, sha256_file


def default_rebuild_options() -> RebuildOptions:
    """Visible Word, maximum fidelity, fast interactive speed, table fast path."""
    return RebuildOptions(
        renderer=RendererChoice.WORD,
        visibility=VisibilityMode.VISIBLE,
        fidelity=FidelityMode.FULL,
        metadata=MetadataMode.FRESH,
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(
            speed_mode=InteractiveSpeedMode.FAST,
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
            result = self.rebuild_service.rebuild(validated.target_snapshot.path, self.rebuild_options)
        else:
            result = self.rebuild_service.rebuild(
                validated.target_snapshot.path,
                self.rebuild_options,
                interactive_control=interactive_control,
                interactive_observer=interactive_observer,
            )
        elapsed = time.monotonic() - started
        return self._finish(validated, result, timing_seconds={"rebuild": elapsed})

    def resume(
        self, request: RepairPackageRequest, job_id: str, *, interactive_control=None, interactive_observer=None
    ) -> RepairCompletionReport:
        now = self.clock()
        validated = load_and_validate_package(request, now=now, engine_version=self.engine_version)
        binding = self.binding_store.validate(
            job_id,
            contract_sha256=validated.contract_sha256,
            source_sha256=validated.original_snapshot.sha256,
            target_sha256=validated.target_snapshot.sha256,
            output_path=validated.output_path,
            engine_version=self.engine_version,
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
        elapsed = time.monotonic() - started
        return self._finish(validated, result, timing_seconds={"resume": elapsed})

    def _original_unchanged(self, validated: ValidatedRepairPackage) -> bool:
        try:
            assert_source_unchanged(validated.original_snapshot)
            return True
        except Exception:
            return False

    def _persist_binding(self, validated: ValidatedRepairPackage, result) -> None:
        if result.project_id is None:
            return
        binding = RepairRunBinding.build(
            job_id=validated.contract.job_id,
            contract_sha256=validated.contract_sha256,
            source_sha256=validated.original_snapshot.sha256,
            target_sha256=validated.target_snapshot.sha256,
            project_id=result.project_id,
            output_path=validated.output_path,
            engine_version=self.engine_version,
        )
        self.binding_store.create(binding)

    def _finish(
        self, validated: ValidatedRepairPackage, result, *, timing_seconds: dict[str, float]
    ) -> RepairCompletionReport:
        self._persist_binding(validated, result)
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
                required_gates=GOLDEN_GATES,
                project_id=result.project_id,
                reconstruction_status=reconstruction_status,
                original_unchanged=original_unchanged,
                reasons=reasons,
                timing_seconds=timing_seconds,
            )

        copy_started = time.monotonic()
        shutil.copyfile(result.output_path, validated.output_path)
        output_sha256 = sha256_file(validated.output_path)
        output_size = validated.output_path.stat().st_size
        if output_sha256 != sha256_file(Path(result.output_path)):
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
            required_gates=GOLDEN_GATES,
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
