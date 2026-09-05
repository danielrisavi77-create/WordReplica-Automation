import hashlib
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

from word_replica.domain.enums import InteractiveSpeedMode, RunStatus
from word_replica.domain.errors import RepairPackageError
from word_replica.repair_contract.binding import RepairRunBinding, RepairRunBindingStore
from word_replica.services.repair_package import (
    RepairPackageService,
    default_rebuild_options,
    repair_contract_required_gates,
)
from tests.unit.test_repair_contract_package import ENGINE_VERSION, NOW, _write_package

ALL_GATES_PASS = {f"G{i}": True for i in range(11)}


def test_paid_repair_default_uses_maximum_speed_without_artificial_text_delay():
    options = default_rebuild_options()

    assert options.interactive.speed_mode is InteractiveSpeedMode.MAXIMUM


class RecordingRunResult:
    def __init__(self, *, status, output_path=None, project_id="project-1", reasons=()):
        self.status = status
        self.output_path = output_path
        self.qa_report_path = None
        self.project_id = project_id
        self.save_count = 1
        self.warnings = []
        self.reasons = list(reasons)


class FakeRebuildService:
    def __init__(self, result, project_root=None):
        self.sources = []
        self.resumed_project_ids = []
        self._result = result
        self.resume_result = None
        self.project_root = project_root
        self.before_start_assertion = None
        self.before_result = None
        self.rebuild_kwargs = []
        self.checkpoint_sha256 = None

    def _prepared(self, source):
        output_dir = (self.project_root or Path(source).parent) / "project-output"
        return SimpleNamespace(
            source_path=str(Path(source).resolve()),
            paths=SimpleNamespace(project_id=self._result.project_id, output_dir=output_dir),
        )

    def rebuild(self, source, options, **kwargs):
        self.sources.append(source)
        self.rebuild_kwargs.append(kwargs)
        callback = kwargs.get("interactive_pre_start")
        if callback is not None:
            prepared = self._prepared(source)
            callback(prepared)
            if self.before_start_assertion is not None:
                self.before_start_assertion(prepared)
        if self.before_result is not None:
            self.before_result()
        return self._result

    def resume_interactive(self, project_id, **kwargs):
        self.resumed_project_ids.append(project_id)
        if self.before_result is not None:
            self.before_result()
        return self.resume_result or self._result

    def describe_interactive_project(self, project_id):
        if project_id != self._result.project_id or not self.sources:
            raise KeyError(project_id)
        prepared = self._prepared(self.sources[0])
        return SimpleNamespace(
            project_id=project_id,
            source_path=Path(prepared.source_path),
            working_output_path=prepared.paths.output_dir / f"{Path(prepared.source_path).stem}_reconstructed.docx",
        )

    def interactive_checkpoint_sha256(self, project_id):
        if project_id != self._result.project_id:
            raise KeyError(project_id)
        return self.checkpoint_sha256


class FakeGateAuditor:
    def __init__(self, gates):
        self.calls = []
        self._gates = gates

    def __call__(self, source_docx, output_docx, qa_dir, **kwargs):
        self.calls.append({"source_docx": source_docx, "output_docx": output_docx, "qa_dir": qa_dir, **kwargs})
        return {"gates": dict(self._gates), "blueprint_fingerprint": "fp-123"}


def _service(tmp_path, *, rebuild_result, gates=None, open_and_repair=False, fields_update=True):
    fake_rebuild = FakeRebuildService(rebuild_result, project_root=tmp_path)
    fake_audit = FakeGateAuditor(gates if gates is not None else ALL_GATES_PASS)
    store = RepairRunBindingStore(app_root=tmp_path / "app_root")
    service = RepairPackageService(
        rebuild_service=fake_rebuild,
        gate_auditor=fake_audit,
        binding_store=store,
        clock=lambda: NOW,
        engine_version=ENGINE_VERSION,
        open_and_repair_checker=lambda path: open_and_repair,
        fields_update_checker=lambda path: fields_update,
    )
    return service, fake_rebuild, fake_audit, store


def test_run_reconstructs_the_target_and_produces_a_full_pass_report(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed)
    service, fake_rebuild, fake_audit, _ = _service(tmp_path, rebuild_result=result)

    report = service.run(request)

    assert fake_rebuild.sources == [request.target_path.resolve()]
    assert fake_rebuild.rebuild_kwargs[0]["expected_source_snapshot"].sha256 == report.target_sha256
    assert report.original_unchanged is True
    assert report.full_pass is True
    assert report.status == "FULL_PASS"
    assert report.output_sha256 == hashlib.sha256(b"reconstructed bytes").hexdigest()
    assert report.project_id == "project-1"
    assert report.blueprint_fingerprint == "fp-123"

    # Audit compares the TARGET to the local output, never the original.
    call = fake_audit.calls[0]
    assert call["source_docx"] == str(request.target_path.resolve())
    assert call["output_docx"] == report.output_path
    assert call["custom_properties_dropped_by_policy"] is True
    assert call["application_properties_rewritten_by_policy"] is True


def test_invalid_package_never_calls_rebuild(tmp_path):
    request = _write_package(tmp_path, source_bytes=b"tampered, does not match signed hash")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=tmp_path / "unused.docx")
    service, fake_rebuild, fake_audit, _ = _service(tmp_path, rebuild_result=result)

    with pytest.raises(RepairPackageError):
        service.run(request)

    assert fake_rebuild.sources == []
    assert fake_audit.calls == []


def test_warn_status_is_retryable_and_never_copies_an_output(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"partial")
    result = RecordingRunResult(status=RunStatus.WARN, output_path=reconstructed, reasons=("some warning",))
    service, fake_rebuild, fake_audit, _ = _service(tmp_path, rebuild_result=result)

    report = service.run(request)

    assert report.status == "RETRYABLE"
    assert report.full_pass is False
    assert report.output_sha256 is None
    assert report.output_path is None
    assert fake_audit.calls == []  # no audit without a copied output


def test_fail_status_is_failed_and_never_copies_an_output(tmp_path):
    request = _write_package(tmp_path)
    result = RecordingRunResult(status=RunStatus.FAIL, output_path=None, reasons=("com_error",))
    service, fake_rebuild, fake_audit, _ = _service(tmp_path, rebuild_result=result)

    report = service.run(request)

    assert report.status == "FAILED"
    assert report.full_pass is False
    assert report.output_path is None
    assert fake_audit.calls == []


def test_rejected_word_com_call_resumes_automatically_once(tmp_path):
    request = _write_package(tmp_path)
    rejected = RecordingRunResult(
        status=RunStatus.FAIL,
        output_path=None,
        project_id="project-rejected-call",
        reasons=("(-2147418111, 'Call was rejected by callee.', None, None)",),
    )
    service, fake_rebuild, _, _ = _service(tmp_path, rebuild_result=rejected)
    cleanup_observations = []
    service.owned_word_cleanup = lambda: cleanup_observations.append(
        tuple(fake_rebuild.resumed_project_ids)
    )
    reconstructed = tmp_path / "reconstructed_after_automatic_resume.docx"
    reconstructed.write_bytes(b"reconstructed after automatic resume")
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=reconstructed,
        project_id="project-rejected-call",
    )

    report = service.run(request)

    assert fake_rebuild.resumed_project_ids == ["project-rejected-call"]
    assert cleanup_observations == [(), ("project-rejected-call",)]
    assert report.status == "FULL_PASS"
    assert report.full_pass is True


def test_a_failing_gate_prevents_full_pass_even_on_a_reconstruction_pass(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed)
    broken_gates = {**ALL_GATES_PASS, "G3": False}
    service, _, _, _ = _service(tmp_path, rebuild_result=result, gates=broken_gates)

    report = service.run(request)

    assert report.full_pass is False
    assert report.status == "RETRYABLE"
    assert report.gates["G3"] is False


def test_unverified_open_and_repair_prevents_full_pass(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed)
    service, _, _, _ = _service(tmp_path, rebuild_result=result, open_and_repair=None)

    report = service.run(request)

    assert report.open_and_repair is None
    assert report.full_pass is False


def test_run_persists_a_binding_usable_for_resume(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed, project_id="project-42")
    service, _, _, store = _service(tmp_path, rebuild_result=result)

    def assert_bound_before_start(prepared):
        binding = store.load("11111111-1111-4111-8111-111111111111")
        assert binding.project_id == prepared.paths.project_id
        assert Path(binding.working_output_path).parent == prepared.paths.output_dir

    service.rebuild_service.before_start_assertion = assert_bound_before_start

    report = service.run(request)
    binding = store.load(report.job_id)
    assert binding.project_id == "project-42"
    assert binding.source_sha256 == report.source_sha256
    assert binding.target_sha256 == report.target_sha256
    assert binding.working_output_path != binding.destination_path
    assert binding.destination_path == report.output_path


def test_resume_validates_binding_and_calls_resume_interactive_with_bound_project_id(tmp_path):
    request = _write_package(tmp_path)
    # First attempt is interrupted: a resumable project exists, but nothing
    # completed, so no output was ever copied to the reserved path yet.
    interrupted = RecordingRunResult(status=RunStatus.WARN, output_path=None, project_id="project-99")
    fake_rebuild = FakeRebuildService(interrupted, project_root=tmp_path)
    fake_audit = FakeGateAuditor(ALL_GATES_PASS)
    store = RepairRunBindingStore(app_root=tmp_path / "app_root")
    service = RepairPackageService(
        rebuild_service=fake_rebuild, gate_auditor=fake_audit, binding_store=store,
        clock=lambda: NOW, engine_version=ENGINE_VERSION,
        open_and_repair_checker=lambda path: False, fields_update_checker=lambda path: True,
    )
    first_report = service.run(request)
    assert first_report.status == "RETRYABLE"

    binding = store.load(first_report.job_id)
    reconstructed = Path(binding.working_output_path)
    reconstructed.parent.mkdir(parents=True, exist_ok=True)
    reconstructed.write_bytes(b"reconstructed bytes")
    fake_rebuild._result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed, project_id="project-99")

    second_report = service.resume(request, first_report.job_id)

    assert fake_rebuild.resumed_project_ids == ["project-99"]
    assert second_report.full_pass is True


def test_resume_reuses_bound_destination_after_post_copy_audit_failure(tmp_path):
    request = _write_package(tmp_path)
    project_output = tmp_path / "project-output"
    project_output.mkdir()
    reconstructed = project_output / f"{request.target_path.stem}_reconstructed.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=reconstructed,
        project_id="project-post-copy",
    )
    fake_rebuild = FakeRebuildService(result, project_root=tmp_path)
    successful_audit = FakeGateAuditor(ALL_GATES_PASS)
    audit_calls = 0

    def fail_once_after_copy(source_docx, output_docx, qa_dir, **kwargs):
        nonlocal audit_calls
        audit_calls += 1
        if audit_calls == 1:
            assert Path(output_docx).read_bytes() == b"reconstructed bytes"
            raise RuntimeError("simulated post-copy audit failure")
        return successful_audit(source_docx, output_docx, qa_dir, **kwargs)

    store = RepairRunBindingStore(app_root=tmp_path / "app_root")
    service = RepairPackageService(
        rebuild_service=fake_rebuild,
        gate_auditor=fail_once_after_copy,
        binding_store=store,
        clock=lambda: NOW,
        engine_version=ENGINE_VERSION,
        open_and_repair_checker=lambda path: False,
        fields_update_checker=lambda path: True,
    )

    with pytest.raises(RuntimeError, match="simulated post-copy audit failure"):
        service.run(request)

    binding = store.load("11111111-1111-4111-8111-111111111111")
    bound_destination = Path(binding.destination_path)
    assert bound_destination.is_file()

    report = service.resume(request, binding.job_id)

    assert report.full_pass is True
    assert Path(report.output_path) == bound_destination
    assert not (bound_destination.parent / f"{bound_destination.stem} (2){bound_destination.suffix}").exists()
    assert bound_destination.read_bytes() == b"reconstructed bytes"


def _resumable_service_with_binding(tmp_path):
    request = _write_package(tmp_path)
    interrupted = RecordingRunResult(
        status=RunStatus.WARN,
        output_path=None,
        project_id="project-bound",
    )
    service, fake_rebuild, fake_audit, store = _service(
        tmp_path, rebuild_result=interrupted
    )
    first_report = service.run(request)
    assert first_report.status == "RETRYABLE"
    binding = store.load(first_report.job_id)
    working_output = Path(binding.working_output_path)
    working_output.parent.mkdir(parents=True, exist_ok=True)
    working_output.write_bytes(b"bound working output")
    return request, service, fake_rebuild, fake_audit, binding, working_output


def test_resume_cleans_owned_word_after_the_attempt(tmp_path):
    request, service, fake_rebuild, _, binding, working_output = (
        _resumable_service_with_binding(tmp_path)
    )
    cleanup_calls = []
    service.owned_word_cleanup = lambda: cleanup_calls.append(binding.project_id)
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=working_output,
        project_id=binding.project_id,
    )

    report = service.resume(request, binding.job_id)

    assert report.full_pass is True
    assert cleanup_calls == [binding.project_id]


@pytest.mark.parametrize(
    "drift",
    ("project_id", "output_path"),
    ids=("result-project-id", "result-working-output"),
)
def test_resume_rejects_result_that_drifts_from_durable_binding(tmp_path, drift):
    request, service, fake_rebuild, fake_audit, binding, working_output = (
        _resumable_service_with_binding(tmp_path)
    )
    result_project_id = binding.project_id
    result_output_path = working_output
    if drift == "project_id":
        result_project_id = "project-other"
    else:
        result_output_path = tmp_path / "other-working-output.docx"
        result_output_path.write_bytes(b"other working output")
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=result_output_path,
        project_id=result_project_id,
    )

    with pytest.raises(RepairPackageError) as excinfo:
        service.resume(request, binding.job_id)

    assert excinfo.value.code == "binding-mismatch"
    assert f"result.{drift}" in str(excinfo.value)
    assert fake_audit.calls == []
    assert not Path(binding.destination_path).exists()


@pytest.mark.parametrize(
    "protected_attribute",
    ("original_path", "target_path"),
    ids=("source-hard-link", "target-hard-link"),
)
def test_resume_rejects_bound_destination_hard_link_to_protected_input(
    tmp_path, protected_attribute
):
    request, service, fake_rebuild, fake_audit, binding, working_output = (
        _resumable_service_with_binding(tmp_path)
    )
    protected_path = getattr(request, protected_attribute)
    working_output.write_bytes(protected_path.read_bytes())
    bound_destination = Path(binding.destination_path)
    os.link(protected_path, bound_destination)
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=working_output,
        project_id=binding.project_id,
    )

    with pytest.raises(RepairPackageError) as excinfo:
        service.resume(request, binding.job_id)

    assert excinfo.value.code == "binding-mismatch"
    assert "destination_path" in str(excinfo.value)
    assert fake_rebuild.resumed_project_ids == []
    assert fake_audit.calls == []


def test_resume_fails_closed_when_bound_destination_identity_cannot_be_checked(
    tmp_path, monkeypatch
):
    request, service, fake_rebuild, fake_audit, binding, working_output = (
        _resumable_service_with_binding(tmp_path)
    )
    bound_destination = Path(binding.destination_path)
    bound_destination.write_bytes(working_output.read_bytes())
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=working_output,
        project_id=binding.project_id,
    )

    def fail_identity_check(self, other):
        raise OSError("simulated identity failure")

    monkeypatch.setattr(Path, "samefile", fail_identity_check)

    with pytest.raises(RepairPackageError) as excinfo:
        service.resume(request, binding.job_id)

    assert excinfo.value.code == "binding-mismatch"
    assert "destination identity check failed" in str(excinfo.value)
    assert fake_rebuild.resumed_project_ids == []
    assert fake_audit.calls == []


@pytest.mark.parametrize(
    "protected_attribute",
    ("original_path", "target_path"),
    ids=("source-hard-link", "target-hard-link"),
)
def test_resume_rechecks_destination_identity_after_interactive_result(
    tmp_path, protected_attribute
):
    request, service, fake_rebuild, fake_audit, binding, working_output = (
        _resumable_service_with_binding(tmp_path)
    )
    protected_path = getattr(request, protected_attribute)
    protected_bytes = protected_path.read_bytes()
    working_output.write_bytes(protected_bytes)
    bound_destination = Path(binding.destination_path)
    bound_destination.write_bytes(protected_bytes)
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=working_output,
        project_id=binding.project_id,
    )

    def replace_destination_with_protected_hard_link():
        bound_destination.unlink()
        os.link(protected_path, bound_destination)

    fake_rebuild.before_result = replace_destination_with_protected_hard_link

    with pytest.raises(RepairPackageError) as excinfo:
        service.resume(request, binding.job_id)

    assert excinfo.value.code == "binding-mismatch"
    assert "destination_path" in str(excinfo.value)
    assert bound_destination.samefile(protected_path)
    assert fake_rebuild.resumed_project_ids == [binding.project_id]
    assert fake_audit.calls == []


def test_resume_fails_closed_when_post_result_identity_recheck_errors(
    tmp_path, monkeypatch
):
    request, service, fake_rebuild, fake_audit, binding, working_output = (
        _resumable_service_with_binding(tmp_path)
    )
    bound_destination = Path(binding.destination_path)
    bound_destination.write_bytes(working_output.read_bytes())
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=working_output,
        project_id=binding.project_id,
    )
    real_samefile = Path.samefile
    identity_checks = 0

    def fail_after_pre_resume_checks(self, other):
        nonlocal identity_checks
        identity_checks += 1
        if identity_checks > 2:
            raise OSError("simulated post-result identity failure")
        return real_samefile(self, other)

    monkeypatch.setattr(Path, "samefile", fail_after_pre_resume_checks)

    with pytest.raises(RepairPackageError) as excinfo:
        service.resume(request, binding.job_id)

    assert excinfo.value.code == "binding-mismatch"
    assert "destination identity check failed" in str(excinfo.value)
    assert identity_checks == 3
    assert fake_rebuild.resumed_project_ids == [binding.project_id]
    assert fake_audit.calls == []


def test_resume_never_overwrites_changed_bound_destination(tmp_path):
    request, service, fake_rebuild, fake_audit, binding, working_output = (
        _resumable_service_with_binding(tmp_path)
    )
    bound_destination = Path(binding.destination_path)
    bound_destination.write_bytes(b"unrelated content")
    fake_rebuild.resume_result = RecordingRunResult(
        status=RunStatus.PASS,
        output_path=working_output,
        project_id=binding.project_id,
    )

    with pytest.raises(RepairPackageError) as excinfo:
        service.resume(request, binding.job_id)

    assert excinfo.value.code == "output-path-collision"
    assert bound_destination.read_bytes() == b"unrelated content"
    assert fake_rebuild.resumed_project_ids == []
    assert fake_audit.calls == []


def test_retry_checkpoint_hash_is_loaded_only_through_the_bound_project(tmp_path):
    request = _write_package(tmp_path)
    interrupted = RecordingRunResult(status=RunStatus.WARN, output_path=None, project_id="project-77")
    service, fake_rebuild, _, _ = _service(tmp_path, rebuild_result=interrupted)
    fake_rebuild.checkpoint_sha256 = "d" * 64

    report = service.run(request)

    assert report.status == "RETRYABLE"
    assert service.retry_checkpoint_sha256(report.job_id) == "d" * 64


def test_resume_fails_closed_when_no_binding_exists_for_the_job(tmp_path):
    request = _write_package(tmp_path)
    result = RecordingRunResult(status=RunStatus.PASS, output_path=tmp_path / "unused.docx")
    service, fake_rebuild, _, _ = _service(tmp_path, rebuild_result=result)

    with pytest.raises(RepairPackageError):
        service.resume(request, "11111111-1111-4111-8111-111111111111")
    assert fake_rebuild.resumed_project_ids == []


def test_output_copy_never_overwrites_source_or_target(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    original_bytes_before = request.original_path.read_bytes()
    target_bytes_before = request.target_path.read_bytes()
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed)
    service, _, _, _ = _service(tmp_path, rebuild_result=result)

    report = service.run(request)

    assert request.original_path.read_bytes() == original_bytes_before
    assert request.target_path.read_bytes() == target_bytes_before
    assert report.output_path not in (str(request.original_path), str(request.target_path))


def test_output_copy_fails_closed_if_another_process_claims_reserved_path(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed)
    service, fake_rebuild, fake_audit, _ = _service(tmp_path, rebuild_result=result)
    collision = request.output_dir / request.target_path.name

    def claim_destination():
        collision.write_bytes(b"other process")

    fake_rebuild.before_result = claim_destination

    with pytest.raises(RepairPackageError) as excinfo:
        service.run(request)

    assert excinfo.value.code == "output-path-collision"
    assert collision.read_bytes() == b"other process"
    assert fake_audit.calls == []


def _real_docx_bytes(text: str) -> bytes:
    import io

    from docx import Document

    buffer = io.BytesIO()
    doc = Document()
    doc.add_paragraph(text)
    doc.save(buffer)
    return buffer.getvalue()


def test_preview_sink_reconstructs_without_any_word_com_dependency(tmp_path, monkeypatch):
    from word_replica.domain.enums import FidelityMode, ReconstructionMode, RendererChoice

    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed)

    fake_rebuild = FakeRebuildService(result)
    fake_audit = FakeGateAuditor(ALL_GATES_PASS)
    store = RepairRunBindingStore(app_root=tmp_path / "app_root")
    preview_events = []
    service = RepairPackageService(
        rebuild_service=fake_rebuild,
        gate_auditor=fake_audit,
        binding_store=store,
        clock=lambda: NOW,
        engine_version=ENGINE_VERSION,
        preview_sink=lambda kind, payload: preview_events.append((kind, payload)),
        open_and_repair_checker=lambda path: False,
        fields_update_checker=lambda path: True,
    )

    # Exercise the real preview player against a real DOCX independently of
    # the immutable signed public package vector.
    preview_docx = tmp_path / "preview.docx"
    preview_docx.write_bytes(_real_docx_bytes("Popravljeni tekst za pregled"))
    service._play_preview(
        SimpleNamespace(target_snapshot=SimpleNamespace(path=preview_docx)),
        None,
    )
    monkeypatch.setattr(service, "_play_preview", lambda validated, control: None)

    report = service.run(request)

    # The preview actually replayed the target's real text.
    preview_text = "".join(payload["text"] for kind, payload in preview_events if kind == "text")
    assert "Popravljeni tekst za pregled" in preview_text
    # The real file still gets produced through the licence-free Pure DOCX path.
    assert fake_rebuild.sources == [request.target_path.resolve()]
    assert service.rebuild_options.renderer is RendererChoice.DOCX
    assert service.rebuild_options.reconstruction_mode is ReconstructionMode.INSTANT
    assert service.rebuild_options.fidelity is FidelityMode.FULL
    assert report.full_pass is True


def test_explicit_rebuild_options_override_the_preview_sink_default(tmp_path):
    from word_replica.config import RebuildOptions
    from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, VisibilityMode

    custom = RebuildOptions(
        renderer=RendererChoice.DOCX, visibility=VisibilityMode.BACKGROUND,
        fidelity=FidelityMode.CLEAN, metadata=MetadataMode.PRESERVE,
    )
    store = RepairRunBindingStore(app_root=tmp_path / "app_root")
    service = RepairPackageService(
        rebuild_service=FakeRebuildService(RecordingRunResult(status=RunStatus.PASS)),
        gate_auditor=FakeGateAuditor(ALL_GATES_PASS),
        binding_store=store,
        clock=lambda: NOW,
        engine_version=ENGINE_VERSION,
        preview_sink=lambda kind, payload: None,
        rebuild_options=custom,
    )
    assert service.rebuild_options is custom


def test_signed_target_contract_requires_g10_but_legacy_policy_does_not():
    assert repair_contract_required_gates(signed_target=False) == tuple(f"G{i}" for i in range(10))
    assert repair_contract_required_gates(signed_target=True) == tuple(f"G{i}" for i in range(11))


def test_signed_target_cannot_full_pass_when_g10_is_missing(tmp_path):
    request = _write_package(tmp_path)
    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed)
    only_legacy_gates = {f"G{i}": True for i in range(10)}
    service, _, fake_audit, _ = _service(tmp_path, rebuild_result=result, gates=only_legacy_gates)

    report = service.run(request)

    assert fake_audit.calls[0]["gate_names"] == tuple(f"G{i}" for i in range(11))
    assert report.gates.get("G10") is not True
    assert report.full_pass is False
