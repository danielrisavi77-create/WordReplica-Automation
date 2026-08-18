import hashlib

import pytest

from word_replica.domain.enums import RunStatus
from word_replica.domain.errors import RepairPackageError
from word_replica.repair_contract.binding import RepairRunBinding, RepairRunBindingStore
from word_replica.services.repair_package import RepairPackageService
from tests.unit.test_repair_contract_package import ENGINE_VERSION, NOW, _write_package

ALL_GATES_PASS = {f"G{i}": True for i in range(10)}


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
    def __init__(self, result):
        self.sources = []
        self.resumed_project_ids = []
        self._result = result

    def rebuild(self, source, options, **kwargs):
        self.sources.append(source)
        return self._result

    def resume_interactive(self, project_id, **kwargs):
        self.resumed_project_ids.append(project_id)
        return self._result


class FakeGateAuditor:
    def __init__(self, gates):
        self.calls = []
        self._gates = gates

    def __call__(self, source_docx, output_docx, qa_dir, **kwargs):
        self.calls.append({"source_docx": source_docx, "output_docx": output_docx, "qa_dir": qa_dir, **kwargs})
        return {"gates": dict(self._gates), "blueprint_fingerprint": "fp-123"}


def _service(tmp_path, *, rebuild_result, gates=None, open_and_repair=False, fields_update=True):
    fake_rebuild = FakeRebuildService(rebuild_result)
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

    report = service.run(request)
    binding = store.load(report.job_id)
    assert binding.project_id == "project-42"
    assert binding.source_sha256 == report.source_sha256
    assert binding.target_sha256 == report.target_sha256


def test_resume_validates_binding_and_calls_resume_interactive_with_bound_project_id(tmp_path):
    request = _write_package(tmp_path)
    # First attempt is interrupted: a resumable project exists, but nothing
    # completed, so no output was ever copied to the reserved path yet.
    interrupted = RecordingRunResult(status=RunStatus.WARN, output_path=None, project_id="project-99")
    fake_rebuild = FakeRebuildService(interrupted)
    fake_audit = FakeGateAuditor(ALL_GATES_PASS)
    store = RepairRunBindingStore(app_root=tmp_path / "app_root")
    service = RepairPackageService(
        rebuild_service=fake_rebuild, gate_auditor=fake_audit, binding_store=store,
        clock=lambda: NOW, engine_version=ENGINE_VERSION,
        open_and_repair_checker=lambda path: False, fields_update_checker=lambda path: True,
    )
    first_report = service.run(request)
    assert first_report.status == "RETRYABLE"

    reconstructed = tmp_path / "reconstructed_project_output.docx"
    reconstructed.write_bytes(b"reconstructed bytes")
    fake_rebuild._result = RecordingRunResult(status=RunStatus.PASS, output_path=reconstructed, project_id="project-99")

    second_report = service.resume(request, first_report.job_id)

    assert fake_rebuild.resumed_project_ids == ["project-99"]
    assert second_report.full_pass is True


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


def _real_docx_bytes(text: str) -> bytes:
    import io

    from docx import Document

    buffer = io.BytesIO()
    doc = Document()
    doc.add_paragraph(text)
    doc.save(buffer)
    return buffer.getvalue()


def test_preview_sink_reconstructs_without_any_word_com_dependency(tmp_path):
    from word_replica.domain.enums import FidelityMode, ReconstructionMode, RendererChoice

    request = _write_package(tmp_path, target_bytes=_real_docx_bytes("Popravljeni tekst za pregled"))
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
