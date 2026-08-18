from pathlib import Path

from word_replica.domain.model import DocumentModel, Paragraph, Run
from word_replica.qa.render import RenderQaResult, VisualMetric
from scripts.codex_automation.audit import audit_docx_pair


def model():
    return DocumentModel("sha", body=[Paragraph("p", [Run("r", "A")])])


class Parser:
    def parse(self, path):
        return model()


def test_audit_docx_pair_builds_all_ten_gates(tmp_path):
    def exporter(docx, pdf, visible=False):
        Path(pdf).write_bytes(b"pdf")
        return Path(pdf)

    render = RenderQaResult(
        available=True,
        within_tolerance=True,
        page_count_match=True,
        source_page_count=1,
        rebuilt_page_count=1,
        metrics=[VisualMetric(True, 0.0, 0.0, (10, 10), (10, 10))],
    )

    report = audit_docx_pair(
        tmp_path / "source.docx",
        tmp_path / "output.docx",
        tmp_path / "qa",
        run_id="r",
        source_sha256="abc",
        commit_sha="def",
        reconstruction_status="PASS",
        parser=Parser(),
        pdf_exporter=exporter,
        pdf_comparer=lambda a, b, q, **kwargs: render,
        page_text_extractor=lambda p: ["A"],
        compatibility_mode_reader=lambda source, output: {"source": 16, "output": 16},
    )

    assert report["full_pass"] is True
    assert report["gates"] == {f"G{i}": True for i in range(10)}
    assert report["compatibility_mode"] == {"source": 16, "output": 16}


def test_default_audit_path_exports_source_and_output_in_one_word_application(tmp_path):
    calls = []

    def pair_exporter(source_docx, source_pdf, output_docx, output_pdf, visible=False):
        calls.append((source_docx, output_docx, visible))
        Path(source_pdf).write_bytes(b"source pdf")
        Path(output_pdf).write_bytes(b"output pdf")

    render = RenderQaResult(
        available=True,
        within_tolerance=True,
        page_count_match=True,
        source_page_count=1,
        rebuilt_page_count=1,
        metrics=[VisualMetric(True, 0.0, 0.0, (10, 10), (10, 10))],
    )

    report = audit_docx_pair(
        tmp_path / "source.docx",
        tmp_path / "output.docx",
        tmp_path / "qa",
        run_id="r",
        source_sha256="abc",
        commit_sha="def",
        reconstruction_status="PASS",
        parser=Parser(),
        pdf_pair_exporter=pair_exporter,
        pdf_comparer=lambda a, b, q, **kwargs: render,
        page_text_extractor=lambda p: ["A"],
        compatibility_mode_reader=lambda source, output: {"source": 16, "output": 16},
    )

    assert report["full_pass"] is True
    assert calls == [(tmp_path / "source.docx", tmp_path / "output.docx", False)]


def test_audit_preserves_model_gates_when_pdf_export_fails(tmp_path):
    def exporter(docx, pdf, visible=False):
        raise RuntimeError("Word PDF failed")

    report = audit_docx_pair(
        tmp_path / "source.docx",
        tmp_path / "output.docx",
        tmp_path / "qa",
        run_id="r",
        source_sha256="abc",
        commit_sha="def",
        reconstruction_status="WARN",
        parser=Parser(),
        pdf_exporter=exporter,
        compatibility_mode_reader=lambda source, output: {"source": 16, "output": 16},
    )

    assert all(report["gates"][f"G{i}"] for i in range(8))
    assert report["gates"]["G8"] is False
    assert report["gates"]["G9"] is False
    assert "Word PDF failed" in report["gate_details"]["G8"]["summary"]


def test_audit_records_compatibility_mode_reader_failure_without_raising(tmp_path):
    def exporter(docx, pdf, visible=False):
        Path(pdf).write_bytes(b"pdf")
        return Path(pdf)

    render = RenderQaResult(
        available=True,
        within_tolerance=True,
        page_count_match=True,
        source_page_count=1,
        rebuilt_page_count=1,
        metrics=[VisualMetric(True, 0.0, 0.0, (10, 10), (10, 10))],
    )

    def failing_reader(source, output):
        raise RuntimeError("no Word here")

    report = audit_docx_pair(
        tmp_path / "source.docx",
        tmp_path / "output.docx",
        tmp_path / "qa",
        run_id="r",
        source_sha256="abc",
        commit_sha="def",
        reconstruction_status="PASS",
        parser=Parser(),
        pdf_exporter=exporter,
        pdf_comparer=lambda a, b, q, **kwargs: render,
        page_text_extractor=lambda p: ["A"],
        compatibility_mode_reader=failing_reader,
    )

    assert report["full_pass"] is True
    assert report["compatibility_mode"]["source"] is None
    assert report["compatibility_mode"]["output"] is None
    assert "no Word here" in report["compatibility_mode"]["error"]
