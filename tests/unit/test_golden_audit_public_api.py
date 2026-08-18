def test_scripts_codex_automation_audit_reexports_the_same_objects():
    from scripts.codex_automation import audit as legacy
    from word_replica.qa import golden_audit as packaged

    for name in ("GateResult", "build_golden_report", "build_model_gates", "build_visual_gate", "audit_docx_pair", "compare_page_text_partitions"):
        assert getattr(legacy, name) is getattr(packaged, name), name


def test_packaged_audit_docx_pair_matches_legacy_report_with_fakes(tmp_path):
    from scripts.codex_automation.audit import audit_docx_pair as legacy_audit_docx_pair
    from word_replica.qa.golden_audit import audit_docx_pair as packaged_audit_docx_pair
    from word_replica.domain.model import DocumentModel, Paragraph, Run

    class FakeParser:
        def parse(self, path):
            return DocumentModel(source_sha256="abc", body=[Paragraph("p1", runs=[Run("r1", text="Hello")])])

    def fake_pdf_exporter(source, destination, *, visible):
        destination.write_bytes(b"%PDF-fake")

    def fake_page_text_extractor(pdf_path):
        return ["Hello"]

    def fake_pdf_comparer(source_pdf, output_pdf, qa_dir, *, dpi, changed_pixel_tolerance, mae_tolerance):
        class Render:
            available = True
            page_count_match = True
            metrics = []
            source_page_count = 1
            rebuilt_page_count = 1
        return Render()

    def fake_compat_reader(source, output):
        return {"source": False, "output": False}

    kwargs = dict(
        run_id="run-1", source_sha256="abc", commit_sha="deadbeef", reconstruction_status="PASS",
        parser=FakeParser(), pdf_exporter=fake_pdf_exporter, page_text_extractor=fake_page_text_extractor,
        pdf_comparer=fake_pdf_comparer, compatibility_mode_reader=fake_compat_reader,
    )
    legacy_report = legacy_audit_docx_pair("source.docx", "output.docx", tmp_path / "legacy", **kwargs)
    packaged_report = packaged_audit_docx_pair("source.docx", "output.docx", tmp_path / "packaged", **kwargs)
    assert legacy_report == packaged_report
    assert packaged_report["full_pass"] is True
