from scripts.remote_harness.static_analysis import analyze_document


def test_static_analysis_reports_blueprint_and_structural_counts(corpus_dir):
    report = analyze_document(corpus_dir / "04_tables_merged.docx")
    assert report["package"]["zip_integrity"] is True
    assert report["package"]["missing_content_types"] == []
    assert report["parser"]["ok"] is True
    assert report["model"]["tables"] >= 1
    assert report["model"]["paragraphs"] >= 1
    assert report["blueprint"]["total_events"] > 0
    assert report["source"]["sha256"]
    assert report["preflight"]["maximum_fidelity_ready"] in {True, False}
