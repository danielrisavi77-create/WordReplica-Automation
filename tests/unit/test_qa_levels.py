from word_replica.domain.model import DocumentModel, Paragraph, Run, Section
from word_replica.qa.policy import run_l0_l3


def _model(text: str = "Hello", bold: bool = True) -> DocumentModel:
    return DocumentModel(
        source_sha256="x",
        body=[Paragraph("p", runs=[Run("r", text=text, properties={"bold": bold})], style_id="Normal")],
        sections=[Section("s", {"orientation": "portrait", "width": "12240"})],
    )


def test_equal_models_pass_l0_l3():
    bundle = run_l0_l3(_model(), _model())
    assert all(result.passed for result in bundle.levels.values())


def test_text_change_fails_l0_with_bounded_evidence():
    source = _model("A" * 500)
    rebuilt = _model("B" * 500)
    bundle = run_l0_l3(source, rebuilt)
    finding = bundle.levels["L0"].findings[0]
    assert bundle.levels["L0"].passed is False
    assert finding.code == "L0_MISMATCH"
    assert len(finding.expected["excerpt"]) <= 201
    assert len(finding.expected["sha256"]) == 64


def test_formatting_change_is_l2_only():
    bundle = run_l0_l3(_model(bold=True), _model(bold=False))
    assert bundle.levels["L0"].passed is True
    assert bundle.levels["L1"].passed is True
    assert bundle.levels["L2"].passed is False


def test_l2_compares_effective_inherited_font_not_only_direct_run_properties():
    source = DocumentModel(source_sha256="x", body=[Paragraph("p", [Run("r", "Hello", {})], style_id="Normal")])
    source.extras["document_defaults"] = {"run_properties": {"font_ascii": "Times New Roman", "size_half_points": "24"}, "paragraph_properties": {}}
    source.extras["style_definitions"] = {"Normal": {"style_id":"Normal","name":"Normal","type":"paragraph","run_properties":{},"paragraph_properties":{}}}

    rebuilt_same = DocumentModel(source_sha256="y", body=[Paragraph("p", [Run("r", "Hello", {"font_ascii":"Times New Roman", "size_half_points":"24"})], style_id="Normal")])
    rebuilt_same.extras["document_defaults"] = {"run_properties": {"font_ascii": "Calibri", "size_half_points": "22"}, "paragraph_properties": {}}
    rebuilt_same.extras["style_definitions"] = {"Normal": {"style_id":"Normal","name":"Normal","type":"paragraph","run_properties":{},"paragraph_properties":{}}}
    assert run_l0_l3(source, rebuilt_same).levels["L2"].passed is True

    rebuilt_wrong = DocumentModel(source_sha256="z", body=[Paragraph("p", [Run("r", "Hello", {"font_ascii":"Calibri", "size_half_points":"24"})], style_id="Normal")])
    rebuilt_wrong.extras.update(rebuilt_same.extras)
    bundle = run_l0_l3(source, rebuilt_wrong)
    assert bundle.levels["L2"].passed is False
    paths = [finding.path for finding in bundle.levels["L2"].findings]
    assert any("font_ascii" in path for path in paths)
