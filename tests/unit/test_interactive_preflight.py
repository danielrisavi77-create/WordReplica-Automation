from word_replica.config import InteractiveOptions
from word_replica.domain.enums import InteractiveFidelity
from word_replica.domain.model import BinaryAsset, DocumentModel, DrawingRef, PreservedPart
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.interactive.preflight import analyze_preflight


def base_model():
    return DocumentModel(source_sha256="a"*64)


def test_word_unavailable_is_blocking():
    model=base_model(); bp=BlueprintCompiler().compile(model)
    report=analyze_preflight(model,bp,InteractiveOptions(),word_probe=lambda:False)
    assert report.word_available is False
    assert report.maximum_fidelity_ready is False
    assert any("Word" in reason for reason in report.blocking_reasons)


def test_missing_drawing_asset_is_blocking():
    model=base_model(); model.drawings.append(DrawingRef("d","missing","word/media/x.png","inline"))
    bp=BlueprintCompiler().compile(model)
    report=analyze_preflight(model,bp,InteractiveOptions(),word_probe=lambda:True)
    assert any("asset" in reason.lower() for reason in report.blocking_reasons)


def test_preserved_object_blocks_maximum_until_real_preservation_is_implemented():
    model=base_model(); model.preserved_parts["word/charts/chart1.xml"] = PreservedPart("word/charts/chart1.xml",None,None,"f"*64,b"x")
    bp=BlueprintCompiler().compile(model)
    blocked=analyze_preflight(model,bp,InteractiveOptions(),word_probe=lambda:True)
    assert blocked.maximum_fidelity_ready is False
    still_blocked=analyze_preflight(model,bp,InteractiveOptions(allow_preserved_objects=True),word_probe=lambda:True)
    assert still_blocked.maximum_fidelity_ready is False
    assert still_blocked.can_proceed is False


def test_standard_fidelity_reports_unsupported_as_warning_not_silent_block():
    model=base_model(); model.drawings.append(DrawingRef("d","a","word/media/x.xyz","inline"))
    model.assets["a"] = BinaryAsset("a","word/media/x.xyz",None,"0"*64,b"x")
    bp=BlueprintCompiler().compile(model)
    opts=InteractiveOptions(fidelity=InteractiveFidelity.STANDARD)
    report=analyze_preflight(model,bp,opts,word_probe=lambda:True)
    assert report.can_proceed is True
    assert report.capability_items
    assert report.warnings


def test_preserved_part_is_unsupported_until_real_preservation_exists():
    model = base_model()
    model.preserved_parts["word/charts/chart1.xml"] = PreservedPart(
        "word/charts/chart1.xml", None, None, "f" * 64, b"x"
    )
    bp = BlueprintCompiler().compile(model)
    report = analyze_preflight(
        model,
        bp,
        InteractiveOptions(allow_preserved_objects=True),
        word_probe=lambda: True,
    )
    assert any(item.source_element_id == "word/charts/chart1.xml" and item.classification.value == "UNSUPPORTED" for item in report.capability_items)
    assert report.can_proceed is False


def test_comments_and_tracked_revisions_are_never_silently_dropped_in_maximum_mode():
    from word_replica.domain.model import Comment, Paragraph, RevisionSpan, Run

    model = base_model()
    model.comments["1"] = Comment("1", "Reviewer", None, [Paragraph("cp", [Run("cr", "Comment")])])
    model.revisions.append(RevisionSpan("7", "insert", "Reviewer", None, "/body/p[1]", "Inserted"))
    bp = BlueprintCompiler().compile(model)
    report = analyze_preflight(model, bp, InteractiveOptions(), word_probe=lambda: True)
    reasons = [item.reason.lower() for item in report.capability_items]
    assert any("comment" in reason for reason in reasons)
    assert any("tracked" in reason or "revision" in reason for reason in reasons)
    assert report.can_proceed is False


def test_preflight_checks_fonts_in_used_paragraph_style_definitions(corpus_dir):
    from word_replica.config import InteractiveOptions
    from word_replica.domain.enums import InteractiveFidelity
    from word_replica.interactive.blueprint import BlueprintCompiler
    from word_replica.parser.parser import DocxParser
    from word_replica.interactive.preflight import analyze_preflight

    model = DocxParser().parse(corpus_dir / "02_headings_styles.docx")
    report = analyze_preflight(
        model, BlueprintCompiler().compile(model),
        InteractiveOptions(fidelity=InteractiveFidelity.MAXIMUM),
        word_probe=lambda: True,
        font_probe=lambda font: font != "Arial",
    )
    assert "Arial" in report.missing_fonts
    assert report.can_proceed is False
