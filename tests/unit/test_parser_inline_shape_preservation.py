"""Inline content the document model cannot represent.

The last defect class G10 found on the corpus, and the largest: every
outstanding divergence is drawing content that is not a picture --
`wordprocessingDrawing` (13 documents), VML (12), `wordprocessingShape` (8),
legacy `office:word` (6), charts (3), groups and diagrams (2 each), plus 7
documents whose `mc:AlternateContent` blocks vanish entirely.

The cause is narrow. `parse_drawings` finds `//w:drawing` anywhere, including
inside `mc:AlternateContent`, but only produces a `DrawingRef` when there is a
`blip` -- that is, a picture. A group shape, a text box or a diagram has no
picture, so it has no model representation at all, and the rebuilt body emits
nothing where it was.

The fix preserves the raw XML fragment and re-emits it. That is only safe
because of a measurement: every `mc:AlternateContent` block sampled from the
corpus carried no `r:` relationship reference. A fragment that *does* carry one
must be refused -- relationship ids are renumbered freely by any writer, so
re-emitting a stale one produces a document Word repairs on open, which is
worse than a missing shape.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
)
from word_replica.parser.parser import DocxParser
from word_replica.qa.preservation import build_preservation_gate, g10_projection
from word_replica.services.rebuild import RebuildService

CONTENT_TYPES = (
    b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    b'.wordprocessingml.document.main+xml"/>'
    b"</Types>"
)
ROOT_RELS = (
    b'<?xml version="1.0"?><Relationships '
    b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Target="word/document.xml" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
    b"</Relationships>"
)
DOC_RELS = (
    b'<?xml version="1.0"?><Relationships '
    b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)

_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

# A group shape with a VML fallback -- the shape Word writes for a drawing
# canvas. No picture anywhere in it, so nothing in the model represents it.
GROUP_SHAPE = (
    "<mc:AlternateContent>"
    '<mc:Choice Requires="wpg">'
    "<w:drawing><wp:anchor><wp:extent cx='3419475' cy='1209675'/>"
    "<wpg:wgp><wpg:grpSpPr/></wpg:wgp></wp:anchor></w:drawing>"
    "</mc:Choice>"
    "<mc:Fallback><w:pict><v:shape id='_x0000_s1026' style='width:269pt'/></w:pict></mc:Fallback>"
    "</mc:AlternateContent>"
)

# The same shape, but the fallback references an image by relationship id.
GROUP_SHAPE_WITH_RELATIONSHIP = (
    "<mc:AlternateContent>"
    '<mc:Choice Requires="wpg">'
    "<w:drawing><wp:anchor><wp:extent cx='100' cy='100'/></wp:anchor></w:drawing>"
    "</mc:Choice>"
    "<mc:Fallback><w:pict><v:shape><v:imagedata r:id='rId42'/></v:shape></w:pict></mc:Fallback>"
    "</mc:AlternateContent>"
)


def _document(inline: str) -> bytes:
    return (
        f'<?xml version="1.0"?><w:document {_NS}>'
        f"<w:body><w:p><w:r><w:t>before</w:t>{inline}</w:r>"
        "<w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
    ).encode("utf-8")


def _source(path, inline: str):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", _document(inline))
    return path


def _tokens(model):
    return [
        token
        for paragraph in model.iter_paragraphs()
        for run in paragraph.runs
        for token in (run.properties.get("content_tokens") or ())
    ]


def _rebuild(tmp_path, source, suffix=""):
    result = RebuildService(app_root=tmp_path / f"app{suffix}").rebuild(
        source,
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )
    assert result.output_path is not None, result.reasons
    return result.output_path


# --- capture ------------------------------------------------------------------

def test_an_unrepresentable_inline_fragment_is_captured(tmp_path):
    model = DocxParser().parse(_source(tmp_path / "src.docx", GROUP_SHAPE))

    preserved = [t for t in _tokens(model) if t["kind"] == "preserved_xml"]
    assert len(preserved) == 1
    assert "AlternateContent" in preserved[0]["value"]


def test_the_surrounding_text_is_untouched(tmp_path):
    model = DocxParser().parse(_source(tmp_path / "src.docx", GROUP_SHAPE))

    assert model.plain_text().replace("\n", "") == "beforeafter"


def test_a_fragment_carrying_a_relationship_id_is_refused(tmp_path):
    # Relationship ids are renumbered freely by any writer. Re-emitting a stale
    # one produces a document Word repairs on open -- worse than a missing shape.
    model = DocxParser().parse(_source(tmp_path / "src.docx", GROUP_SHAPE_WITH_RELATIONSHIP))

    kinds = [t["kind"] for t in _tokens(model)]
    assert "preserved_xml" not in kinds
    assert "unsupported_inline" in kinds


def test_the_refusal_records_why(tmp_path):
    model = DocxParser().parse(_source(tmp_path / "src.docx", GROUP_SHAPE_WITH_RELATIONSHIP))

    refused = next(t for t in _tokens(model) if t["kind"] == "unsupported_inline")
    assert "relationship" in refused["reason"]


def test_a_picture_is_still_modelled_rather_than_preserved_raw(tmp_path, corpus_dir):
    # Pictures have a real model representation; preserving them raw as well
    # would emit them twice.
    model = DocxParser().parse(corpus_dir / "05_images_inline_floating.docx")

    assert model.drawings
    assert not [t for t in _tokens(model) if t["kind"] == "preserved_xml"]


# --- re-emission --------------------------------------------------------------

def test_the_fragment_comes_back_in_the_rebuild(tmp_path):
    output = _rebuild(tmp_path, _source(tmp_path / "src.docx", GROUP_SHAPE))

    with ZipFile(output) as archive:
        document = archive.read("word/document.xml").decode("utf-8")
    assert "AlternateContent" in document
    assert "wgp" in document


def test_g10_sees_the_shape_content_preserved(tmp_path):
    source = _source(tmp_path / "src.docx", GROUP_SHAPE)
    output = _rebuild(tmp_path, source)

    before, after = g10_projection(source), g10_projection(output)
    assert after["markup_compatibility"]["alternate_content_count"] == (
        before["markup_compatibility"]["alternate_content_count"]
    )
    for namespace in ("wordprocessingGroup", "vml", "wordprocessingDrawing"):
        matching = [u for u in before["namespaces"] if namespace.lower() in u.lower()]
        for uri in matching:
            assert uri in after["namespaces"], f"{uri} lost"


def test_a_refused_fragment_is_not_re_emitted(tmp_path):
    output = _rebuild(tmp_path, _source(tmp_path / "src.docx", GROUP_SHAPE_WITH_RELATIONSHIP))

    with ZipFile(output) as archive:
        document = archive.read("word/document.xml").decode("utf-8")
    assert "rId42" not in document


def test_a_rebuild_with_a_preserved_fragment_has_no_dangling_relationships(tmp_path):
    output = _rebuild(tmp_path, _source(tmp_path / "src.docx", GROUP_SHAPE))

    assert g10_projection(output)["dangling_relationships"] == []


def test_the_rebuild_is_still_parseable(tmp_path):
    output = _rebuild(tmp_path, _source(tmp_path / "src.docx", GROUP_SHAPE))

    reparsed = DocxParser().parse(output)
    assert reparsed.plain_text().replace("\n", "") == "beforeafter"


def test_a_second_round_trip_is_stable(tmp_path):
    # A fragment that is captured and re-emitted must survive being captured
    # and re-emitted again, or the document degrades with each rebuild.
    once = _rebuild(tmp_path, _source(tmp_path / "src.docx", GROUP_SHAPE), suffix="1")
    twice = _rebuild(tmp_path, once, suffix="2")

    assert build_preservation_gate(once, twice).passed is True


# --- the interactive path must be unaffected ----------------------------------

def test_the_blueprint_compiler_ignores_a_preserved_fragment(tmp_path):
    from word_replica.interactive.blueprint import BlueprintCompiler

    model = DocxParser().parse(_source(tmp_path / "src.docx", GROUP_SHAPE))
    blueprint = BlueprintCompiler().compile(model)

    assert blueprint.total_events > 0
    assert "beforeafter" in "".join(
        event.payload.get("text", "") for event in blueprint.events
    ).replace(" ", "")
