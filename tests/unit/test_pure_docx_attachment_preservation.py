"""Document-level attachments must survive a rebuild.

Found by G10 on the LibreOffice corpus: reconstructions that every model gate
called perfect had lost the source's custom XML datastore (4 of 18) and its
bibliography sources (4 of 18).

Two gaps combined to cause it. `parser.py` only captured three prefixes into
`preserved_parts` -- word/charts/, word/embeddings/, word/diagrams/ -- so
customXml and bibliography were never read at all. And `pure_docx.py` never
installed preserved parts even when it had them; it emitted an
`UNSUPPORTED_TRANSFER_PART` warning and dropped them.

The distinction that matters is *how a part is reached*. A chart or an embedded
object is referenced from inside the document body, so restoring the part
without the body reference would produce an orphan. A custom XML store, a
bibliography, a people list or an embedded font is attached at document level:
part plus one relationship, self-contained, and safe to carry through.
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
from word_replica.qa.preservation import build_preservation_gate
from word_replica.services.rebuild import RebuildService

CT_HEAD = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
)
ROOT_RELS = (
    '<?xml version="1.0"?><Relationships '
    'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Target="word/document.xml" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
    "</Relationships>"
)
DOC = (
    '<?xml version="1.0"?><w:document '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
)

CUSTOM_XML = b'<?xml version="1.0"?><contract xmlns="urn:acme:contract"><number>A-4417</number></contract>'
CUSTOM_XML_PROPS = (
    b'<?xml version="1.0"?><ds:datastoreItem '
    b'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" '
    b'ds:itemID="{7C4B2A10-0000-0000-0000-000000000001}"/>'
)
BIBLIOGRAPHY = (
    b'<?xml version="1.0"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">'
    b"<b:Source><b:Tag>Kalogjera2020</b:Tag></b:Source></b:Sources>"
)


def _source(path, *, custom_xml=True, bibliography=True):
    content_types = CT_HEAD
    doc_rels = (
        '<?xml version="1.0"?><Relationships '
        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    )
    parts: dict[str, bytes] = {}

    if custom_xml:
        content_types += (
            '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-'
            'officedocument.customXmlProperties+xml"/>'
        )
        parts["customXml/item1.xml"] = CUSTOM_XML
        parts["customXml/itemProps1.xml"] = CUSTOM_XML_PROPS
        # A real customXml item reaches its properties part through its own
        # .rels; without that the props part is an orphan even in the source.
        parts["customXml/_rels/item1.xml.rels"] = (
            b'<?xml version="1.0"?><Relationships '
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Target="itemProps1.xml" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"/>'
            b"</Relationships>"
        )
        doc_rels += (
            '<Relationship Id="rId5" Target="../customXml/item1.xml" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"/>'
        )
    if bibliography:
        content_types += (
            '<Override PartName="/word/bibliography.xml" ContentType="application/vnd.openxmlformats-'
            'officedocument.bibliography+xml"/>'
        )
        parts["word/bibliography.xml"] = BIBLIOGRAPHY
        doc_rels += (
            '<Relationship Id="rId6" Target="bibliography.xml" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography"/>'
        )

    content_types += "</Types>"
    doc_rels += "</Relationships>"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", doc_rels)
        archive.writestr("word/document.xml", DOC)
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


def _rebuild(tmp_path, source):
    result = RebuildService(app_root=tmp_path / "app").rebuild(
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


# --- the parser has to see them at all ---------------------------------------

def test_the_parser_captures_a_custom_xml_store(tmp_path):
    model = DocxParser().parse(_source(tmp_path / "src.docx"))

    assert "customXml/item1.xml" in model.preserved_parts
    assert model.preserved_parts["customXml/item1.xml"].data == CUSTOM_XML


def test_the_parser_captures_bibliography_sources(tmp_path):
    model = DocxParser().parse(_source(tmp_path / "src.docx"))

    assert "word/bibliography.xml" in model.preserved_parts


def test_a_captured_attachment_records_how_it_is_reached(tmp_path):
    # Without the relationship type the renderer could only produce an orphan
    # part, which is worse than dropping it.
    model = DocxParser().parse(_source(tmp_path / "src.docx"))

    assert model.preserved_parts["customXml/item1.xml"].relationship_type.endswith("/customXml")
    assert model.preserved_parts["word/bibliography.xml"].relationship_type.endswith("/bibliography")


# --- and the renderer has to write them back ---------------------------------

def test_the_custom_xml_store_survives_a_rebuild(tmp_path):
    output = _rebuild(tmp_path, _source(tmp_path / "src.docx"))

    with ZipFile(output) as archive:
        names = archive.namelist()
        assert "customXml/item1.xml" in names
        assert archive.read("customXml/item1.xml") == CUSTOM_XML


def test_the_bibliography_survives_a_rebuild(tmp_path):
    output = _rebuild(tmp_path, _source(tmp_path / "src.docx"))

    with ZipFile(output) as archive:
        assert archive.read("word/bibliography.xml") == BIBLIOGRAPHY


def test_a_restored_attachment_is_reachable_rather_than_orphaned(tmp_path):
    # An OPC part nothing references is not preserved content, it is litter.
    from word_replica.qa.preservation import g10_projection

    output = _rebuild(tmp_path, _source(tmp_path / "src.docx"))
    projection = g10_projection(output)

    assert "customXml" in projection["relationship_graph"]
    assert "bibliography" in projection["relationship_graph"]
    assert projection["dangling_relationships"] == []


def test_g10_no_longer_reports_the_attachments_as_lost(tmp_path):
    # The synthetic source above is deliberately minimal -- no styles.xml, no
    # settings.xml -- so a full G10 pass would be measuring the shell template
    # rather than this fix. What this asserts is the specific loss G10 found:
    # the customXml store and the bibliography are no longer missing from the
    # output, and neither is left dangling.
    from word_replica.qa.preservation import g10_projection

    source = _source(tmp_path / "src.docx")
    before = g10_projection(source)
    after = g10_projection(_rebuild(tmp_path, source))

    assert before["opaque_parts"] == after["opaque_parts"]
    for relationship in ("customXml", "bibliography"):
        assert after["relationship_graph"][relationship] == before["relationship_graph"][relationship]
    assert after["dangling_relationships"] == []


def test_a_realistic_document_with_attachments_passes_g10(tmp_path, corpus_dir):
    # Same assertion against a document that has the full set of ordinary parts,
    # so the shell template contributes nothing of its own.
    import shutil
    from zipfile import ZIP_DEFLATED, ZipFile

    base = corpus_dir / "01_plain_text.docx"
    enriched = tmp_path / "enriched.docx"
    with ZipFile(base) as src:
        members = {name: src.read(name) for name in src.namelist()}
    members["word/bibliography.xml"] = BIBLIOGRAPHY
    members["[Content_Types].xml"] = members["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/word/bibliography.xml" ContentType="application/vnd.openxmlformats-'
        b'officedocument.bibliography+xml"/></Types>',
    )
    members["word/_rels/document.xml.rels"] = members["word/_rels/document.xml.rels"].replace(
        b"</Relationships>",
        b'<Relationship Id="rIdBib99" Target="bibliography.xml" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography"/>'
        b"</Relationships>",
    )
    with ZipFile(enriched, "w", ZIP_DEFLATED) as out:
        for name, data in members.items():
            out.writestr(name, data)

    gate = build_preservation_gate(enriched, _rebuild(tmp_path, enriched))

    assert gate.passed is True, f"{gate.summary} {gate.first_divergence}"


def test_a_document_without_attachments_does_not_gain_any(tmp_path):
    source = _source(tmp_path / "bare.docx", custom_xml=False, bibliography=False)
    output = _rebuild(tmp_path, source)

    with ZipFile(output) as archive:
        assert not [n for n in archive.namelist() if n.lower().startswith("customxml/")]
        assert "word/bibliography.xml" not in archive.namelist()


def test_a_body_referenced_part_nothing_refers_to_is_reported(tmp_path):
    """Successor to the transfer-refused warning.

    Charts, diagrams and embedded objects used to be dropped, because a rebuilt
    body could not carry the reference and the part alone would be an orphan.
    Verbatim fragment preservation removed that constraint, so they are now
    restored and the surviving reference points at them.

    What still deserves reporting is the case this test builds: a part with no
    fragment referring to it anywhere. Its bytes would travel with the document
    while nothing displayed them.
    """
    from word_replica.domain.model import PreservedPart
    from word_replica.renderers.pure_docx import PureDocxRenderer

    model = DocxParser().parse(_source(tmp_path / "src.docx"))
    model.preserved_parts["word/charts/chart1.xml"] = PreservedPart(
        "word/charts/chart1.xml", "application/xml", None, "0" * 64, b"<chart/>"
    )
    renderer = PureDocxRenderer()
    renderer.render(model, tmp_path / "out.docx", None)

    codes = {warning.code for warning in renderer._package.warnings}
    assert "PURE_DOCX_ASSET_UNREFERENCED" in codes
