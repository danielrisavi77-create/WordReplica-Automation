"""A preserved fragment may point outside the package.

Inline fragments the model cannot rebuild -- shapes, VML, embedded objects --
are carried verbatim, and every `r:id` inside them is rewritten to the part it
addressed so the id can be reallocated on the way out. A fragment holding an id
that cannot be resolved is refused outright, because a corrupted package is
worse than a missing shape.

`_reference_targets` resolved ids to *parts*, and skipped external
relationships entirely. An external target has no part by definition, so a
shape carrying a hyperlink resolved to nothing and the whole fragment was
refused: the shape, its geometry and its link all disappeared, and the refusal
looked like the safety rule working correctly.

Found as `grouped_link.docx` in the corpus -- a grouped shape whose rId7 is
`https://www.documentfoundation.org`, TargetMode External. G10 recorded five
namespaces lost (wpg, wps, wp, mc, vml), alternate_content_count 1 -> 0, and
both hyperlinks gone.

The safety rule itself must survive this: an id that resolves to neither a part
nor an external target is still refused.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from lxml import etree

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
)
from word_replica.parser.parser import DocxParser
from word_replica.qa.preservation import g10_projection
from word_replica.services.rebuild import RebuildService

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
VML = "urn:schemas-microsoft-com:vml"
TARGET = "https://www.documentfoundation.org"

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")

EXTERNAL_REL = (
    f'<Relationship Id="rId7" Type="{REL}/hyperlink" Target="{TARGET}" TargetMode="External"/>'
)

# A VML shape carrying a link: preserved verbatim, and the only thing inside it
# that needs resolving is external.
SHAPE = (
    "<w:p><w:r><w:pict>"
    f'<v:shape id="_x0000_s1026" style="width:100pt;height:20pt" o:button="t">'
    '<v:fill on="f"/>'
    '<v:path o:connecttype="none"/>'
    "</v:shape>"
    '<w10:wrap type="none"/>'
    "</w:pict></w:r></w:p>"
)
# The link lives on the shape itself, the way a shape hyperlink is authored.
LINKED_SHAPE = SHAPE.replace(
    '<v:fill on="f"/>', '<v:fill on="f"/><o:extrusion r:id="rId7"/>'
)


def _source(path, *, body, doc_rels=""):
    document = (
        '<?xml version="1.0"?><w:document '
        f'xmlns:w="{W}" xmlns:r="{REL}" xmlns:mc="{MC}" xmlns:v="{VML}" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    rels = (
        f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">{doc_rels}</Relationships>'
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", rels)
        archive.writestr("word/document.xml", document)
    return path


@pytest.fixture
def linked_shape(tmp_path):
    return _source(tmp_path / "shape.docx", body=LINKED_SHAPE, doc_rels=EXTERNAL_REL)


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


def _document(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/document.xml"))


def _rels(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/_rels/document.xml.rels"))


# --- capture ------------------------------------------------------------------

def test_the_fragment_is_captured_rather_than_refused(linked_shape):
    model = DocxParser().parse(linked_shape)

    kinds = [
        token.get("kind")
        for paragraph in model.iter_paragraphs()
        for run in paragraph.runs
        for token in run.properties.get("content_tokens") or []
    ]
    assert "preserved_xml" in kinds, "the shape was refused over its external link"
    assert "unsupported_inline" not in kinds


# --- re-emission --------------------------------------------------------------

def test_the_shape_survives(tmp_path, linked_shape):
    assert _document(_rebuild(tmp_path, linked_shape)).find(f".//{{{VML}}}shape") is not None


def test_the_external_target_comes_back(tmp_path, linked_shape):
    rels = _rels(_rebuild(tmp_path, linked_shape, "-target"))

    external = [node for node in rels if node.get("TargetMode") == "External"]
    assert [node.get("Target") for node in external] == [TARGET]


def test_the_reference_points_at_a_relationship_that_exists(tmp_path, linked_shape):
    output = _rebuild(tmp_path, linked_shape, "-rel")
    document, rels = _document(output), _rels(output)

    used = {
        value
        for element in document.iter()
        if isinstance(element.tag, str)
        for name, value in element.attrib.items()
        if name.startswith(f"{{{REL}}}")
    }
    declared = {node.get("Id") for node in rels}
    assert used, "the shape came back with no reference at all"
    assert used <= declared, f"dangling reference: {used - declared}"


def test_g10_no_longer_reports_the_shape_as_lost(tmp_path, linked_shape):
    before = g10_projection(linked_shape)
    after = g10_projection(_rebuild(tmp_path, linked_shape, "-g10"))

    assert VML in after["namespaces"], "the VML namespace disappeared with the shape"
    assert after["relationship_graph"].get("hyperlink") == before["relationship_graph"].get(
        "hyperlink"
    )
    assert after["dangling_relationships"] == []


# --- the safety rule this must not weaken -------------------------------------

def test_a_reference_that_resolves_to_nothing_is_still_refused(tmp_path):
    # rId7 is used but never declared: neither a part nor an external target.
    # Emitting it would produce a package Word repairs, so the fragment is
    # refused exactly as before.
    source = _source(tmp_path / "dangling.docx", body=LINKED_SHAPE, doc_rels="")

    model = DocxParser().parse(source)

    kinds = [
        token.get("kind")
        for paragraph in model.iter_paragraphs()
        for run in paragraph.runs
        for token in run.properties.get("content_tokens") or []
    ]
    assert "unsupported_inline" in kinds
    assert "preserved_xml" not in kinds


def test_a_fragment_with_no_references_is_unaffected(tmp_path):
    source = _source(tmp_path / "plain.docx", body=SHAPE)

    output = _rebuild(tmp_path, source, "-plain")

    assert _document(output).find(f".//{{{VML}}}shape") is not None
    assert [node for node in _rels(output) if node.get("TargetMode") == "External"] == []
