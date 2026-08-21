"""An inline content control still has content.

`parse_paragraph` dispatches on the paragraph's direct children: runs,
bookmarks, and the wrappers it unwraps to reach runs -- w:ins, w:moveTo,
w:fldSimple, w:hyperlink. A `w:sdt` sitting inline in a paragraph matched none
of them, so the control and everything inside it was dropped without a trace:
its text, its formatting, and any fragment its runs carried.

Found as tdf170516_drawingBeforePlainText.docx, where a shape lives in a run
inside `w:sdtContent`:

    <w:sdt><w:sdtContent><w:r><mc:AlternateContent>
      <mc:Choice Requires="wps"><w:drawing>...

G10 reported alternate_content_count going from one to none and choice_requires
losing wps -- the shape was gone, and so was the text of every other inline
control in the corpus, silently, because no gate models a content control.

Unwrapping is what the renderer already does with a block-level control, whose
paragraphs it flattens; this makes the inline case behave the same way. Only
w:sdtContent is searched: w:sdtPr holds the control's own definition, and a run
in there is not document text.
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
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
VML = "urn:schemas-microsoft-com:vml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Target="word/document.xml" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")


def _write(path, body):
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:mc="{MC}" '
        f'xmlns:v="{VML}" xmlns:r="{REL}" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        f'mc:Ignorable="wps">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            b'<?xml version="1.0"?><Relationships '
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        )
        archive.writestr("word/document.xml", document)
    return path


INLINE_TEXT = (
    "<w:p><w:r><w:t>before </w:t></w:r>"
    "<w:sdt>"
    '<w:sdtPr><w:alias w:val="pick one"/></w:sdtPr>'
    "<w:sdtContent><w:r><w:t>inside the control</w:t></w:r></w:sdtContent>"
    "</w:sdt>"
    "<w:r><w:t> after</w:t></w:r></w:p>"
)

SHAPE_IN_CONTROL = (
    "<w:p><w:sdt><w:sdtContent><w:r>"
    "<mc:AlternateContent>"
    '<mc:Choice Requires="wps">'
    '<w:pict><v:shape id="s1" style="width:20pt;height:20pt"><v:fill on="f"/></v:shape></w:pict>'
    "</mc:Choice>"
    '<mc:Fallback><w:pict><v:rect id="r1" style="width:20pt;height:20pt"/></w:pict></mc:Fallback>'
    "</mc:AlternateContent>"
    "</w:r></w:sdtContent></w:sdt></w:p>"
)


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


# --- text ----------------------------------------------------------------------

def test_the_text_inside_an_inline_control_survives(tmp_path):
    source = _write(tmp_path / "text.docx", INLINE_TEXT)

    assert "inside the control" in DocxParser().parse(source).plain_text()


def test_it_keeps_its_place_between_the_runs_around_it(tmp_path):
    source = _write(tmp_path / "order.docx", INLINE_TEXT)

    assert DocxParser().parse(source).plain_text().strip() == "before inside the control after"


def test_the_text_reaches_the_rebuild(tmp_path):
    source = _write(tmp_path / "rebuild.docx", INLINE_TEXT)

    output = _rebuild(tmp_path, source)

    assert "inside the control" in DocxParser().parse(output).plain_text()


# --- a fragment carried by a run inside the control -----------------------------

def test_a_shape_inside_an_inline_control_survives(tmp_path):
    source = _write(tmp_path / "shape.docx", SHAPE_IN_CONTROL)

    output = _rebuild(tmp_path, source, "-shape")

    with ZipFile(output) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
    assert document.find(f".//{{{MC}}}AlternateContent") is not None


def test_g10_still_counts_the_alternate_content(tmp_path):
    source = _write(tmp_path / "g10.docx", SHAPE_IN_CONTROL)

    before = g10_projection(source)
    after = g10_projection(_rebuild(tmp_path, source, "-g10"))

    assert (
        after["markup_compatibility"]["alternate_content_count"]
        == before["markup_compatibility"]["alternate_content_count"]
    )
    assert after["markup_compatibility"]["choice_requires"] == before["markup_compatibility"][
        "choice_requires"
    ]


# --- what must not be harvested -------------------------------------------------

def test_a_run_in_the_control_definition_is_not_document_text(tmp_path):
    # w:sdtPr describes the control -- a placeholder caption there is not text
    # the document contains.
    body = (
        "<w:p><w:sdt>"
        "<w:sdtPr><w:placeholder><w:r><w:t>CAPTION</w:t></w:r></w:placeholder></w:sdtPr>"
        "<w:sdtContent><w:r><w:t>real</w:t></w:r></w:sdtContent>"
        "</w:sdt></w:p>"
    )
    source = _write(tmp_path / "sdtpr.docx", body)

    text = DocxParser().parse(source).plain_text()

    assert "real" in text
    assert "CAPTION" not in text


def test_a_paragraph_without_a_control_is_unaffected(tmp_path):
    source = _write(tmp_path / "plain.docx", "<w:p><w:r><w:t>just text</w:t></w:r></w:p>")

    assert DocxParser().parse(source).plain_text().strip() == "just text"
