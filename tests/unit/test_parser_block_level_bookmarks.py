"""A bookmark need not be inside a paragraph.

`w:bookmarkStart` and `w:bookmarkEnd` are allowed wherever block-level content
is: as direct children of `w:body`, between paragraphs and tables, or inside a
`w:sdtContent`. Bookmark paths were anchored to the enclosing paragraph, so
these had no anchor at all and fell back to the raw lxml path --
`/w:document/w:body/w:bookmarkEnd` -- which carries no position the renderer can
use. Every one of them was dumped into the first paragraph.

It is the whole remaining G7 class in the corpus: 13 of 400 documents, 11 whose
end sits at body level and 2 whose start sits inside a content control.

The anchor is the nearest paragraph, and which side matters. A start standing
before paragraph N begins the bookmark at that paragraph, so it anchors
forward; an end standing after paragraph N closes it there, so it anchors back.
Either way the bookmark covers the same text it covered before -- which is the
only thing about a bookmark that a cross-reference can see.

Paragraphs are counted across the whole document, content controls included,
because the renderer flattens a content control into the paragraphs it holds
and the count has to survive that.
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
from word_replica.qa.golden_audit import build_model_gates
from word_replica.services.rebuild import RebuildService

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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


def _p(text):
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def _start(bookmark_id, name):
    return f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{name}"/>'


def _end(bookmark_id):
    return f'<w:bookmarkEnd w:id="{bookmark_id}"/>'


def _write(path, body):
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}">'
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


def _marks_by_paragraph(path):
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    result = {}
    for index, para in enumerate(root.iter(f"{{{W}}}p"), start=1):
        marks = [
            (child.tag.rsplit("}", 1)[-1], child.get(f"{{{W}}}id"))
            for child in para
            if child.tag.rsplit("}", 1)[-1] in ("bookmarkStart", "bookmarkEnd")
        ]
        if marks:
            result[index] = marks
    return result


def _bookmark(path, name):
    return next(b for b in DocxParser().parse(path).bookmarks if b.name == name)


# --- an end that stands after a paragraph, at body level ------------------------

BODY_LEVEL = _p("one") + _start(0, "spans") + _p("two") + _end(0) + _p("three")


@pytest.fixture
def body_level(tmp_path):
    return _write(tmp_path / "body-level.docx", BODY_LEVEL)


def test_the_end_anchors_to_the_paragraph_it_closes_after(body_level):
    assert _bookmark(body_level, "spans").end_path.startswith("//w:p[2]")


def test_the_start_anchors_to_the_paragraph_it_opens_before(body_level):
    assert _bookmark(body_level, "spans").start_path.startswith("//w:p[2]")


def test_the_rebuild_puts_it_on_that_paragraph(tmp_path, body_level):
    marks = _marks_by_paragraph(_rebuild(tmp_path, body_level))

    assert sorted(marks) == [2], f"expected both marks on paragraph 2, got {marks}"


def test_g7_passes(tmp_path, body_level):
    parser = DocxParser()
    gates = build_model_gates(
        parser.parse(body_level), parser.parse(_rebuild(tmp_path, body_level, "-g7"))
    )

    assert gates["G7"].passed is True, gates["G7"].first_divergence


# --- a start inside a content control ------------------------------------------

@pytest.fixture
def in_sdt(tmp_path):
    body = (
        _p("before")
        + "<w:sdt><w:sdtContent>"
        + _start(0, "incontrol")
        + _p("inside")
        + _end(0)
        + "</w:sdtContent></w:sdt>"
        + _p("after")
    )
    return _write(tmp_path / "sdt.docx", body)


def test_a_start_inside_a_content_control_anchors_to_its_paragraph(in_sdt):
    # The renderer flattens the control into the paragraphs it holds, so the
    # paragraph count is what survives -- "inside" is the second paragraph.
    assert _bookmark(in_sdt, "incontrol").start_path.startswith("//w:p[2]")


def test_g7_passes_on_the_content_control(tmp_path, in_sdt):
    parser = DocxParser()
    gates = build_model_gates(
        parser.parse(in_sdt), parser.parse(_rebuild(tmp_path, in_sdt, "-sdt"))
    )

    assert gates["G7"].passed is True, gates["G7"].first_divergence


# --- edges ---------------------------------------------------------------------

def test_an_end_before_any_paragraph_anchors_to_the_first(tmp_path):
    source = _write(tmp_path / "leading.docx", _end(0) + _start(0, "odd") + _p("only"))

    assert _bookmark(source, "odd").end_path.startswith("//w:p[1]")


def test_a_start_after_every_paragraph_anchors_to_the_last(tmp_path):
    source = _write(tmp_path / "trailing.docx", _p("one") + _p("two") + _start(0, "tail") + _end(0))

    assert _bookmark(source, "tail").start_path.startswith("//w:p[2]")


def test_a_document_with_no_paragraphs_does_not_crash(tmp_path):
    source = _write(tmp_path / "empty.docx", _start(0, "nowhere") + _end(0))

    assert DocxParser().parse(source) is not None


# --- what must not change ------------------------------------------------------

def test_a_bookmark_inside_a_paragraph_is_unaffected(tmp_path):
    source = _write(
        tmp_path / "inside.docx",
        _p("one") + f"<w:p>{_start(0, 'plain')}{_end(0)}<w:r><w:t>two</w:t></w:r></w:p>",
    )

    bookmark = _bookmark(source, "plain")
    assert bookmark.start_path.startswith("//w:p[2]")
    assert sorted(_marks_by_paragraph(_rebuild(tmp_path, source, "-inside"))) == [2]
