"""A bookmark belongs where it was, not in the first paragraph.

`_inject_bookmarks_and_fields` put every bookmark into `paragraphs[0]`: the
start inserted after any `w:pPr`, the end appended to the same paragraph. No
position was consulted, though `Bookmark` has carried `start_path` and
`end_path` all along -- `//w:p[N]/...`, where N is the paragraph's place among
every `w:p` in the document.

The result is that a cross-reference points at the wrong text. Measured on
tdf125298_crossreflink_nonascii_charlimit.docx, a document whose whole purpose
is cross-reference links:

    source  p[1] bookmarkStart 0, bookmarkEnd 0
            p[3] bookmarkStart 1, bookmarkEnd 1
    output  p[1] bookmarkStart 0, bookmarkStart 1, bookmarkEnd 0, bookmarkEnd 1

Bookmark 1 moved from the third paragraph to the first, so a REF to it resolves
to the wrong sentence and a page reference to the wrong page.

It went unnoticed because most documents keep their only bookmark in the first
paragraph anyway, where being put back is indistinguishable from being placed.
It is the whole G7 class in the corpus: 21 of 260 documents, every one
`bookmarks/end_path`, expected `//w:p[1]/w:bookmarkEnd` and got
`//w:p[1]/w:bookmarkEnd[1]` -- lxml only writes that index when a second
bookmarkEnd has arrived in the same paragraph.

Paragraphs are counted the way the parser counts them: every w:p in document
order, including those inside table cells, so a bookmark in a table lands in
the right cell rather than N paragraphs early.
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


def _paragraph(text, *, marks=""):
    return f"<w:p>{marks}<w:r><w:t>{text}</w:t></w:r></w:p>"


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
    """{paragraph index (1-based, all w:p in document order): [(tag, id)]}"""
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


# --- the shape the corpus showed ----------------------------------------------

SPREAD = (
    _paragraph("first", marks=_start(0, "one") + _end(0))
    + _paragraph("second")
    + _paragraph("third", marks=_start(1, "two") + _end(1))
)


@pytest.fixture
def spread(tmp_path):
    return _write(tmp_path / "spread.docx", SPREAD)


def test_a_bookmark_in_the_third_paragraph_comes_back_there(tmp_path, spread):
    marks = _marks_by_paragraph(_rebuild(tmp_path, spread))

    assert 3 in marks, f"bookmark left the third paragraph; marks landed at {sorted(marks)}"
    assert ("bookmarkStart", "1") in marks[3]
    assert ("bookmarkEnd", "1") in marks[3]


def test_the_bookmarks_do_not_collapse_into_the_first_paragraph(tmp_path, spread):
    marks = _marks_by_paragraph(_rebuild(tmp_path, spread, "-collapse"))

    assert sorted(marks) == [1, 3]
    assert len(marks[1]) == 2


def test_g7_passes_on_a_document_whose_bookmarks_are_spread_out(tmp_path, spread):
    output = _rebuild(tmp_path, spread, "-g7")
    parser = DocxParser()

    gates = build_model_gates(parser.parse(spread), parser.parse(output))

    assert gates["G7"].passed is True, gates["G7"].first_divergence


# --- a bookmark that spans paragraphs ------------------------------------------

def test_a_bookmark_spanning_paragraphs_keeps_both_ends(tmp_path):
    source = _write(
        tmp_path / "span.docx",
        _paragraph("intro")
        + _paragraph("start here", marks=_start(0, "range"))
        + _paragraph("middle")
        + _paragraph("end here", marks=_end(0)),
    )

    marks = _marks_by_paragraph(_rebuild(tmp_path, source, "-span"))

    assert marks.get(2) == [("bookmarkStart", "0")]
    assert marks.get(4) == [("bookmarkEnd", "0")]


# --- paragraphs inside tables count too ----------------------------------------

def test_a_bookmark_inside_a_table_cell_lands_in_that_cell(tmp_path):
    # The parser numbers every w:p in document order, table cells included, so
    # the renderer has to count the same way or the bookmark lands early.
    table = (
        "<w:tbl><w:tr>"
        f"<w:tc>{_paragraph('cell one')}</w:tc>"
        f"<w:tc>{_paragraph('cell two', marks=_start(0, 'incell') + _end(0))}</w:tc>"
        "</w:tr></w:tbl>"
    )
    source = _write(tmp_path / "table.docx", _paragraph("before") + table + _paragraph("after"))

    marks = _marks_by_paragraph(_rebuild(tmp_path, source, "-table"))

    assert sorted(marks) == [3], f"expected the third paragraph (the second cell), got {sorted(marks)}"


# --- the parts that must not change --------------------------------------------

def test_a_bookmark_already_in_the_first_paragraph_stays_there(tmp_path):
    source = _write(tmp_path / "first.docx", _paragraph("only", marks=_start(0, "top") + _end(0)))

    assert sorted(_marks_by_paragraph(_rebuild(tmp_path, source, "-first"))) == [1]


def test_a_document_without_bookmarks_gains_none(tmp_path):
    source = _write(tmp_path / "none.docx", _paragraph("plain"))

    assert _marks_by_paragraph(_rebuild(tmp_path, source, "-none")) == {}


def test_the_body_text_is_unchanged_by_the_placement(tmp_path, spread):
    output = _rebuild(tmp_path, spread, "-text")

    assert DocxParser().parse(output).plain_text().split() == ["first", "second", "third"]


def test_a_bookmark_in_an_empty_paragraph_stays_in_it(tmp_path):
    # An lxml element with no children is falsy, so resolving the host with
    # `or` would quietly send this bookmark back to the first paragraph.
    source = _write(
        tmp_path / "empty.docx",
        _paragraph("before") + f"<w:p>{_start(0, 'blank')}{_end(0)}</w:p>" + _paragraph("after"),
    )

    assert sorted(_marks_by_paragraph(_rebuild(tmp_path, source, "-empty"))) == [2]


def test_nested_bookmarks_close_in_the_order_the_source_closed_them(tmp_path):
    # Nested bookmarks close inside-out: start A, start B, end B, end A. Writing
    # the ends in the order the bookmarks were declared would close A first and
    # change what each one covers.
    source = _write(
        tmp_path / "nested.docx",
        _paragraph("outer start", marks=_start(0, "outer"))
        + _paragraph("inner", marks=_start(1, "inner"))
        + f"<w:p>{_end(1)}{_end(0)}<w:r><w:t>both close</w:t></w:r></w:p>",
    )

    marks = _marks_by_paragraph(_rebuild(tmp_path, source, "-nested"))

    assert marks.get(3) == [("bookmarkEnd", "1"), ("bookmarkEnd", "0")]


def test_g7_passes_on_nested_bookmarks(tmp_path):
    source = _write(
        tmp_path / "nested-g7.docx",
        _paragraph("outer start", marks=_start(0, "outer"))
        + _paragraph("inner", marks=_start(1, "inner"))
        + f"<w:p>{_end(1)}{_end(0)}<w:r><w:t>both close</w:t></w:r></w:p>",
    )
    parser = DocxParser()

    gates = build_model_gates(parser.parse(source), parser.parse(_rebuild(tmp_path, source, "-ng7")))

    assert gates["G7"].passed is True, gates["G7"].first_divergence


def test_goback_does_not_shift_the_paths_of_the_bookmarks_that_are_kept(tmp_path):
    # The parser drops _GoBack, Word's volatile "last edit position" mark. Its
    # bookmarkEnd is still a sibling in the source, so a path indexed among all
    # siblings says [2] for a bookmark whose rebuild can only ever write [1] --
    # a divergence describing a bookmark the model never carried.
    source = _write(
        tmp_path / "goback.docx",
        _paragraph("one", marks=_start(0, "real"))
        + f"<w:p>{_start(1, '_GoBack')}{_end(1)}{_end(0)}<w:r><w:t>two</w:t></w:r></w:p>",
    )
    parser = DocxParser()

    gates = build_model_gates(parser.parse(source), parser.parse(_rebuild(tmp_path, source, "-gb")))

    assert gates["G7"].passed is True, gates["G7"].first_divergence
