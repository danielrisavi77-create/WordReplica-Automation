"""Row properties, and the two cell properties the first table fix misread.

The table-geometry commit closed the largest G3 class, and the lab immediately
named what it had missed. Sampling forty table documents that still failed:

    rows/properties/cant_split       18   expected False   actual <missing>
    rows/cells/cell_margins           4   expected {...}   actual <missing>
    rows/cells/vertical_alignment     3   expected center  actual <missing>
    rows/cells/text_direction         1   expected btLr    actual <missing>

The last three are the same defect as before -- a property the parser reads and
the renderer never writes -- except two of them were written, under the wrong
key. `parser/tables.py` stores `cell_margins` and `vertical_alignment`; the
renderer asked for `margins` and `v_align`, so the lookups silently returned
None and the properties vanished. Reading a key that is never present is the
quietest way to lose data, which is why this asserts per property rather than
that "some tcPr came out".

`w:trPr` was not written at all. cant_split is the largest single class because
it is False whenever the source row has any w:trPr, so every row with a height
or a repeated header loses it too -- an empty w:trPr still has to be emitted to
carry that False back.
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

# A header row that repeats, must not split across pages, has a fixed height,
# and whose cell carries every property the sample showed being dropped.
TABLE = (
    "<w:tbl>"
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
    "<w:tr>"
    "<w:trPr>"
    "<w:cantSplit/>"
    '<w:trHeight w:val="720" w:hRule="atLeast"/>'
    "<w:tblHeader/>"
    "</w:trPr>"
    "<w:tc><w:tcPr>"
    '<w:tcW w:w="4000" w:type="dxa"/>'
    '<w:tcMar><w:left w:w="227" w:type="dxa"/><w:right w:w="341" w:type="dxa"/></w:tcMar>'
    '<w:textDirection w:val="btLr"/>'
    '<w:vAlign w:val="center"/>'
    "</w:tcPr><w:p><w:r><w:t>head</w:t></w:r></w:p></w:tc>"
    "</w:tr>"
    # A plain row whose w:trPr exists but is empty: cant_split is False here too,
    # so an empty w:trPr has to survive to say so.
    "<w:tr><w:trPr/><w:tc><w:p><w:r><w:t>body</w:t></w:r></w:p></w:tc></w:tr>"
    "</w:tbl>"
)

BARE_TABLE = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"


def _write(path, body):
    document = (
        '<?xml version="1.0"?><w:document '
        f'xmlns:w="{W}">'
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


@pytest.fixture
def source(tmp_path):
    return _write(tmp_path / "rows.docx", TABLE)


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


def _rows(path):
    with ZipFile(path) as archive:
        table = etree.fromstring(archive.read("word/document.xml")).find(f".//{{{W}}}tbl")
    return table.findall(f"{{{W}}}tr")


def _names(node):
    return [child.tag.rsplit("}", 1)[-1] for child in node]


# --- row properties -----------------------------------------------------------

def test_a_row_that_must_not_split_still_must_not(tmp_path, source):
    tr_pr = _rows(_rebuild(tmp_path, source))[0].find(f"{{{W}}}trPr")

    assert tr_pr is not None, "w:trPr missing; every row property was dropped"
    assert tr_pr.find(f"{{{W}}}cantSplit") is not None


def test_a_repeating_header_row_still_repeats(tmp_path, source):
    tr_pr = _rows(_rebuild(tmp_path, source, "-hdr"))[0].find(f"{{{W}}}trPr")

    assert tr_pr.find(f"{{{W}}}tblHeader") is not None


def test_the_row_height_comes_back(tmp_path, source):
    height = _rows(_rebuild(tmp_path, source, "-h"))[0].find(f"{{{W}}}trPr/{{{W}}}trHeight")

    assert height is not None
    assert (height.get(f"{{{W}}}val"), height.get(f"{{{W}}}hRule")) == ("720", "atLeast")


def test_an_empty_row_properties_element_survives(tmp_path, source):
    assert _rows(_rebuild(tmp_path, source, "-empty"))[1].find(f"{{{W}}}trPr") is not None


def test_row_properties_are_written_in_the_order_the_schema_requires(tmp_path, source):
    tr_pr = _rows(_rebuild(tmp_path, source, "-order"))[0].find(f"{{{W}}}trPr")

    assert _names(tr_pr) == ["cantSplit", "trHeight", "tblHeader"]


# --- the cell properties read under the wrong key ------------------------------

def test_the_cell_margins_come_back(tmp_path, source):
    tc_mar = _rows(_rebuild(tmp_path, source, "-mar"))[0].find(f"{{{W}}}tc/{{{W}}}tcPr/{{{W}}}tcMar")

    assert tc_mar is not None, "read under key 'margins'; the parser stores 'cell_margins'"
    assert tc_mar.find(f"{{{W}}}left").get(f"{{{W}}}w") == "227"


def test_the_vertical_alignment_comes_back(tmp_path, source):
    tc_pr = _rows(_rebuild(tmp_path, source, "-valign"))[0].find(f"{{{W}}}tc/{{{W}}}tcPr")
    v_align = tc_pr.find(f"{{{W}}}vAlign")

    assert v_align is not None, "read under key 'v_align'; the parser stores 'vertical_alignment'"
    assert v_align.get(f"{{{W}}}val") == "center"


def test_the_text_direction_comes_back(tmp_path, source):
    tc_pr = _rows(_rebuild(tmp_path, source, "-dir"))[0].find(f"{{{W}}}tc/{{{W}}}tcPr")

    assert tc_pr.find(f"{{{W}}}textDirection").get(f"{{{W}}}val") == "btLr"


def test_cell_properties_are_written_in_the_order_the_schema_requires(tmp_path, source):
    tc_pr = _rows(_rebuild(tmp_path, source, "-corder"))[0].find(f"{{{W}}}tc/{{{W}}}tcPr")

    assert _names(tc_pr) == ["tcW", "tcMar", "textDirection", "vAlign"]


# --- the gate the lab actually measures ---------------------------------------

def test_g3_passes_on_this_table(tmp_path, source):
    output = _rebuild(tmp_path, source, "-g3")
    parser = DocxParser()

    gates = build_model_gates(parser.parse(source), parser.parse(output))

    assert gates["G3"].passed is True, gates["G3"].first_divergence


def test_a_row_with_no_properties_gains_none(tmp_path):
    source = _write(tmp_path / "bare.docx", BARE_TABLE)

    assert _rows(_rebuild(tmp_path, source, "-bare"))[0].find(f"{{{W}}}trPr") is None
