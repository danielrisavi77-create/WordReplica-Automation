"""Table geometry has to come back, not just table content.

G3 is the largest visible defect class the lab measures: 82 of 137 model-gate
failures across 400 documents. Sampling forty of them showed one shape --
properties the parser reads and the renderer never writes:

    grid_column_widths  18    expected [1710]   actual <missing>
    borders              8
    alignment            5
    cell_margins         4

Losing w:tblGrid is the worst of them. Without it Word has no column widths to
lay the table out from and falls back to its own guess, so a table that was
authored with a narrow label column and a wide value column comes back evenly
split -- content intact, table unrecognisable.

Element order matters here in a way it does not elsewhere in this renderer.
w:tblPr has a required child sequence, and Word repairs a document whose
properties arrive out of order, so a fix that writes everything in the wrong
order is worse than one that writes nothing.
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

TABLE = (
    "<w:tbl>"
    "<w:tblPr>"
    '<w:tblStyle w:val="TableGrid"/>'
    '<w:tblW w:w="9000" w:type="dxa"/>'
    '<w:jc w:val="center"/>'
    "<w:tblBorders>"
    '<w:top w:val="single" w:sz="4" w:color="FF0000"/>'
    '<w:left w:val="dashed" w:sz="8" w:color="00FF00"/>'
    "</w:tblBorders>"
    '<w:tblLayout w:type="fixed"/>'
    "<w:tblCellMar>"
    '<w:left w:w="120" w:type="dxa"/>'
    '<w:right w:w="240" w:type="dxa"/>'
    "</w:tblCellMar>"
    "</w:tblPr>"
    '<w:tblGrid><w:gridCol w:w="1710"/><w:gridCol w:w="7290"/></w:tblGrid>'
    "<w:tr>"
    '<w:tc><w:tcPr><w:tcW w:w="1710" w:type="dxa"/><w:shd w:fill="EEEEEE"/></w:tcPr>'
    "<w:p><w:r><w:t>label</w:t></w:r></w:p></w:tc>"
    '<w:tc><w:tcPr><w:tcW w:w="7290" w:type="dxa"/></w:tcPr>'
    "<w:p><w:r><w:t>value</w:t></w:r></w:p></w:tc>"
    "</w:tr></w:tbl>"
)


@pytest.fixture
def source(tmp_path):
    document = (
        '<?xml version="1.0"?><w:document '
        f'xmlns:w="{W}">'
        f"<w:body>{TABLE}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    path = tmp_path / "table.docx"
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


def _table(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/document.xml")).find(f".//{{{W}}}tbl")


def test_the_column_widths_come_back(tmp_path, source):
    grid = _table(_rebuild(tmp_path, source)).find(f"{{{W}}}tblGrid")

    assert grid is not None, "w:tblGrid missing; Word has nothing to lay the table out from"
    assert [col.get(f"{{{W}}}w") for col in grid] == ["1710", "7290"]


def test_the_table_style_comes_back(tmp_path, source):
    tbl_pr = _table(_rebuild(tmp_path, source, "-style")).find(f"{{{W}}}tblPr")

    assert tbl_pr.find(f"{{{W}}}tblStyle").get(f"{{{W}}}val") == "TableGrid"


def test_the_borders_come_back(tmp_path, source):
    borders = _table(_rebuild(tmp_path, source, "-borders")).find(f"{{{W}}}tblPr/{{{W}}}tblBorders")

    assert borders is not None
    top = borders.find(f"{{{W}}}top")
    assert (top.get(f"{{{W}}}val"), top.get(f"{{{W}}}sz"), top.get(f"{{{W}}}color")) == (
        "single", "4", "FF0000",
    )


def test_the_alignment_comes_back(tmp_path, source):
    tbl_pr = _table(_rebuild(tmp_path, source, "-jc")).find(f"{{{W}}}tblPr")

    assert tbl_pr.find(f"{{{W}}}jc").get(f"{{{W}}}val") == "center"


def test_the_cell_margins_come_back(tmp_path, source):
    margins = _table(_rebuild(tmp_path, source, "-mar")).find(f"{{{W}}}tblPr/{{{W}}}tblCellMar")

    assert margins is not None
    assert margins.find(f"{{{W}}}left").get(f"{{{W}}}w") == "120"


def test_cell_widths_and_shading_come_back(tmp_path, source):
    row = _table(_rebuild(tmp_path, source, "-cells")).find(f"{{{W}}}tr")
    first = row.find(f"{{{W}}}tc/{{{W}}}tcPr")

    assert first.find(f"{{{W}}}tcW").get(f"{{{W}}}w") == "1710"
    assert first.find(f"{{{W}}}shd").get(f"{{{W}}}fill") == "EEEEEE"


def test_table_properties_are_written_in_the_order_the_schema_requires(tmp_path, source):
    # w:tblPr has a required child sequence and Word repairs a document whose
    # properties arrive out of order, so writing them in the wrong order would
    # be worse than not writing them at all.
    tbl_pr = _table(_rebuild(tmp_path, source, "-order")).find(f"{{{W}}}tblPr")
    order = [child.tag.rsplit("}", 1)[-1] for child in tbl_pr]

    expected_sequence = ["tblStyle", "tblW", "jc", "tblBorders", "tblLayout", "tblCellMar"]
    assert order == [name for name in expected_sequence if name in order]


def test_the_grid_follows_the_properties_and_precedes_the_rows(tmp_path, source):
    children = [child.tag.rsplit("}", 1)[-1] for child in _table(_rebuild(tmp_path, source, "-seq"))]

    assert children[:3] == ["tblPr", "tblGrid", "tr"]


def test_g3_passes_on_this_table(tmp_path, source):
    output = _rebuild(tmp_path, source, "-g3")
    parser = DocxParser()

    gates = build_model_gates(parser.parse(source), parser.parse(output))

    assert gates["G3"].passed is True, gates["G3"].first_divergence


def test_a_table_with_no_properties_gains_none(tmp_path):
    document = (
        '<?xml version="1.0"?><w:document '
        f'xmlns:w="{W}"><w:body>'
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        "<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    path = tmp_path / "bare.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            b'<?xml version="1.0"?><Relationships '
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        )
        archive.writestr("word/document.xml", document)

    table = _table(_rebuild(tmp_path, path, "-bare"))
    tbl_pr = table.find(f"{{{W}}}tblPr")

    assert tbl_pr is None or len(tbl_pr) == 0
