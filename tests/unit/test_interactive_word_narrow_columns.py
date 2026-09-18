from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from word_replica.renderers.interactive_word import (
    _WORD_MIN_COM_COLUMN_WIDTH_POINTS,
    _restore_narrow_column_widths,
)

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _document_xml_with_one_table(placeholder_twips: int) -> bytes:
    return (
        b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        b'<w:document xmlns:w="' + _W_NS.encode() + b'">'
        b"<w:body><w:tbl>"
        b'<w:tblGrid><w:gridCol w:w="' + str(placeholder_twips).encode() + b'"/><w:gridCol w:w="2000"/></w:tblGrid>'
        b"<w:tr>"
        b'<w:tc><w:tcPr><w:tcW w:w="' + str(placeholder_twips).encode() + b'" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
        b'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
        b"</w:tr>"
        b"</w:tbl></w:body></w:document>"
    )


def test_restore_narrow_column_widths_corrects_grid_and_cell_width(tmp_path):
    docx_path = tmp_path / "sample.docx"
    placeholder_twips = int(_WORD_MIN_COM_COLUMN_WIDTH_POINTS * 20)
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", _document_xml_with_one_table(placeholder_twips))

    _restore_narrow_column_widths(docx_path, [(0, 1, 96)])

    with ZipFile(docx_path, "r") as archive:
        patched = archive.read("word/document.xml")
    root = etree.fromstring(patched)
    grid_widths = [el.get(f"{{{_W_NS}}}w") for el in root.iter(f"{{{_W_NS}}}gridCol")]
    cell_widths = [el.get(f"{{{_W_NS}}}w") for el in root.iter(f"{{{_W_NS}}}tcW")]
    assert grid_widths == ["96", "2000"]
    assert cell_widths == ["96", "2000"]


def test_restore_narrow_column_widths_is_a_noop_with_no_fixups(tmp_path):
    docx_path = tmp_path / "sample.docx"
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", b"<root/>")
    before = docx_path.read_bytes()

    _restore_narrow_column_widths(docx_path, [])

    assert docx_path.read_bytes() == before


def test_restore_narrow_column_widths_skips_out_of_range_table_sequence(tmp_path):
    docx_path = tmp_path / "sample.docx"
    placeholder_twips = int(_WORD_MIN_COM_COLUMN_WIDTH_POINTS * 20)
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", _document_xml_with_one_table(placeholder_twips))
    before = docx_path.read_bytes()

    _restore_narrow_column_widths(docx_path, [(5, 1, 96)])

    assert docx_path.read_bytes() == before
