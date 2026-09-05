from lxml import etree

from word_replica.parser.tables import parse_table_properties
from word_replica.parser.text import W_NS

_NSMAP = {"w": W_NS}


def _table(xml: bytes):
    return etree.fromstring(xml)


def test_grid_col_without_width_attribute_parses_to_none_not_zero():
    # A gridCol with no w:w is a declared-but-unsized column (autofit), not a
    # zero-width one. Silently defaulting to 0 previously fed a fabricated
    # 0-twips width into Word's COM Columns(n).Width setter, which rejects it
    # outright with a "Value out of range" automation error.
    node = _table(
        b"<w:tbl xmlns:w='" + W_NS.encode() + b"'>"
        b"<w:tblGrid><w:gridCol w:w='1000'/><w:gridCol/><w:gridCol w:w='2000'/></w:tblGrid>"
        b"</w:tbl>"
    )
    props = parse_table_properties(node)
    assert props["grid_column_widths"] == [1000, None, 2000]


def test_grid_col_widths_all_present_parses_as_before():
    node = _table(
        b"<w:tbl xmlns:w='" + W_NS.encode() + b"'>"
        b"<w:tblGrid><w:gridCol w:w='1000'/><w:gridCol w:w='2000'/></w:tblGrid>"
        b"</w:tbl>"
    )
    props = parse_table_properties(node)
    assert props["grid_column_widths"] == [1000, 2000]
