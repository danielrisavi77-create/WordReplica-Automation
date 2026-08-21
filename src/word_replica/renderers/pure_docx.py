from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from typing import Any
import os
import re
import shutil
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from docx import Document
from lxml import etree

from word_replica.domain.model import DocumentModel, Paragraph, Run, Section, Table
from word_replica.domain.results import WarningItem
from word_replica.renderers.base import RenderResult

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
CUST_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
W = f"{{{W_NS}}}"


def _w(tag: str):
    return etree.Element(f"{W}{tag}")


def _set_w(node, name: str, value: str | int | bool | None) -> None:
    if value is not None:
        node.set(f"{W}{name}", str(value))


def _bool_element(parent, tag: str, value) -> None:
    if value is None:
        return
    node = etree.SubElement(parent, f"{W}{tag}")
    if value is False:
        _set_w(node, "val", "0")


_RUN_FONT_ATTRS = (
    ("ascii", "font_ascii"), ("hAnsi", "font_hansi"),
    ("eastAsia", "font_east_asia"), ("cs", "font_cs"),
    ("asciiTheme", "font_ascii_theme"), ("hAnsiTheme", "font_hansi_theme"),
    ("eastAsiaTheme", "font_east_asia_theme"), ("cstheme", "font_cs_theme"),
)
_RUN_PR_KEYS = (
    "bold", "italic", "underline", "strike",
    "font_ascii", "font_hansi", "font_east_asia", "font_cs",
    "font_ascii_theme", "font_hansi_theme", "font_east_asia_theme", "font_cs_theme",
    "color", "highlight", "vert_align", "size_half_points",
    "language", "language_east_asia", "language_bidi",
    "character_spacing", "character_position",
)


HYPERLINK_REFERENCE_PREFIX = "wr-link:"


def _append_runs_grouping_hyperlinks(node, runs) -> None:
    """Emit runs, wrapping consecutive linked ones back in w:hyperlink.

    The parser flattens w:hyperlink to reach the runs inside, so the grouping
    has to be rebuilt from what each run remembers. Consecutive runs sharing a
    group belong to one link; a run with no link ends it.

    An external target cannot become a relationship id yet -- ids are allocated
    once the whole body exists -- so it goes out as a placeholder and is
    resolved with the rest.
    """
    current = None
    container = node
    for run in runs:
        link = run.properties.get("hyperlink") or None
        key = link.get("group") if link else None
        if link is None or key != current:
            current = key
            container = node
            if link is not None:
                container = etree.SubElement(node, f"{W}hyperlink")
                if link.get("anchor"):
                    _set_w(container, "anchor", link["anchor"])
                if link.get("tooltip"):
                    _set_w(container, "tooltip", link["tooltip"])
                if link.get("target"):
                    container.set(
                        f"{{{R_NS}}}id", HYPERLINK_REFERENCE_PREFIX + link["target"]
                    )
        container.append(_run_element(run))


def _run_element(run: Run):
    node = _w("r")
    r_pr = None
    props = run.properties
    if any(key in props for key in _RUN_PR_KEYS) or run.hidden:
        r_pr = etree.SubElement(node, f"{W}rPr")
        if any(props.get(key) is not None for _, key in _RUN_FONT_ATTRS):
            r_fonts = etree.SubElement(r_pr, f"{W}rFonts")
            for attr, key in _RUN_FONT_ATTRS:
                _set_w(r_fonts, attr, props.get(key))
        _bool_element(r_pr, "b", props.get("bold"))
        _bool_element(r_pr, "i", props.get("italic"))
        _bool_element(r_pr, "u", props.get("underline"))
        _bool_element(r_pr, "vanish", run.hidden or props.get("hidden"))
        _bool_element(r_pr, "strike", props.get("strike"))
        if props.get("color") is not None:
            _set_w(etree.SubElement(r_pr, f"{W}color"), "val", props["color"])
        if props.get("character_spacing") is not None:
            _set_w(etree.SubElement(r_pr, f"{W}spacing"), "val", props["character_spacing"])
        if props.get("character_position") is not None:
            _set_w(etree.SubElement(r_pr, f"{W}position"), "val", props["character_position"])
        if props.get("size_half_points") is not None:
            _set_w(etree.SubElement(r_pr, f"{W}sz"), "val", props["size_half_points"])
            _set_w(etree.SubElement(r_pr, f"{W}szCs"), "val", props["size_half_points"])
        if props.get("highlight") is not None:
            _set_w(etree.SubElement(r_pr, f"{W}highlight"), "val", props["highlight"])
        if props.get("vert_align") is not None:
            _set_w(etree.SubElement(r_pr, f"{W}vertAlign"), "val", props["vert_align"])
        if any(props.get(key) is not None for key in ("language", "language_east_asia", "language_bidi")):
            lang = etree.SubElement(r_pr, f"{W}lang")
            _set_w(lang, "val", props.get("language"))
            _set_w(lang, "eastAsia", props.get("language_east_asia"))
            _set_w(lang, "bidi", props.get("language_bidi"))
    content_tokens = props.get("content_tokens") or ()
    has_field_tokens = any(token.get("kind") in _FIELD_TOKEN_KINDS for token in content_tokens)
    # A drawing token carries both: "kind" tells the interactive executor what
    # to do, "value" carries the markup this renderer has to write itself.
    preserved = [token for token in content_tokens if token.get("value") and token.get("kind") in _PRESERVED_KINDS]
    references = [token for token in content_tokens if token.get("kind") in _REFERENCE_TOKENS]
    if not run.text and not has_field_tokens and not preserved and not references:
        return node
    buffer = ""

    def flush() -> None:
        nonlocal buffer
        if buffer:
            t = etree.SubElement(node, f"{W}t")
            if buffer[:1].isspace() or buffer[-1:].isspace():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = buffer
            buffer = ""

    break_types = iter(props.get("break_types", []))
    for char in run.text:
        if char == "\t":
            flush()
            etree.SubElement(node, f"{W}tab")
        elif char == "\n":
            flush()
            br = etree.SubElement(node, f"{W}br")
            break_type = next(break_types, "line")
            if break_type not in {"line", "textWrapping"}:
                _set_w(br, "type", break_type)
        else:
            buffer += char
    flush()
    if has_field_tokens:
        _append_field_tokens(node, content_tokens)
    for token in references:
        tag, key = _REFERENCE_TOKENS[token["kind"]]
        marker = etree.SubElement(node, f"{W}{tag}")
        _set_w(marker, "id", str(token.get(key)))
    _append_preserved_inline(node, preserved)
    return node


ASSET_REFERENCE_PREFIX = "wr-asset:"
# The same idea for a reference that leaves the package: no part to install,
# only a relationship to allocate, and the relationship type comes along
# because a fragment can point outward for more reasons than a hyperlink.
EXTERNAL_REFERENCE_PREFIX = "wr-extref:"


def _append_preserved_inline(node, preserved) -> None:
    """Re-emit inline fragments the model cannot rebuild.

    Relationship ids cannot be written yet: the media parts are installed in a
    later stage, so no id exists here to point at. Each reference is therefore
    replaced with the part name it addressed, marked by ASSET_REFERENCE_PREFIX,
    and resolved to a real id once the assets are in the package.

    A fragment whose reference the parser could not resolve never reaches here;
    it was refused at capture.
    """
    for token in preserved:
        try:
            fragment = etree.fromstring(token["value"])
        except (etree.XMLSyntaxError, KeyError, TypeError):
            # A fragment that will not re-parse is dropped rather than allowed
            # to corrupt the package; G10 reports the resulting loss.
            continue
        targets = token.get("rel_targets") or {}
        if targets:
            for element in fragment.iter():
                if not isinstance(element.tag, str):
                    continue
                for name, value in list(element.attrib.items()):
                    if name.startswith(f"{{{R_NS}}}") and value in targets:
                        target, rel_type, external = targets[value]
                        prefix = (
                            EXTERNAL_REFERENCE_PREFIX if external else ASSET_REFERENCE_PREFIX
                        )
                        element.set(name, f"{prefix}{rel_type}|{target}")
        node.append(fragment)


_PRESERVED_KINDS = {"preserved_xml", "drawing"}


_REFERENCE_TOKENS = {
    "comment_ref": ("commentReference", "comment_id"),
}

_FIELD_CHAR_TYPES = {"field_begin": "begin", "field_separate": "separate", "field_end": "end"}
_FIELD_TOKEN_KINDS = {*_FIELD_CHAR_TYPES, "field_instruction"}


def _append_field_tokens(node, content_tokens) -> None:
    """Recreate a run's own field markers (fldChar/instrText) in place.

    The parser keeps these per-run, at the exact position they were found,
    so replaying them here (instead of the document-level fallback in
    _inject_bookmarks_and_fields) preserves the field's original position —
    e.g. a table of contents stays where it was instead of moving to the
    end of the document.
    """
    for token in content_tokens:
        kind = token.get("kind")
        if kind in _FIELD_CHAR_TYPES:
            _set_w(etree.SubElement(node, f"{W}fldChar"), "fldCharType", _FIELD_CHAR_TYPES[kind])
        elif kind == "field_instruction":
            value = str(token.get("value", ""))
            instr = etree.SubElement(node, f"{W}instrText")
            if value[:1].isspace() or value[-1:].isspace():
                instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            instr.text = value


def _document_has_inline_field_tokens(model: DocumentModel) -> bool:
    for paragraph in model.iter_paragraphs():
        for run in paragraph.runs:
            tokens = run.properties.get("content_tokens") or ()
            if any(token.get("kind") in _FIELD_TOKEN_KINDS for token in tokens):
                return True
    return False


def _section_element(section: Section):
    props = section.properties
    sect = _w("sectPr")
    pg_sz = etree.SubElement(sect, f"{W}pgSz")
    _set_w(pg_sz, "w", props.get("width"))
    _set_w(pg_sz, "h", props.get("height"))
    if props.get("orientation") not in (None, "portrait"):
        _set_w(pg_sz, "orient", props.get("orientation"))
    pg_mar = etree.SubElement(sect, f"{W}pgMar")
    for source, target in (
        ("margin_top", "top"),
        ("margin_right", "right"),
        ("margin_bottom", "bottom"),
        ("margin_left", "left"),
        ("header_distance", "header"),
        ("footer_distance", "footer"),
        ("gutter", "gutter"),
    ):
        _set_w(pg_mar, target, props.get(source))
    if props.get("columns") is not None or props.get("column_space") is not None:
        cols = etree.SubElement(sect, f"{W}cols")
        _set_w(cols, "num", props.get("columns"))
        _set_w(cols, "space", props.get("column_space"))
    if props.get("page_number_start") is not None:
        num = etree.SubElement(sect, f"{W}pgNumType")
        _set_w(num, "start", props.get("page_number_start"))
    return sect


def _paragraph_element(paragraph: Paragraph, sections: list[Section]):
    p = _w("p")
    props = paragraph.properties
    need_ppr = paragraph.style_id is not None or bool(props)
    p_pr = etree.SubElement(p, f"{W}pPr") if need_ppr else None
    if p_pr is not None:
        if paragraph.style_id:
            style = etree.SubElement(p_pr, f"{W}pStyle")
            _set_w(style, "val", paragraph.style_id)
        _bool_element(p_pr, "keepNext", props.get("keepNext"))
        _bool_element(p_pr, "pageBreakBefore", props.get("pageBreakBefore"))
        _bool_element(p_pr, "keepLines", props.get("keepLines"))
        _bool_element(p_pr, "widowControl", props.get("widowControl"))
        if props.get("alignment") is not None:
            jc = etree.SubElement(p_pr, f"{W}jc")
            _set_w(jc, "val", props.get("alignment"))
        spacing_keys = ("spacing_before", "spacing_after", "spacing_line", "spacing_line_rule")
        if any(props.get(key) is not None for key in spacing_keys):
            spacing = etree.SubElement(p_pr, f"{W}spacing")
            for key, attr in (
                ("spacing_before", "before"),
                ("spacing_after", "after"),
                ("spacing_line", "line"),
                ("spacing_line_rule", "lineRule"),
            ):
                _set_w(spacing, attr, props.get(key))
        indent_keys = ("indent_left", "indent_right", "indent_first_line", "indent_hanging")
        if any(props.get(key) is not None for key in indent_keys):
            ind = etree.SubElement(p_pr, f"{W}ind")
            for key, attr in (
                ("indent_left", "left"),
                ("indent_right", "right"),
                ("indent_first_line", "firstLine"),
                ("indent_hanging", "hanging"),
            ):
                _set_w(ind, attr, props.get(key))
        if props.get("numId") is not None:
            num_pr = etree.SubElement(p_pr, f"{W}numPr")
            if props.get("ilvl") is not None:
                ilvl = etree.SubElement(num_pr, f"{W}ilvl")
                _set_w(ilvl, "val", props.get("ilvl"))
            num_id = etree.SubElement(num_pr, f"{W}numId")
            _set_w(num_id, "val", props.get("numId"))
        section_index = props.get("section_index")
        if isinstance(section_index, int) and 0 <= section_index < len(sections):
            p_pr.append(_section_element(sections[section_index]))
    _append_runs_grouping_hyperlinks(p, paragraph.runs)
    return p


def _edge_element(parent, tag: str, edges: dict, keys: tuple[str, ...]) -> None:
    """Write a container of per-edge values (borders, cell margins).

    Edges are written in the schema's own order rather than the dictionary's,
    for the same reason the property sequence is pinned below.
    """
    if not isinstance(edges, dict) or not edges:
        return
    container = etree.SubElement(parent, f"{W}{tag}")
    for edge in ("top", "start", "left", "bottom", "end", "right", "insideH", "insideV"):
        value = edges.get(edge)
        if not isinstance(value, dict):
            continue
        node = etree.SubElement(container, f"{W}{edge}")
        for key in keys:
            if value.get(key) is not None:
                _set_w(node, key, value[key])


# w:tblPr has a required child sequence and Word repairs a document whose
# properties arrive out of order, so the order is pinned here rather than
# following whatever order the model happens to hold them in.
_TABLE_PROPERTY_ORDER = ("tblStyle", "tblW", "jc", "tblBorders", "shd", "tblLayout", "tblCellMar")
_BORDER_KEYS = ("val", "sz", "space", "color")
_MARGIN_KEYS = ("w", "type")


def _table_properties_element(properties: dict):
    written: dict[str, Any] = {}
    holder = _w("tblPr")

    if properties.get("style_id") is not None:
        node = etree.Element(f"{W}tblStyle")
        _set_w(node, "val", properties["style_id"])
        written["tblStyle"] = node
    if properties.get("width") is not None:
        node = etree.Element(f"{W}tblW")
        _set_w(node, "w", properties.get("width"))
        _set_w(node, "type", properties.get("width_type") or "dxa")
        written["tblW"] = node
    if properties.get("alignment") is not None:
        node = etree.Element(f"{W}jc")
        _set_w(node, "val", properties["alignment"])
        written["jc"] = node
    if properties.get("borders"):
        _edge_element(holder, "tblBorders", properties["borders"], _BORDER_KEYS)
        written["tblBorders"] = holder.find(f"{W}tblBorders")
    if properties.get("shading_fill") is not None:
        node = etree.Element(f"{W}shd")
        _set_w(node, "fill", properties["shading_fill"])
        written["shd"] = node
    if properties.get("layout") is not None:
        node = etree.Element(f"{W}tblLayout")
        _set_w(node, "type", properties["layout"])
        written["tblLayout"] = node
    if properties.get("cell_margins"):
        _edge_element(holder, "tblCellMar", properties["cell_margins"], _MARGIN_KEYS)
        written["tblCellMar"] = holder.find(f"{W}tblCellMar")

    tbl_pr = _w("tblPr")
    for name in _TABLE_PROPERTY_ORDER:
        node = written.get(name)
        if node is not None:
            tbl_pr.append(node)
    return tbl_pr


# w:tcPr has its own required sequence.
_CELL_PROPERTY_ORDER = (
    "tcW", "gridSpan", "vMerge", "tcBorders", "shd", "tcMar", "textDirection", "vAlign",
)


def _cell_properties_element(properties: dict):
    written: dict[str, Any] = {}
    holder = _w("tcPr")

    if properties.get("width") is not None:
        node = etree.Element(f"{W}tcW")
        _set_w(node, "w", properties.get("width"))
        _set_w(node, "type", properties.get("width_type") or "dxa")
        written["tcW"] = node
    span = properties.get("grid_span", 1)
    if span and int(span) > 1:
        node = etree.Element(f"{W}gridSpan")
        _set_w(node, "val", span)
        written["gridSpan"] = node
    if properties.get("v_merge") is not None:
        node = etree.Element(f"{W}vMerge")
        if properties.get("v_merge") != "continue":
            _set_w(node, "val", properties.get("v_merge"))
        written["vMerge"] = node
    if properties.get("borders"):
        _edge_element(holder, "tcBorders", properties["borders"], _BORDER_KEYS)
        written["tcBorders"] = holder.find(f"{W}tcBorders")
    if properties.get("shading_fill") is not None:
        node = etree.Element(f"{W}shd")
        _set_w(node, "fill", properties["shading_fill"])
        written["shd"] = node
    # These two keys are "cell_margins" and "vertical_alignment" in the model.
    # Asking for "margins"/"v_align" returned None every time and the properties
    # vanished without any error to notice.
    if properties.get("cell_margins"):
        _edge_element(holder, "tcMar", properties["cell_margins"], _MARGIN_KEYS)
        written["tcMar"] = holder.find(f"{W}tcMar")
    if properties.get("text_direction") is not None:
        node = etree.Element(f"{W}textDirection")
        _set_w(node, "val", properties["text_direction"])
        written["textDirection"] = node
    if properties.get("vertical_alignment") is not None:
        node = etree.Element(f"{W}vAlign")
        _set_w(node, "val", properties["vertical_alignment"])
        written["vAlign"] = node

    tc_pr = _w("tcPr")
    for name in _CELL_PROPERTY_ORDER:
        node = written.get(name)
        if node is not None:
            tc_pr.append(node)
    return tc_pr


# w:trPr is a sequence too.
_ROW_PROPERTY_ORDER = ("cantSplit", "trHeight", "tblHeader")


def _row_properties_element(properties: dict):
    """Rebuild w:trPr, emitting it even when it ends up empty.

    cant_split and repeat_header are False rather than absent whenever the
    source row carried a w:trPr at all, so an empty one still says something:
    dropping it turns False into missing and the row stops round-tripping.
    """
    written: dict[str, Any] = {}

    if properties.get("cant_split"):
        written["cantSplit"] = etree.Element(f"{W}cantSplit")
    if properties.get("height") is not None or properties.get("height_rule") is not None:
        node = etree.Element(f"{W}trHeight")
        _set_w(node, "val", properties.get("height"))
        _set_w(node, "hRule", properties.get("height_rule"))
        written["trHeight"] = node
    if properties.get("repeat_header"):
        written["tblHeader"] = etree.Element(f"{W}tblHeader")

    tr_pr = _w("trPr")
    for name in _ROW_PROPERTY_ORDER:
        node = written.get(name)
        if node is not None:
            tr_pr.append(node)
    return tr_pr


def _table_element(table: Table, sections: list[Section]):
    tbl = _w("tbl")
    tbl_pr = _table_properties_element(table.properties or {})
    if len(tbl_pr):
        tbl.append(tbl_pr)

    # Without w:tblGrid Word has no column widths to lay the table out from and
    # falls back to its own guess, so a table authored with a narrow label
    # column and a wide value column comes back evenly split.
    widths = (table.properties or {}).get("grid_column_widths") or []
    if widths:
        grid = etree.SubElement(tbl, f"{W}tblGrid")
        for width in widths:
            _set_w(etree.SubElement(grid, f"{W}gridCol"), "w", width)

    for row in table.rows:
        tr = etree.SubElement(tbl, f"{W}tr")
        # An empty dict means the source row had no w:trPr; a dict holding only
        # False means it had an empty one, which still has to come back.
        if row.properties:
            tr.append(_row_properties_element(row.properties))
        for cell in row.cells:
            tc = etree.SubElement(tr, f"{W}tc")
            tc.append(_cell_properties_element(cell.properties or {}))
            for block in cell.blocks:
                tc.append(_block_element(block, sections))
            if not cell.blocks:
                tc.append(_w("p"))
    return tbl


def _block_element(block, sections: list[Section]):
    if isinstance(block, Paragraph):
        return _paragraph_element(block, sections)
    if isinstance(block, Table):
        return _table_element(block, sections)
    p = _w("p")
    r = etree.SubElement(p, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = f"[Unsupported block: {type(block).__name__}]"
    return p


_PARAGRAPH_INDEX = re.compile(r"^//w:p\[(\d+)\]")
_SIBLING_INDEX = re.compile(r"\[(\d+)\]$")


def _sibling_index(path: str | None) -> int:
    """Which of its like-named siblings a path points at; 1 when unindexed.

    lxml writes no index when an element is the only one of its name in the
    parent, so a bare path means "the first".
    """
    if not path:
        return 1
    match = _SIBLING_INDEX.search(path)
    return int(match.group(1)) if match else 1


def _paragraph_at(paragraphs: list, path: str | None):
    """The paragraph a Bookmark path names, or None if it names nothing here.

    Paths look like ``//w:p[N]/...`` where N counts every w:p in the document,
    table cells included -- so the list handed in has to be built the same way.
    """
    if not path:
        return None
    match = _PARAGRAPH_INDEX.match(path)
    if match is None:
        return None
    index = int(match.group(1))
    return paragraphs[index - 1] if 1 <= index <= len(paragraphs) else None


def _inject_bookmarks_and_fields(body, model: DocumentModel) -> None:
    paragraphs = body.findall(f"{W}p")
    if not paragraphs:
        return

    # Every w:p in document order, matching how the parser numbered them. Using
    # only the body's direct children would count a bookmark in a table cell as
    # though the whole table were one paragraph.
    all_paragraphs = list(body.iter(f"{W}p"))

    first = paragraphs[0]
    insert_positions: dict[object, int] = {}
    # Ends are appended after every bookmark is placed, in the order the source
    # had them rather than the order the bookmarks were declared. Nested
    # bookmarks close inside-out -- start A, start B, end B, end A -- so
    # following declaration order puts A's end first and moves both.
    pending_ends: list[tuple[object, int, int, object]] = []
    for index, bookmark in enumerate(model.bookmarks, start=1):
        bookmark_id = str(bookmark.bookmark_id or index)
        # Bookmarks used to go into the first paragraph regardless of where they
        # were authored, so a cross-reference resolved to the wrong text. The
        # position was in the model the whole time; it just was not read.
        # "or" would be wrong here: an lxml element with no children is falsy,
        # so a bookmark in an empty paragraph would silently fall back.
        start_host = _paragraph_at(all_paragraphs, bookmark.start_path)
        if start_host is None:
            start_host = first
        end_host = _paragraph_at(all_paragraphs, bookmark.end_path)
        if end_host is None:
            end_host = start_host

        start = etree.Element(f"{W}bookmarkStart")
        _set_w(start, "id", bookmark_id)
        _set_w(start, "name", bookmark.name)
        if start_host not in insert_positions:
            insert_positions[start_host] = (
                1 if len(start_host) and start_host[0].tag == f"{W}pPr" else 0
            )
        start_host.insert(insert_positions[start_host], start)
        insert_positions[start_host] += 1

        end = etree.Element(f"{W}bookmarkEnd")
        _set_w(end, "id", bookmark_id)
        pending_ends.append((end_host, _sibling_index(bookmark.end_path), index, end))

    for host, _position, _declared, end in sorted(pending_ends, key=lambda item: item[1:3]):
        host.append(end)

    if model.fields and not body.findall(f".//{W}fldChar"):
        # Real parsed documents already got their fields rendered in place by
        # _run_element (from each run's own content_tokens), so this branch
        # only fires for models that carry field definitions without any
        # inline token representation (e.g. hand-built in tests). It recreates
        # the instruction sequence without adding new visible text or a new
        # paragraph, so L0/body structure remains unchanged — but, lacking any
        # positional information, can only place it at the end of the body.
        target = paragraphs[-1]
        for field in model.fields:
            begin_run = etree.SubElement(target, f"{W}r")
            begin = etree.SubElement(begin_run, f"{W}fldChar")
            _set_w(begin, "fldCharType", "begin")
            instr_run = etree.SubElement(target, f"{W}r")
            instr = etree.SubElement(instr_run, f"{W}instrText")
            instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            instr.text = f" {field.instruction.strip()} "
            end_run = etree.SubElement(target, f"{W}r")
            end = etree.SubElement(end_run, f"{W}fldChar")
            _set_w(end, "fldCharType", "end")


def build_body_xml(model: DocumentModel):
    body = _w("body")
    for block in model.body:
        body.append(_block_element(block, model.sections))
    _inject_bookmarks_and_fields(body, model)
    if model.sections:
        body.append(_section_element(model.sections[-1]))
    return body


# Pinned so a seed produces byte-identical output across runs and machines.
# The 1980-01-01 epoch is the earliest a ZIP member timestamp can express.
_DETERMINISTIC_ZIP_DATE_TIME = (1980, 1, 1, 0, 0, 0)
_DETERMINISTIC_COMPRESSLEVEL = 6


def _canonical_part_order(parts: dict[str, bytes]) -> list[str]:
    """OPC readers accept any member order; pin one so the bytes are stable.

    [Content_Types].xml must be first for the package to be recognised by
    strict readers, the package relationships follow, then everything sorted.
    """
    lead = [name for name in ("[Content_Types].xml", "_rels/.rels") if name in parts]
    return lead + sorted(name for name in parts if name not in lead)


def _rels_part_for(part_name: str) -> str:
    parent, _, name = part_name.rpartition("/")
    return f"{parent}/_rels/{name}.rels" if parent else f"_rels/{name}.rels"


def _resolve_rel_target(rels_part: str, target: str) -> str:
    owner = PurePosixPath(rels_part).parent
    if owner.name == "_rels":
        owner = owner.parent
    candidate = target.lstrip("/") if target.startswith("/") else str(PurePosixPath(owner, target))
    resolved: list[str] = []
    for piece in PurePosixPath(candidate).parts:
        if piece == "..":
            if resolved:
                resolved.pop()
        elif piece not in (".", ""):
            resolved.append(piece)
    return "/".join(resolved)


def _relative_to_word(part_name: str) -> str:
    """Express a part name as a target relative to word/document.xml."""
    if part_name.startswith("word/"):
        return part_name[len("word/"):]
    return "../" + part_name


IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

_STORY_PART_PREFIXES = ("word/document.xml", "word/header", "word/footer",
                        "word/footnotes.xml", "word/endnotes.xml", "word/comments.xml")


def _is_story_part(name: str) -> bool:
    return name.endswith(".xml") and name.startswith(_STORY_PART_PREFIXES)


def _relative_to(part_name: str, owner_dir: str) -> str:
    """Express a part name as a relationship target relative to owner_dir."""
    owner = PurePosixPath(owner_dir)
    target = PurePosixPath(part_name)
    try:
        return str(target.relative_to(owner))
    except ValueError:
        up = "../" * len(owner.parts)
        return up + str(target)


class MutableDocxPackage:
    def __init__(self, parts: dict[str, bytes]) -> None:
        self.parts = parts
        self.warnings: list[WarningItem] = []

    @classmethod
    def from_file(cls, path: Path) -> "MutableDocxPackage":
        with ZipFile(path, "r") as archive:
            return cls({name: archive.read(name) for name in archive.namelist()})

    def _xml(self, name: str):
        return etree.fromstring(self.parts[name])

    def _write_xml(self, name: str, root) -> None:
        self.parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    def set_document_body(self, body, *, ignorable: str | None = None) -> None:
        root = self._xml("word/document.xml")
        current = root.find(f"{W}body")
        if current is None:
            raise RuntimeError("Fresh DOCX shell has no document body")
        root.replace(current, deepcopy(body))
        # The shell template declares its own mc:Ignorable. Left alone, every
        # rebuilt document inherits it whether or not the source said anything.
        key = f"{{{MC_NS}}}Ignorable"
        if ignorable:
            root.set(key, ignorable)
        elif key in root.attrib:
            del root.attrib[key]
        self._write_xml("word/document.xml", root)

    def _ensure_override(self, part_name: str, content_type: str) -> None:
        root = self._xml("[Content_Types].xml")
        xpath = f"//*[local-name()='Override'][@PartName='/{part_name}']"
        if not root.xpath(xpath):
            node = etree.SubElement(root, f"{{{CT_NS}}}Override")
            node.set("PartName", f"/{part_name}")
            node.set("ContentType", content_type)
            self._write_xml("[Content_Types].xml", root)

    # python-docx's default template is the starting shell for every rebuild,
    # and it ships a custom XML datastore and a thumbnail of its own. Left in
    # place they end up in documents that never had them -- a foreign customXml
    # store is real content, not decoration, and no gate that compares parsed
    # models can see it because the parser does not model custom XML at all.
    # Parts the shell template brings and the renderer never writes. Dropped
    # here so the source's own can be restored in their place; a source that
    # has none ends up with none, rather than inheriting the template's --
    # which also brought the template's mc:Ignorable declaration with it.
    _TEMPLATE_DEBRIS = (
        "customXml/item1.xml",
        "customXml/itemProps1.xml",
        "docProps/thumbnail.jpeg",
        "word/stylesWithEffects.xml",
        "word/theme/theme1.xml",
        "word/webSettings.xml",
        "word/fontTable.xml",
    )

    def drop_template_debris(self) -> list[str]:
        """Remove parts that came from the shell template rather than the source."""
        dropped = [name for name in self._TEMPLATE_DEBRIS if self.drop_part(name)]
        # customXml/_rels/item1.xml.rels has no override and is only reachable
        # from the part just removed.
        for leftover in [n for n in list(self.parts) if n.lower().startswith("customxml/")]:
            del self.parts[leftover]
            dropped.append(leftover)
        return dropped

    def drop_part(self, part_name: str) -> bool:
        """Remove a part together with everything that refers to it.

        Removing the bytes alone would leave a content-type override and a
        relationship pointing at nothing, which Word repairs on open -- a worse
        defect than the one being fixed. Returns whether anything was removed.
        """
        if part_name not in self.parts:
            return False
        del self.parts[part_name]
        self.parts.pop(_rels_part_for(part_name), None)

        root = self._xml("[Content_Types].xml")
        removed_override = False
        for node in root.xpath(f"//*[local-name()='Override'][@PartName='/{part_name}']"):
            root.remove(node)
            removed_override = True
        if removed_override:
            self._write_xml("[Content_Types].xml", root)

        for rels_name in [name for name in self.parts if name.endswith(".rels")]:
            rels_root = self._xml(rels_name)
            changed = False
            for node in list(rels_root):
                if node.get("TargetMode") == "External":
                    continue
                if _resolve_rel_target(rels_name, node.get("Target") or "") == part_name:
                    rels_root.remove(node)
                    changed = True
            if changed:
                self._write_xml(rels_name, rels_root)
        return True

    def _ensure_relationship(
        self, rels_part: str, rel_type: str, target: str, *, external: bool = False
    ) -> str:
        if rels_part not in self.parts:
            self.parts[rels_part] = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
            )
        root = self._xml(rels_part)
        for node in root:
            if node.get("Type") == rel_type and node.get("Target") == target:
                return node.get("Id")
        used = {node.get("Id") for node in root}
        index = 1
        while f"rIdReplica{index}" in used:
            index += 1
        rel_id = f"rIdReplica{index}"
        node = etree.SubElement(root, f"{{{REL_NS}}}Relationship")
        node.set("Id", rel_id)
        node.set("Type", rel_type)
        node.set("Target", target)
        if external:
            node.set("TargetMode", "External")
        self._write_xml(rels_part, root)
        return rel_id


    def restore_verbatim_relationships(
        self, part_name: str, entries
    ) -> list[str]:
        """Rebuild the .rels of a part that was written back byte for byte.

        The ids are written as they were, not reallocated: the bytes that use
        them were preserved unchanged, so a fresh id would point the preserved
        reference somewhere else.

        An internal target whose part did not make it into the package is
        skipped and returned, because restoring it would trade an orphaned part
        for a dangling relationship -- and Word repairs the second one.
        """
        if part_name not in self.parts or not entries:
            return []
        rels_part = _rels_part_for(part_name)
        missing: list[str] = []
        root = etree.fromstring(
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )
        for rel_id, rel_type, target, target_mode in entries:
            external = target_mode == "External"
            if not external and _resolve_rel_target(rels_part, target) not in self.parts:
                missing.append(target)
                continue
            node = etree.SubElement(root, f"{{{REL_NS}}}Relationship")
            node.set("Id", rel_id)
            node.set("Type", rel_type)
            node.set("Target", target)
            if external:
                node.set("TargetMode", "External")
        if len(root):
            self._write_xml(rels_part, root)
        return missing

    def install_structured_part(self, part_name: str, content_type: str, rel_type: str, target: str, data: bytes) -> str:
        self.parts[part_name] = data
        self._ensure_override(part_name, content_type)
        return self._ensure_relationship("word/_rels/document.xml.rels", rel_type, target)

    def attach_default_header_footer(self, kind: str, rel_id: str) -> None:
        root = self._xml("word/document.xml")
        ref_tag = f"{W}{kind}Reference"
        for sect in root.xpath("//w:sectPr", namespaces={"w": W_NS}):
            for existing in list(sect.findall(ref_tag)):
                if existing.get(f"{W}type") == "default":
                    sect.remove(existing)
            ref = etree.Element(ref_tag)
            _set_w(ref, "type", "default")
            ref.set(f"{{{R_NS}}}id", rel_id)
            sect.insert(0, ref)
        self._write_xml("word/document.xml", root)

    def set_styles(self, raw: bytes | None) -> None:
        if raw is None:
            # The source has no styles part. Returning early here left the
            # shell template's in place, so the rebuild silently acquired
            # styles the source never defined.
            self.drop_part("word/styles.xml")
            return
        self.parts["word/styles.xml"] = raw

    def set_numbering(self, raw: bytes | None) -> None:
        if raw is None:
            # Measured on the corpus: documents with no numbering.xml were
            # handed the template's, complete with its own mc:Ignorable
            # declaration -- which is how a "markup compatibility" divergence
            # showed up on documents that use no numbering at all.
            self.drop_part("word/numbering.xml")
            return
        self.parts["word/numbering.xml"] = raw
        self._ensure_override(
            "word/numbering.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        )
        self._ensure_relationship(
            "word/_rels/document.xml.rels",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
            "numbering.xml",
        )

    def set_settings(self, raw: bytes | None, tracked_changes_enabled: bool) -> None:
        if raw is None:
            self.drop_part("word/settings.xml")
            return
        root = etree.fromstring(raw)
        if not tracked_changes_enabled:
            for node in root.findall(f"{W}trackRevisions"):
                root.remove(node)
        self._write_xml("word/settings.xml", root)

    def resolve_asset_references(self) -> list[str]:
        """Turn placeholder asset references into real relationship ids.

        Runs after the media parts are installed. Each story part is resolved
        against its own .rels, so a picture in a header gets a header
        relationship rather than a document one.

        A reference whose part never made it into the package is stripped
        rather than left pointing at nothing: Word repairs a package with a
        broken reference, and a repaired document is a worse outcome than a
        missing picture.
        """
        unresolved: list[str] = []
        for part_name in [name for name in self.parts if _is_story_part(name)]:
            try:
                root = self._xml(part_name)
            except etree.XMLSyntaxError:
                continue
            changed = False
            for element in root.iter():
                if not isinstance(element.tag, str):
                    continue
                for name, value in list(element.attrib.items()):
                    if not value.startswith(
                        (
                            ASSET_REFERENCE_PREFIX,
                            HYPERLINK_REFERENCE_PREFIX,
                            EXTERNAL_REFERENCE_PREFIX,
                        )
                    ):
                        continue
                    if value.startswith(EXTERNAL_REFERENCE_PREFIX):
                        # Points out of the package: allocate the relationship,
                        # install nothing.
                        rel_type, _, target = value[
                            len(EXTERNAL_REFERENCE_PREFIX):
                        ].partition("|")
                        changed = True
                        element.set(
                            name,
                            self._ensure_relationship(
                                _rels_part_for(part_name), rel_type, target, external=True
                            ),
                        )
                        continue
                    if value.startswith(HYPERLINK_REFERENCE_PREFIX):
                        # An external link needs no part, only a relationship.
                        changed = True
                        element.set(
                            name,
                            self._ensure_relationship(
                                _rels_part_for(part_name),
                                HYPERLINK_REL_TYPE,
                                value[len(HYPERLINK_REFERENCE_PREFIX):],
                                external=True,
                            ),
                        )
                        continue
                    rel_type, _, target = value[len(ASSET_REFERENCE_PREFIX):].partition("|")
                    if not target:  # written before the type was recorded
                        rel_type, target = IMAGE_REL_TYPE, rel_type
                    changed = True
                    if target not in self.parts:
                        unresolved.append(target)
                        del element.attrib[name]
                        continue
                    rels_part = _rels_part_for(part_name)
                    owner = PurePosixPath(part_name).parent
                    relative = _relative_to(target, str(owner))
                    element.set(
                        name,
                        self._ensure_relationship(rels_part, rel_type, relative),
                    )
            if changed:
                self._write_xml(part_name, root)
        return unresolved

    def unreferenced_parts(self, prefix: str | tuple[str, ...]) -> list[str]:
        """Parts under ``prefix`` that no relationship reaches.

        An OPC part nothing points at is litter, not content: the bytes ship
        with the document while whatever displayed them is gone.
        """
        reached: set[str] = set()
        for rels_name in [name for name in self.parts if name.endswith(".rels")]:
            try:
                root = self._xml(rels_name)
            except etree.XMLSyntaxError:
                continue
            for node in root:
                if node.get("TargetMode") == "External":
                    continue
                reached.add(_resolve_rel_target(rels_name, node.get("Target") or ""))
        return sorted(name for name in self.parts if name.startswith(prefix) and name not in reached)

    def set_settings_relationships(self, relationships) -> None:
        """Restore the settings part's external relationships.

        Only external ones reach here, so there is no part to install and
        nothing that can end up pointing at a part that is not present.
        """
        for rel_type, target in relationships or ():
            self._ensure_relationship(
                "word/_rels/settings.xml.rels", rel_type, target, external=True
            )

    def install_attachment(
        self,
        part_name: str,
        data: bytes,
        content_type: str | None,
        relationship_type: str,
        *,
        owner_rels: str = "word/_rels/document.xml.rels",
    ) -> None:
        """Restore an attachment and the relationship reaching it.

        Writing the bytes without the relationship would leave an OPC part
        nothing references, which is litter rather than preserved content.
        """
        self.parts[part_name] = data
        if content_type:
            self._ensure_override(part_name, content_type)
        target = _relative_to_word(part_name) if owner_rels.startswith("word/") else part_name
        self._ensure_relationship(owner_rels, relationship_type, target)

    def install_asset(self, part_name: str, data: bytes, content_type: str | None = None) -> None:
        self.parts[part_name] = data
        if content_type:
            self._ensure_override(part_name, content_type)

    def initialize_truthful_lifecycle(self) -> None:
        now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
        if "docProps/core.xml" in self.parts:
            root = self._xml("docProps/core.xml")
            for tag in ("created", "modified"):
                node = root.find(f"{{{DCTERMS_NS}}}{tag}")
                if node is not None:
                    node.text = now
            revision = root.find(f"{{{CP_NS}}}revision")
            if revision is not None:
                revision.text = "1"
            self._write_xml("docProps/core.xml", root)
        if "docProps/app.xml" in self.parts:
            root = self._xml("docProps/app.xml")
            total = root.find(f"{{{EP_NS}}}TotalTime")
            if total is not None:
                total.text = "0"
            self._write_xml("docProps/app.xml", root)

    def set_metadata(self, policy: dict[str, str]) -> None:
        if "docProps/core.xml" in self.parts:
            root = self._xml("docProps/core.xml")
            mappings = {
                "title": (DC_NS, "title"),
                "subject": (DC_NS, "subject"),
                "keywords": (CP_NS, "keywords"),
                "category": (CP_NS, "category"),
                "language": (DC_NS, "language"),
                "creator": (DC_NS, "creator"),
            }
            for key, (ns, tag) in mappings.items():
                if key not in policy:
                    continue
                node = root.find(f"{{{ns}}}{tag}")
                if node is None:
                    node = etree.SubElement(root, f"{{{ns}}}{tag}")
                node.text = policy[key]
            self._write_xml("docProps/core.xml", root)
        if policy.get("company") and "docProps/app.xml" in self.parts:
            root = self._xml("docProps/app.xml")
            node = root.find(f"{{{EP_NS}}}Company")
            if node is None:
                node = etree.SubElement(root, f"{{{EP_NS}}}Company")
            node.text = policy["company"]
            self._write_xml("docProps/app.xml", root)
        for key, value in policy.items():
            if key.startswith("custom:"):
                self.set_custom_property(key.split(":", 1)[1], value)

    def set_custom_property(self, name: str, value: str) -> None:
        part = "docProps/custom.xml"
        if part in self.parts:
            root = self._xml(part)
        else:
            root = etree.Element(f"{{{CUST_NS}}}Properties", nsmap={None: CUST_NS, "vt": VT_NS})
            self._ensure_override(
                part,
                "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            )
            self._ensure_relationship(
                "_rels/.rels",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
                "docProps/custom.xml",
            )
        existing = None
        max_pid = 1
        for prop in root.findall(f"{{{CUST_NS}}}property"):
            try:
                max_pid = max(max_pid, int(prop.get("pid", "1")))
            except ValueError:
                pass
            if prop.get("name") == name:
                existing = prop
        if existing is None:
            existing = etree.SubElement(root, f"{{{CUST_NS}}}property")
            existing.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            existing.set("pid", str(max_pid + 1))
            existing.set("name", name)
        for child in list(existing):
            existing.remove(child)
        text = etree.SubElement(existing, f"{{{VT_NS}}}lpwstr")
        text.text = str(value)
        self._write_xml(part, root)

    def write_deterministic(self, output_path: Path) -> None:
        """Write the package so identical parts always produce identical bytes.

        write_atomic cannot promise this: ZipFile.writestr stamps every member
        with the current local time, and it iterates self.parts in insertion
        order. The fidelity lab addresses generated documents by content hash,
        so a seed has to yield the same bytes forever -- otherwise a Python or
        zlib upgrade silently rebaselines GOLDEN-CORE and orphans every
        regression history attached to it.

        Four things are pinned: a fixed member timestamp, an explicit
        compression level (zlib's default has changed between versions), fixed
        permission bits, and a canonical member order -- OPC readers tolerate
        any order, but a stable one is what makes the bytes stable.
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(prefix=output_path.name, suffix=".tmp", dir=output_path.parent)
        os.close(fd)
        try:
            with ZipFile(temp_name, "w", ZIP_DEFLATED, compresslevel=_DETERMINISTIC_COMPRESSLEVEL) as archive:
                for name in _canonical_part_order(self.parts):
                    info = ZipInfo(name, date_time=_DETERMINISTIC_ZIP_DATE_TIME)
                    info.compress_type = ZIP_DEFLATED
                    info.external_attr = 0o600 << 16
                    archive.writestr(info, self.parts[name])
            Path(temp_name).replace(output_path)
        finally:
            Path(temp_name).unlink(missing_ok=True)

    def write_atomic(self, output_path: Path) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(prefix=output_path.name, suffix=".tmp", dir=output_path.parent)
        # mkstemp returns an open descriptor. Windows will not unlink/replace a
        # file while that descriptor is still open, so close it before any
        # path-level operation. Linux permits unlinking an open file, which is
        # why this bug was invisible in the original non-Windows test run.
        os.close(fd)
        try:
            with ZipFile(temp_name, "w", ZIP_DEFLATED) as archive:
                for name, data in self.parts.items():
                    archive.writestr(name, data)
            Path(temp_name).replace(output_path)
        finally:
            Path(temp_name).unlink(missing_ok=True)


class PureDocxRenderer:
    def __init__(self) -> None:
        self._package: MutableDocxPackage | None = None

    def render(self, model: DocumentModel, output_path: Path, context) -> RenderResult:
        output_path = Path(output_path)
        shell = output_path.with_suffix(".shell.docx")
        shell.parent.mkdir(parents=True, exist_ok=True)
        Document().save(shell)
        self._package = MutableDocxPackage.from_file(shell)
        self._package.drop_template_debris()
        self._package.initialize_truthful_lifecycle()
        if model.fields and not _document_has_inline_field_tokens(model):
            self._package.warnings.append(WarningItem(
                code="PURE_DOCX_FIELD_POSITION_UNAVAILABLE",
                message="Field codes (e.g. table of contents, cross-references) were retained but reinserted at the end of the document body instead of their original position",
                affects_status=True,
            ))
        stages = [
            (
                "body",
                lambda: self._package.set_document_body(
                    build_body_xml(model),
                    ignorable=model.extras.get("document_ignorable"),
                ),
            ),
            ("styles", lambda: self._package.set_styles(model.styles_xml)),
            ("numbering", lambda: self._package.set_numbering(model.numbering_xml)),
            (
                "settings",
                lambda: self._set_settings_and_relationships(model),
            ),
            ("assets", lambda: self._install_assets(model)),
            ("notes_headers", lambda: self._install_notes_headers(model)),
            ("references", lambda: self._resolve_references(model)),
            ("metadata", lambda: self._package.set_metadata(model.extras.get("metadata_policy", {}))),
        ]
        try:
            for index, (stage, mutate) in enumerate(stages, start=1):
                mutate()
                if context is not None:
                    context.mark_content_changed(f"{model.fingerprint()}:{index}:{stage}")
                    context.checkpoint(
                        f"{stage} complete",
                        stage,
                        lambda: self.save(output_path),
                        output_path,
                    )
            if context is None:
                self.save(output_path)
            else:
                context.final_seal(self, output_path)
            return RenderResult(
                output_path=output_path,
                warnings=list(self._package.warnings),
                stages_completed=[stage for stage, _ in stages],
            )
        finally:
            shell.unlink(missing_ok=True)

    def _set_settings_and_relationships(self, model: DocumentModel) -> None:
        assert self._package is not None
        self._package.set_settings(
            model.settings_xml,
            bool(model.extras.get("tracked_changes_enabled", False)),
        )
        if model.settings_xml is not None:
            self._package.set_settings_relationships(
                model.extras.get("settings_relationships")
            )

    def _install_assets(self, model: DocumentModel) -> None:
        assert self._package is not None
        for asset in model.assets.values():
            self._package.install_asset(asset.part_name, asset.bytes_data, asset.content_type)

        for part in model.preserved_parts.values():
            if part.relationship_type:
                # A document-level attachment: part plus one relationship is the
                # whole of it, so it can be restored byte-for-byte.
                self._package.install_attachment(
                    part.part_name,
                    part.data,
                    part.content_type,
                    part.relationship_type,
                    owner_rels=part.owner_rels,
                )
                continue
            if part.sidecar:
                # Reached from the attachment's own .rels, which is itself
                # restored here, so no new relationship is needed.
                self._package.install_asset(part.part_name, part.data, part.content_type)
                continue
            # Body-referenced: a chart, diagram or embedded object. These used
            # to be dropped, because a rebuilt body could not carry the
            # reference and the part alone would have been an orphan. Verbatim
            # fragment preservation removed that constraint -- the reference
            # survives with its id remapped -- so the part is installed and the
            # two ends meet. If the fragment was refused after all, the
            # unreferenced-media check reports the part rather than shipping it
            # silently.
            self._package.install_asset(part.part_name, part.data, part.content_type)

    def _blocks_part_xml(self, root_tag: str, blocks: list[object]) -> bytes:
        root = _w(root_tag)
        for block in blocks:
            root.append(_block_element(block, []))
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    def _notes_part_xml(self, kind: str, notes: dict[str, list[object]]) -> bytes:
        root = _w(f"{kind}s")
        for note_id, blocks in sorted(notes.items(), key=lambda item: item[0]):
            note = etree.SubElement(root, f"{W}{kind}")
            _set_w(note, "id", note_id)
            for block in blocks:
                note.append(_block_element(block, []))
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    def _comments_part_xml(self, model: DocumentModel) -> bytes:
        """Rebuild word/comments.xml from the model.

        Author, date and initials come back because they are how Word
        attributes a comment; a comment whose author is gone reads as somebody
        else's remark.
        """
        root = _w("comments")
        initials = model.extras.get("comment_initials") or {}
        for comment_id, comment in sorted(model.comments.items(), key=lambda item: item[0]):
            node = etree.SubElement(root, f"{W}comment")
            _set_w(node, "id", str(comment_id))
            if comment.author:
                _set_w(node, "author", comment.author)
            if initials.get(comment_id):
                _set_w(node, "initials", initials[comment_id])
            if comment.date:
                _set_w(node, "date", comment.date)
            for block in comment.blocks:
                node.append(_block_element(block, []))
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    def _resolve_references(self, model: DocumentModel) -> None:
        """Turn placeholder asset references into relationship ids.

        Runs after every story part exists, not just the body. Comments,
        headers and notes are written in a later stage than the media is
        installed, so resolving inside the asset stage left their placeholders
        untouched and their pictures orphaned -- which is exactly what the
        orphan check then reported.
        """
        assert self._package is not None
        # styles.xml and numbering.xml were written back unchanged, so the r:id
        # values inside them are still the source's. Their .rels is restored
        # here rather than in the stage that writes them, because the parts
        # those relationships point at are installed in between.
        for part, entries in (model.extras.get("verbatim_part_relationships") or {}).items():
            for target in self._package.restore_verbatim_relationships(part, entries):
                self._package.warnings.append(WarningItem(
                    code="PURE_DOCX_VERBATIM_RELATIONSHIP_UNRESOLVED",
                    message=(
                        f"{part} referenced {target}, which is not in the rebuilt package"
                    ),
                    affects_status=True,
                ))
        for missing in self._package.resolve_asset_references():
            self._package.warnings.append(WarningItem(
                code="PURE_DOCX_ASSET_REFERENCE_UNRESOLVED",
                message=f"A picture referenced {missing}, which is not in the rebuilt package",
                affects_status=True,
            ))
        # The mirror case: media that arrived with the model but that nothing in
        # the rebuilt document points at. Reported rather than shipped silently
        # -- the bytes would travel while the picture itself is gone.
        for orphan in self._package.unreferenced_parts(("word/media/", "word/embeddings/",
                                                       "word/charts/", "word/diagrams/")):
            self._package.warnings.append(WarningItem(
                code="PURE_DOCX_ASSET_UNREFERENCED",
                message=f"Media part {orphan} was retained but nothing in the document refers to it",
                affects_status=True,
            ))

    def _install_notes_headers(self, model: DocumentModel) -> None:
        assert self._package is not None
        for kind, collection, content_type, rel_type in (
            ("header", model.headers, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
            ("footer", model.footers, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
        ):
            installed: list[tuple[str, str]] = []
            for index, (_source_name, blocks) in enumerate(sorted(collection.items()), start=1):
                part_name = f"word/{kind}{index}.xml"
                rel_id = self._package.install_structured_part(
                    part_name, content_type, rel_type, f"{kind}{index}.xml", self._blocks_part_xml("hdr" if kind == "header" else "ftr", blocks)
                )
                installed.append((part_name, rel_id))
            if installed:
                self._package.attach_default_header_footer(kind, installed[0][1])
                if len(installed) > 1:
                    self._package.warnings.append(WarningItem(
                        code="PURE_DOCX_HEADER_FOOTER_MAPPING_APPROXIMATION",
                        message=f"Multiple {kind} parts were reconstructed, but the pure fallback maps only the first as each section's default {kind}",
                        affects_status=True,
                    ))

        if model.comments:
            self._package.install_structured_part(
                "word/comments.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "comments.xml",
                self._comments_part_xml(model),
            )

        for kind, notes, content_type, rel_type in (
            ("footnote", model.footnotes, "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"),
            ("endnote", model.endnotes, "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"),
        ):
            if notes:
                part_name = f"word/{kind}s.xml"
                self._package.install_structured_part(
                    part_name, content_type, rel_type, f"{kind}s.xml", self._notes_part_xml(kind, notes)
                )

    def save(self, output_path: Path) -> None:
        if self._package is None:
            raise RuntimeError("PureDocxRenderer has no active package")
        self._package.write_atomic(output_path)

    def set_custom_property(self, name: str, value: str) -> None:
        if self._package is None:
            raise RuntimeError("PureDocxRenderer has no active package")
        self._package.set_custom_property(name, value)
