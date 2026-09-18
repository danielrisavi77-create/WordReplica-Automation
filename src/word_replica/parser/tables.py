from word_replica.domain.model import Table, TableCell, TableRow
from word_replica.parser.text import NS, W_NS


def _attr(node, name: str) -> str | None:
    if node is None:
        return None
    return node.get(f"{{{W_NS}}}{name}")


def _compact(values: dict) -> dict:
    return {k: v for k, v in values.items() if v is not None}


def _border_dict(parent) -> dict:
    result = {}
    if parent is None:
        return result
    for node in parent:
        result[node.tag.rsplit("}", 1)[-1]] = _compact({
            "val": _attr(node, "val"), "sz": _attr(node, "sz"),
            "color": _attr(node, "color"), "space": _attr(node, "space"),
        })
    return result


def parse_table_properties(node) -> dict:
    tbl_pr = node.find("w:tblPr", namespaces=NS)
    tbl_grid = node.find("w:tblGrid", namespaces=NS)
    if tbl_pr is None:
        tbl_pr = None
    style = tbl_pr.find("w:tblStyle", namespaces=NS) if tbl_pr is not None else None
    width = tbl_pr.find("w:tblW", namespaces=NS) if tbl_pr is not None else None
    layout = tbl_pr.find("w:tblLayout", namespaces=NS) if tbl_pr is not None else None
    shading = tbl_pr.find("w:shd", namespaces=NS) if tbl_pr is not None else None
    borders = tbl_pr.find("w:tblBorders", namespaces=NS) if tbl_pr is not None else None
    align = tbl_pr.find("w:jc", namespaces=NS) if tbl_pr is not None else None
    cell_mar = tbl_pr.find("w:tblCellMar", namespaces=NS) if tbl_pr is not None else None
    grid_widths = []
    if tbl_grid is not None:
        # A gridCol without a w:w attribute means no explicit width was declared
        # (e.g. an autofit table) - keep that as None rather than defaulting to
        # 0, which Word's COM Columns(n).Width setter rejects outright with a
        # "Value out of range" automation error.
        grid_widths = [
            int(raw_width) if (raw_width := _attr(col, "w")) is not None else None
            for col in tbl_grid.findall("w:gridCol", namespaces=NS)
        ]
    margins = {}
    if cell_mar is not None:
        for child in cell_mar:
            margins[child.tag.rsplit("}", 1)[-1]] = _compact({"w": _attr(child, "w"), "type": _attr(child, "type")})
    props = {
        "style_id": _attr(style, "val"),
        "width": _attr(width, "w"), "width_type": _attr(width, "type"),
        "layout": _attr(layout, "type"), "alignment": _attr(align, "val"),
        "grid_column_widths": grid_widths or None,
        "cell_margins": margins or None,
        "shading_fill": _attr(shading, "fill"),
        "borders": _border_dict(borders) or None,
    }
    # Retain the existing raw key for Instant compatibility.
    if borders is not None:
        from lxml import etree
        props["borders_xml"] = etree.tostring(borders, encoding="unicode")
    return _compact(props)


def parse_table(node, ids, path, package) -> Table:
    from word_replica.parser.parser import parse_blocks

    rows: list[TableRow] = []
    for r_index, tr in enumerate(node.findall("w:tr", namespaces=NS)):
        tr_pr = tr.find("w:trPr", namespaces=NS)
        height = tr_pr.find("w:trHeight", namespaces=NS) if tr_pr is not None else None
        row_props = _compact({
            "height": _attr(height, "val"), "height_rule": _attr(height, "hRule"),
            "repeat_header": tr_pr.find("w:tblHeader", namespaces=NS) is not None if tr_pr is not None else None,
            "cant_split": tr_pr.find("w:cantSplit", namespaces=NS) is not None if tr_pr is not None else None,
        })
        cells: list[TableCell] = []
        for c_index, tc in enumerate(tr.findall("w:tc", namespaces=NS)):
            tc_pr = tc.find("w:tcPr", namespaces=NS)
            span = tc_pr.find("w:gridSpan", namespaces=NS) if tc_pr is not None else None
            merge = tc_pr.find("w:vMerge", namespaces=NS) if tc_pr is not None else None
            width = tc_pr.find("w:tcW", namespaces=NS) if tc_pr is not None else None
            valign = tc_pr.find("w:vAlign", namespaces=NS) if tc_pr is not None else None
            shading = tc_pr.find("w:shd", namespaces=NS) if tc_pr is not None else None
            borders = tc_pr.find("w:tcBorders", namespaces=NS) if tc_pr is not None else None
            text_dir = tc_pr.find("w:textDirection", namespaces=NS) if tc_pr is not None else None
            margins_node = tc_pr.find("w:tcMar", namespaces=NS) if tc_pr is not None else None
            margins = {}
            if margins_node is not None:
                for child in margins_node:
                    margins[child.tag.rsplit("}", 1)[-1]] = _compact({"w": _attr(child, "w"), "type": _attr(child, "type")})
            props = _compact({
                "grid_span": int(_attr(span, "val") or "1"),
                "v_merge": _attr(merge, "val") or ("continue" if merge is not None else None),
                "width": _attr(width, "w"), "width_type": _attr(width, "type"),
                "vertical_alignment": _attr(valign, "val"),
                "borders": _border_dict(borders) or None,
                "shading_fill": _attr(shading, "fill"),
                "cell_margins": margins or None,
                "text_direction": _attr(text_dir, "val"),
            })
            cell_path = f"{path}/row/{r_index}/cell/{c_index}"
            cells.append(TableCell(ids.make("cell", cell_path), parse_blocks(tc, ids, cell_path, package), props))
        rows.append(TableRow(ids.make("row", f"{path}/row/{r_index}"), cells, row_props))
    return Table(ids.make("table", path), rows=rows, properties=parse_table_properties(node))
