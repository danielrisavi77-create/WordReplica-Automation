from hashlib import sha256
from pathlib import Path
from typing import Any

from word_replica.domain.errors import PackageReadError
from word_replica.domain.model import (
    Bookmark,
    Comment,
    DocumentModel,
    DrawingRef,
    ElementIdFactory,
    Field,
    Paragraph,
    PreservedPart,
    RevisionSpan,
    Run,
    Section,
)
from word_replica.opc.package_reader import DocxPackage
from word_replica.opc.properties import read_properties
from word_replica.parser.relationships import collect_relationships, content_type_for, extract_asset, resolve_relationship_target
from word_replica.parser.text import NS, W_NS, bool_prop, run_text
from word_replica.services.source_guard import sha256_file


def sha256_file_bytes(data: bytes) -> str:
    return sha256(data).hexdigest()


def _parse_theme_font_scheme(theme_parts: dict[str, bytes]) -> dict[str, str]:
    if not theme_parts:
        return {}
    from lxml import etree

    data = next(iter(theme_parts.values()), None)
    if not data:
        return {}
    root = etree.fromstring(data)
    a_ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:fontScheme", namespaces=a_ns)
    if scheme is None:
        return {}

    result: dict[str, str] = {}
    for prefix, child_name in (("major", "majorFont"), ("minor", "minorFont")):
        font = scheme.find(f"a:{child_name}", namespaces=a_ns)
        if font is None:
            continue
        latin = font.find("a:latin", namespaces=a_ns)
        east_asia = font.find("a:ea", namespaces=a_ns)
        complex_script = font.find("a:cs", namespaces=a_ns)
        latin_name = (latin.get("typeface") if latin is not None else "") or ""
        east_asia_name = (east_asia.get("typeface") if east_asia is not None else "") or latin_name
        complex_script_name = (complex_script.get("typeface") if complex_script is not None else "") or latin_name
        if latin_name:
            result[f"{prefix}HAnsi"] = latin_name
            result[f"{prefix}Ascii"] = latin_name
        if east_asia_name:
            result[f"{prefix}EastAsia"] = east_asia_name
        if complex_script_name:
            result[f"{prefix}Bidi"] = complex_script_name
    return result


def _parse_style_definitions(data: bytes | None) -> dict[str, dict[str, Any]]:
    if not data:
        return {}
    from lxml import etree

    root = etree.fromstring(data)
    result: dict[str, dict[str, Any]] = {}
    for style in root.findall("w:style", namespaces=NS):
        style_id = _attr(style, "styleId")
        if not style_id:
            continue
        style_type = _attr(style, "type") or "paragraph"
        name_node = style.find("w:name", namespaces=NS)
        based_on = style.find("w:basedOn", namespaces=NS)
        next_style = style.find("w:next", namespaces=NS)
        r_pr = style.find("w:rPr", namespaces=NS)
        p_pr = style.find("w:pPr", namespaces=NS)
        r_fonts = r_pr.find("w:rFonts", namespaces=NS) if r_pr is not None else None
        size = r_pr.find("w:sz", namespaces=NS) if r_pr is not None else None
        color = r_pr.find("w:color", namespaces=NS) if r_pr is not None else None
        highlight = r_pr.find("w:highlight", namespaces=NS) if r_pr is not None else None
        vert_align = r_pr.find("w:vertAlign", namespaces=NS) if r_pr is not None else None
        spacing = p_pr.find("w:spacing", namespaces=NS) if p_pr is not None else None
        indent = p_pr.find("w:ind", namespaces=NS) if p_pr is not None else None
        align = p_pr.find("w:jc", namespaces=NS) if p_pr is not None else None
        run_properties = _compact({
            "bold": bool_prop(r_pr, "b"),
            "italic": bool_prop(r_pr, "i"),
            "underline": bool_prop(r_pr, "u"),
            "strike": bool_prop(r_pr, "strike"),
            "font_ascii": _attr(r_fonts, "ascii"),
            "font_hansi": _attr(r_fonts, "hAnsi"),
            "font_east_asia": _attr(r_fonts, "eastAsia"),
            "font_cs": _attr(r_fonts, "cs"),
            "font_ascii_theme": _attr(r_fonts, "asciiTheme"),
            "font_hansi_theme": _attr(r_fonts, "hAnsiTheme"),
            "font_east_asia_theme": _attr(r_fonts, "eastAsiaTheme"),
            "font_cs_theme": _attr(r_fonts, "cstheme"),
            "size_half_points": _attr(size, "val"),
            "color": _attr(color, "val"),
            "highlight": _attr(highlight, "val"),
            "vert_align": _attr(vert_align, "val"),
        })
        paragraph_properties = _compact({
            "keepNext": bool_prop(p_pr, "keepNext"),
            "pageBreakBefore": bool_prop(p_pr, "pageBreakBefore"),
            "keepLines": bool_prop(p_pr, "keepLines"),
            "widowControl": bool_prop(p_pr, "widowControl"),
            "alignment": _attr(align, "val"),
            "spacing_before": _attr(spacing, "before"),
            "spacing_after": _attr(spacing, "after"),
            "spacing_line": _attr(spacing, "line"),
            "spacing_line_rule": _attr(spacing, "lineRule"),
            "indent_left": _attr(indent, "left"),
            "indent_right": _attr(indent, "right"),
            "indent_first_line": _attr(indent, "firstLine"),
            "indent_hanging": _attr(indent, "hanging"),
        })
        result[style_id] = _compact({
            "style_id": style_id,
            "name": _attr(name_node, "val") or style_id,
            "type": style_type,
            "based_on": _attr(based_on, "val"),
            "next_style": _attr(next_style, "val"),
            "run_properties": run_properties,
            "paragraph_properties": paragraph_properties,
        })
    return result



def _parse_default_paragraph_style_id(data: bytes | None) -> str | None:
    if not data:
        return None
    from lxml import etree

    root = etree.fromstring(data)
    for style in root.findall("w:style", namespaces=NS):
        if _attr(style, "type") == "paragraph" and _attr(style, "default") in {"1", "true", "on"}:
            return _attr(style, "styleId")
    return None

def _parse_document_defaults(data: bytes | None) -> dict[str, dict[str, Any]]:
    if not data:
        return {"run_properties": {}, "paragraph_properties": {}}
    from lxml import etree

    root = etree.fromstring(data)
    r_pr = root.find("w:docDefaults/w:rPrDefault/w:rPr", namespaces=NS)
    p_pr = root.find("w:docDefaults/w:pPrDefault/w:pPr", namespaces=NS)
    r_fonts = r_pr.find("w:rFonts", namespaces=NS) if r_pr is not None else None
    size = r_pr.find("w:sz", namespaces=NS) if r_pr is not None else None
    lang = r_pr.find("w:lang", namespaces=NS) if r_pr is not None else None
    spacing = p_pr.find("w:spacing", namespaces=NS) if p_pr is not None else None
    run_properties = _compact({
        "bold": bool_prop(r_pr, "b"),
        "italic": bool_prop(r_pr, "i"),
        "underline": bool_prop(r_pr, "u"),
        "strike": bool_prop(r_pr, "strike"),
        "font_ascii": _attr(r_fonts, "ascii"),
        "font_hansi": _attr(r_fonts, "hAnsi"),
        "font_east_asia": _attr(r_fonts, "eastAsia"),
        "font_cs": _attr(r_fonts, "cs"),
        "font_ascii_theme": _attr(r_fonts, "asciiTheme"),
        "font_hansi_theme": _attr(r_fonts, "hAnsiTheme"),
        "font_east_asia_theme": _attr(r_fonts, "eastAsiaTheme"),
        "font_cs_theme": _attr(r_fonts, "cstheme"),
        "size_half_points": _attr(size, "val"),
        "language": _attr(lang, "val"),
        "language_east_asia": _attr(lang, "eastAsia"),
        "language_bidi": _attr(lang, "bidi"),
    })
    paragraph_properties = _compact({
        "spacing_before": _attr(spacing, "before"),
        "spacing_after": _attr(spacing, "after"),
        "spacing_line": _attr(spacing, "line"),
        "spacing_line_rule": _attr(spacing, "lineRule"),
    })
    return {"run_properties": run_properties, "paragraph_properties": paragraph_properties}


def _attr(node, name: str) -> str | None:
    if node is None:
        return None
    return node.get(f"{{{W_NS}}}{name}")


def _compact(values: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in values.items() if value is not None}


def parse_run(node, ids: ElementIdFactory, path: str) -> Run:
    r_pr = node.find("w:rPr", namespaces=NS)
    r_fonts = r_pr.find("w:rFonts", namespaces=NS) if r_pr is not None else None
    size = r_pr.find("w:sz", namespaces=NS) if r_pr is not None else None
    color = r_pr.find("w:color", namespaces=NS) if r_pr is not None else None
    highlight = r_pr.find("w:highlight", namespaces=NS) if r_pr is not None else None
    vert_align = r_pr.find("w:vertAlign", namespaces=NS) if r_pr is not None else None
    lang = r_pr.find("w:lang", namespaces=NS) if r_pr is not None else None
    char_spacing = r_pr.find("w:spacing", namespaces=NS) if r_pr is not None else None
    position = r_pr.find("w:position", namespaces=NS) if r_pr is not None else None
    properties = _compact(
        {
            "bold": bool_prop(r_pr, "b"),
            "italic": bool_prop(r_pr, "i"),
            "underline": bool_prop(r_pr, "u"),
            "hidden": bool_prop(r_pr, "vanish"),
            "strike": bool_prop(r_pr, "strike"),
            "font_ascii": _attr(r_fonts, "ascii"),
            "font_hansi": _attr(r_fonts, "hAnsi"),
            "font_east_asia": _attr(r_fonts, "eastAsia"),
            "font_cs": _attr(r_fonts, "cs"),
            "font_ascii_theme": _attr(r_fonts, "asciiTheme"),
            "font_hansi_theme": _attr(r_fonts, "hAnsiTheme"),
            "font_east_asia_theme": _attr(r_fonts, "eastAsiaTheme"),
            "font_cs_theme": _attr(r_fonts, "cstheme"),
            "size_half_points": _attr(size, "val"),
            "color": _attr(color, "val"),
            "highlight": _attr(highlight, "val"),
            "vert_align": _attr(vert_align, "val"),
            "language": _attr(lang, "val"),
            "language_east_asia": _attr(lang, "eastAsia"),
            "language_bidi": _attr(lang, "bidi"),
            "character_spacing": _attr(char_spacing, "val"),
            "character_position": _attr(position, "val"),
        }
    )
    break_types: list[str] = []
    content_tokens: list[dict[str, str]] = []
    for child in node:
        local = child.tag.rsplit("}", 1)[-1]
        if local in {"t", "delText"}:
            content_tokens.append({"kind": "text", "value": child.text or ""})
        elif local == "tab":
            content_tokens.append({"kind": "tab"})
        elif local == "br":
            break_type = _attr(child, "type") or "line"
            break_types.append(break_type)
            content_tokens.append({"kind": "page_break" if break_type == "page" else "line_break"})
        elif local == "cr":
            break_types.append("line")
            content_tokens.append({"kind": "line_break"})
        elif local == "drawing":
            blips = child.xpath(".//*[local-name()='blip']")
            if blips:
                relationship_id = blips[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if relationship_id:
                    content_tokens.append({"kind": "drawing", "relationship_id": relationship_id})
        elif local == "footnoteReference":
            note_id = _attr(child, "id")
            if note_id is not None:
                content_tokens.append({"kind": "footnote_ref", "note_id": note_id})
        elif local == "endnoteReference":
            note_id = _attr(child, "id")
            if note_id is not None:
                content_tokens.append({"kind": "endnote_ref", "note_id": note_id})
        elif local == "fldChar":
            field_type = _attr(child, "fldCharType")
            if field_type in {"begin", "separate", "end"}:
                content_tokens.append({"kind": f"field_{field_type}"})
        elif local == "instrText":
            content_tokens.append({"kind": "field_instruction", "value": child.text or ""})
    if break_types:
        properties["break_types"] = break_types
    if content_tokens:
        properties["content_tokens"] = content_tokens
    return Run(
        ids.make("run", path),
        text=run_text(node),
        properties=properties,
        hidden=properties.get("hidden", False) is True,
    )


def parse_paragraph(node, ids: ElementIdFactory, path: str, package: DocxPackage | None = None) -> Paragraph:
    p_pr = node.find("w:pPr", namespaces=NS)
    style = p_pr.find("w:pStyle", namespaces=NS) if p_pr is not None else None
    spacing = p_pr.find("w:spacing", namespaces=NS) if p_pr is not None else None
    indent = p_pr.find("w:ind", namespaces=NS) if p_pr is not None else None
    align = p_pr.find("w:jc", namespaces=NS) if p_pr is not None else None
    num_pr = p_pr.find("w:numPr", namespaces=NS) if p_pr is not None else None
    num_id = num_pr.find("w:numId", namespaces=NS) if num_pr is not None else None
    ilvl = num_pr.find("w:ilvl", namespaces=NS) if num_pr is not None else None
    tabs_node = p_pr.find("w:tabs", namespaces=NS) if p_pr is not None else None
    tab_stops = []
    if tabs_node is not None:
        for tab in tabs_node.findall("w:tab", namespaces=NS):
            tab_stops.append(_compact({"val": _attr(tab, "val"), "pos": _attr(tab, "pos"), "leader": _attr(tab, "leader")}))
    borders_node = p_pr.find("w:pBdr", namespaces=NS) if p_pr is not None else None
    borders = {}
    if borders_node is not None:
        for border in borders_node:
            local = border.tag.rsplit("}", 1)[-1]
            borders[local] = _compact({"val": _attr(border, "val"), "sz": _attr(border, "sz"), "space": _attr(border, "space"), "color": _attr(border, "color")})
    shading = p_pr.find("w:shd", namespaces=NS) if p_pr is not None else None
    properties = _compact(
        {
            "keepNext": bool_prop(p_pr, "keepNext"),
            "pageBreakBefore": bool_prop(p_pr, "pageBreakBefore"),
            "keepLines": bool_prop(p_pr, "keepLines"),
            "widowControl": bool_prop(p_pr, "widowControl"),
            "alignment": _attr(align, "val"),
            "spacing_before": _attr(spacing, "before"),
            "spacing_after": _attr(spacing, "after"),
            "spacing_line": _attr(spacing, "line"),
            "spacing_line_rule": _attr(spacing, "lineRule"),
            "indent_left": _attr(indent, "left"),
            "indent_right": _attr(indent, "right"),
            "indent_first_line": _attr(indent, "firstLine"),
            "indent_hanging": _attr(indent, "hanging"),
            "numId": _attr(num_id, "val"),
            "ilvl": _attr(ilvl, "val"),
            "tab_stops": tab_stops or None,
            "borders": borders or None,
            "shading": _compact({"val": _attr(shading, "val"), "fill": _attr(shading, "fill"), "color": _attr(shading, "color")}) if shading is not None else None,
        }
    )
    runs: list[Run] = []
    inline_markers: list[dict[str, Any]] = []
    r_index = 0
    for child_index, child in enumerate(node):
        local = child.tag.rsplit("}", 1)[-1]
        if local == "bookmarkStart":
            name = _attr(child, "name")
            if name and name != "_GoBack":
                inline_markers.append({"kind": "bookmark_start", "run_index": r_index, "bookmark_id": _attr(child, "id"), "name": name})
        elif local == "bookmarkEnd":
            inline_markers.append({"kind": "bookmark_end", "run_index": r_index, "bookmark_id": _attr(child, "id")})
        elif local == "r":
            runs.append(parse_run(child, ids, f"{path}/run/{r_index}"))
            r_index += 1
        elif local in {"ins", "moveTo", "fldSimple", "hyperlink"}:
            for nested_index, nested_run in enumerate(child.findall(".//w:r", namespaces=NS)):
                runs.append(
                    parse_run(
                        nested_run,
                        ids,
                        f"{path}/{local}/{child_index}/run/{nested_index}",
                    )
                )
                r_index += 1
        # Deleted/move-from revision text is preserved as evidence but never
        # promoted into visible-final Run.text.
    if inline_markers:
        properties["inline_markers"] = inline_markers
    return Paragraph(
        ids.make("paragraph", path),
        runs=runs,
        style_id=_attr(style, "val"),
        properties=properties,
    )


def parse_section(node, ids: ElementIdFactory, path: str) -> Section:
    pg_sz = node.find("w:pgSz", namespaces=NS)
    pg_mar = node.find("w:pgMar", namespaces=NS)
    cols = node.find("w:cols", namespaces=NS)
    pg_num = node.find("w:pgNumType", namespaces=NS)
    section_type = node.find("w:type", namespaces=NS)
    header_refs = [
        {"type": _attr(ref, "type") or "default", "rel_id": ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")}
        for ref in node.findall("w:headerReference", namespaces=NS)
    ]
    footer_refs = [
        {"type": _attr(ref, "type") or "default", "rel_id": ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")}
        for ref in node.findall("w:footerReference", namespaces=NS)
    ]
    properties = _compact(
        {
            "width": _attr(pg_sz, "w"),
            "height": _attr(pg_sz, "h"),
            "orientation": _attr(pg_sz, "orient") or "portrait",
            "margin_top": _attr(pg_mar, "top"),
            "margin_right": _attr(pg_mar, "right"),
            "margin_bottom": _attr(pg_mar, "bottom"),
            "margin_left": _attr(pg_mar, "left"),
            "header_distance": _attr(pg_mar, "header"),
            "footer_distance": _attr(pg_mar, "footer"),
            "gutter": _attr(pg_mar, "gutter"),
            "columns": _attr(cols, "num"),
            "column_space": _attr(cols, "space"),
            "page_number_start": _attr(pg_num, "start"),
            "break_type": _attr(section_type, "val") or "nextPage",
            "title_page": node.find("w:titlePg", namespaces=NS) is not None,
            "header_refs": header_refs or None,
            "footer_refs": footer_refs or None,
        }
    )
    return Section(ids.make("section", path), properties=properties)


def parse_blocks(parent, ids: ElementIdFactory, source_path: str, package: DocxPackage) -> list[object]:
    blocks: list[object] = []
    for index, child in enumerate(parent):
        local = child.tag.rsplit("}", 1)[-1]
        path = f"{source_path}/{index}"
        if local == "p":
            blocks.append(parse_paragraph(child, ids, path, package))
        elif local == "tbl":
            from word_replica.parser.tables import parse_table
            blocks.append(parse_table(child, ids, path, package))
        elif local == "sdt":
            # A content control (e.g. Word's automatic Table of Contents,
            # a rich-text placeholder). Its paragraphs/tables live one level
            # deeper, in sdtContent — flatten it away so that content is not
            # silently dropped from the reconstruction.
            content = child.find("w:sdtContent", namespaces=NS)
            if content is not None:
                blocks.extend(parse_blocks(content, ids, path, package))
    return blocks


_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _int_or_none(value):
    try:
        return int(value) if value is not None else None
    except (TypeError, ValueError):
        return None


def _bool_attr(value):
    if value is None:
        return None
    return str(value).lower() not in {"0", "false", "off", "none"}


def _extract_body_drawings(root, package: DocxPackage, ids: ElementIdFactory, model: DocumentModel) -> None:
    tree = root.getroottree()
    rels = package.relationships("word/document.xml")
    assets_by_part = {asset.part_name: asset for asset in model.assets.values()}
    relation_index: dict[str, list[str]] = {}
    for drawing_node in root.xpath("//w:drawing", namespaces=NS):
        inline = drawing_node.find(f".//{{{_WP_NS}}}inline")
        anchor = drawing_node.find(f".//{{{_WP_NS}}}anchor")
        container = inline if inline is not None else anchor
        if container is None:
            continue
        blips = drawing_node.xpath(".//*[local-name()='blip']")
        if not blips:
            continue
        rel_id = blips[0].get(f"{{{_R_NS}}}embed")
        rel = rels.get(rel_id) if rel_id else None
        if rel is None or (rel.target_mode or "").lower() == "external":
            continue
        target = resolve_relationship_target("word/document.xml", rel.target)
        asset = assets_by_part.get(target)
        if asset is None:
            continue
        extent = container.find(f"{{{_WP_NS}}}extent")
        position_h = anchor.find(f"{{{_WP_NS}}}positionH") if anchor is not None else None
        position_v = anchor.find(f"{{{_WP_NS}}}positionV") if anchor is not None else None
        pos_h_offset = position_h.find(f"{{{_WP_NS}}}posOffset") if position_h is not None else None
        pos_v_offset = position_v.find(f"{{{_WP_NS}}}posOffset") if position_v is not None else None
        wrap_type = None
        if anchor is not None:
            for candidate in ("wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom", "wrapNone"):
                if anchor.find(f"{{{_WP_NS}}}{candidate}") is not None:
                    wrap_type = candidate.replace("wrap", "", 1)
                    wrap_type = wrap_type[:1].lower() + wrap_type[1:]
                    break
        src_rect = drawing_node.find(f".//{{{_A_NS}}}srcRect")
        crop = {}
        if src_rect is not None:
            for key in ("l", "t", "r", "b"):
                val = _int_or_none(src_rect.get(key))
                if val is not None:
                    crop[key] = val
        xfrm = drawing_node.find(f".//{{{_A_NS}}}xfrm")
        rot_raw = _int_or_none(xfrm.get("rot")) if xfrm is not None else None
        rotation = (rot_raw / 60000.0) if rot_raw is not None else 0.0
        locks = drawing_node.xpath(".//*[local-name()='graphicFrameLocks' or local-name()='picLocks']")
        lock_aspect = None
        if locks:
            lock_aspect = _bool_attr(locks[0].get("noChangeAspect"))
        element_id = ids.make("drawing", tree.getpath(drawing_node))
        model.drawings.append(DrawingRef(
            element_id=element_id,
            asset_id=asset.asset_id,
            source_path=target,
            representation="inline" if inline is not None else "floating",
            width_emu=_int_or_none(extent.get("cx")) if extent is not None else None,
            height_emu=_int_or_none(extent.get("cy")) if extent is not None else None,
            lock_aspect_ratio=lock_aspect,
            wrap_type=wrap_type,
            horizontal_relative_from=position_h.get("relativeFrom") if position_h is not None else None,
            horizontal_position_emu=_int_or_none(pos_h_offset.text) if pos_h_offset is not None else None,
            vertical_relative_from=position_v.get("relativeFrom") if position_v is not None else None,
            vertical_position_emu=_int_or_none(pos_v_offset.text) if pos_v_offset is not None else None,
            distance_top_emu=_int_or_none(container.get("distT")),
            distance_bottom_emu=_int_or_none(container.get("distB")),
            distance_left_emu=_int_or_none(container.get("distL")),
            distance_right_emu=_int_or_none(container.get("distR")),
            crop=crop,
            rotation_degrees=rotation,
            behind_text=_bool_attr(anchor.get("behindDoc")) if anchor is not None else False,
            z_order=_int_or_none(anchor.get("relativeHeight")) if anchor is not None else None,
        ))
        relation_index.setdefault(rel_id, []).append(element_id)
    model.extras["drawing_relationship_index"] = relation_index


class DocxParser:
    def parse(self, path: Path) -> DocumentModel:
        source_hash = sha256_file(path)
        ids = ElementIdFactory(source_hash)
        with DocxPackage.open(path) as package:
            if "word/document.xml" not in package.parts:
                raise PackageReadError("DOCX is missing word/document.xml")
            root = package.read_xml("word/document.xml")
            body = root.find("w:body", namespaces=NS)
            if body is None:
                raise PackageReadError("word/document.xml is missing w:body")
            model = DocumentModel(
                source_sha256=source_hash,
                styles_xml=package.read_bytes("word/styles.xml") if "word/styles.xml" in package.parts else None,
                numbering_xml=package.read_bytes("word/numbering.xml") if "word/numbering.xml" in package.parts else None,
                settings_xml=package.read_bytes("word/settings.xml") if "word/settings.xml" in package.parts else None,
            )
            model.extras["source_properties"] = read_properties(package)
            model.extras["style_definitions"] = _parse_style_definitions(model.styles_xml)
            model.extras["default_paragraph_style_id"] = _parse_default_paragraph_style_id(model.styles_xml)
            model.extras["document_defaults"] = _parse_document_defaults(model.styles_xml)
            model.theme_parts = {part: package.read_bytes(part) for part in package.iter_parts("word/theme/")}
            model.extras["theme_font_scheme"] = _parse_theme_font_scheme(model.theme_parts)
            model.body = parse_blocks(body, ids, "body", package)

            for part in package.iter_parts("word/media/"):
                asset = extract_asset(package, part)
                model.assets.setdefault(asset.asset_id, asset)

            _extract_body_drawings(root, package, ids, model)

            relationship_sources = ["word/document.xml"]
            relationship_sources += package.iter_parts("word/header")
            relationship_sources += package.iter_parts("word/footer")
            for source_part in relationship_sources:
                if source_part.endswith(".xml"):
                    model.relationships.update(collect_relationships(package, source_part))

            for part in package.iter_parts("word/header"):
                if part.endswith(".xml"):
                    model.headers[part] = parse_blocks(package.read_xml(part), ids, f"header/{part}", package)
            for part in package.iter_parts("word/footer"):
                if part.endswith(".xml"):
                    model.footers[part] = parse_blocks(package.read_xml(part), ids, f"footer/{part}", package)

            def parse_notes(part_name: str, note_tag: str, prefix: str) -> dict[str, list[object]]:
                result: dict[str, list[object]] = {}
                if part_name not in package.parts:
                    return result
                note_root = package.read_xml(part_name)
                for note in note_root.findall(f"w:{note_tag}", namespaces=NS):
                    note_id = note.get(f"{{{W_NS}}}id")
                    if note_id is not None:
                        result[note_id] = parse_blocks(note, ids, f"{prefix}/{note_id}", package)
                return result

            model.footnotes = parse_notes("word/footnotes.xml", "footnote", "footnote")
            model.endnotes = parse_notes("word/endnotes.xml", "endnote", "endnote")
            model.extras["headers"] = model.headers
            model.extras["footers"] = model.footers
            model.extras["footnotes"] = model.footnotes
            model.extras["endnotes"] = model.endnotes

            tree = root.getroottree()
            end_paths = {
                node.get(f"{{{W_NS}}}id"): tree.getpath(node)
                for node in root.xpath("//w:bookmarkEnd", namespaces=NS)
                if node.get(f"{{{W_NS}}}id") is not None
            }
            for node in root.xpath("//w:bookmarkStart", namespaces=NS):
                bookmark_id = node.get(f"{{{W_NS}}}id")
                name = node.get(f"{{{W_NS}}}name")
                if bookmark_id is not None and name is not None and name != "_GoBack":
                    model.bookmarks.append(
                        Bookmark(
                            bookmark_id,
                            name,
                            tree.getpath(node),
                            end_paths.get(bookmark_id),
                        )
                    )

            revision_kind = {
                "ins": "insert",
                "del": "delete",
                "moveFrom": "move_from",
                "moveTo": "move_to",
            }
            for node in root.xpath("//w:ins | //w:del | //w:moveFrom | //w:moveTo", namespaces=NS):
                local = node.tag.rsplit("}", 1)[-1]
                pieces = node.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS)
                path_text = tree.getpath(node)
                model.revisions.append(
                    RevisionSpan(
                        node.get(f"{{{W_NS}}}id") or ids.make("revision", path_text),
                        revision_kind[local],
                        node.get(f"{{{W_NS}}}author"),
                        node.get(f"{{{W_NS}}}date"),
                        path_text,
                        "".join(pieces),
                    )
                )

            instruction_nodes = root.xpath("//w:instrText", namespaces=NS)
            if instruction_nodes:
                instruction = "".join(node.text or "" for node in instruction_nodes).strip()
                first_path = tree.getpath(instruction_nodes[0])
                model.fields.append(
                    Field(
                        ids.make("field", first_path),
                        instruction,
                        "",
                        False,
                    )
                )

            if "word/comments.xml" in package.parts:
                comments_root = package.read_xml("word/comments.xml")
                for comment_node in comments_root.findall("w:comment", namespaces=NS):
                    comment_id = comment_node.get(f"{{{W_NS}}}id")
                    if comment_id is None:
                        continue
                    model.comments[comment_id] = Comment(
                        comment_id,
                        comment_node.get(f"{{{W_NS}}}author"),
                        comment_node.get(f"{{{W_NS}}}date"),
                        parse_blocks(comment_node, ids, f"comment/{comment_id}", package),
                    )

            tracked_changes = False
            if "word/settings.xml" in package.parts:
                settings_root = package.read_xml("word/settings.xml")
                tracked_changes = settings_root.find("w:trackRevisions", namespaces=NS) is not None

            for part in sorted(package.parts):
                if part.startswith(("word/charts/", "word/embeddings/", "word/diagrams/")) and not part.endswith(".rels"):
                    data = package.read_bytes(part)
                    digest = sha256_file_bytes(data)
                    model.preserved_parts[part] = PreservedPart(
                        part,
                        content_type_for(package, part),
                        None,
                        digest,
                        data,
                    )

            model.extras["comments"] = model.comments
            model.extras["revisions"] = model.revisions
            model.extras["bookmarks"] = model.bookmarks
            model.extras["fields"] = model.fields
            model.extras["preserved_parts"] = model.preserved_parts
            model.extras["tracked_changes_enabled"] = tracked_changes

            section_index = 0
            block_index = 0
            for child_index, child in enumerate(body):
                local = child.tag.rsplit("}", 1)[-1]
                sect_pr = None
                current_block_index: int | None = None
                if local in {"p", "tbl"}:
                    current_block_index = block_index
                    block_index += 1
                if local == "p":
                    p_pr = child.find("w:pPr", namespaces=NS)
                    sect_pr = p_pr.find("w:sectPr", namespaces=NS) if p_pr is not None else None
                elif local == "sectPr":
                    sect_pr = child
                if sect_pr is not None:
                    model.sections.append(parse_section(sect_pr, ids, f"section/{section_index}/body/{child_index}"))
                    if current_block_index is not None and current_block_index < len(model.body):
                        paragraph = model.body[current_block_index]
                        if isinstance(paragraph, Paragraph):
                            paragraph.properties["section_index"] = section_index
                    section_index += 1
            return model
