from contextvars import ContextVar
from dataclasses import replace
from hashlib import sha256
from pathlib import Path
from typing import Any

from word_replica.parser.nodes import local_name
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


def _rels_owner(rels_part: str) -> str:
    """The part a .rels file belongs to; "" for the package root."""
    import posixpath

    directory, _, name = rels_part.rpartition("/")
    owner_name = name[: -len(".rels")]
    parent = posixpath.dirname(directory)  # strip the trailing "_rels"
    return posixpath.join(parent, owner_name) if parent else owner_name


def _parse_theme_font_scheme(theme_parts: dict[str, bytes]) -> dict[str, str]:
    if not theme_parts:
        return {}
    from lxml import etree

    # The first theme part is not necessarily the one that declares fonts: a
    # themeOverride sorts ahead of theme1.xml and may carry no fontScheme at
    # all. Take the first that actually has one.
    a_ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = None
    for data in theme_parts.values():
        if not data:
            continue
        try:
            candidate = etree.fromstring(data).find(".//a:fontScheme", namespaces=a_ns)
        except etree.XMLSyntaxError:
            continue
        if candidate is not None:
            scheme = candidate
            break
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


_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PRESERVABLE_INLINE = {"AlternateContent", "pict", "object"}
_IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def _is_preservable_inline(child, local: str) -> bool:
    """Inline content that must be carried verbatim rather than rebuilt.

    A picture is modelled as a DrawingRef, but only its *bytes* and a few
    measurements; the authored anchor geometry, wrapping, crop and effects live
    only in the original XML. Rebuilding that by hand would keep whatever the
    model happens to represent and drop the rest, so the fragment is preserved
    and re-emitted. Shapes, VML and embedded objects are not modelled at all.
    """
    return local in _PRESERVABLE_INLINE


# Which OPC part the blocks currently being parsed came from. A relationship id
# only means something relative to the part that declares it: in a real corpus
# document rId1 addressed word/styles.xml from the document and
# word/media/image1.png from the comments.
_OWNER_PART: ContextVar[str] = ContextVar("owner_part", default="word/document.xml")


def _reference_targets(
    package: DocxPackage | None, owner_part: str
) -> dict[str, tuple[str, str, bool]]:
    """Relationship id -> (target, type, is_external), for one owning part.

    A relationship id is only meaningful relative to the part that declares it.
    In a real corpus document rId1 addressed word/styles.xml from the document
    and word/media/image1.png from the comments, so resolving an id without
    knowing whose it is can only guess.
    """
    if package is None:
        return {}
    try:
        relationships = package.relationships(owner_part)
    except Exception:
        return {}
    # The relationship *type* travels with the target. An OLE object reached
    # through an image relationship is not the same document: Word uses the
    # type to decide what a reference is for.
    #
    # External relationships are included. They resolve to a URL rather than a
    # part, and excluding them meant any fragment containing one -- a shape with
    # a hyperlink on it -- had an id that resolved to nothing and was refused
    # whole, taking the shape and its geometry with it. The refusal looked like
    # the safety rule working.
    resolved: dict[str, tuple[str, str, bool]] = {}
    for rel_id, rel in relationships.items():
        external = rel.target_mode == "External"
        target = rel.target if external else _resolve_relative_target(owner_part, rel.target)
        resolved[rel_id] = (target, rel.rel_type, external)
    return resolved


def _resolve_relative_target(owner_part: str, target: str) -> str:
    from pathlib import PurePosixPath

    if target.startswith("/"):
        candidate = target.lstrip("/")
    else:
        candidate = str(PurePosixPath(owner_part).parent / target)
    resolved: list[str] = []
    for piece in PurePosixPath(candidate).parts:
        if piece == "..":
            if resolved:
                resolved.pop()
        elif piece not in (".", ""):
            resolved.append(piece)
    return "/".join(resolved)


def _capture_inline(child, package: DocxPackage | None) -> dict:
    """Serialize an inline fragment, resolving the parts its references address.

    A fragment carrying an ``r:`` attribute cannot simply be re-emitted:
    relationship ids are renumbered by any writer, so a stale one points at the
    wrong part or none at all, and Word repairs such a document on open.

    Where the id *can* be resolved -- the parser knows which part it addressed,
    and the renderer owns the new package -- the target is recorded so the id
    can be rewritten on the way out. Where it cannot, the fragment is refused;
    a missing shape is a visible loss, a corrupted package is not.
    """
    from lxml import etree

    known = _reference_targets(package, _OWNER_PART.get())
    targets: dict[str, str] = {}
    for element in child.iter():
        if not isinstance(element.tag, str):
            continue
        for name, value in element.attrib.items():
            if not name.startswith(f"{{{_R_NS}}}"):
                continue
            resolved = known.get(value)
            if resolved is None:
                return {
                    "kind": "unsupported_inline",
                    "reason": "fragment carries a relationship reference that cannot be resolved",
                    "tag": local_name(child),
                }
            targets[value] = resolved
    return {
        "kind": "preserved_xml",
        "value": etree.tostring(child, encoding="unicode"),
        "tag": local_name(child),
        "rel_targets": targets,
    }


def parse_run(node, ids: ElementIdFactory, path: str, package: DocxPackage | None = None) -> Run:
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
        local = local_name(child)
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
            # The kind stays "drawing" because the interactive executor
            # dispatches on it and inserts the picture through Word's own
            # object model. The verbatim fragment is added alongside, for the
            # pure-docx renderer, which has to write the markup itself.
            token = _capture_inline(child, package)
            blips = child.xpath(".//*[local-name()='blip']")
            relationship_id = (
                blips[0].get(f"{{{_R_NS}}}embed") if blips else None
            )
            if relationship_id:
                content_tokens.append({**token, "kind": "drawing", "relationship_id": relationship_id})
            elif token["kind"] == "preserved_xml":
                content_tokens.append(token)
        elif local == "footnoteReference":
            note_id = _attr(child, "id")
            if note_id is not None:
                content_tokens.append({"kind": "footnote_ref", "note_id": note_id})
        elif local == "endnoteReference":
            note_id = _attr(child, "id")
            if note_id is not None:
                content_tokens.append({"kind": "endnote_ref", "note_id": note_id})
        elif local == "commentReference":
            # Without this the comment part can be restored but nothing points
            # at it, and a comment no reader anchors is invisible in Word.
            comment_id = _attr(child, "id")
            if comment_id is not None:
                content_tokens.append({"kind": "comment_ref", "comment_id": comment_id})
        elif local == "fldChar":
            field_type = _attr(child, "fldCharType")
            if field_type in {"begin", "separate", "end"}:
                content_tokens.append({"kind": f"field_{field_type}"})
        elif local == "instrText":
            content_tokens.append({"kind": "field_instruction", "value": child.text or ""})
        elif _is_preservable_inline(child, local):
            content_tokens.append(_capture_inline(child, package))
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


def _hyperlink_properties(node, package: DocxPackage | None) -> dict[str, Any]:
    """Where a hyperlink points, in whichever of the two ways it can.

    An external link carries an r:id resolved through the owning part's
    relationships; an internal one carries only a w:anchor naming a bookmark,
    with no relationship at all.
    """
    anchor = node.get(f"{{{W_NS}}}anchor")
    rel_id = node.get(f"{{{_R_NS}}}id")
    target = None
    if rel_id and package is not None:
        try:
            relationship = package.relationships(_OWNER_PART.get()).get(rel_id)
        except Exception:
            relationship = None
        if relationship is not None:
            target = relationship.target
    return {
        "target": target,
        "anchor": anchor,
        "tooltip": node.get(f"{{{W_NS}}}tooltip"),
    }


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
            local = local_name(border)
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
        local = local_name(child)
        if local == "bookmarkStart":
            name = _attr(child, "name")
            if name and name != "_GoBack":
                inline_markers.append({"kind": "bookmark_start", "run_index": r_index, "bookmark_id": _attr(child, "id"), "name": name})
        elif local == "bookmarkEnd":
            inline_markers.append({"kind": "bookmark_end", "run_index": r_index, "bookmark_id": _attr(child, "id")})
        elif local == "r":
            runs.append(parse_run(child, ids, f"{path}/run/{r_index}", package))
            r_index += 1
        elif local in {"ins", "moveTo", "fldSimple", "hyperlink", "sdt"}:
            # Unwrapping reaches the runs inside, which is what the model
            # represents -- but a hyperlink is more than the words it wraps.
            # Without recording it here the link text survives and the link
            # does not, and G0 sees nothing wrong because the text is exactly
            # what does survive.
            link = _hyperlink_properties(child, package) if local == "hyperlink" else None
            # An inline content control matched none of these, so the control
            # and everything in it was dropped without a trace. Unwrapping is
            # what already happens to a block-level control, whose paragraphs
            # the renderer flattens.
            #
            # Only w:sdtContent is searched: w:sdtPr describes the control
            # itself, and a placeholder caption in there is not document text.
            search = "./w:sdtContent//w:r" if local == "sdt" else ".//w:r"
            for nested_index, nested_run in enumerate(child.findall(search, namespaces=NS)):
                run = parse_run(
                    nested_run,
                    ids,
                    f"{path}/{local}/{child_index}/run/{nested_index}",
                    package,
                )
                if link is not None:
                    run.properties["hyperlink"] = {**link, "group": child_index}
                runs.append(run)
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


_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Parts reached from inside the document body. Capturing them lets a caller
# report what was lost; restoring them without the body reference would only
# produce an orphan.
_BODY_REFERENCED_PREFIXES = ("word/charts/", "word/embeddings/", "word/diagrams/")

# Document-level attachments, keyed by relationship type. Nothing else in the
# model represents what these hold, so they survive verbatim or they are gone.
#
# This is an allow-list on purpose. "Anything with a document-level
# relationship" would also match styles.xml, settings.xml, numbering.xml,
# fontTable.xml and theme1.xml -- parts the renderer builds itself, which
# restoring verbatim would silently overwrite.
_RENDERER_AUTHORED_PARTS = frozenset({
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/settings.xml",
})
_ATTACHMENT_RELATIONSHIP_TYPES = frozenset({
    "customXml",
    "customXmlProps",
    "bibliography",
    "people",
    "commentsExtended",
    "commentsIds",
    "commentsExtensible",
    "glossaryDocument",
    "font",
    # The renderer writes styles, numbering and settings itself, but never
    # these. Left out, every rebuild silently ships the shell template's
    # theme, font table and web settings instead of the source's.
    "stylesWithEffects",
    "webSettings",
    "fontTable",
    "theme",
    # word/customizations.xml -- the key map and toolbar customisations saved
    # with the document. A part with a relationship like any other; it was lost
    # only because this list is a list and the type was not on it.
    "keyMapCustomizations",
})


def _sidecars(package: DocxPackage, part_name: str) -> list[tuple[str, str]]:
    """(target, rels-part) for every internal relationship an attachment owns."""
    parent, _, name = part_name.rpartition("/")
    rels_part = f"{parent}/_rels/{name}.rels" if parent else f"_rels/{name}.rels"
    if rels_part not in package.parts:
        return []
    found: list[tuple[str, str]] = []
    for rel in package.relationships(part_name).values():
        if rel.target_mode == "External":
            continue
        target = rel.target.lstrip("/") if rel.target.startswith("/") else f"{parent}/{rel.target}"
        found.append((target, rels_part))
    return found


def _resolve_document_target(target: str) -> str:
    """Resolve a word/document.xml.rels target to a package part name."""
    from pathlib import PurePosixPath

    if target.startswith("/"):
        candidate = target.lstrip("/")
    else:
        candidate = str(PurePosixPath("word", target))
    resolved: list[str] = []
    for piece in PurePosixPath(candidate).parts:
        if piece == "..":
            if resolved:
                resolved.pop()
        elif piece not in (".", ""):
            resolved.append(piece)
    return "/".join(resolved)


def parse_blocks(parent, ids: ElementIdFactory, source_path: str, package: DocxPackage,
                 owner_part: str | None = None) -> list[object]:
    """Parse a story part's blocks.

    ``owner_part`` names the OPC part these blocks came from. It is ambient
    rather than threaded through every signature because it applies to a whole
    subtree -- paragraphs, runs, table cells and their nested blocks alike --
    and one of those levels lives in another module.
    """
    if owner_part is not None:
        token = _OWNER_PART.set(owner_part)
        try:
            return _parse_blocks(parent, ids, source_path, package)
        finally:
            _OWNER_PART.reset(token)
    return _parse_blocks(parent, ids, source_path, package)


def _parse_blocks(parent, ids: ElementIdFactory, source_path: str, package: DocxPackage) -> list[object]:
    blocks: list[object] = []
    for index, child in enumerate(parent):
        local = local_name(child)
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

            # Sweeping word/media/ finds what nearly every document does, and
            # media nothing points at as well. But the folder is a convention,
            # not a rule: a part is wherever its relationship targets, and a
            # picture kept at media/ in the package root never became an asset,
            # so the drawing had nothing to resolve to and the picture was gone.
            media_parts = list(package.iter_parts("word/media/"))
            seen_media = set(media_parts)
            for rels_part in sorted(package.parts):
                if not rels_part.endswith(".rels"):
                    continue
                try:
                    relationships = package.relationships(_rels_owner(rels_part))
                except Exception:
                    continue
                for rel in relationships.values():
                    if rel.target_mode == "External":
                        continue
                    if rel.rel_type.rsplit("/", 1)[-1] != "image":
                        continue
                    target = resolve_relationship_target(_rels_owner(rels_part) or "x", rel.target)
                    if target in package.parts and target not in seen_media:
                        seen_media.add(target)
                        media_parts.append(target)

            for part in media_parts:
                asset = extract_asset(package, part)
                # The id is the content hash, so two parts holding the same
                # bytes claim the same one and setdefault kept only the first --
                # the second part never reached the model and the rebuild came
                # out a part short. In OPC a part's identity is its name, and a
                # document storing one image under two names is not unusual.
                #
                # Colliding ids are disambiguated rather than the scheme being
                # changed: DrawingRef.asset_id, the interactive preflight and
                # the asset files written to disk all key on them, so every
                # document without a collision keeps exactly the ids it had.
                # iter_parts is sorted, so which part takes the bare id is
                # stable across runs.
                asset_id = asset.asset_id
                suffix = 1
                while asset_id in model.assets:
                    if model.assets[asset_id].part_name == part:
                        break
                    suffix += 1
                    asset_id = f"{asset.asset_id}_{suffix}"
                if asset_id != asset.asset_id:
                    asset = replace(asset, asset_id=asset_id)
                model.assets.setdefault(asset_id, asset)

            _extract_body_drawings(root, package, ids, model)

            relationship_sources = ["word/document.xml"]
            relationship_sources += package.iter_parts("word/header")
            relationship_sources += package.iter_parts("word/footer")
            for source_part in relationship_sources:
                if source_part.endswith(".xml"):
                    model.relationships.update(collect_relationships(package, source_part))

            for part in package.iter_parts("word/header"):
                if part.endswith(".xml"):
                    model.headers[part] = parse_blocks(package.read_xml(part), ids, f"header/{part}", package, part)
            for part in package.iter_parts("word/footer"):
                if part.endswith(".xml"):
                    model.footers[part] = parse_blocks(package.read_xml(part), ids, f"footer/{part}", package, part)

            def parse_notes(part_name: str, note_tag: str, prefix: str) -> dict[str, list[object]]:
                result: dict[str, list[object]] = {}
                if part_name not in package.parts:
                    return result
                note_root = package.read_xml(part_name)
                for note in note_root.findall(f"w:{note_tag}", namespaces=NS):
                    note_id = note.get(f"{{{W_NS}}}id")
                    if note_id is not None:
                        result[note_id] = parse_blocks(note, ids, f"{prefix}/{note_id}", package, part_name)
                return result

            model.footnotes = parse_notes("word/footnotes.xml", "footnote", "footnote")
            model.endnotes = parse_notes("word/endnotes.xml", "endnote", "endnote")
            model.extras["headers"] = model.headers
            model.extras["footers"] = model.footers
            model.extras["footnotes"] = model.footnotes
            model.extras["endnotes"] = model.endnotes

            tree = root.getroottree()

            def _paragraph_relative_path(node) -> str:
                # tree.getpath() positions a node among its parent's direct
                # children - which shifts every element after a <w:sdt> content
                # control (e.g. a TOC) whenever the control's contents get
                # flattened to plain paragraphs (as this renderer does), even
                # though nothing about the document's actual content changed.
                # Anchor the path to the enclosing paragraph's position among
                # ALL <w:p> elements in the document instead, which is stable
                # across that flattening (confirmed: identical total <w:p>
                # count between a source with an sdt-wrapped TOC and the
                # rendered output without one).
                full_path = tree.getpath(node)
                ancestor = node
                while ancestor is not None and local_name(ancestor) != "p":
                    ancestor = ancestor.getparent()
                if ancestor is None:
                    return full_path
                ancestor_path = tree.getpath(ancestor)
                if not full_path.startswith(ancestor_path):
                    return full_path
                paragraph_index = all_paragraph_positions.get(ancestor_path)
                if paragraph_index is None:
                    return full_path
                return f"//w:p[{paragraph_index}]" + full_path[len(ancestor_path):]

            # lxml Element identity (id()) is not reliable for matching a node
            # reached via getparent() back to one reached via a fresh xpath()
            # call - key on the (stable, string) xpath instead.
            all_paragraph_positions = {
                tree.getpath(p): position
                for position, p in enumerate(root.xpath("//w:p", namespaces=NS), start=1)
            }
            # _GoBack is Word's volatile "last edit position" mark and is not
            # modelled, but it is still a sibling in the source. A path indexed
            # among *all* siblings therefore says [2] for a bookmark whose
            # rebuild can only ever write [1], and the gate reports a
            # divergence about a bookmark the model never carried. Index among
            # the bookmarks that are kept instead, so both sides count the same
            # things.
            modelled_ids = {
                node.get(f"{{{W_NS}}}id")
                for node in root.xpath("//w:bookmarkStart", namespaces=NS)
                if node.get(f"{{{W_NS}}}name") not in (None, "_GoBack")
            }

            # A bookmark need not be inside a paragraph: both marks are allowed
            # wherever block-level content is, as direct children of w:body or
            # inside a w:sdtContent. Those had no enclosing paragraph to anchor
            # to, so they fell back to a raw lxml path carrying no position the
            # renderer could use, and every one was dumped into the first
            # paragraph.
            #
            # The anchor is the nearest paragraph, and the side matters: a start
            # standing before paragraph N opens the bookmark there, an end
            # standing after paragraph N closes it there. Either way it covers
            # the text it covered before, which is all a cross-reference sees.
            #
            # One walk in document order does both jobs -- it finds each mark's
            # paragraph and fixes the order marks share one.
            paragraph_total = len(all_paragraph_positions)
            anchors: dict[object, int] = {}
            grouped: dict[tuple[int, str], list[object]] = {}
            seen_paragraphs = 0
            for node in root.iter():
                if not isinstance(node.tag, str):
                    continue
                tag = local_name(node)
                if tag == "p":
                    seen_paragraphs += 1
                    continue
                if tag not in ("bookmarkStart", "bookmarkEnd"):
                    continue
                if node.get(f"{{{W_NS}}}id") not in modelled_ids:
                    continue
                enclosing = node.getparent()
                while enclosing is not None and local_name(enclosing) != "p":
                    enclosing = enclosing.getparent()
                if enclosing is not None or tag == "bookmarkEnd":
                    target = seen_paragraphs
                else:
                    target = seen_paragraphs + 1
                target = min(max(target, 1), paragraph_total) if paragraph_total else 0
                anchors[node] = target
                grouped.setdefault((target, tag), []).append(node)

            def _anchored_path(node, tag: str) -> str:
                target = anchors.get(node)
                if not target:
                    return _paragraph_relative_path(node)
                siblings = grouped.get((target, tag), [])
                if len(siblings) < 2:
                    # lxml omits the index for an only child; match that.
                    return f"//w:p[{target}]/w:{tag}"
                return f"//w:p[{target}]/w:{tag}[{siblings.index(node) + 1}]"

            end_paths = {
                node.get(f"{{{W_NS}}}id"): _anchored_path(node, "bookmarkEnd")
                for node in root.xpath("//w:bookmarkEnd", namespaces=NS)
                if node.get(f"{{{W_NS}}}id") in modelled_ids
            }
            for node in root.xpath("//w:bookmarkStart", namespaces=NS):
                bookmark_id = node.get(f"{{{W_NS}}}id")
                name = node.get(f"{{{W_NS}}}name")
                if bookmark_id is not None and name is not None and name != "_GoBack":
                    model.bookmarks.append(
                        Bookmark(
                            bookmark_id,
                            name,
                            _anchored_path(node, "bookmarkStart"),
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
                local = local_name(node)
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
                    comment = Comment(
                        comment_id,
                        comment_node.get(f"{{{W_NS}}}author"),
                        comment_node.get(f"{{{W_NS}}}date"),
                        parse_blocks(comment_node, ids, f"comment/{comment_id}", package, "word/comments.xml"),
                    )
                    initials = comment_node.get(f"{{{W_NS}}}initials")
                    if initials:
                        model.extras.setdefault("comment_initials", {})[comment_id] = initials
                    model.comments[comment_id] = comment

            tracked_changes = False
            if "word/settings.xml" in package.parts:
                settings_root = package.read_xml("word/settings.xml")
                tracked_changes = settings_root.find("w:trackRevisions", namespaces=NS) is not None

            def _preserve(
                part: str,
                relationship_type: str | None,
                *,
                sidecar: bool = False,
                owner_rels: str = "word/_rels/document.xml.rels",
            ) -> None:
                if part in model.preserved_parts or part not in package.parts:
                    return
                data = package.read_bytes(part)
                model.preserved_parts[part] = PreservedPart(
                    part,
                    content_type_for(package, part),
                    relationship_type,
                    sha256_file_bytes(data),
                    data,
                    sidecar,
                    owner_rels,
                )

            for part in sorted(package.parts):
                if not part.startswith(_BODY_REFERENCED_PREFIXES) or part.endswith(".rels"):
                    continue
                # A chart or diagram is not self-contained: its own .rels points
                # at the style, colour-style and embedded workbook Word renders
                # it from. Restoring the part without them leaves those in the
                # package with nothing pointing at them.
                _preserve(part, None)
                for _target, sidecar_rels in _sidecars(package, part):
                    _preserve(sidecar_rels, None, sidecar=True)

            # Which relationship reached each part in the source. A package
            # routinely relates a part that no element points at -- an image
            # left behind by editing, or a SmartArt drawing, reached from
            # document.xml.rels while dgm:relIds names only the data, layout,
            # colours and quick-style parts. The rebuild kept such a part and
            # dropped its relationship, turning it into an orphan; with this the
            # renderer can put the relationship back and leave the package as it
            # found it.
            part_relationships: dict[str, tuple[str, str]] = {}
            for rels_part in sorted(package.parts):
                if not rels_part.endswith(".rels"):
                    continue
                owner = _rels_owner(rels_part)
                try:
                    relationships = package.relationships(owner)
                except Exception:
                    continue
                for rel in relationships.values():
                    if rel.target_mode == "External":
                        continue
                    # Relative to the owning part's directory, not the _rels
                    # directory the file happens to live in.
                    target = resolve_relationship_target(owner or "x", rel.target)
                    if target and target not in part_relationships:
                        part_relationships[target] = (rels_part, rel.rel_type)
            model.extras["source_part_relationships"] = part_relationships

            # The same idea for relationships that leave the package. Editing a
            # document strips the link text and leaves the relationship behind,
            # and an external one has no part to travel with it, so nothing
            # noticed it was gone. settings.xml already had this fix for its own
            # external relationships; this is the same rule everywhere else.
            external_relationships: list[tuple[str, str, str]] = []
            for rels_part in sorted(package.parts):
                if not rels_part.endswith(".rels"):
                    continue
                try:
                    relationships = package.relationships(_rels_owner(rels_part))
                except Exception:
                    continue
                for rel in relationships.values():
                    if rel.target_mode != "External":
                        continue
                    external_relationships.append((rels_part, rel.rel_type, rel.target))
            model.extras["source_external_relationships"] = sorted(set(external_relationships))

            # And parts no relationship reaches at all. Every mechanism above
            # starts from a relationship, so a package holding an unreachable
            # part -- LibreOffice leaves word/webSettings.xml and
            # word/stylesWithEffects.xml behind this way -- lost it silently.
            #
            # Word does not read a part it cannot reach, which is the argument
            # for carrying it rather than against: tidying away bytes the source
            # shipped is a change to the package, just one whose harmlessness we
            # would be asserting instead of checking.
            for part in sorted(package.parts):
                if part.endswith((".rels", "/")) or part == "[Content_Types].xml":
                    continue
                if part in part_relationships or part in model.preserved_parts:
                    continue
                # The renderer authors these itself; restoring a source copy
                # would overwrite what it wrote.
                if part in _RENDERER_AUTHORED_PARTS or part.startswith("docProps/"):
                    continue
                _preserve(part, None, sidecar=True)

            for rel in package.relationships("word/document.xml").values():
                if rel.target_mode == "External":
                    continue
                if rel.rel_type.rsplit("/", 1)[-1] not in _ATTACHMENT_RELATIONSHIP_TYPES:
                    continue
                anchor = _resolve_document_target(rel.target)
                _preserve(anchor, rel.rel_type)
                # An attachment can own a sidecar: a customXml item points at
                # its properties part through its own .rels. Restoring the item
                # without them would leave that reference dangling.
                for sidecar, sidecar_rels in _sidecars(package, anchor):
                    _preserve(sidecar, None, sidecar=True)
                    _preserve(sidecar_rels, None, sidecar=True)

            # The package thumbnail hangs off _rels/.rels rather than the
            # document, so it needs its own pass. It is the file preview
            # Explorer shows -- shipping the shell template's is user-visible.
            for rel in package.relationships("").values():
                if rel.target_mode == "External":
                    continue
                if rel.rel_type.rsplit("/", 1)[-1] != "thumbnail":
                    continue
                _preserve(
                    rel.target.lstrip("/"),
                    rel.rel_type,
                    owner_rels="_rels/.rels",
                )

            model.extras["comments"] = model.comments
            model.extras["revisions"] = model.revisions
            model.extras["bookmarks"] = model.bookmarks
            model.extras["fields"] = model.fields
            model.extras["preserved_parts"] = model.preserved_parts
            model.extras["tracked_changes_enabled"] = tracked_changes
            # word/_rels/settings.xml.rels carries the attached template the
            # document was authored from -- usually the author's Normal.dotm.
            # The renderer writes settings.xml but never its .rels, so that
            # relationship was simply disappearing.
            #
            # Only external relationships are carried: an external target needs
            # no part in the package, so restoring it cannot leave anything
            # dangling, while an internal one would need its part to travel too.
            settings_relationships = []
            if "word/_rels/settings.xml.rels" in package.parts:
                for rel in package.relationships("word/settings.xml").values():
                    if rel.target_mode == "External":
                        settings_relationships.append((rel.rel_type, rel.target))
            model.extras["settings_relationships"] = sorted(settings_relationships)
            # styles.xml and numbering.xml are written back byte for byte, so
            # any r:id inside them survives the rebuild and has to keep meaning
            # what it meant. Dropping their .rels leaves the preserved bytes
            # pointing at a relationship that no longer exists -- a picture
            # bullet's image becomes an orphaned part and Word repairs the
            # document on open, which is worse than losing the bullet.
            #
            # The ids are carried verbatim rather than reallocated for the same
            # reason: the bytes that use them are verbatim too.
            verbatim_relationships: dict[str, list[tuple[str, str, str, str | None]]] = {}
            for part in ("word/styles.xml", "word/numbering.xml"):
                if f"word/_rels/{part.rsplit('/', 1)[-1]}.rels" not in package.parts:
                    continue
                entries = [
                    (rel_id, rel.rel_type, rel.target, rel.target_mode)
                    for rel_id, rel in package.relationships(part).items()
                ]
                if entries:
                    verbatim_relationships[part] = sorted(entries)
            model.extras["verbatim_part_relationships"] = verbatim_relationships
            # mc:Ignorable on the document root declares which namespace
            # prefixes a reader may skip. The pure-docx shell template carries
            # its own, so without recording the source's, every rebuild
            # silently adopts the template's declaration -- a difference no
            # model gate can see.
            model.extras["document_ignorable"] = root.get(f"{{{_MC_NS}}}Ignorable")

            section_index = 0
            block_index = 0
            for child_index, child in enumerate(body):
                local = local_name(child)
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
