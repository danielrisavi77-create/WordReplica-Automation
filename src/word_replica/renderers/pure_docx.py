from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
import os
import shutil
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from lxml import etree

from word_replica.domain.model import DocumentModel, Paragraph, Run, Section, Table
from word_replica.domain.results import WarningItem
from word_replica.renderers.base import RenderResult

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
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


def _run_element(run: Run):
    node = _w("r")
    r_pr = None
    props = run.properties
    if any(key in props for key in ("bold", "italic", "underline", "hidden")) or run.hidden:
        r_pr = etree.SubElement(node, f"{W}rPr")
        _bool_element(r_pr, "b", props.get("bold"))
        _bool_element(r_pr, "i", props.get("italic"))
        _bool_element(r_pr, "u", props.get("underline"))
        _bool_element(r_pr, "vanish", run.hidden or props.get("hidden"))
    if not run.text:
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
    return node


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
    for run in paragraph.runs:
        p.append(_run_element(run))
    return p


def _table_element(table: Table, sections: list[Section]):
    tbl = _w("tbl")
    if table.properties:
        tbl_pr = etree.SubElement(tbl, f"{W}tblPr")
        if table.properties.get("width") is not None:
            tbl_w = etree.SubElement(tbl_pr, f"{W}tblW")
            _set_w(tbl_w, "w", table.properties.get("width"))
            _set_w(tbl_w, "type", table.properties.get("width_type") or "dxa")
        if table.properties.get("layout") is not None:
            layout = etree.SubElement(tbl_pr, f"{W}tblLayout")
            _set_w(layout, "type", table.properties.get("layout"))
        if table.properties.get("shading_fill") is not None:
            shd = etree.SubElement(tbl_pr, f"{W}shd")
            _set_w(shd, "fill", table.properties.get("shading_fill"))
    for row in table.rows:
        tr = etree.SubElement(tbl, f"{W}tr")
        for cell in row.cells:
            tc = etree.SubElement(tr, f"{W}tc")
            tc_pr = etree.SubElement(tc, f"{W}tcPr")
            span = cell.properties.get("grid_span", 1)
            if span and int(span) > 1:
                grid_span = etree.SubElement(tc_pr, f"{W}gridSpan")
                _set_w(grid_span, "val", span)
            if cell.properties.get("v_merge") is not None:
                v_merge = etree.SubElement(tc_pr, f"{W}vMerge")
                if cell.properties.get("v_merge") != "continue":
                    _set_w(v_merge, "val", cell.properties.get("v_merge"))
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


def _inject_bookmarks_and_fields(body, model: DocumentModel) -> None:
    paragraphs = body.findall(f"{W}p")
    if not paragraphs:
        return

    first = paragraphs[0]
    insert_at = 1 if len(first) and first[0].tag == f"{W}pPr" else 0
    for index, bookmark in enumerate(model.bookmarks, start=1):
        bookmark_id = str(bookmark.bookmark_id or index)
        start = etree.Element(f"{W}bookmarkStart")
        _set_w(start, "id", bookmark_id)
        _set_w(start, "name", bookmark.name)
        first.insert(insert_at, start)
        insert_at += 1
        end = etree.Element(f"{W}bookmarkEnd")
        _set_w(end, "id", bookmark_id)
        first.append(end)

    if model.fields:
        target = paragraphs[-1]
        # The parser models field definitions separately from visible result
        # text. Recreate the instruction sequence without adding new visible
        # text or a new paragraph, so L0/body structure remains unchanged.
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

    def set_document_body(self, body) -> None:
        root = self._xml("word/document.xml")
        current = root.find(f"{W}body")
        if current is None:
            raise RuntimeError("Fresh DOCX shell has no document body")
        root.replace(current, deepcopy(body))
        self._write_xml("word/document.xml", root)

    def _ensure_override(self, part_name: str, content_type: str) -> None:
        root = self._xml("[Content_Types].xml")
        xpath = f"//*[local-name()='Override'][@PartName='/{part_name}']"
        if not root.xpath(xpath):
            node = etree.SubElement(root, f"{{{CT_NS}}}Override")
            node.set("PartName", f"/{part_name}")
            node.set("ContentType", content_type)
            self._write_xml("[Content_Types].xml", root)

    def _ensure_relationship(self, rels_part: str, rel_type: str, target: str) -> str:
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
        self._write_xml(rels_part, root)
        return rel_id


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
        if raw is not None:
            self.parts["word/styles.xml"] = raw

    def set_numbering(self, raw: bytes | None) -> None:
        if raw is None:
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
            return
        root = etree.fromstring(raw)
        if not tracked_changes_enabled:
            for node in root.findall(f"{W}trackRevisions"):
                root.remove(node)
        self._write_xml("word/settings.xml", root)

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
        self._package.initialize_truthful_lifecycle()
        stages = [
            ("body", lambda: self._package.set_document_body(build_body_xml(model))),
            ("styles", lambda: self._package.set_styles(model.styles_xml)),
            ("numbering", lambda: self._package.set_numbering(model.numbering_xml)),
            (
                "settings",
                lambda: self._package.set_settings(
                    model.settings_xml,
                    bool(model.extras.get("tracked_changes_enabled", False)),
                ),
            ),
            ("assets", lambda: self._install_assets(model)),
            ("notes_headers", lambda: self._install_notes_headers(model)),
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

    def _install_assets(self, model: DocumentModel) -> None:
        assert self._package is not None
        for asset in model.assets.values():
            self._package.install_asset(asset.part_name, asset.bytes_data, asset.content_type)
        if model.assets:
            self._package.warnings.append(WarningItem(
                code="PURE_DOCX_ASSET_POSITION_UNAVAILABLE",
                message="Media bytes were retained, but exact inline/anchor positions are not yet represented in the canonical model",
                affects_status=True,
            ))
        for part in model.preserved_parts.values():
            self._package.warnings.append(
                WarningItem(
                    code="UNSUPPORTED_TRANSFER_PART",
                    message=f"Could not transfer {part.part_name}",
                )
            )

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
