from __future__ import annotations

import copy
from dataclasses import asdict
from typing import Any

from word_replica.domain.model import DocumentModel, Paragraph, Run, Table
from word_replica.interactive.tables import build_table_plan
from word_replica.domain.reconstruction import ReconstructionBlueprint, ReconstructionEvent, SemanticLocation


class BlueprintCompiler:
    def compile(self, model: DocumentModel) -> ReconstructionBlueprint:
        self._model = model
        self._drawing_by_id = {drawing.element_id: drawing for drawing in model.drawings}
        self._drawing_queues = {key: list(value) for key, value in model.extras.get("drawing_relationship_index", {}).items()}
        self._bookmark_names = {bookmark.bookmark_id: bookmark.name for bookmark in model.bookmarks}
        events: list[ReconstructionEvent] = []
        location = SemanticLocation(story="body")
        if model.sections:
            self._begin_section(0, events)
        self._compile_blocks(model.body, events, location)
        if model.sections:
            events.append(ReconstructionEvent("EndSection", model.sections[-1].element_id, {"section_index": len(model.sections) - 1}))
        counts = {
            "paragraphs": sum(1 for e in events if e.event_type == "BeginParagraph"),
            "tables": sum(1 for e in events if e.event_type == "BeginTable"),
            "images": sum(1 for e in events if e.event_type == "InsertImage"),
            "sections": len(model.sections),
        }
        return ReconstructionBlueprint.build(
            source_sha256=model.source_sha256,
            source_model_fingerprint=model.fingerprint(),
            events=tuple(events),
            semantic_counts=counts,
        )

    def _compile_blocks(
        self,
        blocks: list[object],
        events: list[ReconstructionEvent],
        location: SemanticLocation,
    ) -> None:
        paragraph_index = 0
        table_index = 0
        for block_index, block in enumerate(blocks):
            if isinstance(block, Paragraph):
                structural_section_boundary = (
                    location.story == "body" and "section_index" in block.properties and not block.runs and not block.text()
                )
                if not structural_section_boundary:
                    self._compile_paragraph(
                        block,
                        events,
                        SemanticLocation(
                            story=location.story, section_index=location.section_index, block_index=block_index,
                            paragraph_index=paragraph_index, table_index=location.table_index,
                            row_index=location.row_index, cell_index=location.cell_index,
                        ),
                    )
                    paragraph_index += 1
                if location.story == "body" and "section_index" in block.properties:
                    current = int(block.properties["section_index"])
                    if current + 1 < len(self._model.sections):
                        events.append(ReconstructionEvent("EndSection", self._model.sections[current].element_id, {"section_index": current}))
                        self._begin_section(current + 1, events)
            elif isinstance(block, Table):
                self._compile_table(
                    block,
                    events,
                    SemanticLocation(
                        story=location.story, section_index=location.section_index, block_index=block_index,
                        table_index=table_index, row_index=location.row_index, cell_index=location.cell_index,
                    ),
                )
                table_index += 1


    def _begin_section(self, section_index: int, events: list[ReconstructionEvent]) -> None:
        section = self._model.sections[section_index]
        events.append(ReconstructionEvent("BeginSection", section.element_id, {"section_index": section_index, "break_type": section.properties.get("break_type", "nextPage")}))
        events.append(ReconstructionEvent("ApplySectionProperties", section.element_id, section.properties))
        self._compile_section_stories(section_index, events)

    def _compile_section_stories(self, section_index: int, events: list[ReconstructionEvent]) -> None:
        section = self._model.sections[section_index]
        for kind, refs, story_map, begin_type, end_type in (
            ("header", section.properties.get("header_refs") or [], self._model.headers, "BeginHeader", "EndHeader"),
            ("footer", section.properties.get("footer_refs") or [], self._model.footers, "BeginFooter", "EndFooter"),
        ):
            for ref in refs:
                rel = self._model.relationships.get(f"word/document.xml:{ref.get('rel_id')}")
                if rel is None or rel.target not in story_map:
                    continue
                story_type = ref.get("type", "default")
                story_id = f"{kind}:{section_index}:{story_type}"
                events.append(ReconstructionEvent(begin_type, story_id, {"story": kind, "story_type": story_type, "section_index": section_index}))
                self._compile_blocks(
                    story_map[rel.target], events,
                    SemanticLocation(story=f"{kind}:{story_type}", section_index=section_index),
                )
                events.append(ReconstructionEvent(end_type, story_id, {"story": kind, "story_type": story_type, "section_index": section_index}))

    def _compile_table(self, table: Table, events: list[ReconstructionEvent], location: SemanticLocation) -> None:
        plan = build_table_plan(table)
        events.append(ReconstructionEvent("BeginTable", table.element_id, {"rows": plan.row_count, "columns": plan.column_count}))
        events.append(ReconstructionEvent("SetTableProperties", table.element_id, plan.properties))
        for index, width in enumerate(plan.properties.get("grid_column_widths", ()) or (), start=1):
            events.append(ReconstructionEvent("SetColumnWidth", table.element_id, {"column": index, "width_twips": width}))
        for row_index, row_props in enumerate(plan.row_properties, start=1):
            if row_props:
                events.append(ReconstructionEvent("SetRowProperties", table.element_id, {"row": row_index, "properties": row_props}))
        for cp in plan.cells:
            if not cp.continuation:
                events.append(ReconstructionEvent("SetCellProperties", cp.cell_id, {"row": cp.row, "column": cp.column, "properties": cp.properties}))
        for merge in plan.merges:
            events.append(ReconstructionEvent("MergeCells", table.element_id, {
                "start_row": merge.start_row, "start_column": merge.start_column,
                "end_row": merge.end_row, "end_column": merge.end_column,
            }))
        for cp in plan.cells:
            if cp.continuation:
                continue
            events.append(ReconstructionEvent("EnterCell", cp.cell_id, {"row": cp.row, "column": cp.column}))
            self._compile_blocks(
                list(cp.blocks), events,
                SemanticLocation(
                    story=location.story, section_index=location.section_index, block_index=location.block_index,
                    table_index=location.table_index, row_index=cp.row, cell_index=cp.column,
                ),
            )
            events.append(ReconstructionEvent("LeaveCell", cp.cell_id, {"row": cp.row, "column": cp.column}))
        events.append(ReconstructionEvent("EndTable", table.element_id, {}))

    def _resolve_theme_fonts(self, props: dict[str, Any]) -> dict[str, Any]:
        resolved = dict(props)
        scheme = self._model.extras.get("theme_font_scheme", {}) or {}
        for font_key, theme_key in (
            ("font_ascii", "font_ascii_theme"),
            ("font_hansi", "font_hansi_theme"),
            ("font_east_asia", "font_east_asia_theme"),
            ("font_cs", "font_cs_theme"),
        ):
            if resolved.get(font_key):
                continue
            theme_ref = resolved.get(theme_key)
            font_name = scheme.get(str(theme_ref)) if theme_ref else None
            if font_name:
                resolved[font_key] = font_name
        return resolved

    def _resolved_style_definition(self, style_id: str | None) -> dict[str, Any] | None:
        if not style_id:
            return None
        definitions = self._model.extras.get("style_definitions", {})
        definition = definitions.get(style_id)
        if not definition:
            return None
        defaults = self._model.extras.get("document_defaults", {})
        run_props: dict[str, Any] = dict(defaults.get("run_properties", {}) or {})
        paragraph_props: dict[str, Any] = {}
        paragraph_props.update(defaults.get("paragraph_properties", {}) or {})
        chain: list[dict[str, Any]] = []
        seen: set[str] = set()
        current = definition
        while current and str(current.get("style_id")) not in seen:
            current_id = str(current.get("style_id"))
            seen.add(current_id)
            chain.append(current)
            based_on = current.get("based_on")
            current = definitions.get(str(based_on)) if based_on else None
        for item in reversed(chain):
            run_props.update(item.get("run_properties", {}) or {})
            paragraph_props.update(item.get("paragraph_properties", {}) or {})
        run_props = self._resolve_theme_fonts(run_props)
        resolved = copy.deepcopy(definition)
        resolved["run_properties"] = run_props
        resolved["paragraph_properties"] = paragraph_props
        return resolved

    def _effective_paragraph_style_id(self, paragraph: Paragraph) -> str | None:
        return paragraph.style_id or self._model.extras.get("default_paragraph_style_id")

    def _effective_paragraph_properties(self, paragraph: Paragraph) -> tuple[dict[str, Any], dict[str, Any] | None]:
        defaults = self._model.extras.get("document_defaults")
        style_id = self._effective_paragraph_style_id(paragraph)
        if not defaults:
            props = copy.deepcopy(paragraph.properties)
            props.pop("inline_markers", None)
            definition = self._model.extras.get("style_definitions", {}).get(style_id) if style_id else None
            return props, copy.deepcopy(definition) if definition else None
        definition = self._resolved_style_definition(style_id)
        props: dict[str, Any] = {}
        props.update(defaults.get("paragraph_properties", {}) or {})
        if definition:
            props.update(definition.get("paragraph_properties", {}) or {})
        direct = copy.deepcopy(paragraph.properties)
        direct.pop("inline_markers", None)
        props.update(direct)
        return props, definition

    def _effective_run_properties(self, paragraph: Paragraph, run: Run, style_definition: dict[str, Any] | None) -> dict[str, Any]:
        defaults = self._model.extras.get("document_defaults")
        direct = {key: copy.deepcopy(value) for key, value in run.properties.items() if key not in {"content_tokens", "break_types"}}
        if not defaults:
            if run.hidden and "hidden" not in direct:
                direct["hidden"] = True
            return direct
        props: dict[str, Any] = {
            "bold": False, "italic": False, "underline": False, "strike": False, "hidden": False,
            "vert_align": "baseline", "character_spacing": "0", "character_position": "0",
        }
        props.update(defaults.get("run_properties", {}) or {})
        if style_definition:
            props.update(style_definition.get("run_properties", {}) or {})
        props.update(direct)
        props = self._resolve_theme_fonts(props)
        if run.hidden:
            props["hidden"] = True
        return props

    def _compile_paragraph(
        self,
        paragraph: Paragraph,
        events: list[ReconstructionEvent],
        location: SemanticLocation,
    ) -> None:
        events.append(ReconstructionEvent("BeginParagraph", paragraph.element_id, {}))
        p_props, style_definition = self._effective_paragraph_properties(paragraph)
        effective_style_id = self._effective_paragraph_style_id(paragraph)
        if effective_style_id is not None:
            p_props = {"style_id": effective_style_id, **p_props}
            if style_definition:
                p_props["style_definition"] = copy.deepcopy(style_definition)
        events.append(ReconstructionEvent("ApplyParagraphProperties", paragraph.element_id, p_props))
        if paragraph.properties.get("numId") is not None or (paragraph.style_id and "List" in paragraph.style_id):
            style_id = paragraph.style_id or ""
            kind = "bullet" if "Bullet" in style_id else "number" if "Number" in style_id else "numbering"
            events.append(ReconstructionEvent("CreateListBinding", paragraph.element_id, {
                "kind": kind, "style_id": paragraph.style_id, "numId": paragraph.properties.get("numId"), "ilvl": paragraph.properties.get("ilvl"),
            }))

        markers_by_index: dict[int, list[dict[str, Any]]] = {}
        for marker in paragraph.properties.get("inline_markers", ()) or ():
            markers_by_index.setdefault(int(marker.get("run_index", 0)), []).append(marker)

        field_state: dict[str, Any] | None = None
        for run_index, run in enumerate(paragraph.runs):
            self._emit_inline_markers(markers_by_index.get(run_index, ()), events, paragraph.element_id)
            run_props = self._effective_run_properties(paragraph, run, style_definition)
            events.append(ReconstructionEvent("ApplyRunProperties", run.element_id, run_props))
            tokens = run.properties.get("content_tokens")
            if not tokens:
                self._compile_legacy_run_content(run, events)
                continue
            for token in tokens:
                kind = token.get("kind")
                if kind == "field_begin":
                    field_state = {"instruction": [], "result": [], "source_element_id": run.element_id, "separated": False}
                    continue
                if kind == "field_instruction" and field_state is not None:
                    field_state["instruction"].append(str(token.get("value", "")))
                    continue
                if kind == "field_separate" and field_state is not None:
                    field_state["separated"] = True
                    continue
                if kind == "field_end" and field_state is not None:
                    instruction = "".join(field_state["instruction"]).strip()
                    cached_result = "".join(field_state["result"])
                    events.append(ReconstructionEvent(
                        "CreateField", field_state["source_element_id"],
                        {"instruction": instruction, "cached_result": cached_result},
                    ))
                    field_state = None
                    continue
                if field_state is not None:
                    # Keep the semantic field while preserving its source cached result.
                    # The cached text is not replayed as ordinary InsertCharacter events.
                    if field_state.get("separated"):
                        if kind == "text": field_state["result"].append(str(token.get("value", "")))
                        elif kind == "tab": field_state["result"].append("\t")
                        elif kind in {"line_break", "page_break"}: field_state["result"].append("\n")
                    continue
                self._compile_token(token, run, events, location)
        self._emit_inline_markers(markers_by_index.get(len(paragraph.runs), ()), events, paragraph.element_id)
        events.append(ReconstructionEvent("EndParagraph", paragraph.element_id, {}))

    def _emit_inline_markers(self, markers, events: list[ReconstructionEvent], paragraph_id: str) -> None:
        for marker in markers:
            bookmark_id = str(marker.get("bookmark_id", ""))
            if marker.get("kind") == "bookmark_start":
                events.append(ReconstructionEvent("BookmarkStart", paragraph_id, {"bookmark_id": bookmark_id, "name": marker.get("name")}))
            elif marker.get("kind") == "bookmark_end":
                name = self._bookmark_names.get(bookmark_id)
                if name:
                    events.append(ReconstructionEvent("CreateBookmark", paragraph_id, {"bookmark_id": bookmark_id, "name": name}))

    def _compile_legacy_run_content(self, run: Run, events: list[ReconstructionEvent]) -> None:
        break_types = iter(run.properties.get("break_types", []))
        text: list[str] = []

        def flush_text() -> None:
            if text:
                events.append(ReconstructionEvent("InsertText", run.element_id, {"text": "".join(text)}))
                text.clear()

        for character in run.text:
            if character == "\t":
                flush_text()
                events.append(ReconstructionEvent("InsertTab", run.element_id, {}))
            elif character == "\n":
                flush_text()
                kind = next(break_types, "line")
                events.append(ReconstructionEvent("InsertPageBreak" if kind == "page" else "InsertLineBreak", run.element_id, {}))
            else:
                text.append(character)
        flush_text()

    def _compile_token(self, token, run: Run, events: list[ReconstructionEvent], location: SemanticLocation) -> None:
        kind = token.get("kind")
        if kind == "text":
            text = str(token.get("value", ""))
            if text:
                events.append(ReconstructionEvent("InsertText", run.element_id, {"text": text}))
        elif kind == "tab": events.append(ReconstructionEvent("InsertTab", run.element_id, {}))
        elif kind == "line_break": events.append(ReconstructionEvent("InsertLineBreak", run.element_id, {}))
        elif kind == "page_break": events.append(ReconstructionEvent("InsertPageBreak", run.element_id, {}))
        elif kind == "drawing":
            relationship_id = str(token.get("relationship_id", "")); queue = self._drawing_queues.get(relationship_id, [])
            if queue: self._compile_drawing(self._drawing_by_id[queue.pop(0)], events)
        elif kind == "footnote_ref": self._compile_note("footnote", str(token.get("note_id")), run.element_id, events, location)
        elif kind == "endnote_ref": self._compile_note("endnote", str(token.get("note_id")), run.element_id, events, location)

    def _compile_note(self, note_kind: str, note_id: str, source_id: str, events: list[ReconstructionEvent], location: SemanticLocation) -> None:
        note_map = self._model.footnotes if note_kind == "footnote" else self._model.endnotes
        blocks = note_map.get(note_id)
        if blocks is None: return
        create = "CreateFootnote" if note_kind == "footnote" else "CreateEndnote"
        begin = "BeginFootnoteStory" if note_kind == "footnote" else "BeginEndnoteStory"
        end = "EndFootnoteStory" if note_kind == "footnote" else "EndEndnoteStory"
        events.append(ReconstructionEvent(create, source_id, {"note_id": note_id}))
        events.append(ReconstructionEvent(begin, source_id, {"note_id": note_id, "story": note_kind}))
        self._compile_blocks(list(blocks), events, SemanticLocation(story=note_kind, section_index=location.section_index))
        events.append(ReconstructionEvent(end, source_id, {"note_id": note_id, "story": note_kind}))

    def _compile_drawing(self, drawing, events: list[ReconstructionEvent]) -> None:
        events.append(ReconstructionEvent("InsertImage", drawing.element_id, {
            "asset_id": drawing.asset_id, "source_path": drawing.source_path, "representation": drawing.representation,
        }))
        events.append(ReconstructionEvent("SetImageSize", drawing.element_id, {
            "width_emu": drawing.width_emu, "height_emu": drawing.height_emu, "lock_aspect_ratio": drawing.lock_aspect_ratio,
        }))
        events.append(ReconstructionEvent("SetImageWrap", drawing.element_id, {
            "wrap_type": drawing.wrap_type, "behind_text": drawing.behind_text,
            "distance_top_emu": drawing.distance_top_emu, "distance_bottom_emu": drawing.distance_bottom_emu,
            "distance_left_emu": drawing.distance_left_emu, "distance_right_emu": drawing.distance_right_emu,
        }))
        events.append(ReconstructionEvent("SetImagePosition", drawing.element_id, {
            "horizontal_relative_from": drawing.horizontal_relative_from, "horizontal_position_emu": drawing.horizontal_position_emu,
            "vertical_relative_from": drawing.vertical_relative_from, "vertical_position_emu": drawing.vertical_position_emu,
        }))
        events.append(ReconstructionEvent("SetImageCrop", drawing.element_id, {"crop": drawing.crop}))
        events.append(ReconstructionEvent("SetImageRotation", drawing.element_id, {"rotation_degrees": drawing.rotation_degrees}))
        events.append(ReconstructionEvent("SetImageZOrder", drawing.element_id, {"z_order": drawing.z_order}))
