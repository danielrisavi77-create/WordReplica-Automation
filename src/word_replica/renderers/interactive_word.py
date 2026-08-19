from __future__ import annotations

from contextlib import suppress
from pathlib import Path
import json
import re
import time
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from word_replica.domain.reconstruction import ExecutionOutcome, ReconstructionBlueprint, ReconstructionEvent, WordStateSnapshot
from word_replica.interactive.speed import SpeedController
from word_replica.config import InteractiveOptions
from word_replica.interactive.verification import state_snapshots_match
from word_replica.renderers.word_ownership import clear_owned_word, record_owned_word, word_process_pids


WD_COLLAPSE_END = 0
WD_CHARACTER = 1
WD_LINE_BREAK = 6
WD_PAGE_BREAK = 7
WD_FORMAT_DOCX = 16

RPC_E_CALL_REJECTED = -2147418111

# Empirically the lowest width Word's COM layer (Columns(i).Width, Cell.Width,
# and Cell/Columns.SetWidth all tried) accepts without raising "Value out of
# range" - confirmed 10pt fails and 12pt succeeds regardless of cell padding.
# A source document can still legitimately declare a narrower column (the
# OOXML format has no such floor); see _event_SetColumnWidth.
_WORD_MIN_COM_COLUMN_WIDTH_POINTS = 12.0

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_VALID_BOOKMARK_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,39}$")
_BOOKMARK_NAME_MAX_LENGTH = 40


def sanitize_word_bookmark_name(name: str, *, taken: set[str]) -> str:
    """Return a name Word's COM Bookmarks.Add will accept.

    Word's automation API rejects any bookmark name containing characters
    other than ASCII letters, digits, and underscores, or not starting with
    a letter/underscore - a stricter rule than the OOXML file format itself
    enforces. A source document can legitimately contain bookmarks that
    violate this (e.g. slug-style names from a non-Word export tool), which
    otherwise crashes interactive reconstruction with "Bad bookmark name".
    Callers must restore the original name in the saved OOXML afterward
    (see `_restore_bookmark_names`) so fidelity to the source is preserved.
    """
    if _VALID_BOOKMARK_NAME.match(name):
        taken.add(name)
        return name
    sanitized = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if not sanitized or not re.match(r"[A-Za-z_]", sanitized[0]):
        sanitized = f"_{sanitized}"
    sanitized = sanitized[:_BOOKMARK_NAME_MAX_LENGTH] or "_bookmark"
    candidate = sanitized
    suffix = 1
    while candidate in taken:
        suffix += 1
        tail = f"_{suffix}"
        candidate = f"{sanitized[: _BOOKMARK_NAME_MAX_LENGTH - len(tail)]}{tail}"
    taken.add(candidate)
    return candidate


def _replace_with_retry(source: Path, destination: Path, *, attempts: int = 60, delay_seconds: float = 0.5) -> None:
    """Word can still hold a brief lock on a file immediately after SaveAs2
    returns, even though the COM call itself already completed. Retry a
    same-machine rename briefly before giving up so a real permission
    problem still surfaces as an error."""
    last: OSError | None = None
    for attempt in range(attempts):
        try:
            source.replace(destination)
            return
        except OSError as exc:
            last = exc
            if attempt + 1 >= attempts:
                raise
            time.sleep(delay_seconds)
    if last is not None:
        raise last


def _restore_bookmark_names(docx_path: str | Path, rewrites: dict[str, str]) -> None:
    """Rewrite `w:bookmarkStart/@w:name` in a saved .docx from the COM-safe
    names `sanitize_word_bookmark_name` assigned back to the original
    source names, across every XML part that may carry bookmarks."""
    if not rewrites:
        return
    docx_path = Path(docx_path)
    with ZipFile(docx_path, "r") as archive:
        infos = archive.infolist()
        members = {info.filename: archive.read(info.filename) for info in infos}
    changed = False
    for filename, data in list(members.items()):
        if not filename.startswith("word/") or not filename.endswith(".xml") or b"bookmarkStart" not in data:
            continue
        root = etree.fromstring(data)
        touched = False
        for element in root.iter(f"{{{_W_NS}}}bookmarkStart"):
            current = element.get(f"{{{_W_NS}}}name")
            if current in rewrites:
                element.set(f"{{{_W_NS}}}name", rewrites[current])
                touched = True
        if touched:
            members[filename] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            changed = True
    if not changed:
        return
    temp = docx_path.with_suffix(docx_path.suffix + ".tmp")
    try:
        with ZipFile(temp, "w", ZIP_DEFLATED) as archive:
            for info in infos:
                archive.writestr(info, members[info.filename])
        _replace_with_retry(temp, docx_path)
    finally:
        temp.unlink(missing_ok=True)


def _restore_narrow_column_widths(
    docx_path: str | Path, fixups: list[tuple[int, int, int]]
) -> None:
    """Correct table columns that were set to a COM-safe placeholder width
    (see `_WORD_MIN_COM_COLUMN_WIDTH_POINTS`) back to their true, narrower
    source width directly in the saved OOXML's `w:tblGrid`/`w:tcW`, where
    there is no such minimum. `fixups` entries are
    `(table_sequence, column, width_twips)`, where `table_sequence` is the
    0-based order tables were inserted in - matched against `w:tbl` elements
    in `word/document.xml` in document order, so this only covers tables in
    the main body, not headers/footers/footnotes.
    """
    if not fixups:
        return
    docx_path = Path(docx_path)
    with ZipFile(docx_path, "r") as archive:
        infos = archive.infolist()
        members = {info.filename: archive.read(info.filename) for info in infos}
    filename = "word/document.xml"
    if filename not in members:
        return
    root = etree.fromstring(members[filename])
    tables = list(root.iter(f"{{{_W_NS}}}tbl"))
    changed = False
    for sequence, column, width_twips in fixups:
        if sequence >= len(tables):
            continue
        table_element = tables[sequence]
        grid = table_element.find(f"{{{_W_NS}}}tblGrid")
        if grid is not None:
            grid_columns = grid.findall(f"{{{_W_NS}}}gridCol")
            if column - 1 < len(grid_columns):
                grid_columns[column - 1].set(f"{{{_W_NS}}}w", str(width_twips))
                changed = True
        for row in table_element.findall(f"{{{_W_NS}}}tr"):
            row_cells = row.findall(f"{{{_W_NS}}}tc")
            if column - 1 >= len(row_cells):
                continue
            cell_properties = row_cells[column - 1].find(f"{{{_W_NS}}}tcPr")
            cell_width = cell_properties.find(f"{{{_W_NS}}}tcW") if cell_properties is not None else None
            if cell_width is not None:
                cell_width.set(f"{{{_W_NS}}}w", str(width_twips))
                changed = True
    if not changed:
        return
    members[filename] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    temp = docx_path.with_suffix(docx_path.suffix + ".tmp")
    try:
        with ZipFile(temp, "w", ZIP_DEFLATED) as archive:
            for info in infos:
                archive.writestr(info, members[info.filename])
        _replace_with_retry(temp, docx_path)
    finally:
        temp.unlink(missing_ok=True)


def _is_rejected_com_call(exc: Exception) -> bool:
    hresult = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        first = exc.args[0]
        if isinstance(first, int):
            hresult = first
    return hresult == RPC_E_CALL_REJECTED


def _retry_rejected_com_call(operation, *, attempts: int = 600, delay_seconds: float = 0.1):
    last = None
    for attempt in range(attempts):
        try:
            return operation()
        except Exception as exc:
            if not _is_rejected_com_call(exc) or attempt + 1 >= attempts:
                raise
            last = exc
            try:
                import pythoncom
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(delay_seconds)
    if last is not None:
        raise last


def _retry_getattr(target: Any, name: str, default: Any = None) -> Any:
    return _retry_rejected_com_call(lambda: getattr(target, name, default))


def _retry_setattr(target: Any, name: str, value: Any) -> None:
    _retry_rejected_com_call(lambda: setattr(target, name, value))


class InteractiveWordController:
    def __init__(self, *, visible: bool = True) -> None:
        self.visible = bool(visible)
        self.application: Any | None = None
        self.document: Any | None = None
        self.active_range: Any | None = None
        self._owns_com = False
        self._paragraph_started = False
        self._page_break_continuation_pending = False
        self._post_table_paragraph_active = False
        self._table_stack: list[dict[str, Any]] = []
        self._table_insertion_sequence = 0
        self._narrow_column_fixups: list[tuple[int, int, int]] = []
        self._asset_resolver = None
        self._active_image: Any | None = None
        self._active_image_representation: str | None = None
        self._section_started = False
        self._active_section: Any | None = None
        self._story_stack: list[tuple[Any, bool]] = []
        self._bookmark_starts: dict[str, int] = {}
        self._bookmark_com_names: set[str] = set()
        self._bookmark_name_rewrites: dict[str, str] = {}
        self._last_saved_path: Path | None = None
        self._pending_note: Any | None = None
        self._active_note_index: int | None = None
        self._active_story = "body"
        self._active_section_index: int | None = None
        self._active_table_element_id: str | None = None
        self._active_cell_element_id: str | None = None
        self._image_transaction_active = False
        self._floating_images: list[tuple[int, Any]] = []
        self._owned_word_pid: int | None = None
        self._formatting_context_generation = 0
        self._last_run_properties_key: tuple[int, str, int] | None = None
        self._active_run_properties: dict[str, Any] | None = None
        self._reset_inherited_run_color_after_insert = False
        self._paragraph_format_context_generation = 0
        self._last_paragraph_properties_key: tuple[int, str, int] | None = None
        self._paragraph_style_cache: dict[tuple[str, str], Any] = {}
        self._table_batch_metrics: dict[str, object] | None = None

    @classmethod
    def for_testing(cls, *, active_range: Any) -> "InteractiveWordController":
        controller = cls()
        controller.active_range = active_range
        return controller

    def open_blank(self) -> None:
        import pythoncom
        import win32com.client

        existing_word_pids = word_process_pids()
        pythoncom.CoInitialize()
        self._owns_com = True
        self.application = win32com.client.DispatchEx("Word.Application")
        self._owned_word_pid = record_owned_word(
            self.application, role="interactive", existing_word_pids=existing_word_pids
        )
        _retry_setattr(self.application, "DisplayAlerts", 0)
        if not self.visible:
            _retry_setattr(self.application, "Visible", False)
        documents = _retry_getattr(self.application, "Documents")
        self.document = _retry_rejected_com_call(lambda: documents.Add())
        self.active_range = _retry_rejected_com_call(lambda: self.document.Range(0, 0))
        if self.visible:
            _retry_setattr(self.application, "Visible", True)
            activate_document = _retry_getattr(self.document, "Activate", None)
            if callable(activate_document):
                _retry_rejected_com_call(activate_document)
            activate_application = _retry_getattr(self.application, "Activate", None)
            if callable(activate_application):
                _retry_rejected_com_call(activate_application)
        self._paragraph_started = False

    def open_existing(self, path: str | Path) -> None:
        import pythoncom
        import win32com.client

        existing_word_pids = word_process_pids()
        pythoncom.CoInitialize()
        self._owns_com = True
        self.application = win32com.client.DispatchEx("Word.Application")
        self._owned_word_pid = record_owned_word(
            self.application, role="interactive", existing_word_pids=existing_word_pids
        )
        self.application.Visible = self.visible
        self.document = self.application.Documents.Open(
            str(Path(path).resolve()), ReadOnly=False, AddToRecentFiles=False
        )
        end = max(int(self.document.Content.Start), int(self.document.Content.End) - 1)
        self.active_range = self.document.Range(end, end)
        self._paragraph_started = True


    def _resume_range_for_story(self, story: str, start: int, end: int, *, section_index: int | None = None, note_index: int | None = None):
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        if story == "body":
            return self.document.Range(int(start), int(end))
        if story.startswith("header:") or story.startswith("footer:"):
            section_i = int(section_index or 0) + 1
            section = self.document.Sections(section_i)
            story_type = story.split(":", 1)[1]
            story_i = {"default": 1, "first": 2, "even": 3}.get(story_type, 1)
            collection = section.Headers if story.startswith("header:") else section.Footers
            target = self._duplicate_range(collection(story_i).Range)
            if hasattr(target, "SetRange"):
                target.SetRange(int(start), int(end))
            else:
                target.Start = int(start); target.End = int(end)
            return target
        if story in {"footnote", "endnote"}:
            if note_index is None:
                raise RuntimeError(f"checkpoint is missing {story} index")
            collection = self.document.Footnotes if story == "footnote" else self.document.Endnotes
            note = collection(int(note_index))
            target = self._duplicate_range(note.Range)
            if hasattr(target, "SetRange"):
                target.SetRange(int(start), int(end))
            else:
                target.Start = int(start); target.End = int(end)
            return target
        raise RuntimeError(f"resume story is not rehydratable: {story}")

    def restore_checkpoint_state(self, checkpoint) -> None:
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        start = getattr(checkpoint, "range_start", None)
        end = getattr(checkpoint, "range_end", None)
        if start is None or end is None:
            raise RuntimeError("checkpoint is missing Word range state")
        state = dict(getattr(checkpoint, "resume_state", {}) or {})
        section_index = state.get("section_index")
        if section_index is None:
            section_index = 0
        self._active_section_index = int(section_index)
        self._section_started = bool(state.get("section_started", False))
        if self._section_started:
            self._active_section = self.document.Sections(self._active_section_index + 1)
        story = str(getattr(checkpoint, "story", "body"))
        note_index = state.get("note_index")
        active_table_element_id = getattr(checkpoint, "table_element_id", None)
        if story == "body" and active_table_element_id is None:
            content = _retry_getattr(self.document, "Content")
            content_start = int(_retry_getattr(content, "Start"))
            content_end = max(content_start, int(_retry_getattr(content, "End")) - 1)
            self.active_range = _retry_rejected_com_call(
                lambda: self.document.Range(content_end, content_end)
            )
        else:
            self.active_range = self._resume_range_for_story(
                story, int(start), int(end), section_index=self._active_section_index, note_index=note_index
            )
        self._paragraph_started = bool(getattr(checkpoint, "paragraph_started", False))
        self._active_story = story
        self._active_note_index = int(note_index) if note_index is not None else None
        self._active_table_element_id = active_table_element_id
        self._active_cell_element_id = getattr(checkpoint, "cell_element_id", None)
        self._story_stack.clear()
        return_start = state.get("return_range_start")
        return_end = state.get("return_range_end")
        if story != "body" and return_start is not None and return_end is not None:
            return_story = str(state.get("return_story", "body"))
            return_range = self._resume_range_for_story(
                return_story, int(return_start), int(return_end),
                section_index=state.get("return_section_index", self._active_section_index),
                note_index=state.get("return_note_index"),
            )
            self._story_stack.append((return_range, bool(state.get("return_paragraph_started", True)), return_story))
        self._table_stack.clear()
        if self._active_table_element_id:
            tables = getattr(self.active_range, "Tables", None)
            if tables is None or int(getattr(tables, "Count", 0)) < 1:
                raise RuntimeError("checkpoint table context cannot be rehydrated from the saved Word range")
            table = tables(1)
            after_range = self._duplicate_range(table.Range)
            with suppress(Exception): after_range.Collapse(WD_COLLAPSE_END)
            self._table_stack.append({
                "table": table, "cells": {}, "parent_range": self.active_range,
                "after_range": after_range, "element_id": self._active_table_element_id,
                "structure_complete": bool(state.get("table_structure_complete", True)),
            })

    def resume_state_snapshot(self) -> dict[str, Any]:
        snapshot = self.current_state_snapshot()
        result = {
            "story": snapshot.story, "range_start": snapshot.range_start, "range_end": snapshot.range_end,
            "section_index": self._active_section_index, "section_started": self._section_started,
            "paragraph_started": self._paragraph_started, "note_index": self._active_note_index,
            "table_depth": len(self._table_stack),
            "table_structure_complete": bool(self._table_stack[-1].get("structure_complete", False)) if self._table_stack else True,
        }
        if self._story_stack:
            return_range, return_paragraph_started, return_story = self._story_stack[-1]
            result.update({
                "return_story": return_story,
                "return_range_start": getattr(return_range, "Start", None),
                "return_range_end": getattr(return_range, "End", None),
                "return_paragraph_started": bool(return_paragraph_started),
                "return_section_index": self._active_section_index,
            })
        return result

    def set_asset_resolver(self, resolver) -> None:
        self._asset_resolver = resolver

    def consume_table_batch_metrics(self) -> dict[str, object] | None:
        metrics = self._table_batch_metrics
        self._table_batch_metrics = None
        return metrics

    @staticmethod
    def _emu_to_points(value: Any) -> float:
        return float(value) / 12700.0

    def _require_range(self) -> Any:
        if self.active_range is None:
            raise RuntimeError("interactive Word document is not open")
        return self.active_range

    def _collapse_end(self) -> None:
        target = self._require_range()
        collapse = _retry_getattr(target, "Collapse", None)
        if callable(collapse):
            _retry_rejected_com_call(lambda: collapse(WD_COLLAPSE_END))
            return
        end = _retry_getattr(target, "End", None)
        if end is None:
            raise AttributeError("Word range exposes neither Collapse nor End")
        _retry_setattr(target, "Start", end)

    def execute_event(self, event: ReconstructionEvent) -> None:
        handler = getattr(self, f"_event_{event.event_type}", None)
        if handler is None:
            # Structural/property events are implemented progressively by later tasks.
            return
        text_events = {"InsertCharacter", "InsertText", "InsertTab", "InsertLineBreak"}
        if event.event_type not in {"ApplyRunProperties", *text_events}:
            self._formatting_context_generation += 1
            self._last_run_properties_key = None
        if event.event_type not in {"ApplyRunProperties", "ApplyParagraphProperties", "BeginParagraph", "EndParagraph", *text_events}:
            self._paragraph_format_context_generation += 1
            self._last_paragraph_properties_key = None
        handler(event)

    @staticmethod
    def _word_bool(value: Any) -> int:
        return -1 if bool(value) else 0

    @staticmethod
    def _twips_to_points(value: Any) -> float:
        return float(value) / 20.0

    @staticmethod
    def _half_points_to_points(value: Any) -> float:
        return float(value) / 2.0

    @staticmethod
    def _word_color(value: str) -> int:
        raw = value.strip().lstrip("#")
        if len(raw) != 6:
            raise ValueError(f"invalid RGB color: {value}")
        red, green, blue = int(raw[0:2], 16), int(raw[2:4], 16), int(raw[4:6], 16)
        return red | (green << 8) | (blue << 16)

    @staticmethod
    def _border_index(name: str) -> int | None:
        return {"top": -1, "left": -2, "bottom": -3, "right": -4, "insideH": -5, "insideV": -6}.get(name)

    @staticmethod
    def _line_style(value: str | None) -> int:
        return {
            "nil": 0, "none": 0, "single": 1, "dotted": 2, "dashed": 3,
            "dashSmallGap": 3, "dashDotStroked": 5, "dotDash": 5, "dotDotDash": 6,
            "double": 7, "triple": 8,
        }.get(str(value or "single"), 1)

    def _apply_borders(self, target, borders) -> None:
        if not borders or not hasattr(target, "Borders"):
            return
        for name, props in borders.items():
            index = self._border_index(str(name))
            if index is None:
                continue
            try:
                border = target.Borders(index)
                border.LineStyle = self._line_style(props.get("val"))
                if props.get("sz") is not None:
                    border.LineWidth = int(props["sz"])
                color = props.get("color")
                if color and str(color).upper() not in {"AUTO", "NONE"}:
                    border.Color = self._word_color(str(color))
            except Exception:
                continue

    def _apply_padding(self, target, margins) -> None:
        if not margins:
            return
        for key, attr in (("top", "TopPadding"), ("bottom", "BottomPadding"), ("left", "LeftPadding"), ("right", "RightPadding")):
            entry = margins.get(key)
            if entry and entry.get("w") is not None:
                with suppress(Exception):
                    setattr(target, attr, self._twips_to_points(entry["w"]))

    def _apply_shading(self, target, fill) -> None:
        if fill and str(fill).upper() not in {"AUTO", "NONE"} and hasattr(target, "Shading"):
            with suppress(Exception):
                target.Shading.BackgroundPatternColor = self._word_color(str(fill))

    def _apply_paragraph_spacing(self, paragraph, props: dict[str, Any], *, defaults: bool = False) -> None:
        if defaults or props.get("spacing_before") is not None:
            _retry_setattr(paragraph, "SpaceBefore", self._twips_to_points(props.get("spacing_before", "0")))
        if defaults or props.get("spacing_after") is not None:
            _retry_setattr(paragraph, "SpaceAfter", self._twips_to_points(props.get("spacing_after", "0")))
        line = props.get("spacing_line")
        rule = props.get("spacing_line_rule")
        if line is None and (defaults or rule == "single"):
            _retry_setattr(paragraph, "LineSpacingRule", 0)
            return
        if line is None:
            return
        if rule in {None, "auto"}:
            _retry_setattr(paragraph, "LineSpacingRule", 5)
            multiple = float(line) / 240.0
            points = multiple * 12.0
            _retry_setattr(paragraph, "LineSpacing", float(points))
        elif rule == "exact":
            _retry_setattr(paragraph, "LineSpacingRule", 4)
            _retry_setattr(paragraph, "LineSpacing", self._twips_to_points(line))
        elif rule == "atLeast":
            _retry_setattr(paragraph, "LineSpacingRule", 3)
            _retry_setattr(paragraph, "LineSpacing", self._twips_to_points(line))
        elif rule == "single":
            _retry_setattr(paragraph, "LineSpacingRule", 0)

    def _event_ApplyDocumentDefaults(self, event: ReconstructionEvent) -> None:
        if self.document is None:
            return
        defaults = event.payload
        styles = _retry_getattr(self.document, "Styles", None)
        if styles is None:
            return
        try:
            normal = _retry_rejected_com_call(lambda: styles("Normal"))
        except Exception:
            return
        run_props = dict(defaults.get("run_properties") or {})
        font = _retry_getattr(normal, "Font", None)
        if font is not None:
            font_name = run_props.get("font_ascii") or run_props.get("font_hansi")
            if font_name:
                _retry_setattr(font, "Name", str(font_name))
            if run_props.get("font_east_asia"):
                with suppress(Exception): _retry_setattr(font, "NameFarEast", str(run_props["font_east_asia"]))
            if run_props.get("font_cs"):
                with suppress(Exception): _retry_setattr(font, "NameBi", str(run_props["font_cs"]))
            if run_props.get("size_half_points") is not None:
                _retry_setattr(font, "Size", self._half_points_to_points(run_props["size_half_points"]))
        paragraph = _retry_getattr(normal, "ParagraphFormat", None)
        if paragraph is not None:
            self._apply_paragraph_spacing(paragraph, dict(defaults.get("paragraph_properties") or {}), defaults=True)

    def _event_ApplyRunProperties(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        props = event.payload
        previous_color = (self._active_run_properties or {}).get("color")
        current_color = props.get("color")
        previous_color_is_explicit = bool(
            previous_color and str(previous_color).lower() not in {"auto", "none"}
        )
        current_color_is_explicit = bool(
            current_color and str(current_color).lower() not in {"auto", "none"}
        )
        self._reset_inherited_run_color_after_insert = (
            previous_color_is_explicit and not current_color_is_explicit
        )
        self._active_run_properties = dict(props)
        properties_key = json.dumps(props, sort_keys=True, ensure_ascii=False, default=str)
        cache_key = (id(target), properties_key, self._formatting_context_generation)
        if cache_key == self._last_run_properties_key:
            return
        font = _retry_getattr(target, "Font")
        reset = _retry_getattr(font, "Reset", None)
        if callable(reset):
            _retry_rejected_com_call(reset)
        if "bold" in props:
            _retry_setattr(font, "Bold", self._word_bool(props["bold"]))
        if "italic" in props:
            _retry_setattr(font, "Italic", self._word_bool(props["italic"]))
        if "underline" in props:
            _retry_setattr(font, "Underline", 1 if props["underline"] else 0)
        if "strike" in props:
            _retry_setattr(font, "StrikeThrough", self._word_bool(props["strike"]))
        font_name = props.get("font_ascii") or props.get("font_hansi")
        if font_name:
            _retry_setattr(font, "Name", str(font_name))
        if props.get("font_east_asia"):
            with suppress(Exception):
                _retry_setattr(font, "NameFarEast", str(props["font_east_asia"]))
        if props.get("font_cs"):
            with suppress(Exception):
                _retry_setattr(font, "NameBi", str(props["font_cs"]))
        if props.get("size_half_points") is not None:
            _retry_setattr(font, "Size", self._half_points_to_points(props["size_half_points"]))
        color = props.get("color")
        if color and str(color).lower() not in {"auto", "none"}:
            _retry_setattr(font, "Color", self._word_color(str(color)))
        if "hidden" in props:
            with suppress(Exception):
                font.Hidden = self._word_bool(props["hidden"])
        vert = props.get("vert_align")
        if vert == "superscript":
            _retry_setattr(font, "Superscript", -1)
            _retry_setattr(font, "Subscript", 0)
        elif vert == "subscript":
            _retry_setattr(font, "Subscript", -1)
            _retry_setattr(font, "Superscript", 0)
        elif vert == "baseline":
            _retry_setattr(font, "Subscript", 0)
            _retry_setattr(font, "Superscript", 0)
        if props.get("character_spacing") is not None:
            with suppress(Exception):
                font.Spacing = float(props["character_spacing"]) / 20.0
        if props.get("character_position") is not None:
            with suppress(Exception):
                font.Position = float(props["character_position"]) / 2.0
        highlight_map = {
            "black": 1, "blue": 2, "turquoise": 3, "brightGreen": 4,
            "pink": 5, "red": 6, "yellow": 7, "white": 8,
            "darkBlue": 9, "teal": 10, "green": 11, "violet": 12,
            "darkRed": 13, "darkYellow": 14, "gray50": 15, "gray25": 16,
        }
        if props.get("highlight") in highlight_map:
            with suppress(Exception):
                target.HighlightColorIndex = highlight_map[props["highlight"]]
        language_ids = {"hr-HR": 1050, "en-US": 1033, "en-GB": 2057, "de-DE": 1031}
        if props.get("language") in language_ids:
            with suppress(Exception):
                font.LanguageID = language_ids[props["language"]]
        self._last_run_properties_key = cache_key

    @staticmethod
    def _contains_east_asia_text(text: str) -> bool:
        return any(
            "\u2e80" <= character <= "\u9fff"
            or "\u3040" <= character <= "\u30ff"
            or "\uac00" <= character <= "\ud7af"
            or "\uf900" <= character <= "\ufaff"
            for character in text
        )

    def _apply_post_insert_run_properties(self, props: dict[str, Any], text: str) -> None:
        """Apply only properties Word does not reliably inherit into inserted text."""
        target = self._require_range()
        font = _retry_getattr(target, "Font")
        reset_inherited_color = self._reset_inherited_run_color_after_insert
        self._reset_inherited_run_color_after_insert = False
        if reset_inherited_color:
            reset = _retry_getattr(font, "Reset", None)
            if callable(reset):
                _retry_rejected_com_call(reset)
        for key, attribute in (
            ("bold", "Bold"),
            ("italic", "Italic"),
            ("strike", "StrikeThrough"),
        ):
            if key in props:
                _retry_setattr(font, attribute, self._word_bool(props[key]))
        if "underline" in props:
            _retry_setattr(font, "Underline", 1 if props["underline"] else 0)

        font_name = props.get("font_ascii") or props.get("font_hansi")
        if font_name:
            _retry_setattr(font, "Name", str(font_name))
        if self._contains_east_asia_text(text) and props.get("font_east_asia"):
            with suppress(Exception):
                _retry_setattr(font, "NameFarEast", str(props["font_east_asia"]))
        if props.get("font_cs"):
            with suppress(Exception):
                _retry_setattr(font, "NameBi", str(props["font_cs"]))
        if props.get("size_half_points") is not None:
            _retry_setattr(font, "Size", self._half_points_to_points(props["size_half_points"]))

        color = props.get("color")
        if color and str(color).lower() not in {"auto", "none"}:
            _retry_setattr(font, "Color", self._word_color(str(color)))
        if props.get("hidden"):
            with suppress(Exception):
                _retry_setattr(font, "Hidden", self._word_bool(props["hidden"]))
        vert = props.get("vert_align")
        if vert == "superscript":
            _retry_setattr(font, "Superscript", -1)
            _retry_setattr(font, "Subscript", 0)
        elif vert == "subscript":
            _retry_setattr(font, "Subscript", -1)
            _retry_setattr(font, "Superscript", 0)
        if props.get("character_spacing") not in {None, "0", 0}:
            with suppress(Exception):
                _retry_setattr(font, "Spacing", float(props["character_spacing"]) / 20.0)
        if props.get("character_position") not in {None, "0", 0}:
            with suppress(Exception):
                _retry_setattr(font, "Position", float(props["character_position"]) / 2.0)

        highlight_map = {
            "black": 1, "blue": 2, "turquoise": 3, "brightGreen": 4,
            "pink": 5, "red": 6, "yellow": 7, "white": 8,
            "darkBlue": 9, "teal": 10, "green": 11, "violet": 12,
            "darkRed": 13, "darkYellow": 14, "gray50": 15, "gray25": 16,
        }
        if props.get("highlight") in highlight_map:
            with suppress(Exception):
                target.HighlightColorIndex = highlight_map[props["highlight"]]
        language_ids = {"hr-HR": 1050, "en-US": 1033, "en-GB": 2057, "de-DE": 1031}
        if props.get("language") in language_ids and props["language"] != "en-US":
            with suppress(Exception):
                font.LanguageID = language_ids[props["language"]]
        self._last_run_properties_key = None

    def _ensure_paragraph_style(self, definition: dict[str, Any] | None, style_id: str | None):
        if self.document is None or not style_id:
            return None
        cache_key = (
            str(style_id),
            json.dumps(definition or {}, sort_keys=True, ensure_ascii=False, default=str),
        )
        if cache_key in self._paragraph_style_cache:
            return self._paragraph_style_cache[cache_key]
        styles = _retry_getattr(self.document, "Styles", None)
        if styles is None:
            return None
        name = str((definition or {}).get("name") or style_id)
        style = None
        for key in (style_id, name):
            try:
                style = _retry_rejected_com_call(lambda key=key: styles(key))
                break
            except Exception:
                continue
        if style is None and definition and str(definition.get("type", "paragraph")) == "paragraph":
            try:
                style = _retry_rejected_com_call(lambda: styles.Add(name, 1))
            except Exception:
                return None
        if style is None:
            return None
        run_props = dict((definition or {}).get("run_properties") or {})
        font = _retry_getattr(style, "Font", None)
        if font is not None:
            if "bold" in run_props: _retry_setattr(font, "Bold", self._word_bool(run_props["bold"]))
            if "italic" in run_props: _retry_setattr(font, "Italic", self._word_bool(run_props["italic"]))
            if "underline" in run_props: _retry_setattr(font, "Underline", 1 if run_props["underline"] else 0)
            if "strike" in run_props: _retry_setattr(font, "StrikeThrough", self._word_bool(run_props["strike"]))
            font_name = run_props.get("font_ascii") or run_props.get("font_hansi")
            if font_name: _retry_setattr(font, "Name", str(font_name))
            if run_props.get("font_cs"): _retry_setattr(font, "NameBi", str(run_props["font_cs"]))
            if run_props.get("size_half_points") is not None: _retry_setattr(font, "Size", self._half_points_to_points(run_props["size_half_points"]))
            color = run_props.get("color")
            if color and str(color).lower() not in {"auto", "none"}: _retry_setattr(font, "Color", self._word_color(str(color)))
        paragraph = _retry_getattr(style, "ParagraphFormat", None)
        p_props = dict((definition or {}).get("paragraph_properties") or {})
        if paragraph is not None:
            alignment_map = {"left": 0, "center": 1, "right": 2, "both": 3, "justify": 3, "distribute": 4}
            if p_props.get("alignment") in alignment_map: _retry_setattr(paragraph, "Alignment", alignment_map[p_props["alignment"]])
            if p_props.get("indent_left") is not None: _retry_setattr(paragraph, "LeftIndent", self._twips_to_points(p_props["indent_left"]))
            if p_props.get("indent_right") is not None: _retry_setattr(paragraph, "RightIndent", self._twips_to_points(p_props["indent_right"]))
            if p_props.get("indent_first_line") is not None: _retry_setattr(paragraph, "FirstLineIndent", self._twips_to_points(p_props["indent_first_line"]))
            elif p_props.get("indent_hanging") is not None: _retry_setattr(paragraph, "FirstLineIndent", -self._twips_to_points(p_props["indent_hanging"]))
            self._apply_paragraph_spacing(paragraph, p_props)
            if "keepNext" in p_props: _retry_setattr(paragraph, "KeepWithNext", self._word_bool(p_props["keepNext"]))
            if "keepLines" in p_props: _retry_setattr(paragraph, "KeepTogether", self._word_bool(p_props["keepLines"]))
            if "widowControl" in p_props: _retry_setattr(paragraph, "WidowControl", self._word_bool(p_props["widowControl"]))
            if "pageBreakBefore" in p_props: _retry_setattr(paragraph, "PageBreakBefore", self._word_bool(p_props["pageBreakBefore"]))
        based_on = (definition or {}).get("based_on")
        if based_on:
            with suppress(Exception): style.BaseStyle = styles(str(based_on))
        next_style = (definition or {}).get("next_style")
        if next_style:
            with suppress(Exception): style.NextParagraphStyle = styles(str(next_style))
        self._paragraph_style_cache[cache_key] = style
        return style

    def _event_ApplyParagraphProperties(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        props = event.payload
        try:
            if int(_retry_getattr(target, "Start")) == int(_retry_getattr(target, "End")):
                expanded = self._duplicate_range(target)
                if expanded is not target:
                    start = int(_retry_getattr(expanded, "Start"))
                    end = int(_retry_getattr(expanded, "End"))
                    if end == start:
                        _retry_setattr(expanded, "End", end + 1)
                    target = expanded
        except Exception:
            target = self._require_range()
        properties_key = json.dumps(props, sort_keys=True, ensure_ascii=False, default=str)
        cache_key = (id(target), properties_key, self._paragraph_format_context_generation)
        if cache_key == self._last_paragraph_properties_key:
            return
        if props.get("style_id"):
            style = self._ensure_paragraph_style(props.get("style_definition"), str(props["style_id"]))
            _retry_setattr(target, "Style", style if style is not None else props["style_id"])
        paragraph = _retry_getattr(target, "ParagraphFormat")
        reset = _retry_getattr(paragraph, "Reset", None)
        if callable(reset):
            _retry_rejected_com_call(reset)
        alignment_map = {"left": 0, "center": 1, "right": 2, "both": 3, "justify": 3, "distribute": 4}
        if props.get("alignment") in alignment_map:
            _retry_setattr(paragraph, "Alignment", alignment_map[props["alignment"]])
        if props.get("indent_left") is not None:
            _retry_setattr(paragraph, "LeftIndent", self._twips_to_points(props["indent_left"]))
        if props.get("indent_right") is not None:
            _retry_setattr(paragraph, "RightIndent", self._twips_to_points(props["indent_right"]))
        if props.get("indent_first_line") is not None:
            _retry_setattr(paragraph, "FirstLineIndent", self._twips_to_points(props["indent_first_line"]))
        elif props.get("indent_hanging") is not None:
            _retry_setattr(paragraph, "FirstLineIndent", -self._twips_to_points(props["indent_hanging"]))
        self._apply_paragraph_spacing(paragraph, props)
        if "keepNext" in props:
            _retry_setattr(paragraph, "KeepWithNext", self._word_bool(props["keepNext"]))
        if "keepLines" in props:
            _retry_setattr(paragraph, "KeepTogether", self._word_bool(props["keepLines"]))
        if "widowControl" in props:
            _retry_setattr(paragraph, "WidowControl", self._word_bool(props["widowControl"]))
        if "pageBreakBefore" in props:
            _retry_setattr(paragraph, "PageBreakBefore", self._word_bool(props["pageBreakBefore"]))
        tab_stops = props.get("tab_stops") or ()
        if tab_stops:
            tab_collection = _retry_getattr(paragraph, "TabStops", None)
            if tab_collection is not None:
                clear_all = _retry_getattr(tab_collection, "ClearAll", None)
                if callable(clear_all):
                    _retry_rejected_com_call(clear_all)
                tab_align = {"left": 0, "center": 1, "right": 2, "decimal": 3, "bar": 4}
                tab_leader = {"none": 0, "dot": 1, "hyphen": 2, "underscore": 3, "heavy": 4, "middleDot": 5}
                for tab in tab_stops:
                    if tab.get("val") == "clear":
                        continue
                    _retry_rejected_com_call(lambda tab=tab: tab_collection.Add(
                        Position=self._twips_to_points(tab["pos"]),
                        Alignment=tab_align.get(tab.get("val"), 0),
                        Leader=tab_leader.get(tab.get("leader"), 0),
                    ))
        self._last_paragraph_properties_key = cache_key

    def _event_InsertCharacter(self, event: ReconstructionEvent) -> None:
        character = event.payload.get("character")
        if not isinstance(character, str) or len(character) != 1:
            raise ValueError("InsertCharacter payload must contain exactly one character")
        self._insert_text(character)

    def _event_InsertText(self, event: ReconstructionEvent) -> None:
        text = event.payload.get("text")
        if not isinstance(text, str) or not text:
            raise ValueError("InsertText payload must contain non-empty text")
        self._insert_text(text)

    def _insert_text(self, text: str) -> None:
        target = self._require_range()
        _retry_rejected_com_call(lambda: target.InsertAfter(text))
        if self._active_run_properties is not None:
            self._apply_post_insert_run_properties(self._active_run_properties, text)
        self._collapse_end()
        self._page_break_continuation_pending = False
        self._post_table_paragraph_active = False

    def _event_InsertTab(self, event: ReconstructionEvent) -> None:
        self._insert_text("\t")

    def _event_InsertLineBreak(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        _retry_rejected_com_call(lambda: target.InsertBreak(Type=WD_LINE_BREAK))
        self._collapse_end()
        self._page_break_continuation_pending = False
        self._post_table_paragraph_active = False

    def _event_InsertPageBreak(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        _retry_rejected_com_call(lambda: target.InsertAfter("\f"))
        if self._active_run_properties is not None:
            self._apply_post_insert_run_properties(self._active_run_properties, "\f")
        self._collapse_end()
        self._page_break_continuation_pending = False
        self._post_table_paragraph_active = False

    @staticmethod
    def _duplicate_range(value: Any) -> Any:
        duplicate = getattr(value, "Duplicate", value)
        # pywin32 COM dispatch objects can be callable because __call__ invokes
        # the object's default property. A Word Range.Duplicate property is
        # already a Range; calling it can coerce it to its default text value.
        if hasattr(duplicate, "_oleobj_"):
            return duplicate
        return duplicate() if callable(duplicate) else duplicate

    def _current_table(self) -> dict[str, Any]:
        if not self._table_stack:
            raise RuntimeError("table event outside active table")
        return self._table_stack[-1]

    def _event_CreateListBinding(self, event: ReconstructionEvent) -> None:
        list_format = getattr(self._require_range(), "ListFormat", None)
        if list_format is None: return
        kind = event.payload.get("kind")
        if kind == "bullet":
            list_format.ApplyBulletDefault()
        else:
            list_format.ApplyNumberDefault()

    def _event_BookmarkStart(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        start = getattr(target, "Start", None)
        if start is not None:
            self._bookmark_starts[str(event.payload.get("bookmark_id", ""))] = int(start)

    def _event_CreateBookmark(self, event: ReconstructionEvent) -> None:
        if self.document is None: raise RuntimeError("interactive Word document is not open")
        bookmark_id = str(event.payload.get("bookmark_id", "")); start = self._bookmark_starts.pop(bookmark_id, None)
        end = getattr(self._require_range(), "Start", None)
        if start is None or end is None: return
        rng = _retry_rejected_com_call(lambda: self.document.Range(int(start), int(end)))
        name = str(event.payload["name"])
        com_name = sanitize_word_bookmark_name(name, taken=self._bookmark_com_names)
        if com_name != name:
            self._bookmark_name_rewrites[com_name] = name
        _retry_rejected_com_call(lambda: self.document.Bookmarks.Add(Name=com_name, Range=rng))

    def _event_CreateField(self, event: ReconstructionEvent) -> None:
        if self.document is None: raise RuntimeError("interactive Word document is not open")
        field_range = self._require_range()
        instruction = str(event.payload.get("instruction", ""))
        field = _retry_rejected_com_call(lambda: self.document.Fields.Add(
            Range=field_range, Type=-1, Text=instruction, PreserveFormatting=True,
        ))
        cached_result = event.payload.get("cached_result")
        cached_text = None
        if cached_result is not None:
            cached_text = str(cached_result)
            with suppress(Exception): field.Result.Text = cached_text
        # Do not refresh a field while later source content may not exist yet.
        # The source cached result remains visible and the output remains a real field.
        result = self._duplicate_range(field.Result)
        if cached_text is not None and self._active_run_properties is not None:
            self.active_range = result
            self._apply_post_insert_run_properties(self._active_run_properties, cached_text)
        with suppress(Exception): result.Collapse(WD_COLLAPSE_END)
        move = _retry_getattr(result, "Move", None)
        if not callable(move):
            raise RuntimeError("Word field result range does not support Move")
        _retry_rejected_com_call(lambda: move(WD_CHARACTER, 1))
        self.active_range = result

    def _create_note(self, *, footnote: bool) -> None:
        if self.document is None: raise RuntimeError("interactive Word document is not open")
        target = self._require_range()
        collection = self.document.Footnotes if footnote else self.document.Endnotes
        note = _retry_rejected_com_call(lambda: collection.Add(Range=target))
        self._active_note_index = int(getattr(note, "Index", getattr(collection, "Count", 1)))
        reference = self._duplicate_range(note.Reference)
        with suppress(Exception): reference.Collapse(WD_COLLAPSE_END)
        self._pending_note = (note, reference, "footnote" if footnote else "endnote")

    def _begin_note_story(self) -> None:
        if self._pending_note is None: raise RuntimeError("note story without created note")
        note, return_range, note_kind = self._pending_note
        self._story_stack.append((return_range, self._paragraph_started, self._active_story))
        self._active_story = note_kind
        target = self._duplicate_range(note.Range)
        with suppress(Exception): target.End = max(target.Start, target.End - 1)
        self.active_range = target; self._paragraph_started = False

    def _end_note_story(self) -> None:
        self._end_story(); self._pending_note = None; self._active_note_index = None

    def _event_CreateFootnote(self, event: ReconstructionEvent) -> None: self._create_note(footnote=True)
    def _event_BeginFootnoteStory(self, event: ReconstructionEvent) -> None: self._begin_note_story()
    def _event_EndFootnoteStory(self, event: ReconstructionEvent) -> None: self._end_note_story()
    def _event_CreateEndnote(self, event: ReconstructionEvent) -> None: self._create_note(footnote=False)
    def _event_BeginEndnoteStory(self, event: ReconstructionEvent) -> None: self._begin_note_story()
    def _event_EndEndnoteStory(self, event: ReconstructionEvent) -> None: self._end_note_story()

    def _event_BeginSection(self, event: ReconstructionEvent) -> None:
        if self.document is None: raise RuntimeError("interactive Word document is not open")
        self._active_section_index = int(event.payload.get("section_index", 0))
        if not self._section_started:
            self._active_section = _retry_rejected_com_call(lambda: self.document.Sections(1))
            self._section_started = True
        else:
            break_type = {"continuous": 0, "newColumn": 1, "nextPage": 2, "evenPage": 3, "oddPage": 4}.get(event.payload.get("break_type"), 2)
            range_for_break = self._require_range()
            self._active_section = _retry_rejected_com_call(
                lambda: self.document.Sections.Add(Range=range_for_break, Start=break_type)
            )
            start = int(_retry_rejected_com_call(lambda: self._active_section.Range.Start))
            self.active_range = _retry_rejected_com_call(lambda: self.document.Range(start, start))
        self._paragraph_started = False

    def _apply_page_number_start(self, section: Any, start: Any) -> None:
        last_error: Exception | None = None
        for collection_name in ("Footers", "Headers"):
            collection = _retry_getattr(section, collection_name, None)
            if collection is None:
                continue
            try:
                story = _retry_rejected_com_call(lambda collection=collection: collection(1))
                page_numbers = _retry_getattr(story, "PageNumbers", None)
                if page_numbers is None:
                    continue
                _retry_setattr(page_numbers, "RestartNumberingAtSection", self._word_bool(True))
                _retry_setattr(page_numbers, "StartingNumber", int(start))
                return
            except Exception as exc:
                last_error = exc
        if last_error is not None:
            raise last_error
        raise RuntimeError("Word section does not expose a PageNumbers collection")

    def _event_ApplySectionProperties(self, event: ReconstructionEvent) -> None:
        section = self._active_section
        if section is None: raise RuntimeError("section properties outside section")
        setup = section.PageSetup; props = event.payload
        if props.get("orientation") is not None:
            _retry_setattr(setup, "Orientation", 1 if props["orientation"] == "landscape" else 0)
        for key, attr in (("width","PageWidth"),("height","PageHeight"),("margin_top","TopMargin"),("margin_right","RightMargin"),("margin_bottom","BottomMargin"),("margin_left","LeftMargin"),("header_distance","HeaderDistance"),("footer_distance","FooterDistance"),("gutter","Gutter")):
            if props.get(key) is not None:
                with suppress(Exception): setattr(setup, attr, self._twips_to_points(props[key]))
        if props.get("columns") is not None:
            with suppress(Exception): setup.TextColumns.SetCount(int(props["columns"]))
        if props.get("column_space") is not None:
            with suppress(Exception): setup.TextColumns.Spacing = self._twips_to_points(props["column_space"])
        if props.get("title_page") is not None:
            with suppress(Exception): setup.DifferentFirstPageHeaderFooter = self._word_bool(props["title_page"])
        if props.get("page_number_start") is not None:
            self._apply_page_number_start(section, props["page_number_start"])

    def _event_EndSection(self, event: ReconstructionEvent) -> None:
        return

    def _begin_story(self, story_type: str, *, header: bool, link_to_previous: bool = False) -> None:
        section = self._active_section
        if section is None: raise RuntimeError("story event outside active section")
        index = {"default": 1, "first": 2, "even": 3}.get(story_type, 1)
        collection = section.Headers if header else section.Footers
        story = collection(index)
        with suppress(Exception): story.LinkToPrevious = self._word_bool(link_to_previous)
        self._story_stack.append((self._require_range(), self._paragraph_started, self._active_story))
        self._active_story = f"{'header' if header else 'footer'}:{story_type}"
        target = self._duplicate_range(story.Range)
        with suppress(Exception): target.End = max(target.Start, target.End - 1)
        self.active_range = target
        self._paragraph_started = False

    def _end_story(self) -> None:
        self.active_range, self._paragraph_started, self._active_story = self._story_stack.pop()

    def _event_BeginHeader(self, event: ReconstructionEvent) -> None:
        self._begin_story(
            str(event.payload.get("story_type", "default")),
            header=True,
            link_to_previous=bool(event.payload.get("link_to_previous", False)),
        )

    def _event_EndHeader(self, event: ReconstructionEvent) -> None:
        self._end_story()

    def _event_BeginFooter(self, event: ReconstructionEvent) -> None:
        self._begin_story(
            str(event.payload.get("story_type", "default")),
            header=False,
            link_to_previous=bool(event.payload.get("link_to_previous", False)),
        )

    def _event_EndFooter(self, event: ReconstructionEvent) -> None:
        self._end_story()

    def _event_InsertImage(self, event: ReconstructionEvent) -> None:
        self._image_transaction_active = True
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        if self._asset_resolver is None:
            raise RuntimeError("interactive image asset resolver is not configured")
        asset_path = self._asset_resolver(str(event.payload["asset_id"]))
        image_range = self._require_range()
        file_name = str(asset_path)
        inline = _retry_rejected_com_call(lambda: self.document.InlineShapes.AddPicture(
            FileName=file_name, LinkToFile=False, SaveWithDocument=True, Range=image_range
        ))
        font_cs = (self._active_run_properties or {}).get("font_cs")
        if font_cs:
            with suppress(Exception):
                image_font = _retry_getattr(_retry_getattr(inline, "Range"), "Font")
                _retry_setattr(image_font, "NameBi", str(font_cs))
        after_range = self._duplicate_range(inline.Range)
        with suppress(Exception): after_range.Collapse(WD_COLLAPSE_END)
        representation = str(event.payload.get("representation", "inline"))
        if representation == "floating":
            self._active_image = inline.ConvertToShape()
        else:
            self._active_image = inline
        self._active_image_representation = representation
        self.active_range = after_range
        self._page_break_continuation_pending = False

    def _event_SetImageSize(self, event: ReconstructionEvent) -> None:
        image = self._active_image
        if image is None: raise RuntimeError("image geometry event without active image")
        lock_aspect_ratio = event.payload.get("lock_aspect_ratio")
        width_emu = event.payload.get("width_emu")
        height_emu = event.payload.get("height_emu")
        unlock_for_exact_size = width_emu is not None and height_emu is not None
        restore_lock = None
        if unlock_for_exact_size:
            restore_lock = (
                self._word_bool(lock_aspect_ratio)
                if lock_aspect_ratio is not None
                else _retry_getattr(image, "LockAspectRatio", None)
            )
            _retry_setattr(image, "LockAspectRatio", self._word_bool(False))
        elif lock_aspect_ratio is not None:
            _retry_setattr(image, "LockAspectRatio", self._word_bool(lock_aspect_ratio))
        try:
            if width_emu is not None:
                image.Width = self._emu_to_points(width_emu)
            if height_emu is not None:
                image.Height = self._emu_to_points(height_emu)
        finally:
            if unlock_for_exact_size and restore_lock is not None:
                _retry_setattr(image, "LockAspectRatio", restore_lock)

    def _event_SetImageWrap(self, event: ReconstructionEvent) -> None:
        image = self._active_image
        if image is None: raise RuntimeError("image wrap event without active image")
        if self._active_image_representation != "floating": return
        wrap = image.WrapFormat
        if event.payload.get("behind_text"):
            wrap.Type = 5
        elif event.payload.get("wrap_type"):
            wrap.Type = {"square": 0, "tight": 1, "through": 2, "none": 3, "topAndBottom": 4}.get(event.payload["wrap_type"], 0)
        for key, attr in (("distance_top_emu","DistanceTop"),("distance_bottom_emu","DistanceBottom"),("distance_left_emu","DistanceLeft"),("distance_right_emu","DistanceRight")):
            value = event.payload.get(key)
            if value is not None:
                with suppress(Exception): setattr(wrap, attr, self._emu_to_points(value))

    def _event_SetImagePosition(self, event: ReconstructionEvent) -> None:
        image = self._active_image
        if image is None: raise RuntimeError("image position event without active image")
        if self._active_image_representation != "floating": return
        hmap = {"margin": 0, "page": 1, "column": 2, "character": 3, "leftMargin": 4, "rightMargin": 5, "insideMargin": 6, "outsideMargin": 7}
        vmap = {"margin": 0, "page": 1, "paragraph": 2, "line": 3, "topMargin": 4, "bottomMargin": 5, "insideMargin": 6, "outsideMargin": 7}
        if event.payload.get("horizontal_relative_from") in hmap:
            with suppress(Exception): image.RelativeHorizontalPosition = hmap[event.payload["horizontal_relative_from"]]
        if event.payload.get("vertical_relative_from") in vmap:
            with suppress(Exception): image.RelativeVerticalPosition = vmap[event.payload["vertical_relative_from"]]
        if event.payload.get("horizontal_position_emu") is not None:
            image.Left = self._emu_to_points(event.payload["horizontal_position_emu"])
        if event.payload.get("vertical_position_emu") is not None:
            image.Top = self._emu_to_points(event.payload["vertical_position_emu"])

    def _event_SetImageCrop(self, event: ReconstructionEvent) -> None:
        image = self._active_image
        if image is None: raise RuntimeError("image crop event without active image")
        crop = event.payload.get("crop") or {}
        picture = getattr(image, "PictureFormat", None)
        if picture is None: return
        width = float(getattr(image, "Width", 0) or 0); height = float(getattr(image, "Height", 0) or 0)
        for key, attr, base in (("l","CropLeft",width),("r","CropRight",width),("t","CropTop",height),("b","CropBottom",height)):
            if key in crop:
                with suppress(Exception): setattr(picture, attr, base * (float(crop[key]) / 100000.0))

    def _event_SetImageRotation(self, event: ReconstructionEvent) -> None:
        image = self._active_image
        if image is None: raise RuntimeError("image rotation event without active image")
        if event.payload.get("rotation_degrees") is not None:
            with suppress(Exception): image.Rotation = float(event.payload["rotation_degrees"])

    def _event_SetImageZOrder(self, event: ReconstructionEvent) -> None:
        image = self._active_image
        if image is None: raise RuntimeError("image z-order event without active image")
        if self._active_image_representation == "floating":
            z_value = event.payload.get("z_order")
            if z_value is None:
                z_value = len(self._floating_images)
            self._floating_images.append((int(z_value), image))
            # Word does not expose DrawingML relativeHeight as a writable absolute
            # value. Recreate the same relative stacking order among reconstructed
            # floating images by bringing them to front in ascending source order.
            for _, shape in sorted(self._floating_images, key=lambda item: item[0]):
                with suppress(Exception):
                    shape.ZOrder(0)  # Office msoBringToFront
        self._image_transaction_active = False

    def _event_BeginTable(self, event: ReconstructionEvent) -> None:
        self._active_table_element_id = event.source_element_id
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        rows = int(event.payload["rows"]); columns = int(event.payload["columns"])
        parent_range = self._require_range()
        table = _retry_rejected_com_call(lambda: self.document.Tables.Add(Range=parent_range, NumRows=rows, NumColumns=columns))
        cells = {(r, c): table.Cell(r, c) for r in range(1, rows + 1) for c in range(1, columns + 1)}
        after_range = self._duplicate_range(table.Range)
        with suppress(Exception):
            after_range.Collapse(WD_COLLAPSE_END)
        sequence = self._table_insertion_sequence
        self._table_insertion_sequence += 1
        self._table_stack.append({"table": table, "cells": cells, "parent_range": parent_range, "after_range": after_range, "element_id": event.source_element_id, "structure_complete": False, "sequence": sequence})
        self._paragraph_started = False

    def _event_InsertTableBatch(self, event: ReconstructionEvent) -> None:
        started = time.perf_counter()
        cells = list(event.payload.get("cells") or ())
        metrics: dict[str, object] = {
            "table_id": event.source_element_id,
            "row_count": int(event.payload.get("rows") or 0),
            "cell_count": len(cells),
            "run_count": sum(len(cell.get("runs") or ()) for cell in cells),
            "insert_seconds": 0.0,
            "convert_seconds": 0.0,
            "geometry_seconds": 0.0,
            "formatting_seconds": 0.0,
            "verification_seconds": 0.0,
            "total_seconds": 0.0,
            "failed_phase": None,
            "success": False,
            "_phase": "validation",
            "_phase_started": started,
        }
        self._table_batch_metrics = None
        try:
            self._execute_table_batch(event, metrics)
        except Exception:
            failed_phase = str(metrics.pop("_phase", "unknown"))
            phase_started = float(metrics.pop("_phase_started", started))
            phase_metric = f"{failed_phase}_seconds"
            if phase_metric in metrics and float(metrics[phase_metric]) == 0.0:
                metrics[phase_metric] = max(0.0, time.perf_counter() - phase_started)
            metrics["failed_phase"] = failed_phase
            metrics["total_seconds"] = max(0.0, time.perf_counter() - started)
            self._table_batch_metrics = metrics
            raise
        metrics.pop("_phase", None)
        metrics.pop("_phase_started", None)
        metrics["success"] = True
        metrics["total_seconds"] = max(0.0, time.perf_counter() - started)
        self._table_batch_metrics = metrics

    def _execute_table_batch(
        self,
        event: ReconstructionEvent,
        metrics: dict[str, object],
    ) -> None:
        def start_phase(name: str) -> None:
            metrics["_phase"] = name
            metrics["_phase_started"] = time.perf_counter()

        def finish_phase(name: str) -> None:
            metrics[f"{name}_seconds"] = max(
                0.0,
                time.perf_counter() - float(metrics["_phase_started"]),
            )

        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        rows = int(event.payload["rows"])
        columns = int(event.payload["columns"])
        cells = list(event.payload["cells"])
        if len(cells) != rows * columns:
            raise ValueError("InsertTableBatch payload must contain every table cell")
        expected_coordinates = [
            (row, column)
            for row in range(1, rows + 1)
            for column in range(1, columns + 1)
        ]
        actual_coordinates = [
            (int(cell["row"]), int(cell["column"]))
            for cell in cells
        ]
        if actual_coordinates != expected_coordinates:
            raise ValueError("InsertTableBatch cells must be in complete row-major order")

        payload_text = "\r".join(
            "\t".join(
                str(cells[(row - 1) * columns + column - 1]["text"])
                for column in range(1, columns + 1)
            )
            for row in range(1, rows + 1)
        )
        start_phase("insert")
        parent_range = self._require_range()
        insertion_start = int(_retry_getattr(parent_range, "Start"))
        _retry_rejected_com_call(lambda: parent_range.InsertAfter(payload_text))
        inserted_range = _retry_rejected_com_call(
            lambda: self.document.Range(insertion_start, insertion_start + len(payload_text))
        )
        finish_phase("insert")
        start_phase("convert")
        table = _retry_rejected_com_call(
            lambda: inserted_range.ConvertToTable(
                Separator=1,
                NumRows=rows,
                NumColumns=columns,
            )
        )
        finish_phase("convert")
        start_phase("geometry")
        # Word's Table.Cell(row, column) re-resolves from the table root on
        # every call - O(n) or worse per call, so a per-cell loop is O(n^2)
        # over the whole table and gets dramatically slower as the table (and
        # the surrounding document) grows. A single pass over the table's own
        # Range.Cells collection returns every cell in row-major order in one
        # COM enumeration instead of `rows * columns` individual round trips.
        table_range_for_cells = _retry_getattr(table, "Range")
        cells_collection = _retry_rejected_com_call(lambda: list(table_range_for_cells.Cells))
        if len(cells_collection) != len(expected_coordinates):
            raise RuntimeError("InsertTableBatch cell collection size mismatch")
        word_cells = dict(zip(expected_coordinates, cells_collection))
        table_range = _retry_getattr(table, "Range")
        after_range = self._duplicate_range(table_range)
        collapse = _retry_getattr(after_range, "Collapse", None)
        if callable(collapse):
            _retry_rejected_com_call(lambda: collapse(WD_COLLAPSE_END))
        sequence = self._table_insertion_sequence
        self._table_insertion_sequence += 1
        context = {
            "table": table,
            "cells": word_cells,
            "parent_range": parent_range,
            "after_range": after_range,
            "element_id": event.source_element_id,
            "structure_complete": True,
            "sequence": sequence,
        }
        self._table_stack.append(context)
        self._active_table_element_id = event.source_element_id
        try:
            self._event_SetTableProperties(ReconstructionEvent(
                "SetTableProperties",
                event.source_element_id,
                dict(event.payload.get("table_properties") or {}),
            ))
            for width in event.payload.get("column_widths") or ():
                self._event_SetColumnWidth(ReconstructionEvent(
                    "SetColumnWidth", event.source_element_id, dict(width)
                ))
            for row_record in event.payload.get("row_properties") or ():
                self._event_SetRowProperties(ReconstructionEvent(
                    "SetRowProperties", event.source_element_id, dict(row_record)
                ))
            for cell_record in cells:
                cell_id = str(cell_record["source_element_id"])
                row = int(cell_record["row"])
                column = int(cell_record["column"])
                self._event_SetCellProperties(ReconstructionEvent(
                    "SetCellProperties",
                    cell_id,
                    {
                        "row": row,
                        "column": column,
                        "properties": dict(cell_record.get("properties") or {}),
                    },
                ))

            finish_phase("geometry")
            start_phase("formatting")
            paragraph_keys = [
                json.dumps(
                    dict(cell.get("paragraph_properties") or {}),
                    sort_keys=True,
                    ensure_ascii=False,
                    default=str,
                )
                for cell in cells
            ]
            common_paragraph_properties = (
                dict(cells[0].get("paragraph_properties") or {})
                if len(cells) > 1 and len(set(paragraph_keys)) == 1
                else None
            )
            run_properties_by_key: dict[str, dict[str, Any]] = {}
            run_property_counts: dict[str, int] = {}
            for cell_record in cells:
                for run_record in cell_record.get("runs") or ():
                    properties = dict(run_record.get("properties") or {})
                    key = json.dumps(
                        properties,
                        sort_keys=True,
                        ensure_ascii=False,
                        default=str,
                    )
                    run_properties_by_key.setdefault(key, properties)
                    run_property_counts[key] = run_property_counts.get(key, 0) + 1
            baseline_run_key = None
            if run_property_counts:
                candidate = max(run_property_counts, key=run_property_counts.get)
                if run_property_counts[candidate] > 1:
                    baseline_run_key = candidate

            common_range = None
            if common_paragraph_properties is not None or baseline_run_key is not None:
                common_range = self._duplicate_range(table_range)
            if common_paragraph_properties is not None:
                self.active_range = common_range
                self._event_ApplyParagraphProperties(ReconstructionEvent(
                    "ApplyParagraphProperties",
                    event.source_element_id,
                    common_paragraph_properties,
                ))

            for cell_record in cells:
                cell_id = str(cell_record["source_element_id"])
                key = (int(cell_record["row"]), int(cell_record["column"]))
                self._active_cell_element_id = cell_id
                word_cell_range = _retry_getattr(word_cells[key], "Range")
                cell_start = int(_retry_getattr(word_cell_range, "Start"))
                if common_paragraph_properties is None:
                    cell_range = self._duplicate_range(word_cell_range)
                    cell_end = int(_retry_getattr(cell_range, "End"))
                    _retry_setattr(cell_range, "End", max(cell_start, cell_end - 1))
                    self.active_range = cell_range
                    self._event_ApplyParagraphProperties(ReconstructionEvent(
                        "ApplyParagraphProperties",
                        cell_id,
                        dict(cell_record.get("paragraph_properties") or {}),
                    ))

            if baseline_run_key is not None:
                self.active_range = common_range
                self._event_ApplyRunProperties(ReconstructionEvent(
                    "ApplyRunProperties",
                    event.source_element_id,
                    run_properties_by_key[baseline_run_key],
                ))

            for cell_record in cells:
                cell_id = str(cell_record["source_element_id"])
                key = (int(cell_record["row"]), int(cell_record["column"]))
                self._active_cell_element_id = cell_id
                word_cell_range = _retry_getattr(word_cells[key], "Range")
                cell_start = int(_retry_getattr(word_cell_range, "Start"))
                for run_record in cell_record.get("runs") or ():
                    run_properties = dict(run_record.get("properties") or {})
                    run_key = json.dumps(
                        run_properties,
                        sort_keys=True,
                        ensure_ascii=False,
                        default=str,
                    )
                    if run_key == baseline_run_key:
                        continue
                    run_start = cell_start + int(run_record["start"])
                    run_end = cell_start + int(run_record["end"])
                    run_range = _retry_rejected_com_call(
                        lambda run_start=run_start, run_end=run_end: self.document.Range(
                            run_start, run_end
                        )
                    )
                    self.active_range = run_range
                    self._event_ApplyRunProperties(ReconstructionEvent(
                        "ApplyRunProperties",
                        str(run_record["source_element_id"]),
                        run_properties,
                    ))
            finish_phase("formatting")
        except Exception:
            self._table_stack.pop()
            self._active_table_element_id = None
            self._active_cell_element_id = None
            raise

        self._table_stack.pop()
        self.active_range = after_range
        self._active_table_element_id = None
        self._active_cell_element_id = None
        self._active_run_properties = None
        self._paragraph_started = False
        self._post_table_paragraph_active = True
        start_phase("verification")
        if len(word_cells) != rows * columns or self.active_range is not after_range:
            raise RuntimeError("InsertTableBatch postcondition failed")
        finish_phase("verification")

    def _event_SetTableProperties(self, event: ReconstructionEvent) -> None:
        table = self._current_table()["table"]; props = event.payload
        if props.get("style_id"):
            with suppress(Exception): table.Style = str(props["style_id"])
        alignment = {"left": 0, "center": 1, "right": 2}.get(props.get("alignment"))
        if alignment is not None:
            with suppress(Exception): table.Rows.Alignment = alignment
        layout = props.get("layout")
        if layout == "fixed":
            with suppress(Exception): table.AllowAutoFit = False
        elif layout == "autofit":
            with suppress(Exception): table.AllowAutoFit = True
        width = props.get("width"); width_type = props.get("width_type")
        if width is not None:
            if width_type in {None, "dxa"}:
                with suppress(Exception):
                    table.PreferredWidthType = 3
                    table.PreferredWidth = self._twips_to_points(width)
            elif width_type == "pct":
                with suppress(Exception):
                    table.PreferredWidthType = 2
                    table.PreferredWidth = float(width) / 50.0
            elif width_type == "auto":
                with suppress(Exception): table.PreferredWidthType = 1
        self._apply_padding(table, props.get("cell_margins"))
        self._apply_shading(table, props.get("shading_fill"))
        self._apply_borders(table, props.get("borders"))

    def _event_SetColumnWidth(self, event: ReconstructionEvent) -> None:
        ctx = self._current_table()
        table = ctx["table"]
        column = int(event.payload["column"])
        width_twips = int(event.payload["width_twips"])
        width_points = self._twips_to_points(width_twips)
        if width_points < _WORD_MIN_COM_COLUMN_WIDTH_POINTS:
            # Word's Columns(i).Width setter (and Cell.Width/SetWidth - all
            # tried) reject anything below ~11pt with "Value out of range",
            # even with zero cell padding - a hard floor in the COM layer,
            # not a margins issue. The OOXML format itself has no such floor,
            # so set a COM-safe width now and correct the true value directly
            # in the saved file's tblGrid/tcW after close() (see
            # _restore_narrow_column_widths), the same pattern used for
            # bookmark names.
            self._narrow_column_fixups.append((int(ctx["sequence"]), column, width_twips))
            width_points = _WORD_MIN_COM_COLUMN_WIDTH_POINTS
        table.Columns(column).Width = width_points

    def _event_SetRowProperties(self, event: ReconstructionEvent) -> None:
        table = self._current_table()["table"]; row = table.Rows(int(event.payload["row"])); props = event.payload.get("properties", {})
        if props.get("height") is not None:
            row.Height = self._twips_to_points(props["height"])
        if props.get("height_rule"):
            row.HeightRule = {"auto": 0, "atLeast": 1, "exact": 2}.get(props["height_rule"], 0)
        if "repeat_header" in props:
            with suppress(Exception): row.HeadingFormat = self._word_bool(props["repeat_header"])
        if "cant_split" in props:
            with suppress(Exception): row.AllowBreakAcrossPages = not bool(props["cant_split"])

    def _event_SetCellProperties(self, event: ReconstructionEvent) -> None:
        ctx = self._current_table(); cell = ctx["cells"][(int(event.payload["row"]), int(event.payload["column"]))]
        props = event.payload.get("properties", {})
        if props.get("vertical_alignment") is not None:
            mapping = {"top": 0, "center": 1, "bottom": 3}
            if props["vertical_alignment"] in mapping:
                cell.VerticalAlignment = mapping[props["vertical_alignment"]]
        if props.get("width") is not None and props.get("width_type") in {None, "dxa"}:
            with suppress(Exception):
                cell.PreferredWidthType = 3
                cell.PreferredWidth = self._twips_to_points(props["width"])
        self._apply_padding(cell, props.get("cell_margins"))
        self._apply_shading(cell, props.get("shading_fill"))
        self._apply_borders(cell, props.get("borders"))
        direction = props.get("text_direction")
        if direction is not None:
            orientation = {"lrTb": 0, "tbRl": 3, "btLr": 2, "tbRlV": 1}.get(str(direction))
            if orientation is not None:
                with suppress(Exception): cell.Range.Orientation = orientation

    def _event_MergeCells(self, event: ReconstructionEvent) -> None:
        ctx = self._current_table(); p = event.payload
        start = ctx["cells"][(int(p["start_row"]), int(p["start_column"]))]
        end = ctx["cells"][(int(p["end_row"]), int(p["end_column"]))]
        _retry_rejected_com_call(lambda: start.Merge(end))

    def _event_EnterCell(self, event: ReconstructionEvent) -> None:
        self._active_cell_element_id = event.source_element_id
        ctx = self._current_table(); p = event.payload
        ctx["structure_complete"] = True
        key = (int(p["row"]), int(p["column"]))
        cell = ctx["cells"].get(key)
        if cell is None:
            cell = ctx["table"].Cell(*key)
            ctx["cells"][key] = cell
        target = self._duplicate_range(cell.Range)
        with suppress(Exception):
            target.End = max(target.Start, target.End - 1)
        self.active_range = target
        self._paragraph_started = False

    def _event_LeaveCell(self, event: ReconstructionEvent) -> None:
        self._active_cell_element_id = None
        self._paragraph_started = False

    def _event_EndTable(self, event: ReconstructionEvent) -> None:
        ctx = self._current_table()
        self.active_range = ctx["after_range"]
        self._table_stack.pop()
        self._active_table_element_id = self._table_stack[-1].get("element_id") if self._table_stack else None
        self._active_cell_element_id = None
        # Word keeps a real paragraph immediately after a table. Reuse it for
        # the next source paragraph instead of inserting an extra blank one.
        self._paragraph_started = False
        self._post_table_paragraph_active = True

    def _event_BeginParagraph(self, event: ReconstructionEvent) -> None:
        if self._page_break_continuation_pending:
            self._page_break_continuation_pending = False
            self._paragraph_started = True
            return
        if not self._paragraph_started:
            self._paragraph_started = True
            return
        target = self._require_range()
        _retry_rejected_com_call(lambda: target.InsertParagraphAfter())
        self._collapse_end()
        self._post_table_paragraph_active = False

    def _event_EndParagraph(self, event: ReconstructionEvent) -> None:
        return

    def is_restart_safe(self) -> bool:
        if self._image_transaction_active:
            return False
        # A persisted checkpoint can currently rehydrate one active Word table
        # from the saved insertion range. Nested table stacks are therefore
        # restart-safe only after the innermost table has been exited.
        if len(self._table_stack) > 1:
            return False
        return all(bool(ctx.get("structure_complete", False)) for ctx in self._table_stack)

    def save(self, path: str | Path) -> None:
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        destination = Path(path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        _retry_rejected_com_call(lambda: self.document.SaveAs2(str(destination), FileFormat=WD_FORMAT_DOCX))
        # Bookmark-name restoration is deferred to close() rather than done here:
        # Word keeps this exact file open (and locked) as its active document for
        # as long as this session runs, so a rename-based rewrite against it here
        # deterministically fails with WinError 5 (confirmed - even hidden/hasn't-
        # yet-been-touched-again, the file stays locked the entire time the
        # document stays open, not just briefly after SaveAs2 returns).
        self._last_saved_path = destination

    def restore_pending_bookmark_names(self) -> None:
        """Rewrite sanitized COM-safe bookmark names back to their source names
        in the last-saved file. Must run after close() - Word only releases its
        lock on the file once the document is closed."""
        if self._last_saved_path is None or not self._bookmark_name_rewrites:
            return
        _restore_bookmark_names(self._last_saved_path, self._bookmark_name_rewrites)

    def restore_pending_narrow_column_widths(self) -> None:
        """Correct table columns that were widened to a COM-safe placeholder
        (see `_WORD_MIN_COM_COLUMN_WIDTH_POINTS`) back to their true source
        width in the last-saved file. Must run after close() for the same
        reason as `restore_pending_bookmark_names`."""
        if self._last_saved_path is None or not self._narrow_column_fixups:
            return
        _restore_narrow_column_widths(self._last_saved_path, self._narrow_column_fixups)

    def set_custom_property(self, name: str, value: Any) -> None:
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        properties = self.document.CustomDocumentProperties
        with suppress(Exception):
            _retry_rejected_com_call(lambda: properties(name).Delete())
        # Office MsoDocProperties.msoPropertyTypeString = 4
        _retry_rejected_com_call(lambda: properties.Add(name, False, 4, str(value)))

    def apply_metadata(self, policy: dict[str, str]) -> None:
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        built_in = {
            "title": "Title", "subject": "Subject", "keywords": "Keywords",
            "category": "Category", "creator": "Author",
        }
        for key, word_name in built_in.items():
            if key not in policy:
                continue
            with suppress(Exception):
                self.document.BuiltInDocumentProperties(word_name).Value = str(policy[key])
        for key, value in policy.items():
            if key.startswith("custom:"):
                self.set_custom_property(key.split(":", 1)[1], value)

    def current_state_snapshot(self) -> WordStateSnapshot:
        target = self.active_range
        return WordStateSnapshot(
            document_identity=f"com:{id(self.document)}" if self.document is not None else "closed",
            story=self._active_story,
            range_start=_retry_getattr(target, "Start", None) if target is not None else None,
            range_end=_retry_getattr(target, "End", None) if target is not None else None,
            section_index=self._active_section_index,
            table_element_id=self._active_table_element_id,
            cell_element_id=self._active_cell_element_id,
            last_completed_event_index=None,
            active_range_type=type(target).__name__ if target is not None else None,
        )

    def close(self) -> None:
        if self.document is not None:
            with suppress(Exception):
                _retry_rejected_com_call(lambda: self.document.Close(False))
            self.document = None
            # Only reachable once Word has actually released its lock on the
            # saved file (see restore_pending_bookmark_names / save()).
            with suppress(Exception):
                self.restore_pending_bookmark_names()
            with suppress(Exception):
                self.restore_pending_narrow_column_widths()
        application_quit = self.application is None
        if self.application is not None:
            try:
                _retry_rejected_com_call(lambda: self.application.Quit())
                application_quit = True
            except Exception:
                application_quit = False
            self.application = None
        self.active_range = None
        if self._owns_com:
            with suppress(Exception):
                import pythoncom
                pythoncom.CoUninitialize()
            self._owns_com = False
        if application_quit and (
            self._owned_word_pid is None
            or self._owned_word_pid not in word_process_pids()
        ):
            clear_owned_word(self._owned_word_pid)
            self._owned_word_pid = None


class InteractiveWordRenderer:
    """Thin owner for the dedicated visible interactive Word controller."""

    def __init__(
        self,
        controller: InteractiveWordController | None = None,
        *,
        options: InteractiveOptions | None = None,
        speed: SpeedController | None = None,
    ) -> None:
        self.controller = controller or InteractiveWordController()
        self.options = options or InteractiveOptions()
        self.speed = speed or SpeedController(self.options)

    def open_blank(self) -> None:
        self.controller.open_blank()

    def open_existing(self, path: str | Path) -> None:
        self.controller.open_existing(path)

    def restore_checkpoint_state(self, checkpoint) -> None:
        self.controller.restore_checkpoint_state(checkpoint)

    def execute_event(self, event: ReconstructionEvent) -> None:
        self.controller.execute_event(event)


    def execute_blueprint(self, blueprint: ReconstructionBlueprint, control, observer=None, *, start_index: int = 0) -> ExecutionOutcome:
        if start_index < 0 or start_index > blueprint.total_events:
            raise ValueError("start_index out of range")
        snapshot_fn = getattr(self.controller, "current_state_snapshot", None)

        def take_snapshot(index, event, phase):
            if not callable(snapshot_fn):
                return None
            try:
                return snapshot_fn()
            except Exception as exc:
                callback = getattr(observer, "state_snapshot_failed", None) if observer is not None else None
                if callback is not None:
                    callback(index, event, phase, exc)
                raise

        def emit_table_batch_profile(index, event) -> None:
            if event.event_type != "InsertTableBatch":
                return
            consume = getattr(self.controller, "consume_table_batch_metrics", None)
            metrics = consume() if callable(consume) else None
            callback = getattr(observer, "table_batch_profile", None) if observer is not None else None
            if metrics is not None and callback is not None:
                callback(index, event, metrics)

        initial_event = blueprint.events[start_index] if start_index < blueprint.total_events else None
        expected_state = take_snapshot(start_index, initial_event, "initial_state") if initial_event is not None else None
        previous_event_type = None
        for index in range(start_index, blueprint.total_events):
            event = blueprint.events[index]
            actual_state = None
            text_events = {"InsertCharacter", "InsertText", "InsertTab", "InsertLineBreak", "InsertPageBreak"}
            reuse_post_snapshot = (
                expected_state is not None
                and event.event_type in text_events
                and previous_event_type == "ApplyRunProperties"
            )
            if expected_state is not None and callable(snapshot_fn) and not reuse_post_snapshot:
                actual_state = take_snapshot(index, event, "pre_state_check")
                if not state_snapshots_match(expected_state, actual_state):
                    control.pause()
                    callback = getattr(observer, "state_mismatch", None) if observer is not None else None
                    if callback is not None:
                        callback(index, event, expected_state, actual_state)
                    return ExecutionOutcome("PAUSED_STATE_MISMATCH", index - 1, max(0, index))
            decision = control.before_next_event(index - 1)
            if decision.stop_requested:
                restart_safe = getattr(self.controller, "is_restart_safe", lambda: True)()
                if restart_safe:
                    return ExecutionOutcome.stopped(index - 1)
            # The verified pre-event snapshot is also the observer's before-state.
            # Avoid a redundant COM Range poll for every single character event.
            before_state = (
                actual_state
                if actual_state is not None
                else expected_state
                if reuse_post_snapshot
                else take_snapshot(index, event, "before_event")
            )
            started = getattr(observer, "event_started", None) if observer is not None else None
            if started is not None:
                started(index, event, before_state)
            letter_by_letter = event.event_type == "InsertText" and self.options.letter_by_letter
            try:
                if letter_by_letter:
                    self._execute_insert_text_letter_by_letter(event, control)
                else:
                    self.execute_event(event)
            except Exception as exc:
                emit_table_batch_profile(index, event)
                failed = getattr(observer, "event_failed", None) if observer is not None else None
                if failed is not None:
                    failed(index, event, before_state, exc)
                raise
            emit_table_batch_profile(index, event)
            expected_state = take_snapshot(index, event, "post_event")
            previous_event_type = event.event_type
            if observer is not None:
                observer.event_completed(index, event)
                finished = getattr(observer, "event_finished", None)
                if finished is not None:
                    finished(index, event, expected_state)
                halt_status = getattr(observer, "halt_status", None)
                if halt_status:
                    return ExecutionOutcome(str(halt_status), index, min(blueprint.total_events, index + 1))
            if not letter_by_letter:
                character_count = len(str(event.payload.get("text", ""))) if event.event_type == "InsertText" else 1
                self.speed.delay_after(event.event_type, character_count=character_count)
        return ExecutionOutcome.completed(blueprint.total_events)

    def _execute_insert_text_letter_by_letter(self, event: ReconstructionEvent, control) -> None:
        """Replay one InsertText blueprint event as individual on-screen keystrokes.

        The blueprint keeps whole-run InsertText events (stable checkpoints, compact
        replay); only this replay step fans a run out into single characters so a
        paused run can react between letters. A stop request is intentionally not
        honored mid-word: the outer loop only ever resumes at whole-event boundaries,
        so stopping partway through a run would leave the document with half of an
        event applied and nothing to resume from.
        """
        text = str(event.payload.get("text", ""))
        for character in text:
            control.before_next_event(-1)
            self.execute_event(ReconstructionEvent("InsertCharacter", event.source_element_id, {"character": character}))
            self.speed.delay_after("InsertCharacter", character_count=1)

    def resume_state_snapshot(self) -> dict[str, Any]:
        return dict(getattr(self.controller, "resume_state_snapshot", self.current_state_snapshot)())

    def is_restart_safe(self) -> bool:
        return bool(getattr(self.controller, "is_restart_safe", lambda: True)())

    def save(self, path: str | Path) -> None:
        self.controller.save(path)

    def current_state_snapshot(self) -> dict[str, Any]:
        return self.controller.current_state_snapshot()

    def set_custom_property(self, name: str, value: Any) -> None:
        self.controller.set_custom_property(name, value)

    def apply_metadata(self, policy: dict[str, str]) -> None:
        if not policy:
            return
        self.controller.apply_metadata(policy)

    def close(self) -> None:
        self.controller.close()
