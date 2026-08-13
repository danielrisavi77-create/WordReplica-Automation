from __future__ import annotations

from contextlib import suppress
from pathlib import Path
import json
import time
from typing import Any

from word_replica.domain.reconstruction import ExecutionOutcome, ReconstructionBlueprint, ReconstructionEvent, WordStateSnapshot
from word_replica.interactive.speed import SpeedController
from word_replica.config import InteractiveOptions
from word_replica.interactive.verification import state_snapshots_match
from word_replica.renderers.word_ownership import clear_owned_word, record_owned_word, word_process_pids


WD_COLLAPSE_END = 0
WD_LINE_BREAK = 6
WD_PAGE_BREAK = 7
WD_FORMAT_DOCX = 16

RPC_E_CALL_REJECTED = -2147418111


def _is_rejected_com_call(exc: Exception) -> bool:
    hresult = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        first = exc.args[0]
        if isinstance(first, int):
            hresult = first
    return hresult == RPC_E_CALL_REJECTED


def _retry_rejected_com_call(operation, *, attempts: int = 300, delay_seconds: float = 0.1):
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
        self._asset_resolver = None
        self._active_image: Any | None = None
        self._active_image_representation: str | None = None
        self._section_started = False
        self._active_section: Any | None = None
        self._story_stack: list[tuple[Any, bool]] = []
        self._bookmark_starts: dict[str, int] = {}
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
        self._paragraph_format_context_generation = 0
        self._last_paragraph_properties_key: tuple[int, str, int] | None = None
        self._paragraph_style_cache: dict[tuple[str, str], Any] = {}

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
        self.application.Visible = self.visible
        self.document = self.application.Documents.Add()
        self.active_range = self.document.Range(0, 0)
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
        self.active_range = self._resume_range_for_story(
            story, int(start), int(end), section_index=self._active_section_index, note_index=note_index
        )
        self._paragraph_started = bool(getattr(checkpoint, "paragraph_started", False))
        self._active_story = story
        self._active_note_index = int(note_index) if note_index is not None else None
        self._active_table_element_id = getattr(checkpoint, "table_element_id", None)
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
        self._collapse_end()
        self._page_break_continuation_pending = False
        self._post_table_paragraph_active = False

    def _event_InsertTab(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        _retry_rejected_com_call(lambda: target.InsertAfter("\t"))
        self._collapse_end()
        self._page_break_continuation_pending = False
        self._post_table_paragraph_active = False

    def _event_InsertLineBreak(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        _retry_rejected_com_call(lambda: target.InsertBreak(Type=WD_LINE_BREAK))
        self._collapse_end()
        self._page_break_continuation_pending = False
        self._post_table_paragraph_active = False

    def _event_InsertPageBreak(self, event: ReconstructionEvent) -> None:
        target = self._require_range()
        post_table = self._post_table_paragraph_active
        if post_table:
            _retry_rejected_com_call(lambda: target.InsertAfter("\f"))
        else:
            _retry_rejected_com_call(lambda: target.InsertBreak(Type=WD_PAGE_BREAK))
        self._collapse_end()
        self._page_break_continuation_pending = not post_table
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
        rng = self.document.Range(int(start), int(end))
        self.document.Bookmarks.Add(Name=str(event.payload["name"]), Range=rng)

    def _event_CreateField(self, event: ReconstructionEvent) -> None:
        if self.document is None: raise RuntimeError("interactive Word document is not open")
        field = self.document.Fields.Add(
            Range=self._require_range(), Type=-1,
            Text=str(event.payload.get("instruction", "")), PreserveFormatting=True,
        )
        cached_result = event.payload.get("cached_result")
        if cached_result is not None:
            with suppress(Exception): field.Result.Text = str(cached_result)
        # Do not refresh a field while later source content may not exist yet.
        # The source cached result remains visible and the output remains a real field.
        result = self._duplicate_range(field.Result)
        with suppress(Exception): result.Collapse(WD_COLLAPSE_END)
        self.active_range = result

    def _create_note(self, *, footnote: bool) -> None:
        if self.document is None: raise RuntimeError("interactive Word document is not open")
        target = self._require_range()
        collection = self.document.Footnotes if footnote else self.document.Endnotes
        note = collection.Add(Range=target)
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
            self._active_section = self.document.Sections.Add(Range=self._require_range(), Start=break_type)
            with suppress(Exception):
                start = int(self._active_section.Range.Start)
                self.active_range = self.document.Range(start, start)
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

    def _begin_story(self, story_type: str, *, header: bool) -> None:
        section = self._active_section
        if section is None: raise RuntimeError("story event outside active section")
        index = {"default": 1, "first": 2, "even": 3}.get(story_type, 1)
        collection = section.Headers if header else section.Footers
        story = collection(index)
        with suppress(Exception): story.LinkToPrevious = False
        self._story_stack.append((self._require_range(), self._paragraph_started, self._active_story))
        self._active_story = f"{'header' if header else 'footer'}:{story_type}"
        target = self._duplicate_range(story.Range)
        with suppress(Exception): target.End = max(target.Start, target.End - 1)
        self.active_range = target
        self._paragraph_started = False

    def _end_story(self) -> None:
        self.active_range, self._paragraph_started, self._active_story = self._story_stack.pop()

    def _event_BeginHeader(self, event: ReconstructionEvent) -> None:
        self._begin_story(str(event.payload.get("story_type", "default")), header=True)

    def _event_EndHeader(self, event: ReconstructionEvent) -> None:
        self._end_story()

    def _event_BeginFooter(self, event: ReconstructionEvent) -> None:
        self._begin_story(str(event.payload.get("story_type", "default")), header=False)

    def _event_EndFooter(self, event: ReconstructionEvent) -> None:
        self._end_story()

    def _event_InsertImage(self, event: ReconstructionEvent) -> None:
        self._image_transaction_active = True
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        if self._asset_resolver is None:
            raise RuntimeError("interactive image asset resolver is not configured")
        asset_path = self._asset_resolver(str(event.payload["asset_id"]))
        inline = self.document.InlineShapes.AddPicture(
            FileName=str(asset_path), LinkToFile=False, SaveWithDocument=True, Range=self._require_range()
        )
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
        if event.payload.get("lock_aspect_ratio") is not None:
            with suppress(Exception): image.LockAspectRatio = self._word_bool(event.payload["lock_aspect_ratio"])
        if event.payload.get("width_emu") is not None:
            image.Width = self._emu_to_points(event.payload["width_emu"])
        if event.payload.get("height_emu") is not None:
            image.Height = self._emu_to_points(event.payload["height_emu"])

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
        table = self.document.Tables.Add(Range=parent_range, NumRows=rows, NumColumns=columns)
        cells = {(r, c): table.Cell(r, c) for r in range(1, rows + 1) for c in range(1, columns + 1)}
        after_range = self._duplicate_range(table.Range)
        with suppress(Exception):
            after_range.Collapse(WD_COLLAPSE_END)
        self._table_stack.append({"table": table, "cells": cells, "parent_range": parent_range, "after_range": after_range, "element_id": event.source_element_id, "structure_complete": False})
        self._paragraph_started = False

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
        table = self._current_table()["table"]
        table.Columns(int(event.payload["column"])).Width = self._twips_to_points(event.payload["width_twips"])

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
            with suppress(Exception): cell.Width = self._twips_to_points(props["width"])
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
        start.Merge(end)

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
        self.document.SaveAs2(str(destination), FileFormat=WD_FORMAT_DOCX)

    def set_custom_property(self, name: str, value: Any) -> None:
        if self.document is None:
            raise RuntimeError("interactive Word document is not open")
        properties = self.document.CustomDocumentProperties
        with suppress(Exception):
            properties(name).Delete()
        # Office MsoDocProperties.msoPropertyTypeString = 4
        properties.Add(name, False, 4, str(value))

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
                self.document.Close(False)
            self.document = None
        if self.application is not None:
            with suppress(Exception):
                self.application.Quit()
            self.application = None
        clear_owned_word(self._owned_word_pid)
        self._owned_word_pid = None
        self.active_range = None
        if self._owns_com:
            with suppress(Exception):
                import pythoncom
                pythoncom.CoUninitialize()
            self._owns_com = False


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

        initial_event = blueprint.events[start_index] if start_index < blueprint.total_events else None
        expected_state = take_snapshot(start_index, initial_event, "initial_state") if initial_event is not None else None
        for index in range(start_index, blueprint.total_events):
            event = blueprint.events[index]
            actual_state = None
            if expected_state is not None and callable(snapshot_fn):
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
            before_state = actual_state if actual_state is not None else take_snapshot(index, event, "before_event")
            started = getattr(observer, "event_started", None) if observer is not None else None
            if started is not None:
                started(index, event, before_state)
            try:
                self.execute_event(event)
            except Exception as exc:
                failed = getattr(observer, "event_failed", None) if observer is not None else None
                if failed is not None:
                    failed(index, event, before_state, exc)
                raise
            expected_state = take_snapshot(index, event, "post_event")
            if observer is not None:
                observer.event_completed(index, event)
                finished = getattr(observer, "event_finished", None)
                if finished is not None:
                    finished(index, event, expected_state)
                halt_status = getattr(observer, "halt_status", None)
                if halt_status:
                    return ExecutionOutcome(str(halt_status), index, min(blueprint.total_events, index + 1))
            character_count = len(str(event.payload.get("text", ""))) if event.event_type == "InsertText" else 1
            self.speed.delay_after(event.event_type, character_count=character_count)
        return ExecutionOutcome.completed(blueprint.total_events)

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
