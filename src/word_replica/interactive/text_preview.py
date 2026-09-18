from __future__ import annotations

from collections.abc import Callable
from typing import Any

from word_replica.config import InteractiveOptions
from word_replica.domain.reconstruction import ExecutionOutcome, ReconstructionBlueprint, ReconstructionEvent
from word_replica.interactive.speed import SpeedController


_PLACEHOLDER_EVENTS = {
    "InsertImage": "[slika]",
    "CreateFootnote": "[fusnota]",
    "CreateEndnote": "[endnota]",
}


class TextPreviewPlayer:
    """Replays a blueprint into a plain-text preview sink instead of Microsoft Word.

    This exists so the visual "watch it get written" experience does not depend
    on a licensed, activated Word install: it only needs the same canonical
    blueprint the Word renderer already compiles, and paces itself with the
    same SpeedController used for Interactive Reconstruction. It does not
    write the .docx itself and has no bearing on the real output's fidelity;
    the actual file still comes from the normal Pure DOCX rebuild. Structures
    the preview can't represent (tables, images, notes) show as short bracketed
    placeholders rather than being silently dropped, so the preview stays
    honest about what it can't show.
    """

    def __init__(
        self,
        options: InteractiveOptions,
        sink: Callable[[str, dict[str, Any]], None],
        *,
        speed: SpeedController | None = None,
    ) -> None:
        self.options = options
        self._sink = sink
        self.speed = speed or SpeedController(options)
        self._bold = False
        self._italic = False
        self._underline = False
        self._strike = False

    def play(self, blueprint: ReconstructionBlueprint, control) -> ExecutionOutcome:
        for index in range(blueprint.total_events):
            decision = control.before_next_event(index - 1)
            if decision.stop_requested:
                return ExecutionOutcome.stopped(index - 1)
            self._play_event(blueprint.events[index])
        return ExecutionOutcome.completed(blueprint.total_events)

    def _play_event(self, event: ReconstructionEvent) -> None:
        event_type = event.event_type
        if event_type == "InsertText" and self.options.letter_by_letter:
            for character in str(event.payload.get("text", "")):
                self._emit_text(character)
                self.speed.delay_after("InsertCharacter", character_count=1)
            return
        if event_type == "InsertText":
            text = str(event.payload.get("text", ""))
            self._emit_text(text)
            self.speed.delay_after(event_type, character_count=len(text))
            return
        if event_type == "InsertCharacter":
            self._emit_text(str(event.payload.get("character", "")))
            self.speed.delay_after(event_type, character_count=1)
            return
        if event_type == "InsertTab":
            self._emit_text("\t")
            return
        if event_type in ("BeginParagraph", "InsertLineBreak"):
            self._sink("newline", {})
            return
        if event_type == "InsertPageBreak":
            self._sink("pagebreak", {})
            self.speed.delay_after(event_type)
            return
        if event_type == "ApplyRunProperties":
            props = event.payload
            if "bold" in props: self._bold = bool(props["bold"])
            if "italic" in props: self._italic = bool(props["italic"])
            if "underline" in props: self._underline = bool(props["underline"])
            if "strike" in props: self._strike = bool(props["strike"])
            return
        if event_type == "InsertTableBatch":
            rows = event.payload.get("rows", "?")
            columns = event.payload.get("columns", "?")
            self._sink("placeholder", {"label": f"[tablica {rows}×{columns}]"})
            self.speed.delay_after(event_type)
            return
        if event_type == "BeginTable":
            self._sink("placeholder", {"label": "[tablica]"})
            self.speed.delay_after(event_type)
            return
        if event_type in _PLACEHOLDER_EVENTS:
            self._sink("placeholder", {"label": _PLACEHOLDER_EVENTS[event_type]})
            return
        if event_type == "BeginSection":
            self._sink("sectionbreak", {})
            self.speed.delay_after(event_type)
            return
        # Remaining structural/property events (styles, sections properties,
        # bookmarks, fields, headers/footers, checkpoints) have no direct
        # plain-text analogue and are intentionally silent in the preview.

    def _emit_text(self, text: str) -> None:
        if not text:
            return
        self._sink("text", {
            "text": text,
            "bold": self._bold,
            "italic": self._italic,
            "underline": self._underline,
            "strike": self._strike,
        })
