from __future__ import annotations

from typing import Any

from word_replica.domain.model import DocumentModel, Paragraph, Run


PARAGRAPH_KEYS = (
    "keepNext", "pageBreakBefore", "keepLines", "widowControl", "alignment",
    "spacing_before", "spacing_after", "spacing_line", "spacing_line_rule",
    "indent_left", "indent_right", "indent_first_line", "indent_hanging",
    "numId", "ilvl",
)
RUN_KEYS = (
    "bold", "italic", "underline", "strike", "hidden",
    "font_ascii", "font_hansi", "font_east_asia", "font_cs",
    "size_half_points", "color", "highlight", "vert_align",
    "language", "language_east_asia", "language_bidi",
    "character_spacing", "character_position",
)


def _style_chain(model: DocumentModel, style_id: str | None) -> list[dict[str, Any]]:
    if not style_id:
        return []
    definitions = model.extras.get("style_definitions", {})
    chain: list[dict[str, Any]] = []
    seen: set[str] = set()
    current = definitions.get(style_id)
    while current:
        current_id = str(current.get("style_id", style_id))
        if current_id in seen:
            break
        seen.add(current_id)
        chain.append(current)
        based_on = current.get("based_on")
        current = definitions.get(str(based_on)) if based_on else None
    return list(reversed(chain))


def _effective_paragraph_properties(model: DocumentModel, paragraph: Paragraph) -> dict[str, Any]:
    defaults = model.extras.get("document_defaults", {})
    props: dict[str, Any] = {
        "spacing_before": "0",
        "spacing_after": "0",
        "spacing_line_rule": "single",
    }
    props.update(defaults.get("paragraph_properties", {}) or {})
    for definition in _style_chain(model, paragraph.style_id):
        props.update(definition.get("paragraph_properties", {}) or {})
    props.update(paragraph.properties)
    return {key: props.get(key) for key in PARAGRAPH_KEYS if key in props}


def _effective_run_properties(model: DocumentModel, paragraph: Paragraph, run: Run) -> dict[str, Any]:
    defaults = model.extras.get("document_defaults", {})
    props: dict[str, Any] = {
        "bold": False,
        "italic": False,
        "underline": False,
        "strike": False,
        "hidden": False,
        "vert_align": "baseline",
        "character_spacing": "0",
        "character_position": "0",
    }
    props.update(defaults.get("run_properties", {}) or {})
    for definition in _style_chain(model, paragraph.style_id):
        props.update(definition.get("run_properties", {}) or {})
    props.update({k: v for k, v in run.properties.items() if k not in {"content_tokens", "break_types"}})
    if run.hidden:
        props["hidden"] = True
    return {key: props.get(key) for key in RUN_KEYS if key in props}


def normalize_formatting(model: DocumentModel, paragraph: Paragraph) -> dict:
    return {
        "style_id": paragraph.style_id,
        "paragraph": _effective_paragraph_properties(model, paragraph),
        "runs": [
            {
                **_effective_run_properties(model, paragraph, run),
                "text_len": len(run.text),
            }
            for run in paragraph.runs
        ],
    }


def l2_projection(model: DocumentModel) -> list[dict]:
    return [normalize_formatting(model, paragraph) for paragraph in model.iter_paragraphs()]
