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
FONT_PROPERTY_PAIRS = (
    ("font_ascii", "font_ascii_theme"),
    ("font_hansi", "font_hansi_theme"),
    ("font_east_asia", "font_east_asia_theme"),
    ("font_cs", "font_cs_theme"),
)


def _merge_run_properties(
    properties: dict[str, Any], overrides: dict[str, Any]
) -> None:
    for direct_key, theme_key in FONT_PROPERTY_PAIRS:
        if theme_key in overrides:
            properties.pop(direct_key, None)
        elif direct_key in overrides:
            properties.pop(theme_key, None)
    properties.update(overrides)
    for direct_key, theme_key in FONT_PROPERTY_PAIRS:
        if theme_key in overrides:
            properties.pop(direct_key, None)


def _uses_east_asia_font(text: str) -> bool:
    return any(
        "\u2e80" <= character <= "\u9fff"
        or "\u3040" <= character <= "\u30ff"
        or "\uac00" <= character <= "\ud7af"
        or "\uf900" <= character <= "\ufaff"
        for character in text
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
    style_id = paragraph.style_id or model.extras.get("default_paragraph_style_id")
    for definition in _style_chain(model, style_id):
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
    _merge_run_properties(props, defaults.get("run_properties", {}) or {})
    style_id = paragraph.style_id or model.extras.get("default_paragraph_style_id")
    for definition in _style_chain(model, style_id):
        _merge_run_properties(props, definition.get("run_properties", {}) or {})
    _merge_run_properties(
        props,
        {
            key: value
            for key, value in run.properties.items()
            if key not in {"content_tokens", "break_types"}
        },
    )
    scheme = model.extras.get("theme_font_scheme", {}) or {}
    for font_key, theme_key in FONT_PROPERTY_PAIRS:
        if not props.get(font_key) and props.get(theme_key) in scheme:
            props[font_key] = scheme[props[theme_key]]
    if run.hidden:
        props["hidden"] = True
    return {key: props.get(key) for key in RUN_KEYS if key in props}


def normalize_formatting(model: DocumentModel, paragraph: Paragraph) -> dict:
    runs = []
    for run in paragraph.runs:
        properties = _effective_run_properties(model, paragraph, run)
        if not _uses_east_asia_font(run.text):
            properties.pop("font_east_asia", None)
        if runs and all(runs[-1].get(key) == value for key, value in properties.items()) and all(
            key == "text_len" or key in properties for key in runs[-1]
        ):
            runs[-1]["text_len"] += len(run.text)
        else:
            runs.append({**properties, "text_len": len(run.text)})
    return {
        "style_id": paragraph.style_id,
        "paragraph": _effective_paragraph_properties(model, paragraph),
        "runs": runs,
    }


def l2_projection(model: DocumentModel) -> list[dict]:
    return [normalize_formatting(model, paragraph) for paragraph in model.iter_paragraphs()]
