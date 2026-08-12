from __future__ import annotations

import sys
from collections.abc import Callable

from word_replica.config import InteractiveOptions
from word_replica.domain.enums import CapabilityClass, InteractiveFidelity
from word_replica.domain.model import DocumentModel
from word_replica.domain.reconstruction import CapabilityDecision, PreflightReport, ReconstructionBlueprint
from word_replica.interactive.capabilities import classify_drawing


def _default_word_probe() -> bool:
    return sys.platform == "win32"


def analyze_preflight(
    model: DocumentModel,
    blueprint: ReconstructionBlueprint,
    options: InteractiveOptions,
    *,
    word_probe: Callable[[], bool] = _default_word_probe,
    font_probe: Callable[[str], bool] | None = None,
    asset_probe: Callable[[object], bool] | None = None,
) -> PreflightReport:
    word_ok = bool(word_probe())
    blocking: list[str] = []
    warnings: list[str] = []
    if not word_ok:
        blocking.append("Microsoft Word desktop automation is unavailable")

    missing_assets: list[str] = []
    for drawing in model.drawings:
        asset = model.assets.get(drawing.asset_id)
        if asset is None or not asset.bytes_data or (asset_probe is not None and not asset_probe(asset)):
            missing_assets.append(drawing.asset_id)
    if missing_assets:
        blocking.append("Missing required image asset(s): " + ", ".join(sorted(set(missing_assets))))

    fonts_set = {
        str(run.properties[key])
        for paragraph in model.iter_paragraphs()
        for run in paragraph.runs
        for key in ("font_ascii", "font_hansi", "font_east_asia", "font_cs")
        if run.properties.get(key)
    }
    style_definitions = model.extras.get("style_definitions", {})
    used_style_ids = {paragraph.style_id for paragraph in model.iter_paragraphs() if paragraph.style_id}
    pending = list(used_style_ids)
    visited: set[str] = set()
    while pending:
        style_id = str(pending.pop())
        if style_id in visited:
            continue
        visited.add(style_id)
        definition = style_definitions.get(style_id) or {}
        run_props = definition.get("run_properties") or {}
        for key in ("font_ascii", "font_hansi", "font_east_asia", "font_cs"):
            if run_props.get(key):
                fonts_set.add(str(run_props[key]))
        based_on = definition.get("based_on")
        if based_on:
            pending.append(str(based_on))
    fonts = sorted(fonts_set)
    missing_fonts = [font for font in fonts if font_probe is not None and not font_probe(font)]
    if missing_fonts:
        message = "Missing required font(s): " + ", ".join(missing_fonts)
        if options.fidelity is InteractiveFidelity.MAXIMUM:
            blocking.append(message)
        else:
            warnings.append(message)

    capability_items: list[CapabilityDecision] = [classify_drawing(drawing, asset=model.assets.get(drawing.asset_id)) for drawing in model.drawings]
    # Interactive mode must never claim that a complex OOXML part was preserved
    # unless there is a concrete preservation path in the executor. Today these
    # parts are not re-injected into the newly-authored Word document, so classify
    # them honestly as unsupported rather than silently dropping them.
    for part_name in sorted(model.preserved_parts):
        capability_items.append(CapabilityDecision(
            CapabilityClass.UNSUPPORTED,
            f"Complex OOXML part is not yet reconstructable or preservable in Interactive mode: {part_name}",
            part_name,
        ))
    for comment_id in sorted(model.comments):
        capability_items.append(CapabilityDecision(
            CapabilityClass.UNSUPPORTED,
            f"Word comment is not yet reconstructable in Interactive mode: {comment_id}",
            f"comment:{comment_id}",
        ))
    for revision in model.revisions:
        capability_items.append(CapabilityDecision(
            CapabilityClass.UNSUPPORTED,
            f"Tracked revision is not yet reconstructable in Interactive mode: {revision.revision_id}",
            f"revision:{revision.revision_id}",
        ))

    unsupported = [item for item in capability_items if item.classification is CapabilityClass.UNSUPPORTED]
    preserved = [item for item in capability_items if item.classification is CapabilityClass.PRESERVED]
    if unsupported:
        message = "; ".join(item.reason for item in unsupported)
        if options.fidelity is InteractiveFidelity.MAXIMUM and options.block_on_unsupported:
            blocking.append(message)
        else:
            warnings.append(message)
    if preserved:
        message = "; ".join(item.reason for item in preserved)
        if options.fidelity is InteractiveFidelity.MAXIMUM and not options.allow_preserved_objects:
            blocking.append("Preservation requires explicit permission: " + message)
        else:
            warnings.append("Preserved rather than reconstructed: " + message)

    counts = {
        "events": blueprint.total_events,
        "characters": blueprint.total_visible_characters,
        "paragraphs": blueprint.semantic_counts.get("paragraphs", 0),
        "tables": blueprint.semantic_counts.get("tables", 0),
        "images": len(model.drawings),
        "sections": len(model.sections),
        "fields": len(model.fields),
        "footnotes": len(model.footnotes),
        "endnotes": len(model.endnotes),
    }
    ready = not blocking
    return PreflightReport(
        source_sha256=model.source_sha256,
        blueprint_fingerprint=blueprint.fingerprint,
        blueprint_schema_version=blueprint.schema_version,
        word_available=word_ok,
        missing_assets=tuple(sorted(set(missing_assets))),
        missing_fonts=tuple(missing_fonts),
        counts=counts,
        capability_items=tuple(capability_items),
        maximum_fidelity_ready=ready if options.fidelity is InteractiveFidelity.MAXIMUM else not any("Microsoft Word" in r or "asset" in r.lower() for r in blocking),
        can_proceed=ready,
        blocking_reasons=tuple(blocking),
        warnings=tuple(warnings),
    )
