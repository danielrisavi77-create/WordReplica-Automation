import os
from pathlib import Path

import pytest

from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.parser.parser import DocxParser


GOLDEN_ENV = "WORD_REPLICA_GOLDEN_DOCX"


def _golden_path() -> Path:
    value = os.environ.get(GOLDEN_ENV)
    if not value:
        pytest.skip(f"{GOLDEN_ENV} is not set")
    path = Path(value)
    if not path.exists():
        pytest.skip(f"Golden DOCX does not exist: {path}")
    return path


def test_golden_rektorova_source_and_blueprint_preserve_confirmed_fidelity_invariants():
    model = DocxParser().parse(_golden_path())
    blueprint = BlueprintCompiler().compile(model)

    assert model.extras["default_paragraph_style_id"] == "Normal"
    normal = model.extras["style_definitions"]["Normal"]
    assert normal["run_properties"]["font_ascii"] == "Times New Roman"
    assert normal["paragraph_properties"]["alignment"] == "both"
    assert normal["paragraph_properties"]["spacing_after"] == "120"
    assert normal["paragraph_properties"]["spacing_line"] == "360"
    assert normal["paragraph_properties"]["spacing_line_rule"] == "auto"

    first_cell = next(i for i, event in enumerate(blueprint.events) if event.event_type == "EnterCell")
    cell_p = next(event for event in blueprint.events[first_cell:] if event.event_type == "ApplyParagraphProperties")
    cell_r = next(event for event in blueprint.events[first_cell:] if event.event_type == "ApplyRunProperties")
    assert cell_p.payload["style_id"] == "Normal"
    assert cell_p.payload["spacing_after"] == "0"
    assert cell_p.payload["spacing_line"] == "240"
    assert cell_r.payload["font_ascii"] == "Times New Roman"
    assert cell_r.payload["size_half_points"] == "20"

    header = next(i for i, event in enumerate(blueprint.events) if event.event_type == "BeginHeader")
    header_p = next(event for event in blueprint.events[header:] if event.event_type == "ApplyParagraphProperties")
    header_r = next(event for event in blueprint.events[header:] if event.event_type == "ApplyRunProperties")
    assert header_p.payload["style_id"] == "Header"
    assert header_r.payload["font_ascii"] == "Times New Roman"
    assert header_r.payload["italic"] is True
    assert header_r.payload["size_half_points"] == "17"
    assert header_r.payload["color"] == "5C5C5A"

    section_events = [event for event in blueprint.events if event.event_type == "ApplySectionProperties"]
    assert section_events[1].payload["page_number_start"] == "1"

    ref_field = next(
        event for event in blueprint.events
        if event.event_type == "CreateField" and "ref_tab_1" in event.payload.get("instruction", "").lower()
    )
    assert ref_field.payload["cached_result"] == "1"

    assert all(
        event.payload.get("style_id")
        for event in blueprint.events
        if event.event_type == "ApplyParagraphProperties"
    )
