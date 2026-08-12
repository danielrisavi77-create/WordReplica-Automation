import os
from pathlib import Path

import pytest
from PIL import Image

from word_replica.domain.reconstruction import ReconstructionEvent
from word_replica.parser.parser import DocxParser
from word_replica.renderers.interactive_word import InteractiveWordController

pytestmark = [
    pytest.mark.word,
    pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word integration disabled"),
]


def test_real_word_creates_floating_image_as_new_shape(tmp_path):
    image = tmp_path / "image.png"
    Image.new("RGB", (80, 40), "white").save(image)
    output = tmp_path / "floating.docx"
    controller = InteractiveWordController()
    try:
        controller.open_blank()
        controller.set_asset_resolver(lambda asset_id: image)
        events = [
            ReconstructionEvent("BeginSection", "s1", {"section_index": 0}),
            ReconstructionEvent("BeginParagraph", "p1", {}),
            ReconstructionEvent("InsertCharacter", "r1", {"character": "A"}),
            ReconstructionEvent("InsertImage", "d1", {"asset_id": "a1", "representation": "floating"}),
            ReconstructionEvent("SetImageSize", "d1", {"width_emu": 914400, "height_emu": 457200, "lock_aspect_ratio": True}),
            ReconstructionEvent("SetImageWrap", "d1", {"wrap_type": "square", "behind_text": False}),
            ReconstructionEvent("SetImagePosition", "d1", {
                "horizontal_relative_from": "column", "horizontal_position_emu": 12700,
                "vertical_relative_from": "paragraph", "vertical_position_emu": 25400,
            }),
            ReconstructionEvent("SetImageCrop", "d1", {"crop": {}}),
            ReconstructionEvent("SetImageRotation", "d1", {"rotation_degrees": 0.0}),
            ReconstructionEvent("SetImageZOrder", "d1", {"z_order": 1}),
        ]
        for event in events:
            controller.execute_event(event)
        controller.save(output)
    finally:
        controller.close()
    model = DocxParser().parse(output)
    assert len(model.drawings) == 1
    drawing = model.drawings[0]
    assert drawing.representation == "floating"
    assert drawing.width_emu is not None and drawing.height_emu is not None


def test_real_word_merged_table_cell_text_is_recreated_as_table_content(tmp_path):
    output = tmp_path / "table.docx"
    controller = InteractiveWordController()
    try:
        controller.open_blank()
        events = [
            ReconstructionEvent("BeginSection", "s1", {"section_index": 0}),
            ReconstructionEvent("BeginTable", "t1", {"rows": 2, "columns": 2}),
            ReconstructionEvent("MergeCells", "t1", {"start_row": 1, "start_column": 1, "end_row": 1, "end_column": 2}),
            ReconstructionEvent("EnterCell", "c1", {"row": 1, "column": 1}),
            ReconstructionEvent("BeginParagraph", "p1", {}),
            ReconstructionEvent("InsertCharacter", "r1", {"character": "A"}),
            ReconstructionEvent("InsertCharacter", "r1", {"character": "B"}),
            ReconstructionEvent("LeaveCell", "c1", {}),
            ReconstructionEvent("EnterCell", "c2", {"row": 2, "column": 1}),
            ReconstructionEvent("BeginParagraph", "p2", {}),
            ReconstructionEvent("InsertCharacter", "r2", {"character": "C"}),
            ReconstructionEvent("LeaveCell", "c2", {}),
            ReconstructionEvent("EndTable", "t1", {}),
        ]
        for event in events:
            controller.execute_event(event)
        controller.save(output)
    finally:
        controller.close()
    model = DocxParser().parse(output)
    tables = [block for block in model.body if block.__class__.__name__ == "Table"]
    assert len(tables) == 1
    table = tables[0]
    assert table.rows[0].cells[0].text().strip() == "AB"
    assert table.rows[1].cells[0].text().strip() == "C"
