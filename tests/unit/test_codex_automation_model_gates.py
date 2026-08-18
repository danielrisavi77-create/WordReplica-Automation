from word_replica.domain.model import (
    BinaryAsset, Bookmark, DocumentModel, DrawingRef, Field, Paragraph, Run,
    Section, Table, TableCell, TableRow,
)
from scripts.codex_automation.audit import build_model_gates
from word_replica.qa.structure import l1_projection


def para(eid, text, *, style=None, props=None, run_props=None):
    return Paragraph(eid, [Run(eid + "r", text, properties=run_props or {})], style_id=style, properties=props or {})


def base_model():
    model = DocumentModel("sha")
    model.extras["default_paragraph_style_id"] = "Normal"
    model.extras["style_definitions"] = {
        "Normal": {"style_id": "Normal", "paragraph_properties": {}, "run_properties": {"font_ascii": "Times New Roman", "size_half_points": "22"}}
    }
    model.body = [para("p1", "Hello", style="Normal")]
    model.sections = [Section("s1", {"width": "11906", "height": "16838", "page_number_start": "1"})]
    model.headers = {"default": [para("h1", "Header", style="Normal")]}
    model.footers = {"default": [para("f1", "1", style="Normal")]}
    model.fields = [Field("field1", "PAGE", "1")]
    model.bookmarks = [Bookmark("1", "bm", "body/0", "body/0")]
    asset = BinaryAsset("a1", "word/media/image1.png", "image/png", "assetsha", b"x")
    model.assets = {"a1": asset}
    model.drawings = [DrawingRef("d1", "a1", "body/0/run/0", "inline", width_emu=100, height_emu=200)]
    cell = TableCell("c1", [para("cp", "X", style="Normal")], {"width": "1000", "grid_span": 1})
    model.body.append(Table("t1", [TableRow("r1", [cell], {"repeat_header": True})], {"grid_column_widths": [1000], "layout": "fixed"}))
    return model


def test_identical_models_pass_g0_to_g7():
    source = base_model()
    output = base_model()

    gates = build_model_gates(source, output)

    assert all(gates[f"G{i}"].passed for i in range(8))


def test_table_width_difference_is_first_g3_divergence():
    source = base_model()
    output = base_model()
    output.body[1].properties["grid_column_widths"] = [1200]

    gates = build_model_gates(source, output)

    assert gates["G0"].passed
    assert gates["G1"].passed
    assert gates["G2"].passed
    assert gates["G3"].passed is False
    assert "grid_column_widths" in gates["G3"].first_divergence["path"]


def test_g3_ignores_word_generated_table_grid_when_source_omits_it():
    source = base_model()
    source.body[1].properties.pop("grid_column_widths")
    source.body[1].properties.pop("layout")
    output = base_model()

    gates = build_model_gates(source, output)

    assert gates["G3"].passed


def test_g3_ignores_word_generated_cell_width_when_source_omits_it():
    source = base_model()
    source.body[1].properties.pop("grid_column_widths")
    source.body[1].properties.pop("layout")
    source.body[1].rows[0].cells[0].properties.pop("width", None)
    source.body[1].rows[0].cells[0].properties.pop("width_type", None)
    output = base_model()
    output.body[1].rows[0].cells[0].properties.update({"width": "1000", "width_type": "dxa"})

    gates = build_model_gates(source, output)

    assert gates["G3"].passed


def test_header_typography_difference_is_g6_failure():
    source = base_model()
    output = base_model()
    output.headers["default"][0].runs[0].properties["italic"] = True

    gates = build_model_gates(source, output)

    assert gates["G6"].passed is False
    assert "headers" in gates["G6"].first_divergence["path"]


def test_field_instruction_difference_is_g7_failure():
    source = base_model()
    output = base_model()
    output.fields[0].instruction = "NUMPAGES"

    gates = build_model_gates(source, output)

    assert gates["G7"].passed is False
    assert "fields" in gates["G7"].first_divergence["path"]


def test_l1_ignores_unreferenced_media_assets():
    source = base_model()
    source.drawings = []
    output = base_model()
    output.drawings = []
    output.assets = {}

    assert l1_projection(source)["asset_hashes"] == []
    assert l1_projection(source) == l1_projection(output)
