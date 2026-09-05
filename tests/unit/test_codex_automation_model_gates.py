from word_replica.domain.model import (
    BinaryAsset, Bookmark, DocumentModel, DrawingRef, Field, Paragraph,
    RelationshipRef, Run, Section, Table, TableCell, TableRow,
)
from scripts.codex_automation.audit import build_model_gates
from word_replica.qa.golden_audit import _header_footer_projection, _semantic_projection
from word_replica.qa.policy import compare_projection
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


def test_g7_ignores_mergeformat_switch_word_adds_unconditionally_on_save():
    # \* MERGEFORMAT is a cosmetic "preserve formatting on update" switch, not
    # part of a field's identity or result - Word appends it (even a second
    # time, if source already had one) whenever it serializes a field as
    # fldSimple on save, regardless of what source specified.
    source = base_model()
    output = base_model()
    output.fields[0].instruction = "PAGE \\* MERGEFORMAT"

    gates = build_model_gates(source, output)

    assert gates["G7"].passed


def test_g7_matches_fields_by_content_not_list_position():
    # Regression: a multi-paragraph field (e.g. a TOC) can open and close
    # across paragraphs the interactive renderer can never fully replay (see
    # BlueprintCompiler._paragraph_closing_field_begins), so source can have
    # a field with no counterpart in output at all. Comparing by list index
    # used to shift every same-index comparison after the missing entry,
    # burying the one real finding under noise for every field that follows
    # it. Matching by content instead means only the missing field itself
    # shows up as a finding.
    source = base_model()
    source.fields = [
        Field("f1", 'TOC \\o "1-3"', "1. Uvod\t1", False),
        Field("f2", "REF _Ref_tab1 \\h", "1", False),
        Field("f3", "REF _Ref_tab2 \\h", "2", False),
    ]
    output = base_model()
    output.fields = [
        Field("f2", "REF _Ref_tab1 \\h", "1", False),
        Field("f3", "REF _Ref_tab2 \\h", "2", False),
    ]

    findings = compare_projection("G7", _semantic_projection(source), _semantic_projection(output))

    assert len(findings) == 1
    assert "TOC" in str(findings[0].expected)


def test_g6_matches_footers_by_section_not_by_swapped_part_name():
    # Regression: physical footer part numbering (footer1.xml, footer2.xml, ...)
    # is independent of section order and a rebuild can freely renumber parts
    # (Word materializes its own even/first/default variants on save) without
    # changing which section a story belongs to. Confirmed live: this alone
    # accounted for a golden document's entire G6 footer mismatch even though
    # both packages' actual footer content was identical - source's section 0
    # footer lived in footer1.xml, but the rebuilt output's semantically same
    # footer ended up in a different physically-numbered part.
    def model_with_swapped_footer_names(swap: bool) -> DocumentModel:
        model = DocumentModel("sha")
        model.sections = [
            Section("s0", {"footer_refs": [{"type": "default", "rel_id": "rId1"}]}),
            Section("s1", {"footer_refs": [{"type": "default", "rel_id": "rId2"}]}),
        ]
        first_name, second_name = (
            ("word/footer2.xml", "word/footer1.xml") if swap else ("word/footer1.xml", "word/footer2.xml")
        )
        model.relationships = {
            "word/document.xml:rId1": RelationshipRef("rId1", "footer", first_name),
            "word/document.xml:rId2": RelationshipRef("rId2", "footer", second_name),
        }
        model.footers = {
            first_name: [para("f0", "Section zero footer")],
            second_name: [para("f1", "Section one footer")],
        }
        return model

    source = model_with_swapped_footer_names(swap=False)
    output = model_with_swapped_footer_names(swap=True)

    findings = compare_projection("G6", _header_footer_projection(source), _header_footer_projection(output))
    assert findings == []


def test_l1_ignores_unreferenced_media_assets():
    source = base_model()
    source.drawings = []
    output = base_model()
    output.drawings = []
    output.assets = {}

    assert l1_projection(source)["asset_hashes"] == []
    assert l1_projection(source) == l1_projection(output)


def test_l1_counts_referenced_footer_slots_not_word_materialized_parts():
    def model_with_footers(*, shared: bool) -> DocumentModel:
        model = DocumentModel("sha")
        model.sections = [
            Section("s0", {"footer_refs": [{"type": "default", "rel_id": "rId1"}]}),
            Section("s1", {"footer_refs": [{"type": "default", "rel_id": "rId2"}]}),
        ]
        if shared:
            targets = ("word/footer1.xml", "word/footer1.xml")
            model.footers = {
                "word/footer1.xml": [],
                "word/footer2.xml": [],
                "word/footer3.xml": [],
            }
        else:
            targets = ("word/footer2.xml", "word/footer4.xml")
            model.footers = {
                "word/footer1.xml": [],
                "word/footer2.xml": [],
                "word/footer3.xml": [],
                "word/footer4.xml": [],
            }
        model.relationships = {
            "word/document.xml:rId1": RelationshipRef("rId1", "footer", targets[0]),
            "word/document.xml:rId2": RelationshipRef("rId2", "footer", targets[1]),
        }
        return model

    source = model_with_footers(shared=True)
    output = model_with_footers(shared=False)

    assert l1_projection(source)["footers"] == 2
    assert l1_projection(output)["footers"] == 2
    assert l1_projection(source) == l1_projection(output)
