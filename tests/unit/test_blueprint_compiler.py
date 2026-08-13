from lxml import etree

from word_replica.domain.model import DocumentModel, Paragraph, Run
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.domain.model import ElementIdFactory
from word_replica.parser.parser import parse_run
from word_replica.parser.text import W_NS


def test_compiler_legacy_character_assertion_is_superseded_by_text_batching():
    model = DocumentModel(
        source_sha256="a" * 64,
        body=[Paragraph("p1", [Run("r1", "Až B", {"bold": True})])],
    )
    blueprint = BlueprintCompiler().compile(model)
    text_events = [event for event in blueprint.events if event.event_type == "InsertText"]
    assert [event.payload["text"] for event in text_events] == [model.body[0].runs[0].text]
    assert not any(event.event_type == "InsertCharacter" for event in blueprint.events)


def test_compiler_batches_contiguous_text_within_a_run_for_word_insertion():
    model = DocumentModel(
        source_sha256="a" * 64,
        body=[Paragraph("p1", [Run("r1", "AÅ¾ B", {"bold": True})])],
    )
    events = BlueprintCompiler().compile(model).events
    assert [(event.event_type, event.payload) for event in events if event.source_element_id == "r1"] == [
        ("ApplyRunProperties", {"bold": True}),
        ("InsertText", {"text": "AÅ¾ B"}),
    ]


def test_run_properties_precede_first_character_of_each_run():
    model = DocumentModel(
        source_sha256="a" * 64,
        body=[Paragraph(
            "p1",
            runs=[
                Run("r1", "A", {"bold": False}),
                Run("r2", "B", {"bold": True}),
            ],
        )],
    )
    events = BlueprintCompiler().compile(model).events
    trace = [
        (event.event_type, event.source_element_id, event.payload)
        for event in events
        if event.event_type in {"ApplyRunProperties", "InsertText"}
    ]
    assert trace == [
        ("ApplyRunProperties", "r1", {"bold": False}),
        ("InsertText", "r1", {"text": "A"}),
        ("ApplyRunProperties", "r2", {"bold": True}),
        ("InsertText", "r2", {"text": "B"}),
    ]


def test_tabs_and_breaks_use_semantic_events_in_source_order():
    run = Run(
        "r1",
        text="A\tB\n\n",
        properties={
            "content_tokens": [
                {"kind": "text", "value": "A"},
                {"kind": "tab"},
                {"kind": "text", "value": "B"},
                {"kind": "line_break"},
                {"kind": "page_break"},
            ]
        },
    )
    model = DocumentModel(source_sha256="a" * 64, body=[Paragraph("p1", [run])])
    events = BlueprintCompiler().compile(model).events
    trace = [
        (event.event_type, event.payload)
        for event in events
        if event.event_type.startswith("Insert")
    ]
    assert trace == [
        ("InsertText", {"text": "A"}),
        ("InsertTab", {}),
        ("InsertText", {"text": "B"}),
        ("InsertLineBreak", {}),
        ("InsertPageBreak", {}),
    ]


def test_parse_run_preserves_exact_token_order_and_extended_properties():
    xml = f'''<w:r xmlns:w="{W_NS}">
      <w:rPr>
        <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
        <w:sz w:val="24"/><w:strike/><w:color w:val="FF0000"/>
        <w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/>
        <w:lang w:val="hr-HR"/><w:spacing w:val="10"/><w:position w:val="2"/>
      </w:rPr>
      <w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:br w:type="page"/>
    </w:r>'''
    node = etree.fromstring(xml.encode())
    run = parse_run(node, ElementIdFactory("a" * 64), "p/r")
    assert run.properties["content_tokens"] == [
        {"kind": "text", "value": "A"}, {"kind": "tab"},
        {"kind": "text", "value": "B"}, {"kind": "line_break"},
        {"kind": "page_break"},
    ]
    assert run.properties["font_ascii"] == "Aptos"
    assert run.properties["size_half_points"] == "24"
    assert run.properties["strike"] is True
    assert run.properties["color"] == "FF0000"
    assert run.properties["highlight"] == "yellow"
    assert run.properties["vert_align"] == "superscript"
    assert run.properties["language"] == "hr-HR"
    assert run.properties["character_spacing"] == "10"
    assert run.properties["character_position"] == "2"


def test_drawing_token_emits_insert_then_geometry_events():
    from word_replica.domain.model import BinaryAsset, DrawingRef
    run = Run("r1", "", {"content_tokens": [{"kind":"drawing", "relationship_id":"rId5"}]})
    model = DocumentModel(source_sha256="a"*64, body=[Paragraph("p1", [run])])
    model.drawings.append(DrawingRef(
        element_id="d1", asset_id="asset1", source_path="word/media/image1.png",
        representation="floating", width_emu=914400, height_emu=457200, lock_aspect_ratio=True,
        wrap_type="square", horizontal_relative_from="column", horizontal_position_emu=100,
        vertical_relative_from="paragraph", vertical_position_emu=200, distance_top_emu=0,
        distance_bottom_emu=0, distance_left_emu=0, distance_right_emu=0, crop={},
        rotation_degrees=0.0, behind_text=False, z_order=5,
    ))
    model.extras["drawing_relationship_index"] = {"rId5": ["d1"]}
    events = BlueprintCompiler().compile(model).events
    trace = [e.event_type for e in events if "Image" in e.event_type]
    assert trace == ["InsertImage", "SetImageSize", "SetImageWrap", "SetImagePosition", "SetImageCrop", "SetImageRotation", "SetImageZOrder"]


def test_section_transition_occurs_at_canonical_boundary_not_before_body(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_core_fixture
    model = DocxParser().parse(build_core_fixture(tmp_path / "sections.docx"))
    events = BlueprintCompiler().compile(model).events
    begin_sections = [i for i,e in enumerate(events) if e.event_type == "BeginSection"]
    assert len(begin_sections) == 2
    first_text = next(i for i,e in enumerate(events) if e.event_type == "InsertText")
    assert begin_sections[0] < first_text
    # second section is not pre-created; it follows the first section's content/boundary.
    assert begin_sections[1] > first_text
    assert [e.payload["orientation"] for e in events if e.event_type == "ApplySectionProperties"][-1] == "landscape"


def test_empty_section_boundary_paragraph_is_preserved_before_section_transition():
    from word_replica.domain.model import Section

    model = DocumentModel(
        source_sha256="a" * 64,
        body=[
            Paragraph("before", [Run("before-run", "Before")]),
            Paragraph("section-boundary", properties={"section_index": 0}),
            Paragraph("after", [Run("after-run", "After")]),
        ],
        sections=[Section("section-0"), Section("section-1", {"break_type": "nextPage"})],
    )

    events = BlueprintCompiler().compile(model).events
    boundary_start = next(
        index for index, event in enumerate(events)
        if event.event_type == "BeginParagraph" and event.source_element_id == "section-boundary"
    )
    boundary_end = next(
        index for index, event in enumerate(events)
        if event.event_type == "EndParagraph" and event.source_element_id == "section-boundary"
    )
    section_end = next(index for index, event in enumerate(events) if event.event_type == "EndSection")

    assert boundary_start < boundary_end < section_end


def test_header_footer_stories_use_same_character_events(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_extended_fixture
    model = DocxParser().parse(build_extended_fixture(tmp_path / "stories.docx"))
    events = BlueprintCompiler().compile(model).events
    assert any(e.event_type == "BeginHeader" for e in events)
    assert any(e.event_type == "BeginFooter" for e in events)
    header_start = next(i for i,e in enumerate(events) if e.event_type == "BeginHeader")
    header_end = next(i for i,e in enumerate(events) if e.event_type == "EndHeader")
    header_text = [e.payload["text"] for e in events[header_start:header_end] if e.event_type == "InsertText"]
    assert "".join(header_text) == "Header"


def test_list_paragraph_emits_list_binding_without_literal_marker(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_lists
    model = DocxParser().parse(build_lists(tmp_path / "lists.docx"))
    events = BlueprintCompiler().compile(model).events
    assert sum(1 for e in events if e.event_type == "CreateListBinding") == 5
    chars = "".join(e.payload["text"] for e in events if e.event_type == "InsertText")
    assert "First" in chars and "Alpha" in chars
    assert not chars.startswith("1.")


def test_notes_compile_at_reference_with_note_story_characters(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_footnotes_endnotes
    model = DocxParser().parse(build_footnotes_endnotes(tmp_path / "notes.docx"))
    events = BlueprintCompiler().compile(model).events
    assert any(e.event_type == "CreateFootnote" and e.payload["note_id"] == "1" for e in events)
    assert any(e.event_type == "CreateEndnote" and e.payload["note_id"] == "1" for e in events)
    f0 = next(i for i,e in enumerate(events) if e.event_type == "BeginFootnoteStory")
    f1 = next(i for i,e in enumerate(events) if e.event_type == "EndFootnoteStory")
    assert "".join(e.payload["text"] for e in events[f0:f1] if e.event_type == "InsertText") == "Footnote evidence"


def test_field_result_is_not_typed_as_static_text_when_semantic_field_is_compiled(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_toc_fields
    model = DocxParser().parse(build_toc_fields(tmp_path / "fields.docx"))
    events = BlueprintCompiler().compile(model).events
    fields = [e for e in events if e.event_type == "CreateField"]
    assert any("TOC" in e.payload["instruction"] for e in fields)
    # Result text comes from Word field update, not a static character replay.
    chars = "".join(e.payload["text"] for e in events if e.event_type == "InsertText")
    assert "Chapter One .... 1" not in chars


def test_bookmark_anchor_compiles_start_and_real_create_event(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_bookmarks_crossrefs
    model = DocxParser().parse(build_bookmarks_crossrefs(tmp_path / "bookmarks.docx"))
    events = BlueprintCompiler().compile(model).events
    assert any(e.event_type == "BookmarkStart" and e.payload["name"] == "TargetBookmark" for e in events)
    assert any(e.event_type == "CreateBookmark" and e.payload["name"] == "TargetBookmark" for e in events)
    assert any(e.event_type == "CreateField" and "REF TargetBookmark" in e.payload["instruction"] for e in events)


def test_blueprint_carries_source_style_definition_for_custom_paragraph(corpus_dir):
    from word_replica.parser.parser import DocxParser
    from word_replica.interactive.blueprint import BlueprintCompiler

    model = DocxParser().parse(corpus_dir / "02_headings_styles.docx")
    events = BlueprintCompiler().compile(model).events
    styled = next(
        event for event in events
        if event.event_type == "ApplyParagraphProperties" and event.payload.get("style_id") == "ReplicaBody"
    )
    assert styled.payload["style_definition"]["name"] == "ReplicaBody"
    assert styled.payload["style_definition"]["run_properties"]["font_ascii"] == "Arial"


def test_semantic_field_preserves_source_cached_result_without_typing_it(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_toc_fields
    model = DocxParser().parse(build_toc_fields(tmp_path / "fields_cached.docx"))
    field = next(e for e in BlueprintCompiler().compile(model).events if e.event_type == "CreateField" and "TOC" in e.payload["instruction"])
    assert field.payload["cached_result"] == "Chapter One .... 1"


def test_parsed_document_defaults_are_resolved_into_every_run_and_reset_sticky_formatting():
    model = DocumentModel(
        source_sha256="a" * 64,
        body=[Paragraph("p1", [Run("r1", "A", {"bold": True}), Run("r2", "B", {})], style_id="Heading1")],
    )
    model.extras["document_defaults"] = {
        "run_properties": {"font_ascii": "Times New Roman", "font_hansi": "Times New Roman", "size_half_points": "24"},
        "paragraph_properties": {},
    }
    model.extras["style_definitions"] = {
        "Heading1": {"style_id": "Heading1", "name": "heading 1", "type": "paragraph", "run_properties": {"size_half_points": "28"}, "paragraph_properties": {}}
    }
    events = BlueprintCompiler().compile(model).events
    assert not any(e.event_type == "ApplyDocumentDefaults" for e in events)
    runs = [e for e in events if e.event_type == "ApplyRunProperties"]
    assert runs[0].payload["font_ascii"] == "Times New Roman"
    assert runs[0].payload["size_half_points"] == "28"
    assert runs[0].payload["bold"] is True
    assert runs[1].payload["font_ascii"] == "Times New Roman"
    assert runs[1].payload["size_half_points"] == "28"
    assert runs[1].payload["bold"] is False
    paragraph = next(e for e in events if e.event_type == "ApplyParagraphProperties")
    assert "spacing_before" not in paragraph.payload
    assert "spacing_after" not in paragraph.payload
    assert "spacing_line_rule" not in paragraph.payload


def test_theme_font_references_resolve_to_concrete_font_names_before_word_events():
    model = DocumentModel(
        source_sha256="b" * 64,
        body=[Paragraph("p1", [Run("r1", "A", {})])],
    )
    model.extras["document_defaults"] = {
        "run_properties": {
            "font_ascii_theme": "minorHAnsi",
            "font_hansi_theme": "minorHAnsi",
            "font_east_asia_theme": "minorEastAsia",
            "font_cs_theme": "minorBidi",
            "size_half_points": "22",
        },
        "paragraph_properties": {},
    }
    model.extras["theme_font_scheme"] = {
        "minorHAnsi": "Cambria",
        "minorEastAsia": "MS Mincho",
        "minorBidi": "Times New Roman",
    }

    event = next(e for e in BlueprintCompiler().compile(model).events if e.event_type == "ApplyRunProperties")
    assert event.payload["font_ascii"] == "Cambria"
    assert event.payload["font_hansi"] == "Cambria"
    assert event.payload["font_east_asia"] == "MS Mincho"
    assert event.payload["font_cs"] == "Times New Roman"


def test_paragraph_without_explicit_style_resolves_document_default_paragraph_style():
    model = DocumentModel(
        source_sha256="c" * 64,
        body=[Paragraph("p1", [Run("r1", "A", {})], style_id=None)],
    )
    model.extras["default_paragraph_style_id"] = "Normal"
    model.extras["document_defaults"] = {
        "run_properties": {"font_ascii": "Cambria", "size_half_points": "22"},
        "paragraph_properties": {"spacing_after": "200"},
    }
    model.extras["style_definitions"] = {
        "Normal": {
            "style_id": "Normal", "name": "Normal", "type": "paragraph",
            "run_properties": {"font_ascii": "Times New Roman"},
            "paragraph_properties": {"alignment": "both", "spacing_after": "120", "spacing_line": "360", "spacing_line_rule": "auto"},
        }
    }
    events = BlueprintCompiler().compile(model).events
    paragraph = next(e for e in events if e.event_type == "ApplyParagraphProperties")
    run = next(e for e in events if e.event_type == "ApplyRunProperties")
    assert paragraph.payload["style_id"] == "Normal"
    assert paragraph.payload["alignment"] == "both"
    assert paragraph.payload["spacing_after"] == "120"
    assert paragraph.payload["spacing_line"] == "360"
    assert run.payload["font_ascii"] == "Times New Roman"


def test_table_cell_without_explicit_style_resolves_normal_after_heading():
    from word_replica.domain.model import Table, TableCell, TableRow
    heading = Paragraph("h", [Run("hr", "Heading", {})], style_id="Heading1")
    cell_paragraph = Paragraph("cp", [Run("cr", "Cell", {"size_half_points": "20"})], style_id=None,
                               properties={"spacing_after": "0", "spacing_line": "240", "spacing_line_rule": "auto"})
    table = Table("t", [TableRow("row", [TableCell("cell", [cell_paragraph])])])
    model = DocumentModel(source_sha256="d" * 64, body=[heading, table])
    model.extras["default_paragraph_style_id"] = "Normal"
    model.extras["document_defaults"] = {"run_properties": {"size_half_points": "22"}, "paragraph_properties": {"spacing_after": "200"}}
    model.extras["style_definitions"] = {
        "Normal": {"style_id":"Normal","name":"Normal","type":"paragraph","run_properties":{"font_ascii":"Times New Roman"},"paragraph_properties":{"alignment":"both","spacing_after":"120","spacing_line":"360","spacing_line_rule":"auto"}},
        "Heading1": {"style_id":"Heading1","name":"heading 1","type":"paragraph","based_on":"Normal","run_properties":{"bold":True,"size_half_points":"28"},"paragraph_properties":{"keepNext":True}},
    }
    events = BlueprintCompiler().compile(model).events
    enter = next(i for i,e in enumerate(events) if e.event_type == "EnterCell")
    cell_props = next(e for e in events[enter:] if e.event_type == "ApplyParagraphProperties")
    cell_run = next(e for e in events[enter:] if e.event_type == "ApplyRunProperties")
    assert cell_props.payload["style_id"] == "Normal"
    assert cell_props.payload["alignment"] == "both"
    assert cell_props.payload["spacing_after"] == "0"
    assert cell_props.payload["spacing_line"] == "240"
    assert cell_run.payload["font_ascii"] == "Times New Roman"
    assert cell_run.payload["bold"] is False
