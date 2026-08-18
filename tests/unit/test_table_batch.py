import pytest

from word_replica.domain.reconstruction import ReconstructionEvent


def _resolved_table_events():
    return (
        ReconstructionEvent("BeginTable", "t1", {"rows": 1, "columns": 1}),
        ReconstructionEvent("SetTableProperties", "t1", {"layout": "fixed"}),
        ReconstructionEvent("SetColumnWidth", "t1", {"column": 1, "width_twips": "2400"}),
        ReconstructionEvent(
            "SetCellProperties",
            "c1",
            {"row": 1, "column": 1, "properties": {"vertical_alignment": "top"}},
        ),
        ReconstructionEvent("EnterCell", "c1", {"row": 1, "column": 1}),
        ReconstructionEvent("BeginParagraph", "p1", {}),
        ReconstructionEvent(
            "ApplyParagraphProperties",
            "p1",
            {"style_id": "Normal", "alignment": "both"},
        ),
        ReconstructionEvent(
            "ApplyRunProperties",
            "r1",
            {"bold": False, "font_ascii": "Times New Roman"},
        ),
        ReconstructionEvent("InsertText", "r1", {"text": "Alpha "}),
        ReconstructionEvent(
            "ApplyRunProperties",
            "r2",
            {"bold": True, "font_ascii": "Times New Roman"},
        ),
        ReconstructionEvent("InsertText", "r2", {"text": "Beta"}),
        ReconstructionEvent("EndParagraph", "p1", {}),
        ReconstructionEvent("LeaveCell", "c1", {"row": 1, "column": 1}),
        ReconstructionEvent("EndTable", "t1", {}),
    )


def test_resolved_text_only_table_becomes_one_batch_event():
    from word_replica.interactive.table_batch import build_table_batch_event

    batch = build_table_batch_event(_resolved_table_events())

    assert batch is not None
    assert batch.event_type == "InsertTableBatch"
    assert batch.source_element_id == "t1"
    assert batch.payload["rows"] == 1
    assert batch.payload["columns"] == 1
    assert batch.payload["paragraph_count"] == 1
    assert batch.payload["text_projection"] == "Alpha Beta"
    assert batch.payload["cells"][0]["text"] == "Alpha Beta"
    assert batch.payload["cells"][0]["runs"] == (
        {
            "source_element_id": "r1",
            "start": 0,
            "end": 6,
            "properties": {"bold": False, "font_ascii": "Times New Roman"},
        },
        {
            "source_element_id": "r2",
            "start": 6,
            "end": 10,
            "properties": {"bold": True, "font_ascii": "Times New Roman"},
        },
    )


@pytest.mark.parametrize(
    "unsupported_type",
    [
        "MergeCells",
        "InsertTab",
        "InsertLineBreak",
        "InsertPageBreak",
        "CreateField",
        "BookmarkStart",
        "CreateBookmark",
        "InsertImage",
        "CreateListBinding",
        "BeginTable",
    ],
)
def test_unsupported_table_content_keeps_the_legacy_path(unsupported_type):
    from word_replica.interactive.table_batch import build_table_batch_event

    events = list(_resolved_table_events())
    events.insert(-3, ReconstructionEvent(unsupported_type, "unsupported", {}))

    assert build_table_batch_event(events) is None


def test_second_paragraph_in_one_cell_keeps_the_legacy_path():
    from word_replica.interactive.table_batch import build_table_batch_event

    events = list(_resolved_table_events())
    events.insert(-3, ReconstructionEvent("BeginParagraph", "p2", {}))

    assert build_table_batch_event(events) is None


@pytest.mark.parametrize("text", ["Alpha\tBeta", "Alpha\rBeta", "Alpha\x07Beta"])
def test_word_table_delimiter_in_source_text_keeps_the_legacy_path(text):
    from word_replica.interactive.table_batch import build_table_batch_event

    events = [
        ReconstructionEvent(event.event_type, event.source_element_id, {"text": text})
        if event.event_type == "InsertText" and event.source_element_id == "r1"
        else event
        for event in _resolved_table_events()
    ]

    assert build_table_batch_event(events) is None


def test_missing_cell_or_dimension_mismatch_keeps_the_legacy_path():
    from word_replica.interactive.table_batch import build_table_batch_event

    events = list(_resolved_table_events())
    events[0] = ReconstructionEvent("BeginTable", "t1", {"rows": 1, "columns": 2})

    assert build_table_batch_event(events) is None
