from word_replica.config import InteractiveOptions
from word_replica.domain.reconstruction import ReconstructionBlueprint, ReconstructionEvent
from word_replica.interactive.control import InteractiveRunControl
from word_replica.interactive.text_preview import TextPreviewPlayer


def _blueprint(events):
    return ReconstructionBlueprint.build(source_sha256="a" * 64, source_model_fingerprint="m", events=tuple(events))


def _play(options, events):
    sink_calls = []
    player = TextPreviewPlayer(options, lambda kind, payload: sink_calls.append((kind, payload)))
    player.speed._sleep = lambda seconds: None
    control = InteractiveRunControl(); control.start()
    outcome = player.play(_blueprint(events), control)
    return sink_calls, outcome


def test_plain_text_run_emits_one_sink_call_without_letter_by_letter():
    calls, outcome = _play(InteractiveOptions(), [
        ReconstructionEvent("InsertText", "r", {"text": "Sažetak"}),
    ])
    assert calls == [("text", {"text": "Sažetak", "bold": False, "italic": False, "underline": False, "strike": False})]
    assert outcome.status == "COMPLETED"


def test_letter_by_letter_emits_one_sink_call_per_character():
    calls, _ = _play(InteractiveOptions(letter_by_letter=True), [
        ReconstructionEvent("InsertText", "r", {"text": "AB"}),
    ])
    assert [c[1]["text"] for c in calls] == ["A", "B"]


def test_run_properties_carry_into_subsequent_text_until_reset():
    calls, _ = _play(InteractiveOptions(), [
        ReconstructionEvent("ApplyRunProperties", "r", {"bold": True, "italic": True}),
        ReconstructionEvent("InsertText", "r", {"text": "naslov"}),
        ReconstructionEvent("ApplyRunProperties", "r", {"bold": False}),
        ReconstructionEvent("InsertText", "r", {"text": "tekst"}),
    ])
    assert calls[0] == ("text", {"text": "naslov", "bold": True, "italic": True, "underline": False, "strike": False})
    assert calls[1] == ("text", {"text": "tekst", "bold": False, "italic": True, "underline": False, "strike": False})


def test_table_image_and_note_events_become_bracketed_placeholders_not_silence():
    calls, _ = _play(InteractiveOptions(), [
        ReconstructionEvent("InsertTableBatch", "t", {"rows": 2, "columns": 3}),
        ReconstructionEvent("InsertImage", "d", {"asset_id": "a1"}),
        ReconstructionEvent("CreateFootnote", "f", {}),
    ])
    assert calls == [
        ("placeholder", {"label": "[tablica 2×3]"}),
        ("placeholder", {"label": "[slika]"}),
        ("placeholder", {"label": "[fusnota]"}),
    ]


def test_paragraph_and_section_breaks_emit_structural_markers():
    calls, _ = _play(InteractiveOptions(), [
        ReconstructionEvent("BeginParagraph", "p", {}),
        ReconstructionEvent("InsertText", "p", {"text": "x"}),
        ReconstructionEvent("BeginSection", "s", {"section_index": 1}),
    ])
    assert calls[0] == ("newline", {})
    assert calls[-1] == ("sectionbreak", {})


def test_stop_request_halts_playback_and_reports_stopped_status():
    options = InteractiveOptions()
    sink_calls = []
    player = TextPreviewPlayer(options, lambda kind, payload: sink_calls.append((kind, payload)))
    player.speed._sleep = lambda seconds: None
    control = InteractiveRunControl(); control.start(); control.stop()
    outcome = player.play(_blueprint([
        ReconstructionEvent("InsertText", "r", {"text": "never played"}),
    ]), control)
    assert outcome.status == "STOPPED"
    assert sink_calls == []
