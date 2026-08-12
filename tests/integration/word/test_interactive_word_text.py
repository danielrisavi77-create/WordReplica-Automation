import os
import pytest

from word_replica.domain.model import DocumentModel, Paragraph, Run
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.renderers.interactive_word import InteractiveWordController
from word_replica.parser.parser import DocxParser


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word integration disabled")
def test_interactive_word_types_unicode_one_character_at_a_time(tmp_path):
    model = DocumentModel(source_sha256="a" * 64, body=[Paragraph("p1", [Run("r1", "Až B")])])
    blueprint = BlueprintCompiler().compile(model)
    out = tmp_path / "interactive.docx"
    controller = InteractiveWordController()
    try:
        controller.open_blank()
        for event in blueprint.events:
            controller.execute_event(event)
        controller.save(out)
    finally:
        controller.close()
    reparsed = DocxParser().parse(out)
    assert reparsed.plain_text().rstrip("\n") == "Až B"


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word integration disabled")
def test_real_word_pause_resume_mid_paragraph_blocks_between_characters(tmp_path):
    import threading
    from word_replica.config import InteractiveOptions
    from word_replica.interactive.control import InteractiveRunControl
    from word_replica.renderers.interactive_word import InteractiveWordRenderer

    model = DocumentModel(source_sha256="a" * 64, body=[Paragraph("p1", [Run("r1", "ABCDE")])])
    blueprint = BlueprintCompiler().compile(model)
    output = tmp_path / "pause_resume.docx"
    control = InteractiveRunControl(); control.start()
    paused = threading.Event()
    done = threading.Event()
    result_box = {}

    class Observer:
        def event_completed(self, index, event):
            if event.event_type == "InsertCharacter" and event.payload.get("character") == "B" and not paused.is_set():
                control.pause()
                paused.set()

    def worker():
        controller = InteractiveWordController()
        renderer = InteractiveWordRenderer(controller=controller, options=InteractiveOptions(object_step_delay_ms=0))
        renderer.speed._sleep = lambda seconds: None
        try:
            controller.open_blank()
            result_box["outcome"] = renderer.execute_blueprint(blueprint, control, Observer())
            renderer.save(output)
        finally:
            renderer.close(); done.set()

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    assert paused.wait(20), "Interactive renderer never reached pause point"
    assert thread.is_alive(), "Renderer did not actually block while paused"
    control.resume()
    assert done.wait(30), "Interactive renderer did not resume and finish"
    thread.join(timeout=1)
    assert result_box["outcome"].status == "COMPLETED"
    assert DocxParser().parse(output).plain_text().rstrip("\n") == "ABCDE"


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word integration disabled")
def test_real_word_state_mismatch_pauses_before_second_character(tmp_path):
    from word_replica.config import InteractiveOptions
    from word_replica.interactive.control import InteractiveRunControl
    from word_replica.renderers.interactive_word import InteractiveWordRenderer

    model = DocumentModel(source_sha256="a" * 64, body=[Paragraph("p1", [Run("r1", "AB")])])
    blueprint = BlueprintCompiler().compile(model)
    controller = InteractiveWordController()
    renderer = InteractiveWordRenderer(controller=controller, options=InteractiveOptions(object_step_delay_ms=0))
    renderer.speed._sleep = lambda seconds: None
    control = InteractiveRunControl(); control.start()
    output = tmp_path / "state_mismatch.docx"

    class Observer:
        def event_completed(self, index, event):
            if event.event_type == "InsertCharacter" and event.payload.get("character") == "A":
                controller.active_range = controller.document.Range(0, 0)
        def state_mismatch(self, index, event, expected, actual):
            pass

    try:
        controller.open_blank()
        outcome = renderer.execute_blueprint(blueprint, control, Observer())
        renderer.save(output)
    finally:
        renderer.close()
    assert outcome.status == "PAUSED_STATE_MISMATCH"
    assert control.state.value == "PAUSED"
    assert DocxParser().parse(output).plain_text().rstrip("\n") == "A"
