from word_replica.desktop import DesktopState, options_from_desktop_state
from word_replica.domain.enums import (
    InteractiveFidelity, InteractiveSpeedMode, ReconstructionMode, RendererChoice, VisibilityMode,
)


def test_desktop_interactive_state_maps_exact_engine_options():
    state=DesktopState(reconstruction_mode="interactive", renderer="word", visibility="background",
        interactive_speed_mode="custom", interactive_characters_per_second=12.5,
        interactive_object_step_delay_ms=275, interactive_fidelity="maximum",
        checkpoint_after_tables=True, checkpoint_after_images=False, checkpoint_after_sections=True,
        checkpoint_event_interval=321, verify_during_run=True, allow_preserved_objects=False)
    options=options_from_desktop_state(state)
    assert options.reconstruction_mode is ReconstructionMode.INTERACTIVE
    assert options.renderer is RendererChoice.WORD
    assert options.visibility is VisibilityMode.VISIBLE
    assert options.interactive.speed_mode is InteractiveSpeedMode.CUSTOM
    assert options.interactive.characters_per_second == 12.5
    assert options.interactive.object_step_delay_ms == 275
    assert options.interactive.fidelity is InteractiveFidelity.MAXIMUM
    assert options.interactive.checkpoint_after_images is False
    assert options.interactive.checkpoint_event_interval == 321


def test_desktop_interactive_pure_docx_is_a_licence_free_preview_routed_to_instant_docx():
    from word_replica.desktop import is_pure_docx_preview_state
    state=DesktopState(reconstruction_mode="interactive", renderer="docx",
        interactive_speed_mode="custom", interactive_characters_per_second=12.5,
        interactive_letter_by_letter=True)
    assert is_pure_docx_preview_state(state) is True
    options=options_from_desktop_state(state)
    # The real file still comes from the ordinary, well-tested Instant Pure DOCX
    # path — Interactive+Pure DOCX never drives Word, so it works without a
    # licensed/activated Word install.
    assert options.reconstruction_mode is ReconstructionMode.INSTANT
    assert options.renderer is RendererChoice.DOCX
    assert options.visibility is VisibilityMode.BACKGROUND
    # Pacing options for the local text-preview animation still carry through.
    assert options.interactive.speed_mode is InteractiveSpeedMode.CUSTOM
    assert options.interactive.characters_per_second == 12.5
    assert options.interactive.letter_by_letter is True


def test_is_pure_docx_preview_state_requires_both_interactive_and_docx():
    from word_replica.desktop import is_pure_docx_preview_state
    assert is_pure_docx_preview_state(DesktopState(reconstruction_mode="interactive", renderer="word")) is False
    assert is_pure_docx_preview_state(DesktopState(reconstruction_mode="instant", renderer="docx")) is False


def test_desktop_defaults_remain_instant():
    state=DesktopState()
    options=options_from_desktop_state(state)
    assert state.reconstruction_mode == "instant"
    assert options.reconstruction_mode is ReconstructionMode.INSTANT


def test_transparency_copy_describes_automated_character_reconstruction_not_manual_authorship():
    from word_replica.desktop import AUTOMATION_DISCLOSURE
    text=AUTOMATION_DISCLOSURE.lower()
    assert "automat" in text
    assert "ručno autorstvo" in text


def test_progress_view_model_is_determinate_and_semantic():
    from word_replica.desktop import progress_view_model
    from word_replica.domain.enums import InteractiveRunState
    from word_replica.domain.reconstruction import InteractiveProgress, SemanticLocation
    progress=InteractiveProgress(state=InteractiveRunState.RUNNING,completed_events=50,total_events=100,
        completed_characters=20,total_characters=40,location=SemanticLocation(story="body",section_index=1,row_index=2,cell_index=3),
        source_element_id="cell-3",completed_tables=2,completed_images=1,completed_sections=2)
    view=progress_view_model(progress)
    assert view["percent"] == 50.0
    assert "20/40" in view["status"]
    assert "tablice 2" in view["status"]
    assert "red 2" in view["location"] and "ćelija 3" in view["location"]


def test_live_speed_update_routes_to_thread_safe_control():
    from word_replica.desktop import update_interactive_control_speed
    calls=[]
    class Control:
        def set_speed(self,mode,cps=None): calls.append((mode,cps))
    update_interactive_control_speed(Control(),"custom",17.5)
    assert calls[0][0] is InteractiveSpeedMode.CUSTOM
    assert calls[0][1] == 17.5
