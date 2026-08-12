from word_replica.gui.controller import UiState, options_from_state


def test_initial_ui_state_is_safe():
    state = UiState()
    assert state.source_path is None
    assert state.running is False
    assert state.renderer == 'auto'
    assert state.fidelity == 'clean'
    assert state.metadata == 'fresh'


def test_options_are_derived_without_mutating_state():
    state = UiState(renderer='docx', fidelity='full')
    options = options_from_state(state)
    assert options.renderer.value == 'docx'
    assert options.fidelity.value == 'full'
    assert state.running is False
