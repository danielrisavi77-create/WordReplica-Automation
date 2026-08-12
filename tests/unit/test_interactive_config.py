import pytest
from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import (
    InteractiveFidelity,
    InteractiveSpeedMode,
    ReconstructionMode,
)


def test_interactive_options_have_deterministic_release_defaults():
    options = InteractiveOptions()
    assert options.speed_mode is InteractiveSpeedMode.FAST
    assert options.characters_per_second == 25.0
    assert options.object_step_delay_ms == 150
    assert options.fidelity is InteractiveFidelity.MAXIMUM
    assert options.checkpoint_after_tables is True
    assert options.checkpoint_after_images is True
    assert options.checkpoint_after_sections is True
    assert options.checkpoint_event_interval == 500
    assert options.verify_during_run is True
    assert options.block_on_unsupported is True


def test_rebuild_options_default_to_existing_instant_behavior():
    options = RebuildOptions()
    assert options.reconstruction_mode is ReconstructionMode.INSTANT
    assert options.interactive == InteractiveOptions()


def test_interactive_options_reject_invalid_rates():
    with pytest.raises(ValueError, match="characters_per_second"):
        InteractiveOptions(characters_per_second=0)
    with pytest.raises(ValueError, match="object_step_delay_ms"):
        InteractiveOptions(object_step_delay_ms=-1)
    with pytest.raises(ValueError, match="checkpoint_event_interval"):
        InteractiveOptions(checkpoint_event_interval=0)
