import pytest
from word_replica.config import RebuildOptions
from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, VisibilityMode


def test_defaults_are_safe_and_truthful():
    options = RebuildOptions()
    assert options.renderer is RendererChoice.AUTO
    assert options.visibility is VisibilityMode.BACKGROUND
    assert options.fidelity is FidelityMode.CLEAN
    assert options.metadata is MetadataMode.FRESH
    assert options.allow_source_overwrite is False
    assert options.preserve_author_fields is False
    assert options.custom_metadata_allowlist == ()
    assert options.periodic_save_seconds == 60


def test_periodic_save_must_be_positive():
    with pytest.raises(ValueError, match="periodic_save_seconds"):
        RebuildOptions(periodic_save_seconds=0)
