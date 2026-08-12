from dataclasses import dataclass, field
from word_replica.domain.enums import (
    FidelityMode,
    InteractiveFidelity,
    InteractiveSpeedMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
    VisibilityMode,
)


@dataclass(frozen=True, slots=True)
class InteractiveOptions:
    speed_mode: InteractiveSpeedMode = InteractiveSpeedMode.FAST
    characters_per_second: float = 25.0
    object_step_delay_ms: int = 150
    fidelity: InteractiveFidelity = InteractiveFidelity.MAXIMUM
    checkpoint_after_tables: bool = True
    checkpoint_after_images: bool = True
    checkpoint_after_sections: bool = True
    checkpoint_event_interval: int = 500
    verify_during_run: bool = True
    block_on_unsupported: bool = True
    allow_preserved_objects: bool = False

    def __post_init__(self) -> None:
        if self.characters_per_second <= 0:
            raise ValueError("characters_per_second must be > 0")
        if self.object_step_delay_ms < 0:
            raise ValueError("object_step_delay_ms must be >= 0")
        if self.checkpoint_event_interval <= 0:
            raise ValueError("checkpoint_event_interval must be > 0")


@dataclass(frozen=True, slots=True)
class RebuildOptions:
    renderer: RendererChoice = RendererChoice.AUTO
    visibility: VisibilityMode = VisibilityMode.BACKGROUND
    fidelity: FidelityMode = FidelityMode.CLEAN
    metadata: MetadataMode = MetadataMode.FRESH
    allow_source_overwrite: bool = False
    preserve_author_fields: bool = False
    custom_metadata_allowlist: tuple[str, ...] = ()
    periodic_save_seconds: int = 60
    reconstruction_mode: ReconstructionMode = ReconstructionMode.INSTANT
    interactive: InteractiveOptions = field(default_factory=InteractiveOptions)

    def __post_init__(self) -> None:
        if self.periodic_save_seconds <= 0:
            raise ValueError("periodic_save_seconds must be > 0")
