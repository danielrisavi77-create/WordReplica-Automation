from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
import json
from typing import Any

from word_replica.domain.enums import CapabilityClass, InteractiveRunState


def _freeze_payload(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(k): _freeze_payload(v) for k, v in sorted(value.items(), key=lambda item: str(item[0]))}
    if isinstance(value, (list, tuple)):
        return tuple(_freeze_payload(v) for v in value)
    return value


def _json_safe(value: Any) -> Any:
    if isinstance(value, tuple):
        return [_json_safe(v) for v in value]
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    return value


@dataclass(frozen=True, slots=True)
class ReconstructionEvent:
    event_type: str
    source_element_id: str
    payload: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "payload", _freeze_payload(dict(self.payload)))


@dataclass(frozen=True, slots=True)
class SemanticLocation:
    story: str = "body"
    section_index: int = 0
    block_index: int = 0
    paragraph_index: int | None = None
    run_index: int | None = None
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None


@dataclass(frozen=True, slots=True)
class InteractiveProgress:
    state: InteractiveRunState = InteractiveRunState.CREATED
    completed_events: int = 0
    total_events: int = 0
    completed_characters: int = 0
    total_characters: int = 0
    location: SemanticLocation = field(default_factory=SemanticLocation)
    source_element_id: str | None = None
    completed_tables: int = 0
    completed_images: int = 0
    completed_sections: int = 0
    last_checkpoint_event_index: int | None = None
    last_checkpoint_timestamp_utc: str | None = None
    verification_status: str = "NOT_RUN"


@dataclass(frozen=True, slots=True)
class WordStateSnapshot:
    document_identity: str
    story: str
    range_start: int | None
    range_end: int | None
    section_index: int | None
    table_element_id: str | None
    cell_element_id: str | None
    last_completed_event_index: int | None = None
    active_range_type: str | None = None


@dataclass(frozen=True, slots=True)
class LiveVerificationResult:
    status: str
    boundary: str
    passed: bool
    reasons: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ReconstructionBlueprint:
    schema_version: int
    source_sha256: str
    source_model_fingerprint: str
    events: tuple[ReconstructionEvent, ...]
    fingerprint: str
    total_events: int
    total_visible_characters: int
    semantic_counts: dict[str, int] = field(default_factory=dict)

    @classmethod
    def build(
        cls,
        *,
        source_sha256: str,
        source_model_fingerprint: str,
        events: tuple[ReconstructionEvent, ...] | list[ReconstructionEvent],
        semantic_counts: dict[str, int] | None = None,
    ) -> "ReconstructionBlueprint":
        frozen_events = tuple(events)
        counts = dict(semantic_counts or {})
        canonical = {
            "schema_version": 1,
            "source_sha256": source_sha256,
            "source_model_fingerprint": source_model_fingerprint,
            "events": [
                {
                    "event_type": e.event_type,
                    "source_element_id": e.source_element_id,
                    "payload": _json_safe(e.payload),
                }
                for e in frozen_events
            ],
            "semantic_counts": counts,
        }
        encoded = json.dumps(
            canonical,
            sort_keys=True,
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        return cls(
            schema_version=1,
            source_sha256=source_sha256,
            source_model_fingerprint=source_model_fingerprint,
            events=frozen_events,
            fingerprint=sha256(encoded).hexdigest(),
            total_events=len(frozen_events),
            total_visible_characters=sum(1 for event in frozen_events if event.event_type == "InsertCharacter"),
            semantic_counts=counts,
        )


@dataclass(frozen=True, slots=True)
class ControlDecision:
    stop_requested: bool
    last_completed_index: int


@dataclass(frozen=True, slots=True)
class ExecutionOutcome:
    status: str
    last_completed_index: int
    completed_events: int

    @classmethod
    def stopped(cls, last_completed_index: int) -> "ExecutionOutcome":
        return cls(
            status="STOPPED",
            last_completed_index=last_completed_index,
            completed_events=max(0, last_completed_index + 1),
        )

    @classmethod
    def completed(cls, total_events: int) -> "ExecutionOutcome":
        return cls(
            status="COMPLETED",
            last_completed_index=total_events - 1,
            completed_events=total_events,
        )


@dataclass(frozen=True, slots=True)
class CapabilityDecision:
    classification: CapabilityClass
    reason: str
    source_element_id: str


@dataclass(frozen=True, slots=True)
class PreflightReport:
    source_sha256: str
    blueprint_fingerprint: str
    blueprint_schema_version: int
    word_available: bool
    missing_assets: tuple[str, ...]
    missing_fonts: tuple[str, ...]
    counts: dict[str, int]
    capability_items: tuple[CapabilityDecision, ...]
    maximum_fidelity_ready: bool
    can_proceed: bool
    blocking_reasons: tuple[str, ...]
    warnings: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class PreparedInteractiveRun:
    source_path: str
    model: Any
    blueprint: ReconstructionBlueprint
    preflight: PreflightReport
    paths: Any
    options: Any = None
