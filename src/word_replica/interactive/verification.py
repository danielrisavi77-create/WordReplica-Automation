from __future__ import annotations

from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any

from word_replica.domain.enums import InteractiveRunState
from word_replica.domain.reconstruction import (
    InteractiveProgress,
    LiveVerificationResult,
    ReconstructionBlueprint,
    SemanticLocation,
    WordStateSnapshot,
)


def _snapshot_dict(snapshot: WordStateSnapshot | dict[str, Any]) -> dict[str, Any]:
    if is_dataclass(snapshot):
        return asdict(snapshot)
    return dict(snapshot)


def state_snapshots_match(expected, actual) -> bool:
    left = _snapshot_dict(expected)
    right = _snapshot_dict(actual)
    keys = (
        "document_identity", "story", "range_start", "range_end",
        "section_index", "table_element_id", "cell_element_id",
    )
    return all(left.get(key) == right.get(key) for key in keys)


class InteractiveProgressTracker:
    def __init__(self, blueprint: ReconstructionBlueprint) -> None:
        self.blueprint = blueprint
        self.completed_events = 0
        self.completed_characters = 0
        self.completed_tables = 0
        self.completed_images = 0
        self.completed_sections = 0
        self.source_element_id: str | None = None
        self.story = "body"
        self.section_index = 0
        self.table_index = -1
        self.row_index: int | None = None
        self.cell_index: int | None = None
        self.last_checkpoint_event_index: int | None = None
        self.last_checkpoint_timestamp_utc: str | None = None
        self.verification_status = "NOT_RUN"

    def event_completed(self, index, event) -> None:
        self.completed_events = index + 1
        self.source_element_id = event.source_element_id
        if event.event_type == "InsertCharacter":
            self.completed_characters += 1
        elif event.event_type == "InsertText":
            self.completed_characters += len(str(event.payload.get("text", "")))
        elif event.event_type == "BeginSection":
            self.section_index = int(event.payload.get("section_index", self.section_index))
            self.completed_sections += 1
        elif event.event_type == "BeginHeader":
            self.story = f"header:{event.payload.get('story_type', 'default')}"
        elif event.event_type == "BeginFooter":
            self.story = f"footer:{event.payload.get('story_type', 'default')}"
        elif event.event_type in {"EndHeader", "EndFooter", "EndFootnoteStory", "EndEndnoteStory"}:
            self.story = "body"
        elif event.event_type == "BeginFootnoteStory":
            self.story = "footnote"
        elif event.event_type == "BeginEndnoteStory":
            self.story = "endnote"
        elif event.event_type == "BeginTable":
            self.table_index += 1
        elif event.event_type == "EndTable":
            self.completed_tables += 1
        elif event.event_type == "EnterCell":
            self.row_index = int(event.payload.get("row", 0))
            self.cell_index = int(event.payload.get("column", 0))
        elif event.event_type == "LeaveCell":
            self.row_index = None; self.cell_index = None
        elif event.event_type == "SetImageZOrder":
            self.completed_images += 1

    def prime_through(self, last_completed_index: int) -> None:
        if last_completed_index < 0:
            return
        end = min(last_completed_index, self.blueprint.total_events - 1)
        for index in range(0, end + 1):
            self.event_completed(index, self.blueprint.events[index])

    def checkpoint_saved(self, event_index: int, timestamp_utc: str) -> None:
        self.last_checkpoint_event_index = event_index
        self.last_checkpoint_timestamp_utc = timestamp_utc

    def verification_updated(self, status: str) -> None:
        self.verification_status = status

    def snapshot(self, *, state: InteractiveRunState) -> InteractiveProgress:
        return InteractiveProgress(
            state=state,
            completed_events=self.completed_events,
            total_events=self.blueprint.total_events,
            completed_characters=self.completed_characters,
            total_characters=self.blueprint.total_visible_characters,
            location=SemanticLocation(
                story=self.story,
                section_index=self.section_index,
                table_index=max(0, self.table_index) if self.table_index >= 0 else None,
                row_index=self.row_index,
                cell_index=self.cell_index,
            ),
            source_element_id=self.source_element_id,
            completed_tables=self.completed_tables,
            completed_images=self.completed_images,
            completed_sections=self.completed_sections,
            last_checkpoint_event_index=self.last_checkpoint_event_index,
            last_checkpoint_timestamp_utc=self.last_checkpoint_timestamp_utc,
            verification_status=self.verification_status,
        )


class LiveVerifier:
    """Safe-boundary canonical prefix verifier. It never mutates either model."""

    def __init__(self, parser) -> None:
        self.parser = parser

    @staticmethod
    def _paragraph_text(model) -> str:
        return "".join(paragraph.text() for paragraph in model.iter_paragraphs())

    def verify_boundary(self, source_model, output_path: Path, boundary: dict[str, Any]) -> LiveVerificationResult:
        try:
            output_model = self.parser.parse(Path(output_path)) if hasattr(self.parser, "parse") else self.parser(Path(output_path))
        except Exception as exc:
            return LiveVerificationResult("PAUSED_FIDELITY_MISMATCH", str(boundary.get("name", "boundary")), False, (f"output parse failed: {exc}",))
        expected_prefix = str(boundary.get("expected_text_prefix", ""))
        actual_text = self._paragraph_text(output_model)
        reasons: list[str] = []
        if expected_prefix and not actual_text.startswith(expected_prefix):
            reasons.append("text prefix mismatch")
        expected_tables = boundary.get("completed_tables")
        if expected_tables is not None:
            from word_replica.domain.model import Table
            actual_tables = sum(1 for block in output_model.body if isinstance(block, Table))
            if actual_tables < int(expected_tables): reasons.append("table count mismatch")
        status = "PASS" if not reasons else "PAUSED_FIDELITY_MISMATCH"
        return LiveVerificationResult(status, str(boundary.get("name", "boundary")), not reasons, tuple(reasons))
