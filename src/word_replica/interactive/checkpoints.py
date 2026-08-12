from __future__ import annotations

from dataclasses import asdict, dataclass, field, is_dataclass
from datetime import datetime, timezone
import json
import os
from pathlib import Path
from typing import Any

from word_replica.domain.reconstruction import ReconstructionBlueprint
from word_replica.services.source_guard import sha256_file


@dataclass(frozen=True, slots=True)
class InteractiveCheckpoint:
    project_id: str
    source_sha256: str
    source_model_fingerprint: str
    blueprint_fingerprint: str
    blueprint_schema_version: int
    output_path: str
    output_sha256: str
    last_completed_event_index: int
    section_element_id: str | None
    block_element_id: str | None
    table_element_id: str | None
    cell_element_id: str | None
    save_sequence: int
    settings: dict[str, Any]
    timestamp_utc: str
    status: str
    story: str = "body"
    range_start: int | None = None
    range_end: int | None = None
    paragraph_started: bool = False
    resume_state: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def now(cls, **kwargs) -> "InteractiveCheckpoint":
        return cls(timestamp_utc=datetime.now(timezone.utc).isoformat(), **kwargs)


@dataclass(frozen=True, slots=True)
class ResumeValidation:
    valid: bool
    reason: str = ""


class InteractiveCheckpointStore:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def write(self, checkpoint: InteractiveCheckpoint) -> None:
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        tmp.write_text(json.dumps(asdict(checkpoint), indent=2, ensure_ascii=False, default=str), encoding="utf-8")
        os.replace(tmp, self.path)

    def load_latest(self) -> InteractiveCheckpoint:
        payload = json.loads(self.path.read_text(encoding="utf-8"))
        return InteractiveCheckpoint(**payload)

    def validate_resume(
        self,
        checkpoint: InteractiveCheckpoint,
        *,
        source_path: Path,
        blueprint: ReconstructionBlueprint,
    ) -> ResumeValidation:
        source_path = Path(source_path)
        output = Path(checkpoint.output_path)
        if not source_path.exists() or sha256_file(source_path) != checkpoint.source_sha256:
            return ResumeValidation(False, "source hash mismatch")
        if checkpoint.blueprint_schema_version != blueprint.schema_version:
            return ResumeValidation(False, "schema mismatch")
        if checkpoint.blueprint_fingerprint != blueprint.fingerprint:
            return ResumeValidation(False, "blueprint fingerprint mismatch")
        if checkpoint.source_model_fingerprint != blueprint.source_model_fingerprint:
            return ResumeValidation(False, "source model fingerprint mismatch")
        if not output.exists():
            return ResumeValidation(False, "output missing")
        if sha256_file(output) != checkpoint.output_sha256:
            return ResumeValidation(False, "output hash mismatch")
        if checkpoint.last_completed_event_index < -1 or checkpoint.last_completed_event_index >= blueprint.total_events:
            return ResumeValidation(False, "event index out of range")
        return ResumeValidation(True, "")


class InteractiveCheckpointCoordinator:
    """Turns safe interactive boundaries into truthful Word saves + resume metadata."""

    def __init__(
        self,
        *,
        project_id: str,
        source_path: Path,
        blueprint: ReconstructionBlueprint,
        output_path: Path,
        settings: dict[str, Any],
        renderer,
        save_manager,
        checkpoint_store: InteractiveCheckpointStore,
    ) -> None:
        self.project_id = project_id
        self.source_path = Path(source_path)
        self.blueprint = blueprint
        self.output_path = Path(output_path)
        self.settings = dict(settings)
        self.renderer = renderer
        self.save_manager = save_manager
        self.checkpoint_store = checkpoint_store
        self.section_element_id: str | None = None
        self.block_element_id: str | None = None
        self.table_element_id: str | None = None
        self.cell_element_id: str | None = None
        self.last_checkpoint_event_index: int | None = None
        self.last_checkpoint_timestamp_utc: str | None = None

    def _update_context(self, event) -> None:
        et = event.event_type
        sid = event.source_element_id
        if et == "BeginSection":
            self.section_element_id = sid
        if et in {"BeginParagraph", "BeginTable"}:
            self.block_element_id = sid
        if et == "BeginTable":
            self.table_element_id = sid
        elif et == "EndTable":
            self.table_element_id = None
            self.cell_element_id = None
        elif et == "EnterCell":
            self.cell_element_id = sid
        elif et == "LeaveCell":
            self.cell_element_id = None

    def _boundary_reason(self, index: int, event) -> str | None:
        if not bool(getattr(self.renderer, "is_restart_safe", lambda: True)()):
            return None
        et = event.event_type
        if et == "EndTable" and self.settings.get("checkpoint_after_tables", False):
            return "table_boundary"
        if et == "SetImageZOrder" and self.settings.get("checkpoint_after_images", False):
            return "image_boundary"
        if et == "EndSection" and self.settings.get("checkpoint_after_sections", False):
            return "section_boundary"
        interval = int(self.settings.get("checkpoint_event_interval", 0) or 0)
        if interval > 0 and (index + 1) % interval == 0:
            return "event_interval"
        return None

    def event_completed(self, index: int, event) -> None:
        self._update_context(event)
        reason = self._boundary_reason(index, event)
        if reason is not None:
            self.checkpoint(index, status="RUNNING", reason=reason)

    def stop(self, last_completed_index: int) -> InteractiveCheckpoint:
        return self.checkpoint(last_completed_index, status="STOPPED", reason="stop")

    def milestone(self, last_completed_index: int, reason: str) -> InteractiveCheckpoint:
        return self.checkpoint(last_completed_index, status="RUNNING", reason=reason)

    def checkpoint(self, last_completed_index: int, *, status: str, reason: str) -> InteractiveCheckpoint:
        save_event = self.save_manager.save(
            reason=reason,
            stage="interactive",
            save_callable=lambda: self.renderer.save(self.output_path),
            document_path=self.output_path,
        )
        raw_snapshot = self.renderer.current_state_snapshot()
        snapshot = asdict(raw_snapshot) if is_dataclass(raw_snapshot) else dict(raw_snapshot)
        resume_snapshot_fn = getattr(self.renderer, "resume_state_snapshot", None)
        resume_state = dict(resume_snapshot_fn()) if callable(resume_snapshot_fn) else dict(snapshot)
        checkpoint = InteractiveCheckpoint.now(
            project_id=self.project_id,
            source_sha256=sha256_file(self.source_path),
            source_model_fingerprint=self.blueprint.source_model_fingerprint,
            blueprint_fingerprint=self.blueprint.fingerprint,
            blueprint_schema_version=self.blueprint.schema_version,
            output_path=str(self.output_path),
            output_sha256=save_event.document_sha256,
            last_completed_event_index=last_completed_index,
            section_element_id=self.section_element_id,
            block_element_id=self.block_element_id,
            table_element_id=self.table_element_id,
            cell_element_id=self.cell_element_id,
            save_sequence=save_event.sequence,
            settings=self.settings,
            status=status,
            story=str(snapshot.get("story", "body")),
            range_start=snapshot.get("range_start"),
            range_end=snapshot.get("range_end"),
            paragraph_started=bool(snapshot.get("paragraph_started", False)),
            resume_state=resume_state,
        )
        self.checkpoint_store.write(checkpoint)
        self.last_checkpoint_event_index = last_completed_index
        self.last_checkpoint_timestamp_utc = checkpoint.timestamp_utc
        return checkpoint
