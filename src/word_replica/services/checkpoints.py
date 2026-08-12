from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Callable
import uuid

from word_replica.services.audit import AuditLog
from word_replica.services.source_guard import sha256_file


@dataclass(frozen=True, slots=True)
class SaveEvent:
    event_id: str
    sequence: int
    reason: str
    stage: str
    timestamp_utc: str
    timestamp_local: str
    document_path: str
    document_sha256: str


class CheckpointManager:
    def __init__(self, history_path: Path, audit: AuditLog) -> None:
        self.history_path = history_path
        self.history_path.parent.mkdir(parents=True, exist_ok=True)
        self.audit = audit
        self.sequence = 0

    @property
    def next_sequence(self) -> int:
        return self.sequence + 1

    def save(
        self,
        reason: str,
        stage: str,
        save_callable: Callable[[], None],
        document_path: Path,
    ) -> SaveEvent:
        save_callable()
        self.sequence += 1
        now_utc = datetime.now(timezone.utc)
        event = SaveEvent(
            uuid.uuid4().hex,
            self.sequence,
            reason,
            stage,
            now_utc.isoformat(),
            now_utc.astimezone().isoformat(),
            str(document_path),
            sha256_file(document_path),
        )
        with self.history_path.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(asdict(event), ensure_ascii=False) + "\n")
        self.audit.append("DOCUMENT_SAVED", asdict(event))
        return event
