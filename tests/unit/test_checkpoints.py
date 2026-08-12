import json
from pathlib import Path

import pytest

from word_replica.services.audit import AuditLog
from word_replica.services.checkpoints import CheckpointManager


def test_checkpoint_logs_only_after_real_save(tmp_path: Path):
    document = tmp_path / "out.docx"
    history = tmp_path / "save_history.jsonl"
    audit = AuditLog(tmp_path / "audit.jsonl")
    manager = CheckpointManager(history, audit)

    def real_save() -> None:
        document.write_bytes(b"saved")

    event = manager.save("chapter complete", "chapter_1", real_save, document)
    row = json.loads(history.read_text(encoding="utf-8").splitlines()[0])
    assert row["sequence"] == 1
    assert row["event_id"]
    assert row["timestamp_utc"]
    assert row["timestamp_local"]
    assert row["reason"] == "chapter complete"
    assert row["document_sha256"] == event.document_sha256


def test_failed_save_does_not_increment_or_log(tmp_path: Path):
    history = tmp_path / "save_history.jsonl"
    manager = CheckpointManager(history, AuditLog(tmp_path / "audit.jsonl"))

    def failed_save() -> None:
        raise OSError("disk full")

    with pytest.raises(OSError, match="disk full"):
        manager.save("x", "x", failed_save, tmp_path / "missing.docx")
    assert manager.sequence == 0
    assert not history.exists()
