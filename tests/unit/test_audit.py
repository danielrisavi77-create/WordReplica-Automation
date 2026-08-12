import json
from pathlib import Path

from word_replica.services.audit import AuditLog


def test_audit_is_append_only_and_uses_event_ids(tmp_path: Path):
    path = tmp_path / "audit.jsonl"
    log = AuditLog(path)
    first = log.append("PROJECT_CREATED", {"project_id": "p1"})
    second = log.append("REBUILD_STARTED", {"renderer": "docx"})
    rows = [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]
    assert [row["event_type"] for row in rows] == ["PROJECT_CREATED", "REBUILD_STARTED"]
    assert first["event_id"] != second["event_id"]
    assert all(row["timestamp_utc"] for row in rows)
    assert rows[0]["payload"] == {"project_id": "p1"}
