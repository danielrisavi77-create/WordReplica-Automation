from __future__ import annotations

from dataclasses import asdict, is_dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Any


def _snapshot_dict(snapshot: Any) -> dict[str, Any]:
    if snapshot is None:
        return {}
    if is_dataclass(snapshot):
        return asdict(snapshot)
    if isinstance(snapshot, dict):
        return dict(snapshot)
    return {
        "story": getattr(snapshot, "story", None),
        "range_start": getattr(snapshot, "range_start", getattr(snapshot, "Start", None)),
        "range_end": getattr(snapshot, "range_end", getattr(snapshot, "End", None)),
        "section_index": getattr(snapshot, "section_index", None),
        "table_element_id": getattr(snapshot, "table_element_id", None),
        "cell_element_id": getattr(snapshot, "cell_element_id", None),
    }


class JsonlEventTraceObserver:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._last_started: dict[int, dict[str, Any]] = {}

    def _write(self, payload: dict[str, Any]) -> None:
        payload = {"timestamp_utc": datetime.now(timezone.utc).isoformat(), **payload}
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")

    def _base(self, index, event, snapshot) -> dict[str, Any]:
        state = _snapshot_dict(snapshot)
        return {
            "event_index": int(index),
            "event_type": event.event_type,
            "source_element_id": event.source_element_id,
            "story": state.get("story"),
            "section_index": state.get("section_index"),
            "table_element_id": state.get("table_element_id"),
            "cell_element_id": state.get("cell_element_id"),
            "word_range_start": state.get("range_start"),
            "word_range_end": state.get("range_end"),
            "active_range_type": state.get("active_range_type") or type(snapshot).__name__,
        }

    def event_started(self, index, event, snapshot) -> None:
        row = self._base(index, event, snapshot)
        self._last_started[int(index)] = row
        self._write({**row, "status": "before"})

    def event_completed(self, index, event) -> None:
        # Existing callback is retained for production progress observers. The
        # harness writes the richer after record from event_finished().
        return None

    def event_finished(self, index, event, snapshot) -> None:
        self._write({**self._base(index, event, snapshot), "status": "after"})

    def event_failed(self, index, event, snapshot, exc) -> None:
        self._write({
            **self._base(index, event, snapshot),
            "status": "error",
            "exception_type": type(exc).__name__,
            "exception_message": str(exc),
        })

    def state_snapshot_failed(self, index, event, phase, exc) -> None:
        self._write({
            **self._base(index, event, None),
            "status": "snapshot_error",
            "phase": str(phase),
            "exception_type": type(exc).__name__,
            "exception_message": str(exc),
        })

    def state_mismatch(self, index, event, expected, actual) -> None:
        self._write({
            **self._base(index, event, actual),
            "status": "state_mismatch",
            "expected": _snapshot_dict(expected),
            "actual": _snapshot_dict(actual),
        })
