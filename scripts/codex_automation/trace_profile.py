from __future__ import annotations

from collections import defaultdict
from datetime import datetime
import json
from pathlib import Path
from typing import Any


TABLE_EVENT_TYPES = {
    "BeginTable", "SetTableProperties", "SetColumnWidth", "SetRowProperties",
    "SetCellProperties", "MergeCells", "EnterCell", "LeaveCell", "EndTable",
    "InsertTableBatch",
}
TABLE_BATCH_PROFILE_FIELDS = {
    "event_index", "table_id", "row_count", "cell_count", "run_count",
    "insert_seconds", "convert_seconds", "geometry_seconds",
    "formatting_seconds", "verification_seconds", "total_seconds",
    "success", "failed_phase",
}


def _event_key(row: dict[str, Any]) -> tuple[int, str, str]:
    return (
        int(row["event_index"]),
        str(row["event_type"]),
        str(row["source_element_id"]),
    )


def _rounded(value: float) -> float:
    return round(float(value), 6)


def profile_event_trace(path: Path, *, slowest_limit: int = 20) -> dict[str, Any]:
    pending: dict[tuple[int, str, str], dict[str, Any]] = {}
    durations: list[dict[str, Any]] = []
    table_batch_profiles: list[dict[str, Any]] = []
    malformed_line_count = 0

    with Path(path).open("r", encoding="utf-8") as handle:
        for line in handle:
            try:
                row = json.loads(line)
                key = _event_key(row)
                status = str(row["status"])
                timestamp = datetime.fromisoformat(str(row["timestamp_utc"]))
            except (KeyError, TypeError, ValueError, json.JSONDecodeError):
                malformed_line_count += 1
                continue

            if status == "table_batch_profile":
                table_batch_profiles.append({
                    key: row[key]
                    for key in TABLE_BATCH_PROFILE_FIELDS
                    if key in row
                })
                continue

            if status == "before":
                pending[key] = {"row": row, "timestamp": timestamp}
                continue
            if status != "after" or key not in pending:
                continue

            before = pending.pop(key)
            seconds = (timestamp - before["timestamp"]).total_seconds()
            source = before["row"]
            durations.append({
                "event_index": key[0],
                "event_type": key[1],
                "source_element_id": key[2],
                "table_element_id": source.get("table_element_id"),
                "cell_element_id": source.get("cell_element_id"),
                "seconds": seconds,
            })

    grouped: dict[str, list[float]] = defaultdict(list)
    for duration in durations:
        grouped[duration["event_type"]].append(float(duration["seconds"]))

    by_event_type = [
        {
            "event_type": event_type,
            "count": len(values),
            "total_seconds": _rounded(sum(values)),
            "mean_seconds": _rounded(sum(values) / len(values)),
            "max_seconds": _rounded(max(values)),
        }
        for event_type, values in grouped.items()
    ]
    by_event_type.sort(key=lambda row: (-row["total_seconds"], row["event_type"]))

    slowest = sorted(
        durations,
        key=lambda row: (-float(row["seconds"]), int(row["event_index"])),
    )[:max(0, int(slowest_limit))]
    slowest_events = [
        {**row, "seconds": _rounded(row["seconds"])}
        for row in slowest
    ]
    table_seconds = sum(
        float(row["seconds"])
        for row in durations
        if row.get("table_element_id") is not None
        or row["event_type"] in TABLE_EVENT_TYPES
    )
    total_seconds = sum(float(row["seconds"]) for row in durations)

    return {
        "paired_event_count": len(durations),
        "unmatched_before_count": len(pending),
        "malformed_line_count": malformed_line_count,
        "total_seconds": _rounded(total_seconds),
        "table_seconds": _rounded(table_seconds),
        "non_table_seconds": _rounded(total_seconds - table_seconds),
        "table_batch_profiles": table_batch_profiles,
        "threshold_counts": {
            "over_1s": sum(float(row["seconds"]) > 1 for row in durations),
            "over_5s": sum(float(row["seconds"]) > 5 for row in durations),
            "over_30s": sum(float(row["seconds"]) > 30 for row in durations),
        },
        "by_event_type": by_event_type,
        "slowest_events": slowest_events,
    }
