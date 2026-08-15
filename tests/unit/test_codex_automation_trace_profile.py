import json

from scripts.codex_automation.trace_profile import profile_event_trace


def _row(index, event_type, source_id, status, timestamp):
    return {
        "event_index": index,
        "event_type": event_type,
        "source_element_id": source_id,
        "status": status,
        "timestamp_utc": timestamp,
        "table_element_id": "table-1" if index == 2 else None,
        "cell_element_id": "cell-1" if index == 2 else None,
    }


def test_profile_pairs_complete_events_and_reports_unmatched_before(tmp_path):
    trace = tmp_path / "event_trace.jsonl"
    rows = [
        _row(1, "BeginParagraph", "p1", "before", "2026-08-14T00:00:00+00:00"),
        _row(1, "BeginParagraph", "p1", "after", "2026-08-14T00:00:02+00:00"),
        _row(2, "InsertText", "r1", "before", "2026-08-14T00:00:02+00:00"),
        _row(2, "InsertText", "r1", "after", "2026-08-14T00:00:08.500000+00:00"),
        _row(3, "InsertText", "r2", "before", "2026-08-14T00:00:09+00:00"),
    ]
    trace.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")

    profile = profile_event_trace(trace)

    assert profile["paired_event_count"] == 2
    assert profile["unmatched_before_count"] == 1
    assert profile["malformed_line_count"] == 0
    assert profile["total_seconds"] == 8.5
    assert profile["table_seconds"] == 6.5
    assert profile["non_table_seconds"] == 2.0
    assert profile["threshold_counts"] == {"over_1s": 2, "over_5s": 1, "over_30s": 0}
    assert profile["by_event_type"] == [
        {
            "event_type": "InsertText",
            "count": 1,
            "total_seconds": 6.5,
            "mean_seconds": 6.5,
            "max_seconds": 6.5,
        },
        {
            "event_type": "BeginParagraph",
            "count": 1,
            "total_seconds": 2.0,
            "mean_seconds": 2.0,
            "max_seconds": 2.0,
        },
    ]
    assert profile["slowest_events"][0] == {
        "event_index": 2,
        "event_type": "InsertText",
        "source_element_id": "r1",
        "table_element_id": "table-1",
        "cell_element_id": "cell-1",
        "seconds": 6.5,
    }


def test_profile_counts_malformed_lines_without_failing(tmp_path):
    trace = tmp_path / "event_trace.jsonl"
    trace.write_text("not-json\n{}\n", encoding="utf-8")

    profile = profile_event_trace(trace)

    assert profile["paired_event_count"] == 0
    assert profile["malformed_line_count"] == 2


def test_profile_preserves_sanitized_table_batch_phase_rows(tmp_path):
    trace = tmp_path / "event_trace.jsonl"
    rows = [
        _row(4, "InsertTableBatch", "t2", "before", "2026-08-14T00:00:00+00:00"),
        {
            **_row(4, "InsertTableBatch", "t2", "table_batch_profile", "2026-08-14T00:00:01+00:00"),
            "table_id": "t2", "cell_count": 48, "total_seconds": 0.8,
            "formatting_seconds": 0.5, "success": True, "failed_phase": None,
        },
        _row(4, "InsertTableBatch", "t2", "after", "2026-08-14T00:00:01+00:00"),
    ]
    trace.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")

    profile = profile_event_trace(trace)

    assert profile["table_seconds"] == 1.0
    assert profile["table_batch_profiles"] == [{
        "event_index": 4,
        "table_id": "t2",
        "cell_count": 48,
        "total_seconds": 0.8,
        "formatting_seconds": 0.5,
        "success": True,
        "failed_phase": None,
    }]
