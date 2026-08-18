from __future__ import annotations

import os
from pathlib import Path

import pytest

from scripts.codex_automation.word_microbenchmark import (
    build_table_batch_fixture,
    run_table_batch_comparison,
)
from word_replica.domain.model import Table
from word_replica.parser.parser import DocxParser


pytestmark = pytest.mark.skipif(
    os.environ.get("WORD_REPLICA_RUN_WORD_PERF") != "1",
    reason="set WORD_REPLICA_RUN_WORD_PERF=1 to run the real-Word table benchmark",
)


def _effective_runs(paragraph) -> tuple:
    segments: list[tuple[str, bool, bool]] = []
    for run in paragraph.runs:
        key = (bool(run.properties.get("bold")), bool(run.properties.get("italic")))
        if segments and segments[-1][1:] == key:
            previous = segments[-1]
            segments[-1] = (previous[0] + run.text, *key)
        else:
            segments.append((run.text, *key))
    return tuple(segments)


def _table_projection(model) -> tuple:
    result = []
    for table in (block for block in model.body if isinstance(block, Table)):
        rows = []
        for row in table.rows:
            cells = []
            for cell in row.cells:
                paragraphs = []
                for paragraph in cell.blocks:
                    paragraphs.append((
                        paragraph.text(),
                        _effective_runs(paragraph),
                        paragraph.properties.get("alignment"),
                    ))
                cells.append((tuple(paragraphs), {
                    "grid_span": cell.properties.get("grid_span", 1),
                    "width": cell.properties.get("width"),
                    "width_type": cell.properties.get("width_type"),
                    "vertical_alignment": cell.properties.get("vertical_alignment", "top"),
                }))
            rows.append(tuple(cells))
        width_type = table.properties.get("width_type")
        if width_type in {None, "auto", "nil"}:
            width_type = "auto"
        result.append((tuple(rows), {
            "alignment": table.properties.get("alignment"),
            "layout": table.properties.get("layout"),
            "width": table.properties.get("width"),
            "width_type": width_type,
        }))
    return tuple(result)


def test_real_word_table_fast_path_preserves_fidelity_and_beats_legacy_median(tmp_path):
    repo_root = Path(__file__).resolve().parents[3]
    source = build_table_batch_fixture(
        tmp_path / "table-batch-source.docx",
        tables=2,
        rows=12,
        columns=4,
    )
    report = run_table_batch_comparison(
        source,
        tmp_path / "comparison",
        repo_root=repo_root,
        runs=int(os.environ.get("WORD_REPLICA_TABLE_PERF_RUNS", "3")),
        timeout_seconds=1800,
    )

    assert report["evaluable"] is True
    assert report["fidelity_equal"] is True
    assert report["improvement_percent"] >= 50.0
    assert report["status"] == "PASS"

    parser = DocxParser()
    source_model = parser.parse(source)
    source_projection = _table_projection(source_model)
    for pair in report["pairs"]:
        legacy = pair["legacy"]
        fast = pair["fast"]
        assert fast["l0_l3"] == legacy["l0_l3"]
        assert all(level["passed"] for level in fast["l0_l3"].values())
        fast_model = parser.parse(Path(fast["output"]))
        assert fast_model.plain_text() == source_model.plain_text()
        assert _table_projection(fast_model) == source_projection
        assert any(
            row["event_type"] == "InsertTableBatch"
            for row in fast["performance_profile"]["by_event_type"]
        )
