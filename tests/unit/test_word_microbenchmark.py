import json
import shutil

from scripts.codex_automation.process import ChildResult
from scripts.codex_automation.word_microbenchmark import (
    build_table_batch_fixture,
    build_text_growth_fixture,
    run_table_batch_comparison,
    run_word_microbenchmark,
)
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.domain.model import Paragraph, Table
from word_replica.parser.parser import DocxParser


def test_text_growth_fixture_has_exact_text_table_shape_and_bold_spans(tmp_path):
    source = build_text_growth_fixture(
        tmp_path / "text-growth.docx",
        prefix_paragraphs=3,
        measured_paragraphs=2,
        table_cells=4,
    )

    model = DocxParser().parse(source)
    body_paragraphs = [block for block in model.body if isinstance(block, Paragraph)]
    tables = [block for block in model.body if isinstance(block, Table)]

    assert [type(block).__name__ for block in model.body] == [
        "Paragraph", "Paragraph", "Paragraph", "Table", "Paragraph", "Paragraph",
    ]
    assert model.sections[0].properties["column_space"] == "708"
    assert [paragraph.text() for paragraph in body_paragraphs] == [
        "Prefix 0000: ordinary document growth text.",
        "Prefix 0001: ordinary document growth text.",
        "Prefix 0002: ordinary document growth text.",
        "Measured 0000: normal bold normal.",
        "Measured 0001: normal bold normal.",
    ]
    assert len(tables) == 1
    cells = [cell for row in tables[0].rows for cell in row.cells]
    assert len(cells) == 4
    assert [cell.blocks[0].text() for cell in cells] == [
        "Cell 0000: normal bold normal.",
        "Cell 0001: normal bold normal.",
        "Cell 0002: normal bold normal.",
        "Cell 0003: normal bold normal.",
    ]

    measured_and_cell_paragraphs = body_paragraphs[3:] + [cell.blocks[0] for cell in cells]
    assert [
        [(run.text, run.properties.get("bold")) for run in paragraph.runs]
        for paragraph in measured_and_cell_paragraphs
    ] == [
        [("Measured 0000: normal ", None), ("bold", True), (" normal.", None)],
        [("Measured 0001: normal ", None), ("bold", True), (" normal.", None)],
        [("Cell 0000: normal ", None), ("bold", True), (" normal.", None)],
        [("Cell 0001: normal ", None), ("bold", True), (" normal.", None)],
        [("Cell 0002: normal ", None), ("bold", True), (" normal.", None)],
        [("Cell 0003: normal ", None), ("bold", True), (" normal.", None)],
    ]


def test_word_microbenchmark_reports_process_profile_and_l0_l3(tmp_path):
    source = build_text_growth_fixture(
        tmp_path / "source.docx",
        prefix_paragraphs=1,
        measured_paragraphs=1,
        table_cells=1,
    )
    output_dir = tmp_path / "benchmark"

    def executor(command, **kwargs):
        run_dir = command[command.index("--run-dir") + 1]
        run_dir = type(source)(run_dir)
        run_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, run_dir / "output.docx")
        (run_dir / "result.json").write_text(json.dumps({
            "run_status": "PASS",
            "elapsed_seconds": 1.75,
            "reasons": [],
        }), encoding="utf-8")
        (run_dir / "event_trace.jsonl").write_text("\n".join([
            json.dumps({
                "event_index": 0,
                "event_type": "InsertText",
                "source_element_id": "r1",
                "status": "before",
                "timestamp_utc": "2026-08-14T00:00:00+00:00",
                "table_element_id": None,
                "cell_element_id": None,
            }),
            json.dumps({
                "event_index": 0,
                "event_type": "InsertText",
                "source_element_id": "r1",
                "status": "after",
                "timestamp_utc": "2026-08-14T00:00:01.500000+00:00",
                "table_element_id": None,
                "cell_element_id": None,
            }),
        ]), encoding="utf-8")
        return ChildResult(0, False, 1.8, [])

    report = run_word_microbenchmark(
        source,
        output_dir,
        repo_root=tmp_path / "repo",
        child_executor=executor,
    )

    assert report["status"] == "PASS"
    assert report["reconstruction_elapsed_seconds"] == 1.75
    assert report["process"]["timed_out"] is False
    assert report["performance_profile"]["paired_event_count"] == 1
    assert report["performance_profile"]["by_event_type"][0]["total_seconds"] == 1.5
    assert report["l0_l3"] == {
        "L0": {"passed": True, "finding_count": 0},
        "L1": {"passed": True, "finding_count": 0},
        "L2": {"passed": True, "finding_count": 0},
        "L3": {"passed": True, "finding_count": 0},
    }
    assert (output_dir / "benchmark_report.json").exists()


def test_table_batch_fixture_has_exact_shape_runs_and_batch_blueprint(tmp_path):
    source = build_table_batch_fixture(
        tmp_path / "table-batch.docx",
        tables=2,
        rows=12,
        columns=4,
    )

    model = DocxParser().parse(source)
    tables = [block for block in model.body if isinstance(block, Table)]
    cells = [cell for table in tables for row in table.rows for cell in row.cells]

    assert [type(block).__name__ for block in model.body] == [
        "Table", "Paragraph", "Table", "Paragraph",
    ]
    assert model.body[1].text() == "Table benchmark separator 00"
    assert model.body[3].text() == "Table benchmark end"
    assert len(tables) == 2
    assert len(cells) == 96
    assert all(len(cell.blocks) == 1 for cell in cells)
    assert cells[0].blocks[0].text() == "Table 00 Cell 00,00: normal bold normal."
    assert cells[-1].blocks[0].text() == "Table 01 Cell 11,03: normal bold normal."
    for cell in cells:
        paragraph = cell.blocks[0]
        assert [(run.text, run.properties.get("bold")) for run in paragraph.runs] == [
            (paragraph.text().split("bold")[0], None),
            ("bold", True),
            (" normal.", None),
        ]

    blueprint = BlueprintCompiler(enable_table_fast_path=True).compile(model)
    batches = [event for event in blueprint.events if event.event_type == "InsertTableBatch"]
    assert len(batches) == 2
    assert all(len(event.payload["cells"]) == 48 for event in batches)
    assert all(event.payload["paragraph_count"] == 48 for event in batches)
    assert not any(event.event_type == "BeginTable" for event in blueprint.events)


def test_table_batch_comparison_uses_three_pairs_and_median_acceptance_gate(tmp_path):
    source = build_table_batch_fixture(
        tmp_path / "source.docx", tables=1, rows=2, columns=2,
    )
    commands = []

    def executor(command, **kwargs):
        commands.append(command)
        run_dir = type(source)(command[command.index("--run-dir") + 1])
        run_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, run_dir / "output.docx")
        (run_dir / "result.json").write_text(json.dumps({
            "run_status": "PASS", "elapsed_seconds": 12.0, "reasons": [],
        }), encoding="utf-8")
        legacy = "--disable-table-fast-path" in command
        event_type = "InsertText" if legacy else "InsertTableBatch"
        end_second = 10 if legacy else 2
        (run_dir / "event_trace.jsonl").write_text("\n".join([
            json.dumps({
                "event_index": 0, "event_type": event_type,
                "source_element_id": "t", "status": "before",
                "timestamp_utc": "2026-08-14T00:00:00+00:00",
                "table_element_id": "t", "cell_element_id": "c",
            }),
            json.dumps({
                "event_index": 0, "event_type": event_type,
                "source_element_id": "t", "status": "after",
                "timestamp_utc": f"2026-08-14T00:00:{end_second:02d}+00:00",
                "table_element_id": "t", "cell_element_id": "c",
            }),
        ]), encoding="utf-8")
        return ChildResult(0, False, float(end_second), [])

    report = run_table_batch_comparison(
        source,
        tmp_path / "comparison",
        repo_root=tmp_path / "repo",
        runs=3,
        child_executor=executor,
    )

    assert len(commands) == 6
    assert sum("--disable-table-fast-path" in command for command in commands) == 3
    assert report["legacy_median_table_seconds"] == 10.0
    assert report["fast_median_table_seconds"] == 2.0
    assert report["improvement_percent"] == 80.0
    assert report["fidelity_equal"] is True
    assert report["status"] == "PASS"
