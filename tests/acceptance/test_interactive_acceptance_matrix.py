import json
import os

import pytest

from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import (
    InteractiveFidelity,
    InteractiveSpeedMode,
    ReconstructionMode,
    RunStatus,
)
from word_replica.domain.model import Table
from word_replica.opc.package_reader import DocxPackage
from word_replica.opc.properties import read_properties
from word_replica.parser.parser import DocxParser
from word_replica.services.rebuild import RebuildService
from word_replica.services.source_guard import sha256_file

INTERACTIVE_SUPPORTED = [
    "01_plain_text.docx",
    "02_headings_styles.docx",
    "03_lists.docx",
    "04_tables_merged.docx",
    "05_images_inline_floating.docx",
    "06_sections_orientations.docx",
    "07_headers_footers_numbers.docx",
    "08_footnotes_endnotes.docx",
    "09_toc_fields.docx",
    "11_bookmarks_crossrefs.docx",
    "14_interactive_mixed_breaks.docx",
    "15_nested_table.docx",
    "16_table_shading_borders.docx",
    "17_academic_supported.docx",
]


def _count_tables(blocks):
    total = 0
    for block in blocks:
        if isinstance(block, Table):
            total += 1
            for row in block.rows:
                for cell in row.cells:
                    total += _count_tables(cell.blocks)
    return total


def _jsonl(path):
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def maximum_speed_options(*, verify_during_run: bool = False) -> RebuildOptions:
    return RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(
            speed_mode=InteractiveSpeedMode.MAXIMUM,
            characters_per_second=250,
            object_step_delay_ms=0,
            checkpoint_event_interval=100000,
            verify_during_run=verify_during_run,
        ),
    )


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word acceptance disabled")
@pytest.mark.parametrize("fixture_name", INTERACTIVE_SUPPORTED)
def test_interactive_supported_corpus_reconstructs_in_real_word(corpus_dir, fixture_name):
    source = corpus_dir / fixture_name
    before = sha256_file(source)
    result = RebuildService.default_for_tests().rebuild(source, maximum_speed_options())
    assert sha256_file(source) == before
    assert result.status in {RunStatus.PASS, RunStatus.WARN}, result.reasons
    assert result.output_path is not None and result.output_path.exists()
    assert result.qa_report_path is not None and result.qa_report_path.exists()
    html = result.qa_report_path.read_text(encoding="utf-8")
    assert "L0 — PASS" in html
    assert "L1 — PASS" in html
    assert "interactive-word" in html
    assert "Not verified in this run" not in html, "real Word L4 export/comparison must run in acceptance"

    project_root = result.output_path.parent.parent
    audit_rows = _jsonl(project_root / "logs" / "audit.jsonl")
    save_rows = _jsonl(project_root / "logs" / "save_history.jsonl")
    blueprint = json.loads((project_root / "working" / "blueprint.json").read_text(encoding="utf-8"))
    metrics = [row for row in audit_rows if row.get("event_type") == "INTERACTIVE_METRICS"]
    assert metrics and metrics[-1]["payload"]["completed_characters"] == blueprint["total_visible_characters"]
    assert metrics[-1]["payload"]["completed_events"] == blueprint["total_events"]
    assert any(row.get("event_type") == "INTERACTIVE_EXECUTION_STARTED" for row in audit_rows)
    assert any(row.get("event_type") == "INTERACTIVE_FINAL_QA_COMPLETED" for row in audit_rows)

    output_model = DocxParser().parse(result.output_path)
    assert _count_tables(output_model.body) == int(blueprint["semantic_counts"].get("tables", 0))
    assert len(output_model.drawings) == int(blueprint["semantic_counts"].get("images", 0))
    with DocxPackage.open(result.output_path) as package:
        props = read_properties(package)
    assert int(props.custom["WordReplicaActualSaveCount"]) == len(save_rows) == result.save_count


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word acceptance disabled")
def test_interactive_live_verification_runs_on_real_table_boundary(corpus_dir):
    source = corpus_dir / "04_tables_merged.docx"
    result = RebuildService.default_for_tests().rebuild(
        source, maximum_speed_options(verify_during_run=True)
    )
    assert result.status in {RunStatus.PASS, RunStatus.WARN}, result.reasons
    assert result.qa_report_path is not None


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word acceptance disabled")
def test_maximum_fidelity_blocks_academic_complex_before_silent_downgrade(corpus_dir):
    source = corpus_dir / "13_academic_complex.docx"
    before = sha256_file(source)
    result = RebuildService.default_for_tests().rebuild(source, maximum_speed_options())
    assert sha256_file(source) == before
    assert result.status is RunStatus.FAIL
    assert result.output_path is None
    assert any("comment" in reason.lower() or "revision" in reason.lower() or "ooxml" in reason.lower() for reason in result.reasons)


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word acceptance disabled")
def test_standard_fidelity_academic_complex_is_explicit_warn_not_silent(corpus_dir):
    source = corpus_dir / "13_academic_complex.docx"
    options = RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        interactive=InteractiveOptions(
            speed_mode=InteractiveSpeedMode.MAXIMUM,
            object_step_delay_ms=0,
            fidelity=InteractiveFidelity.STANDARD,
            checkpoint_event_interval=100000,
            verify_during_run=False,
        ),
    )
    result = RebuildService.default_for_tests().rebuild(source, options)
    assert result.status is RunStatus.WARN, result.reasons
    assert result.output_path is not None and result.output_path.exists()
    assert any(w.code == "CAPABILITY_UNSUPPORTED" for w in result.warnings)
