import json
import os
from pathlib import Path

import pytest

from word_replica.config import RebuildOptions
from word_replica.domain.enums import FidelityMode, RendererChoice, RunStatus
from word_replica.opc.package_reader import DocxPackage
from word_replica.opc.properties import read_properties
from word_replica.services.rebuild import RebuildService
from word_replica.services.source_guard import sha256_file

PURE_REQUIRED = [
    "01_plain_text.docx", "02_headings_styles.docx", "03_lists.docx",
    "04_tables_merged.docx", "06_sections_orientations.docx",
    "07_headers_footers_numbers.docx", "08_footnotes_endnotes.docx",
    "18_academic_citations.docx",
]

WORD_REQUIRED = [
    f"{i:02d}_{name}.docx" for i, name in [
        (1,"plain_text"),(2,"headings_styles"),(3,"lists"),(4,"tables_merged"),(5,"images_inline_floating"),
        (6,"sections_orientations"),(7,"headers_footers_numbers"),(8,"footnotes_endnotes"),(9,"toc_fields"),
        (10,"comments_tracked_changes"),(11,"bookmarks_crossrefs"),(12,"charts_embedded"),(13,"academic_complex"),
        (18,"academic_citations"),
    ]
]


@pytest.mark.parametrize("fixture_name", PURE_REQUIRED)
def test_required_fallback_fixtures_do_not_fail(corpus_dir, fixture_name):
    source = corpus_dir / fixture_name
    before = sha256_file(source)
    result = RebuildService.default_for_tests().rebuild(
        source, RebuildOptions(renderer=RendererChoice.DOCX)
    )
    assert sha256_file(source) == before
    assert result.status in {RunStatus.PASS, RunStatus.WARN}
    assert result.qa_report_path and result.qa_report_path.exists()


@pytest.mark.word
@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word acceptance disabled")
@pytest.mark.parametrize("fidelity", [FidelityMode.CLEAN, FidelityMode.FULL])
@pytest.mark.parametrize("fixture_name", WORD_REQUIRED)
def test_word_acceptance_matrix(corpus_dir, fixture_name, fidelity):
    source = corpus_dir / fixture_name
    before = sha256_file(source)
    result = RebuildService.default().rebuild(
        source, RebuildOptions(renderer=RendererChoice.WORD, fidelity=fidelity)
    )
    assert sha256_file(source) == before
    assert result.status in {RunStatus.PASS, RunStatus.WARN}
    assert result.qa_report_path and result.qa_report_path.exists()


def test_academic_complex_truthful_save_count_and_fresh_metadata(corpus_dir):
    source = corpus_dir / "13_academic_complex.docx"
    result = RebuildService.default_for_tests().rebuild(
        source, RebuildOptions(renderer=RendererChoice.DOCX)
    )
    assert result.output_path is not None
    project_root = result.output_path.parent.parent
    history = project_root / "logs" / "save_history.jsonl"
    rows = [json.loads(line) for line in history.read_text(encoding="utf-8").splitlines() if line]
    assert len(rows) >= 2
    with DocxPackage.open(result.output_path) as package:
        props = read_properties(package)
    assert int(props.custom["WordReplicaActualSaveCount"]) == len(rows)
    with DocxPackage.open(source) as source_package:
        source_props = read_properties(source_package)
    assert props.created != source_props.created
    assert props.modified != source_props.modified
    assert props.revision != source_props.revision
    assert props.total_editing_time != source_props.total_editing_time
