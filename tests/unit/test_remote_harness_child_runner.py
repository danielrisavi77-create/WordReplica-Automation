import json
from pathlib import Path

from scripts.remote_harness.child_runner import _interactive_service_for, _options_for, serialize_run_result, classify_run_result
from word_replica.domain.enums import RunStatus, VisibilityMode
from word_replica.domain.results import RunResult, WarningItem
from scripts.remote_harness.contracts import HarnessStatus


def test_child_serializes_run_result_and_normalizes_com_failure(tmp_path):
    result=RunResult(RunStatus.FAIL, None, None, project_id="p1", reasons=["Call was rejected by callee."])
    payload=serialize_run_result("x.docx","interactive_maximum",result,1.5)
    assert classify_run_result(payload) is HarnessStatus.COM_FAIL
    assert payload["project_id"] == "p1"
    assert payload["elapsed_seconds"] == 1.5


def test_expected_block_is_detectable_from_static_preflight():
    payload={"run_status":"FAIL","reasons":["unsupported"],"preflight_can_proceed":False,"stage":"interactive_maximum"}
    assert classify_run_result(payload) is HarnessStatus.EXPECTED_BLOCK


def test_interactive_maximum_golden_stage_runs_word_in_background():
    assert _options_for("interactive_maximum").visibility is VisibilityMode.BACKGROUND


def test_golden_child_can_defer_internal_l4_qa():
    service = _interactive_service_for("interactive_maximum", defer_l4_qa=True)

    assert service is not None
    assert service.run_l4_qa is False


def test_interactive_options_can_make_word_visible():
    assert _options_for("interactive_maximum", visible_word=True).visibility is VisibilityMode.VISIBLE


def test_benchmark_can_explicitly_disable_table_fast_path_without_changing_default():
    assert _options_for("interactive_maximum").interactive.enable_table_fast_path is True
    assert _options_for(
        "interactive_maximum", disable_table_fast_path=True,
    ).interactive.enable_table_fast_path is False
