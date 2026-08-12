import json
from pathlib import Path

from scripts.remote_harness.child_runner import _options_for, serialize_run_result, classify_run_result
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
