import json

import pytest

from scripts.codex_automation.autonomous_loop import AutonomousLoopLock, decide_next_step, run_loop


def report(*, full_pass=False, stop_required=False, promotion_ready=False):
    return {
        "full_pass": full_pass,
        "first_divergent_gate": "G0",
        "first_divergence": {"error": "evidence"},
        "automation_decision": {
            "stop_required": stop_required,
            "reason": "three consecutive iterations without gate improvement" if stop_required else None,
            "promotion_ready": promotion_ready,
        },
    }


def test_failed_golden_requires_a_production_fix_before_repeating():
    decision = decide_next_step(report())
    assert decision.action == "PRODUCTION_FIX_REQUIRED"
    assert decision.exit_code == 2


def test_fail_safe_has_priority_over_automatic_retries():
    decision = decide_next_step(report(stop_required=True))
    assert decision.action == "STOP_FAIL_SAFE"
    assert decision.exit_code == 3


def test_first_full_pass_requests_same_commit_confirmation():
    decision = decide_next_step(report(full_pass=True))
    assert decision.action == "REPEAT_FULL_PASS"
    assert decision.exit_code == 0


def test_promotion_ready_ends_the_loop():
    decision = decide_next_step(report(full_pass=True, promotion_ready=True))
    assert decision.action == "COMPLETE"
    assert decision.exit_code == 0


def test_loop_lock_is_exclusive_and_cleans_up(tmp_path):
    lock_path = tmp_path / "golden_autonomous_loop.lock"
    with AutonomousLoopLock(lock_path, pid=123):
        assert json.loads(lock_path.read_text(encoding="utf-8")) == {"pid": 123}
        with pytest.raises(RuntimeError, match="already active"):
            with AutonomousLoopLock(lock_path, pid=456):
                pass
    assert not lock_path.exists()


def _config_file(tmp_path):
    path = tmp_path / "config.json"
    path.write_text(json.dumps({"local_root": str(tmp_path / "local")}), encoding="utf-8")
    return path


def test_loop_does_not_start_golden_when_full_pytest_fails(tmp_path):
    calls = []

    def pytest_runner(repo_root, python_executable):
        calls.append("pytest")
        return 1

    def golden_runner_factory(**kwargs):
        calls.append("golden")
        raise AssertionError("Golden must not run after a red full suite")

    assert run_loop(
        repo_root=tmp_path,
        config_path=_config_file(tmp_path),
        local_root=tmp_path / "local",
        python_executable="python",
        pytest_runner=pytest_runner,
        golden_runner_factory=golden_runner_factory,
    ) == 2
    assert calls == ["pytest"]


def test_loop_does_not_bypass_existing_fail_safe_state(tmp_path):
    local_root = tmp_path / "local"
    state_dir = local_root / "state"
    state_dir.mkdir(parents=True)
    (state_dir / "automation_state.json").write_text(
        json.dumps({"no_improvement_count": 3}), encoding="utf-8"
    )
    calls = []
    assert run_loop(
        repo_root=tmp_path,
        config_path=_config_file(tmp_path),
        local_root=local_root,
        python_executable="python",
        pytest_runner=lambda *_: calls.append("pytest") or 0,
        golden_runner_factory=lambda **_: calls.append("golden"),
    ) == 3
    assert calls == []


def test_loop_requires_explicit_review_acknowledgement_to_resume(tmp_path):
    local_root = tmp_path / "local"
    state_dir = local_root / "state"
    state_dir.mkdir(parents=True)
    (state_dir / "automation_state.json").write_text(
        json.dumps({"no_improvement_count": 3}), encoding="utf-8"
    )
    report_path = tmp_path / "report.json"
    report_path.write_text(json.dumps(report()), encoding="utf-8")
    calls = []

    class FakeGolden:
        def __init__(self, **kwargs): pass
        def run(self):
            calls.append("golden")
            return report_path

    assert run_loop(
        repo_root=tmp_path,
        config_path=_config_file(tmp_path),
        local_root=local_root,
        python_executable="python",
        pytest_runner=lambda *_: calls.append("pytest") or 0,
        golden_runner_factory=FakeGolden,
        resume_after_review=True,
    ) == 2
    assert calls == ["pytest", "golden"]


def test_loop_repeats_only_for_first_full_pass_then_stops_on_promotion(tmp_path):
    calls = []
    reports = []
    for index in range(2):
        report_path = tmp_path / f"report-{index}.json"
        report_path.write_text(json.dumps(report(full_pass=True, promotion_ready=index == 1)), encoding="utf-8")
        reports.append(report_path)

    class FakeGolden:
        def __init__(self, **kwargs):
            pass
        def run(self):
            calls.append("golden")
            return reports.pop(0)

    assert run_loop(
        repo_root=tmp_path,
        config_path=_config_file(tmp_path),
        local_root=tmp_path / "local",
        python_executable="python",
        pytest_runner=lambda *_: calls.append("pytest") or 0,
        golden_runner_factory=FakeGolden,
    ) == 0
    assert calls == ["pytest", "golden", "pytest", "golden"]
