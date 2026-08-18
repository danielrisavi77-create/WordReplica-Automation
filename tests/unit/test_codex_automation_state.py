from scripts.codex_automation.state import AutomationState, evaluate_run


def _gates(*passed):
    return {f"G{i}": (f"G{i}" in passed) for i in range(10)}


def test_no_improvement_stops_after_three_consecutive_iterations():
    state = AutomationState()
    report = {"gates": _gates("G0", "G1")}
    first = evaluate_run(state, report)
    assert first.stop_required is False
    for _ in range(2):
        decision = evaluate_run(state, report)
        assert decision.stop_required is False
    decision = evaluate_run(state, report)
    assert decision.stop_required is True
    assert decision.reason == "three consecutive iterations without gate improvement"


def test_regression_of_previously_passing_gate_stops_immediately():
    state = AutomationState(last_gate_status=_gates("G0", "G1", "G2"), best_score=3)
    report = {"gates": _gates("G0", "G2")}

    decision = evaluate_run(state, report)

    assert decision.stop_required is True
    assert "G1" in decision.reason


def test_two_full_passes_on_same_commit_make_promotion_ready():
    state = AutomationState()
    full = {"gates": _gates(*[f"G{i}" for i in range(10)]), "full_pass": True, "commit_sha": "abc"}

    first = evaluate_run(state, full)
    second = evaluate_run(state, full)

    assert first.promotion_ready is False
    assert second.promotion_ready is True
