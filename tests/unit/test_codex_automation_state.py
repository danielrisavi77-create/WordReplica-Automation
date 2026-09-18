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


def test_first_clean_full_pass_resets_stale_no_improvement_fail_safe():
    gates = _gates(*[f"G{i}" for i in range(10)])
    state = AutomationState(
        last_gate_status=gates,
        best_score=10,
        no_improvement_count=21,
    )
    full = {
        "gates": gates,
        "full_pass": True,
        "commit_sha": "new-clean-commit",
        "worktree_clean": True,
    }

    decision = evaluate_run(state, full)

    assert decision.stop_required is False
    assert decision.promotion_ready is False
    assert state.no_improvement_count == 0
    assert state.consecutive_full_pass_same_commit == 1


def test_dirty_worktree_full_pass_never_counts_toward_promotion():
    state = AutomationState()
    full = {
        "gates": _gates(*[f"G{i}" for i in range(10)]),
        "full_pass": True,
        "commit_sha": "abc",
        "worktree_clean": False,
    }

    assert evaluate_run(state, full).promotion_ready is False
    assert evaluate_run(state, full).promotion_ready is False
    assert state.consecutive_full_pass_same_commit == 0


def test_full_pass_on_different_word_build_does_not_count_as_consecutive():
    state = AutomationState()
    gates = _gates(*[f"G{i}" for i in range(10)])

    def full_pass_report(build):
        return {
            "gates": gates,
            "full_pass": True,
            "commit_sha": "abc",
            "environment": {"word": {"build": build}},
        }

    first = evaluate_run(state, full_pass_report("16.0.1"))
    second = evaluate_run(state, full_pass_report("16.0.2"))
    third = evaluate_run(state, full_pass_report("16.0.2"))

    assert first.promotion_ready is False
    assert second.promotion_ready is False
    assert third.promotion_ready is True
    assert state.consecutive_full_pass_same_commit == 2


def test_ten_gate_report_without_required_gates_still_promotes_exactly_as_before():
    # Golden #1's reports predate `required_gates` and must keep promoting on
    # exactly the ten gates they always had. This pins the default so the
    # G10 work cannot silently change the main-promotion gate.
    state = AutomationState()
    full = {"gates": _gates(*[f"G{i}" for i in range(10)]), "full_pass": True, "commit_sha": "abc"}

    assert evaluate_run(state, full).promotion_ready is False
    assert evaluate_run(state, full).promotion_ready is True


def test_eleven_gate_report_promotes_only_when_all_declared_gates_pass():
    state = AutomationState()
    gates = {f"G{i}": True for i in range(11)}
    required = [f"G{i}" for i in range(11)]
    full = {"gates": gates, "full_pass": True, "commit_sha": "abc", "required_gates": required}

    assert evaluate_run(state, full).promotion_ready is False
    assert evaluate_run(state, full).promotion_ready is True


def test_eleven_gate_report_missing_the_declared_extra_gate_does_not_promote():
    # A report that declares eleven required gates but only carries ten must
    # never reach FULL PASS -- otherwise adding G10 would let an unevaluated
    # gate count as passing.
    state = AutomationState()
    full = {
        "gates": _gates(*[f"G{i}" for i in range(10)]),
        "full_pass": True,
        "commit_sha": "abc",
        "required_gates": [f"G{i}" for i in range(11)],
    }

    evaluate_run(state, full)
    assert evaluate_run(state, full).promotion_ready is False
