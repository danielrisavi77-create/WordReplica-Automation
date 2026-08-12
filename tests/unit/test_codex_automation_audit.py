from scripts.codex_automation.audit import GateResult, build_golden_report, compare_page_text_partitions


def test_page_partition_requires_same_page_count_and_text():
    good = compare_page_text_partitions(["A", "B"], ["A", "B"])
    assert good.passed is True

    moved = compare_page_text_partitions(["A", "B"], ["A B"])
    assert moved.passed is False
    assert moved.details["source_page_count"] == 2
    assert moved.details["output_page_count"] == 1


def test_report_selects_first_failed_gate_in_g0_to_g9_order():
    gates = {
        f"G{i}": GateResult(name=f"G{i}", passed=True, summary="ok")
        for i in range(10)
    }
    gates["G3"] = GateResult(name="G3", passed=False, summary="table mismatch", first_divergence={"path": "table/0"})
    gates["G8"] = GateResult(name="G8", passed=False, summary="pagination mismatch")

    report = build_golden_report(
        run_id="r1",
        source_sha256="abc",
        commit_sha="deadbeef",
        reconstruction_status="WARN",
        gates=gates,
        source_page_count=72,
        output_page_count=73,
    )

    assert report["full_pass"] is False
    assert report["first_divergent_gate"] == "G3"
    assert report["first_divergence"]["path"] == "table/0"


def test_visual_gate_reports_first_page_over_tolerance():
    from word_replica.qa.render import RenderQaResult, VisualMetric
    from scripts.codex_automation.audit import build_visual_gate

    render = RenderQaResult(
        available=True,
        within_tolerance=False,
        page_count_match=True,
        source_page_count=2,
        rebuilt_page_count=2,
        metrics=[
            VisualMetric(True, 0.0001, 0.1, (100, 100), (100, 100)),
            VisualMetric(True, 0.01, 2.0, (100, 100), (100, 100)),
        ],
    )

    gate = build_visual_gate(render, changed_pixel_tolerance=0.001, mae_tolerance=0.25)

    assert gate.passed is False
    assert gate.first_divergence["page"] == 2
    assert gate.first_divergence["changed_pixel_ratio"] == 0.01
