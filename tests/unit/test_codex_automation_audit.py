from scripts.codex_automation.audit import GateResult, build_golden_report, compare_page_text_partitions


def test_page_partition_requires_same_page_count_and_text():
    good = compare_page_text_partitions(["A", "B"], ["A", "B"])
    assert good.passed is True

    moved = compare_page_text_partitions(["A", "B"], ["A B"])
    assert moved.passed is False
    assert moved.details["source_page_count"] == 2
    assert moved.details["output_page_count"] == 1


def test_page_partition_ignores_pdf_glyph_spacing_but_detects_moved_characters():
    glyph_spacing = compare_page_text_partitions(
        ["čini kvalitet a odluka", "sljedeća stranica"],
        ["čini kvaliteta odluka", "sljedeća stranica"],
    )
    assert glyph_spacing.passed is True

    moved = compare_page_text_partitions(
        ["ABC", "DEF"],
        ["AB", "CDEF"],
    )
    assert moved.passed is False
    assert moved.first_divergence["page"] == 1


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


def test_visual_gate_allows_low_mae_antialiasing_noise():
    from word_replica.qa.render import RenderQaResult, VisualMetric
    from scripts.codex_automation.audit import build_visual_gate

    render = RenderQaResult(
        available=True,
        within_tolerance=False,
        page_count_match=True,
        source_page_count=1,
        rebuilt_page_count=1,
        metrics=[VisualMetric(True, 0.0029, 0.111, (100, 100), (100, 100))],
    )

    gate = build_visual_gate(render, changed_pixel_tolerance=0.001, mae_tolerance=0.25)

    assert gate.passed is True
    assert gate.first_divergence is None


def test_visual_gate_allows_measured_word_text_rasterization_noise():
    from word_replica.qa.render import RenderQaResult, VisualMetric
    from scripts.codex_automation.audit import build_visual_gate

    render = RenderQaResult(
        available=True,
        within_tolerance=False,
        page_count_match=True,
        source_page_count=1,
        rebuilt_page_count=1,
        metrics=[VisualMetric(True, 0.0234, 0.805, (100, 100), (100, 100))],
    )

    gate = build_visual_gate(render, changed_pixel_tolerance=0.001, mae_tolerance=0.25)

    assert gate.passed is True


def test_visual_gate_uses_blurred_mae_to_separate_word_antialiasing_from_layout_error():
    from word_replica.qa.render import RenderQaResult, VisualMetric
    from scripts.codex_automation.audit import build_visual_gate

    render = RenderQaResult(
        available=True,
        within_tolerance=False,
        page_count_match=True,
        source_page_count=2,
        rebuilt_page_count=2,
        metrics=[
            VisualMetric(
                True,
                0.0287,
                1.804,
                (1191, 1684),
                (1191, 1684),
                blurred_mean_absolute_error=0.914,
            ),
            VisualMetric(
                True,
                0.0885,
                7.7,
                (1191, 1684),
                (1191, 1684),
                blurred_mean_absolute_error=4.8,
            ),
        ],
    )

    gate = build_visual_gate(render, changed_pixel_tolerance=0.001, mae_tolerance=0.25)

    assert gate.passed is False
    assert gate.first_divergence["page"] == 2
    assert gate.first_divergence["blurred_mean_absolute_error"] == 4.8
    assert gate.details["blurred_antialiasing_policy"] == {
        "changed_pixel_ratio_allowance": 0.04,
        "blurred_mae_allowance": 1.0,
        "gaussian_blur_radius": 1.0,
    }
    assert gate.details["blur_assisted_pages"] == [1]
    assert gate.details["acceptance_mode_counts"] == {
        "strict": 0,
        "legacy_antialiasing": 0,
        "blurred_antialiasing": 1,
        "failed": 1,
    }


def _passing(name):
    return GateResult(name=name, passed=True, summary="ok")


def _report(gates, **kwargs):
    return build_golden_report(
        run_id="r", source_sha256="s", commit_sha="c",
        reconstruction_status="PASS", gates=gates, **kwargs,
    )


def test_default_report_declares_the_original_ten_gates():
    # Golden #1's contract: ten gates, and a report that does not opt in to a
    # different gate set must keep declaring exactly those ten.
    gates = {f"G{i}": _passing(f"G{i}") for i in range(10)}
    report = _report(gates)

    assert report["required_gates"] == [f"G{i}" for i in range(10)]
    assert report["full_pass"] is True
    assert list(report["gates"]) == [f"G{i}" for i in range(10)]


def test_extra_gate_is_ignored_unless_it_is_declared():
    # G10 present in the gates dict but not declared must not appear in the
    # report, so wiring the lab's gate in cannot leak into a Golden run.
    gates = {f"G{i}": _passing(f"G{i}") for i in range(11)}
    report = _report(gates)

    assert "G10" not in report["gates"]
    assert report["required_gates"] == [f"G{i}" for i in range(10)]


def test_declared_gate_names_drive_reporting_and_full_pass():
    names = [f"G{i}" for i in range(11)]
    gates = {name: _passing(name) for name in names}
    report = _report(gates, gate_names=names)

    assert report["required_gates"] == names
    assert report["gates"]["G10"] is True
    assert report["full_pass"] is True


def test_declared_gate_missing_from_results_is_not_a_full_pass():
    names = [f"G{i}" for i in range(11)]
    gates = {name: _passing(name) for name in names[:-1]}
    report = _report(gates, gate_names=names)

    assert report["full_pass"] is False
    assert report["first_divergent_gate"] == "G10"
