import json
from pathlib import Path

from scripts.codex_automation.retention import prune_diagnostics, prune_stale_work


def _run(root: Path, name: str, *, full_pass: bool, pinned: bool = False):
    path = root / name
    path.mkdir(parents=True)
    (path / "golden_report.json").write_text(json.dumps({"full_pass": full_pass, "pinned": pinned}), encoding="utf-8")
    return path


def test_retention_keeps_latest_success_two_failures_and_pinned(tmp_path):
    root = tmp_path / "diagnostics"
    root.mkdir()
    pinned = _run(root, "001", full_pass=False, pinned=True)
    _run(root, "002", full_pass=True)
    _run(root, "003", full_pass=False)
    _run(root, "004", full_pass=False)
    latest_success = _run(root, "005", full_pass=True)
    fail1 = _run(root, "006", full_pass=False)
    fail2 = _run(root, "007", full_pass=False)

    kept = prune_diagnostics(root, keep_success=1, keep_failures=2)

    assert set(p.name for p in kept) == {pinned.name, latest_success.name, fail1.name, fail2.name}
    assert not (root / "002").exists()
    assert not (root / "003").exists()
    assert not (root / "004").exists()


def test_retention_ignores_non_run_diagnostic_directories(tmp_path):
    root = tmp_path / "diagnostics"
    root.mkdir()
    older_failure = _run(root, "20260816T010000Z_aaaaaaaa", full_pass=False)
    latest_failure = _run(root, "20260816T020000Z_bbbbbbbb", full_pass=False)
    pytest_temp = root / "pytest-full-field-format"
    pytest_temp.mkdir()

    kept = prune_diagnostics(root, keep_success=0, keep_failures=1)

    assert not older_failure.exists()
    assert latest_failure.exists()
    assert pytest_temp.exists()
    assert pytest_temp not in kept


def test_stale_work_cleanup_removes_only_golden_run_directories(tmp_path):
    work = tmp_path / "work"
    stale = work / "20260904T223835Z_deadbeef"
    suffixed = work / "20260904T223835Z_deadbeef_2"
    manual = work / "manual-investigation"
    for path in (stale, suffixed, manual):
        path.mkdir(parents=True)
        (path / "evidence.txt").write_text("keep-or-delete", encoding="utf-8")

    removed = prune_stale_work(work)

    assert removed == [stale, suffixed]
    assert not stale.exists()
    assert not suffixed.exists()
    assert manual.exists()
