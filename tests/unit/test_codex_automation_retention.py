import json
from pathlib import Path

from scripts.codex_automation.retention import prune_diagnostics


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
