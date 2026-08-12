import json
from pathlib import Path

from scripts.codex_automation.config import CodexAutomationConfig
from scripts.codex_automation.process import ChildResult
from scripts.codex_automation.runner import GoldenRunner


def _arg(command, name):
    return Path(command[command.index(name) + 1])


def test_runner_produces_report_moves_run_and_preserves_golden(tmp_path):
    root = tmp_path / "local"
    golden_dir = root / "golden"
    golden_dir.mkdir(parents=True)
    golden = golden_dir / "golden.docx"
    golden.write_bytes(b"golden")
    repo = tmp_path / "repo"
    (repo / "scripts" / "remote_harness").mkdir(parents=True)
    (repo / "scripts" / "codex_automation").mkdir(parents=True)
    config = CodexAutomationConfig(local_root=root, golden_filename="golden.docx")
    calls = []

    def executor(command, **kwargs):
        calls.append(list(command))
        if "--stage" in command:
            stage = command[command.index("--stage") + 1]
            run_dir = _arg(command, "--run-dir")
            run_dir.mkdir(parents=True, exist_ok=True)
            payload = {"stage": stage, "status": "PASS", "run_status": "PASS", "reasons": []}
            if stage == "interactive_maximum":
                (run_dir / "output.docx").write_bytes(b"output")
            (run_dir / "result.json").write_text(json.dumps(payload), encoding="utf-8")
        else:
            report_path = _arg(command, "--report")
            report_path.write_text(json.dumps({
                "run_id": "x",
                "source_sha256": "abc",
                "commit_sha": "deadbeef",
                "reconstruction_status": "PASS",
                "gates": {f"G{i}": False for i in range(10)},
                "gate_details": {},
                "full_pass": False,
                "first_divergent_gate": "G0",
                "first_divergence": {"path": "/"},
            }), encoding="utf-8")
        return ChildResult(0, False, 0.1, [])

    runner = GoldenRunner(
        repo_root=repo,
        config=config,
        child_executor=executor,
        git_info=lambda root: ("automation-dev", "deadbeef"),
    )
    report_path = runner.run()

    assert report_path.parent.parent == root / "diagnostics"
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["commit_sha"] == "deadbeef"
    assert report["source_unchanged"] is True
    assert golden.read_bytes() == b"golden"
    assert len(calls) == 3
    assert not any((root / "work").iterdir())


def test_runner_refuses_autonomous_run_on_main(tmp_path):
    root = tmp_path / "local"
    (root / "golden").mkdir(parents=True)
    (root / "golden" / "golden.docx").write_bytes(b"golden")
    repo = tmp_path / "repo"
    repo.mkdir()
    runner = GoldenRunner(
        repo_root=repo,
        config=CodexAutomationConfig(local_root=root, golden_filename="golden.docx"),
        child_executor=lambda *a, **k: None,
        git_info=lambda root: ("main", "abc"),
    )

    try:
        runner.run()
    except RuntimeError as exc:
        assert "automation-dev" in str(exc)
    else:
        raise AssertionError("runner should refuse main")
