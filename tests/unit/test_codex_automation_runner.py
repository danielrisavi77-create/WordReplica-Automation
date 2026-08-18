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
                (run_dir / "event_trace.jsonl").write_text("\n".join([
                    json.dumps({
                        "event_index": 0,
                        "event_type": "InsertText",
                        "source_element_id": "r1",
                        "status": "before",
                        "timestamp_utc": "2026-08-14T00:00:00+00:00",
                        "table_element_id": None,
                        "cell_element_id": None,
                    }),
                    json.dumps({
                        "event_index": 0,
                        "event_type": "InsertText",
                        "source_element_id": "r1",
                        "status": "after",
                        "timestamp_utc": "2026-08-14T00:00:01.250000+00:00",
                        "table_element_id": None,
                        "cell_element_id": None,
                    }),
                ]), encoding="utf-8")
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
        environment_capture=lambda: {"os": {"system": "TestOS"}, "word": {"available": True, "build": "99.0"}},
    )
    report_path = runner.run()

    assert report_path.parent.parent == root / "diagnostics" / "golden_1"
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["golden_id"] == "golden_1"
    assert report["gates_main_promotion"] is True
    assert report["commit_sha"] == "deadbeef"
    assert report["source_unchanged"] is True
    assert report["environment"] == {"os": {"system": "TestOS"}, "word": {"available": True, "build": "99.0"}}
    assert report["performance_profile"]["paired_event_count"] == 1
    assert report["performance_profile"]["by_event_type"][0]["total_seconds"] == 1.25
    assert golden.read_bytes() == b"golden"
    assert len(calls) == 3
    interactive_call = next(call for call in calls if "interactive_maximum" in call)
    assert "--defer-l4-qa" in interactive_call
    assert not any((root / "work").iterdir())


def test_runner_runs_secondary_golden_document_independently(tmp_path):
    root = tmp_path / "local"
    golden_dir = root / "golden"
    golden_dir.mkdir(parents=True)
    (golden_dir / "golden.docx").write_bytes(b"golden-1")
    (golden_dir / "golden2.docx").write_bytes(b"golden-2")
    repo = tmp_path / "repo"
    (repo / "scripts" / "remote_harness").mkdir(parents=True)
    (repo / "scripts" / "codex_automation").mkdir(parents=True)
    config = CodexAutomationConfig(
        local_root=root, golden_filename="golden.docx", additional_golden_filenames=["golden2.docx"],
    )
    seen_sources = []

    def executor(command, **kwargs):
        if "--stage" in command:
            run_dir = _arg(command, "--run-dir")
            run_dir.mkdir(parents=True, exist_ok=True)
            seen_sources.append(_arg(command, "--source").read_bytes())
            (run_dir / "result.json").write_text(json.dumps({"status": "COM_FAIL", "reasons": ["no output"]}), encoding="utf-8")
        return ChildResult(0, False, 0.1, [])

    runner = GoldenRunner(
        repo_root=repo,
        config=config,
        child_executor=executor,
        git_info=lambda root: ("automation-dev", "deadbeef"),
    )
    report_path = runner.run(golden_id="golden_2")
    report = json.loads(report_path.read_text(encoding="utf-8"))

    assert seen_sources == [b"golden-2", b"golden-2"]
    assert report["golden_id"] == "golden_2"
    assert report["golden_filename"] == "golden2.docx"
    assert report["gates_main_promotion"] is False
    assert report_path.parent.parent == root / "diagnostics" / "golden_2"
    assert not (root / "diagnostics" / "golden_1").exists()
    assert (root / "state" / "automation_state_golden_2.json").exists()
    assert not (root / "state" / "automation_state_golden_1.json").exists()


def test_runner_survives_environment_capture_failure(tmp_path):
    root = tmp_path / "local"
    golden_dir = root / "golden"
    golden_dir.mkdir(parents=True)
    golden = golden_dir / "golden.docx"
    golden.write_bytes(b"golden")
    repo = tmp_path / "repo"
    (repo / "scripts" / "remote_harness").mkdir(parents=True)
    (repo / "scripts" / "codex_automation").mkdir(parents=True)
    config = CodexAutomationConfig(local_root=root, golden_filename="golden.docx")

    def executor(command, **kwargs):
        if "--stage" in command:
            run_dir = _arg(command, "--run-dir")
            run_dir.mkdir(parents=True, exist_ok=True)
            (run_dir / "result.json").write_text(json.dumps({"status": "COM_FAIL", "reasons": ["no output"]}), encoding="utf-8")
        return ChildResult(0, False, 0.1, [])

    def failing_environment_capture():
        raise RuntimeError("font enumeration blew up")

    runner = GoldenRunner(
        repo_root=repo,
        config=config,
        child_executor=executor,
        git_info=lambda root: ("automation-dev", "deadbeef"),
        environment_capture=failing_environment_capture,
    )
    report_path = runner.run()

    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["environment"]["error"] == "font enumeration blew up"
    assert report["environment"]["exception_type"] == "RuntimeError"


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


def test_visible_word_is_forwarded_only_to_interactive_child(tmp_path):
    root = tmp_path / "local"
    (root / "golden").mkdir(parents=True)
    (root / "golden" / "golden.docx").write_bytes(b"golden")
    repo = tmp_path / "repo"
    (repo / "scripts" / "remote_harness").mkdir(parents=True)
    (repo / "scripts" / "codex_automation").mkdir(parents=True)
    calls = []

    def executor(command, **kwargs):
        calls.append(list(command))
        if "--stage" in command:
            stage = command[command.index("--stage") + 1]
            run_dir = _arg(command, "--run-dir")
            run_dir.mkdir(parents=True, exist_ok=True)
            if stage == "interactive_maximum":
                (run_dir / "output.docx").write_bytes(b"output")
            (run_dir / "result.json").write_text(
                json.dumps({"stage": stage, "run_status": "PASS", "reasons": []}),
                encoding="utf-8",
            )
        else:
            _arg(command, "--report").write_text(
                json.dumps({
                    "gates": {f"G{i}": False for i in range(10)},
                    "gate_details": {},
                    "full_pass": False,
                }),
                encoding="utf-8",
            )
        return ChildResult(0, False, 0.1, [])

    GoldenRunner(
        repo_root=repo,
        config=CodexAutomationConfig(local_root=root, golden_filename="golden.docx"),
        child_executor=executor,
        git_info=lambda root: ("automation-dev", "deadbeef"),
        visible_word=True,
    ).run()

    static_call = next(call for call in calls if "static" in call)
    interactive_call = next(call for call in calls if "interactive_maximum" in call)
    assert "--visible-word" not in static_call
    assert "--visible-word" in interactive_call


def test_runner_restarts_word_once_and_resumes_checkpoint_after_com_failure(tmp_path):
    root = tmp_path / "local"
    (root / "golden").mkdir(parents=True)
    (root / "golden" / "golden.docx").write_bytes(b"golden")
    repo = tmp_path / "repo"
    (repo / "scripts" / "remote_harness").mkdir(parents=True)
    (repo / "scripts" / "codex_automation").mkdir(parents=True)
    calls = []

    def executor(command, **kwargs):
        calls.append(list(command))
        if "--stage" in command:
            stage = command[command.index("--stage") + 1]
            run_dir = _arg(command, "--run-dir")
            run_dir.mkdir(parents=True, exist_ok=True)
            if stage == "static":
                payload = {"stage": stage, "status": "PASS", "run_status": "PASS", "reasons": []}
            elif "--resume-project-id" not in command:
                payload = {
                    "stage": stage,
                    "status": "COM_FAIL",
                    "run_status": "FAIL",
                    "project_id": "project-1",
                    "reasons": ["Call was rejected by callee."],
                }
            else:
                assert command[command.index("--resume-project-id") + 1] == "project-1"
                payload = {"stage": stage, "status": "PASS", "run_status": "PASS", "project_id": "project-1", "reasons": []}
                (run_dir / "output.docx").write_bytes(b"resumed-output")
            (run_dir / "result.json").write_text(json.dumps(payload), encoding="utf-8")
        else:
            _arg(command, "--report").write_text(json.dumps({
                "gates": {f"G{i}": True for i in range(10)},
                "gate_details": {},
                "full_pass": True,
            }), encoding="utf-8")
        return ChildResult(0, False, 0.1, [])

    report_path = GoldenRunner(
        repo_root=repo,
        config=CodexAutomationConfig(local_root=root, golden_filename="golden.docx"),
        child_executor=executor,
        git_info=lambda _root: ("automation-dev", "deadbeef"),
    ).run()

    interactive_calls = [call for call in calls if "interactive_maximum" in call]
    assert len(interactive_calls) == 2
    assert "--resume-project-id" not in interactive_calls[0]
    assert "--resume-project-id" in interactive_calls[1]
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["full_pass"] is True
    assert len(report["interactive_processes"]) == 2
