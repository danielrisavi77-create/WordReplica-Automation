from copy import deepcopy

import pytest

from scripts.codex_automation.word_gate_autofix import (
    StopRepair, check_changes, check_golden, check_red, eligible,
)


def golden():
    return {"gates": {f"G{i}": i < 8 for i in range(10)},
            "source_sha256": "abc", "source_unchanged": True,
            "automation_decision": {"stop_required": False}}


def test_only_human_automation_branch_can_start_repair():
    assert eligible("automation-dev", "fix document")
    assert not eligible("main", "fix document")
    assert not eligible("release/rc1", "fix document")
    assert not eligible("automation-dev", "fix\n\nWordReplica-Autofix: true")


def test_agent_cannot_weaken_existing_tests_or_edit_gate():
    regression = "tests/unit/test_autofix_abc.py"
    for path in (".github/workflows/ci.yml", "AGENTS.md", "tests/unit/test_existing.py",
                 "src/word_replica/repair_contract/signing.py", "src/word_replica/renderers/word_ownership.py", "RUN_WINDOWS_RELEASE_GATE.ps1",
                 "src/word_replica/parser/../../qa/gates.py"):
        with pytest.raises(StopRepair):
            check_changes([path, regression], regression, phase="fix")
    check_changes([regression, "src/word_replica/parser/fidelity.py"], regression, phase="fix")
    with pytest.raises(StopRepair):
        check_changes([regression, "src/word_replica/parser/fidelity.py"], regression, phase="test")


def test_red_requires_actual_failed_assertion_not_collection_or_skip(tmp_path):
    path = tmp_path / "red.xml"
    for xml, code in (("<testsuites><testsuite errors='1'/></testsuites>", 1),
                      ("<testsuites><testsuite skipped='1'/></testsuites>", 1),
                      ("<testsuites><testsuite failures='1'/></testsuites>", 2)):
        path.write_text(xml)
        with pytest.raises(StopRepair):
            check_red(code, path)
    path.write_text("<testsuites><testsuite tests='1' failures='1' errors='0'/></testsuites>")
    check_red(1, path)


def test_golden_regression_or_integrity_failure_stops():
    before = golden()
    for change in (lambda r: r["gates"].update(G0=False),
                   lambda r: r.update(source_sha256="changed"),
                   lambda r: r.update(source_unchanged=False),
                   lambda r: r["automation_decision"].update(stop_required=True),
                   lambda r: r["gates"].pop("G9")):
        after = deepcopy(before)
        change(after)
        with pytest.raises(StopRepair):
            check_golden(before, after)
    after = deepcopy(before)
    after["gates"]["G8"] = True
    check_golden(before, after)
    check_golden(before, before)  # release gate improvement is separately required


@pytest.fixture
def repair_harness(tmp_path, monkeypatch):
    import json
    import subprocess
    from pathlib import Path
    from scripts.codex_automation import word_gate_autofix as repair

    source = tmp_path / "source"
    source.mkdir()
    subprocess.run(["git", "init", "-b", "automation-dev", str(source)], check=True, capture_output=True)
    root = tmp_path / "local"
    (root / "golden").mkdir(parents=True)
    (root / "golden/source.docx").write_bytes(b"immutable fixture")
    (root / ".venv/Scripts").mkdir(parents=True)
    (root / ".venv/Scripts/python.exe").touch()
    (source / ".gitignore").write_text(".venv_gate/\n")
    (source / "codex_automation.json").write_text(json.dumps({"golden_filename": "source.docx"}))
    (source / "src/word_replica/parser").mkdir(parents=True)
    (source / "src/word_replica/parser/example.py").write_text("VALUE = 0\n")
    (source / "tests/unit").mkdir(parents=True)
    (source / "tests/unit/test_existing.py").write_text("def test_existing(): pass\n")
    repair.git(source, "add", ".")
    repair.git(source, "-c", "user.name=Test", "-c", "user.email=test@example.com", "commit", "-m", "human change")
    base = repair.git(source, "rev-parse", "HEAD")
    (source / ".venv_gate/Scripts").mkdir(parents=True)
    (source / ".venv_gate/Scripts/python.exe").touch()
    baseline = golden()
    baseline["source_sha256"] = repair.digest(root / "golden/source.docx")
    baseline["commit_sha"] = base
    report_dir = root / "diagnostics/baseline"
    report_dir.mkdir(parents=True)
    (report_dir / "golden_report.json").write_text(json.dumps(baseline))
    run = root / "autofix/run"
    run.mkdir(parents=True)
    failure = run / "word-gate.log"
    failure.write_text("FAILED acceptance: concrete fixture failure\nWORD REPLICA WINDOWS RELEASE GATE: TEST FAIL")
    monkeypatch.setenv("GITHUB_REF_NAME", "automation-dev")
    monkeypatch.setenv("GH_TOKEN", "test-do-not-inherit")
    monkeypatch.setattr(repair.shutil, "which", lambda _: "codex")
    calls = []
    mode = {"value": "success"}

    def fake_run(args, *, repo, log, env, prompt=None, timeout=5400):
        assert "GH_TOKEN" not in env
        log.write_text("verification feedback")
        calls.append(args)
        if args[0] == "codex":
            if "login" in args:
                return 0
            test = repo / f"tests/unit/test_autofix_{base[:12]}.py"
            if not test.exists():
                test.write_text("def test_regression(): assert False\n")
            else:
                (repo / "src/word_replica/parser/example.py").write_text("VALUE = 1\n")
                if mode["value"] == "tamper":
                    test.write_text("def test_regression(): pass\n")
            return 0
        xml = next((a.split("=", 1)[1] for a in args if a.startswith("--junitxml=")), None)
        if xml:
            Path(xml).write_text("<testsuites><testsuite failures='1'/></testsuites>")
            return 1
        if "pytest" in args:
            if mode["value"] == "exhausted":
                return 1
            if mode["value"] == "suite_failure" and args[-1] == "-q":
                return 1
        if args[0] == "powershell" and "-Golden" in args:
            after = deepcopy(baseline)
            if mode["value"] == "regressed":
                after["gates"]["G0"] = False
            new_dir = root / "diagnostics/new"
            new_dir.mkdir()
            (new_dir / "golden_report.json").write_text(json.dumps(after))
            return 1
        return 0

    monkeypatch.setattr(repair, "run_logged", fake_run)
    return repair, source, root, run, failure, calls, mode


def test_supervisor_commits_only_after_independent_checks(repair_harness):
    repair, source, root, run, failure, calls, mode = repair_harness
    sha = repair.supervise(source, root, run, failure)
    assert sha == repair.git(run / "candidate", "rev-parse", "HEAD")
    assert repair.MARKER in repair.git(run / "candidate", "log", "-1", "--format=%B")
    assert any("-Golden" in call for call in calls)
    assert repair.git(run / "candidate", "remote") == ""
    with pytest.raises(StopRepair, match="already ran"):
        repair.supervise(source, root, run.parent / "another", failure)


@pytest.mark.parametrize("mode_name,reason", [
    ("tamper", "altered"), ("regressed", "regressed"),
    ("suite_failure", "Full suite"), ("exhausted", "Three repair attempts"),
])
def test_supervisor_never_commits_failed_or_weakened_candidate(repair_harness, mode_name, reason):
    repair, source, root, run, failure, calls, mode = repair_harness
    mode["value"] = mode_name
    with pytest.raises(StopRepair, match=reason):
        repair.supervise(source, root, run, failure)
    assert repair.git(run / "candidate", "rev-parse", "HEAD") == repair.git(source, "rev-parse", "HEAD")
    if mode_name == "exhausted":
        assert sum(call[0] == "codex" and "exec" in call for call in calls) == 4
