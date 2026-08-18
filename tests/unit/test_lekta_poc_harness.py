import json
from datetime import datetime, timezone
from pathlib import Path

import pytest

from scripts.run_lekta_poc import HarnessPreflightError, preflight, run_harness


class FakeCompletedProcess:
    def __init__(self, *, stdout="", returncode=0):
        self.stdout = stdout
        self.returncode = returncode


class FakeGit:
    def __init__(self, *, branch="automation-dev", dirty_lines=()):
        self.branch = branch
        self.dirty_lines = dirty_lines
        self.calls = []

    def __call__(self, args, **kwargs):
        self.calls.append(args)
        if args[:2] == ["git", "rev-parse"] and "--abbrev-ref" in args:
            return FakeCompletedProcess(stdout=self.branch + "\n")
        if args[:2] == ["git", "rev-parse"]:
            return FakeCompletedProcess(stdout="deadbeef\n")
        if args[:2] == ["git", "status"]:
            body = "\n".join(self.dirty_lines)
            return FakeCompletedProcess(stdout=body + ("\n" if body else ""))
        raise AssertionError(f"unexpected git call: {args}")


def _write_package(tmp_path: Path) -> Path:
    package_dir = tmp_path / "package"
    package_dir.mkdir()
    (package_dir / "original.docx").write_bytes(b"original")
    (package_dir / "target.docx").write_bytes(b"target")
    (package_dir / "contract.json").write_text("{}", encoding="utf-8")
    (package_dir / "public-key.spki.b64url").write_text("abc", encoding="utf-8")
    return package_dir


def test_preflight_passes_on_the_required_branch_with_a_clean_tree_and_complete_package(tmp_path):
    package_dir = _write_package(tmp_path)
    info = preflight(repo_root=tmp_path, package_dir=package_dir, run_subprocess=FakeGit())
    assert info["branch"] == "automation-dev"
    assert info["commit_sha"] == "deadbeef"


def test_preflight_rejects_the_wrong_branch(tmp_path):
    package_dir = _write_package(tmp_path)
    with pytest.raises(HarnessPreflightError, match="automation-dev"):
        preflight(repo_root=tmp_path, package_dir=package_dir, run_subprocess=FakeGit(branch="main"))


def test_preflight_rejects_a_dirty_production_file(tmp_path):
    package_dir = _write_package(tmp_path)
    git = FakeGit(dirty_lines=[" M src/word_replica/cli.py"])
    with pytest.raises(HarnessPreflightError, match="dirty"):
        preflight(repo_root=tmp_path, package_dir=package_dir, run_subprocess=git)


def test_preflight_allows_dirty_design_and_plan_docs(tmp_path):
    package_dir = _write_package(tmp_path)
    git = FakeGit(dirty_lines=[" M docs/superpowers/plans/2026-08-18-lekta-target-docx-local-poc.md"])
    info = preflight(repo_root=tmp_path, package_dir=package_dir, run_subprocess=git)
    assert info["branch"] == "automation-dev"


def test_preflight_rejects_a_missing_package_file(tmp_path):
    package_dir = _write_package(tmp_path)
    (package_dir / "target.docx").unlink()
    with pytest.raises(HarnessPreflightError, match="target.docx"):
        preflight(repo_root=tmp_path, package_dir=package_dir, run_subprocess=FakeGit())


def test_run_harness_calls_the_cli_and_writes_diagnostics_outside_the_repo(tmp_path):
    package_dir = _write_package(tmp_path)
    diagnostics_root = tmp_path / "diagnostics"
    git = FakeGit()
    cli_calls = []

    def run_subprocess(args, **kwargs):
        if args and args[0] == "git":
            return git(args, **kwargs)
        cli_calls.append(args)
        return FakeCompletedProcess(returncode=0)

    fixed_now = datetime(2026, 8, 18, 12, 0, 0, tzinfo=timezone.utc)

    exit_code = run_harness(
        repo_root=tmp_path,
        python_executable=Path("python.exe"),
        package_dir=package_dir,
        diagnostics_root=diagnostics_root,
        renderer="pure-docx",
        now=fixed_now,
        run_subprocess=run_subprocess,
    )

    assert exit_code == 0
    diagnostics_dir = diagnostics_root / "repair-poc-20260818T120000Z"
    assert diagnostics_dir.is_dir()
    data = json.loads((diagnostics_dir / "harness_diagnostics.json").read_text(encoding="utf-8"))
    assert data["branch"] == "automation-dev"
    assert data["renderer"] == "pure-docx"
    assert data["exit_code"] == 0
    assert "source_sha256" in data

    called_args = cli_calls[0]
    assert "repair-poc" in called_args
    assert str(package_dir / "original.docx") in called_args
    assert "--renderer" in called_args and "pure-docx" in called_args
    assert diagnostics_root != tmp_path  # diagnostics never land inside the repo


def test_run_harness_never_calls_the_cli_when_preflight_fails(tmp_path):
    package_dir = _write_package(tmp_path)
    git = FakeGit(branch="main")
    cli_calls = []

    def run_subprocess(args, **kwargs):
        if args and args[0] == "git":
            return git(args, **kwargs)
        cli_calls.append(args)
        return FakeCompletedProcess(returncode=0)

    with pytest.raises(HarnessPreflightError):
        run_harness(
            repo_root=tmp_path, python_executable=Path("python.exe"), package_dir=package_dir,
            diagnostics_root=tmp_path / "diagnostics", renderer="word", run_subprocess=run_subprocess,
        )
    assert cli_calls == []


def test_powershell_wrapper_runs_in_the_foreground_and_never_kills_word_broadly():
    script = Path(r"C:\WordReplica-Automation\repo\RUN_LEKTA_POC.ps1").read_text(encoding="utf-8")
    assert "Start-Process" not in script
    assert "taskkill" not in script.lower()
    assert "Stop-Process" not in script
    assert "-NoNewWindow" not in script
    assert "& $VenvPython" in script
