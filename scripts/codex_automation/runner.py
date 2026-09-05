from __future__ import annotations

from dataclasses import asdict
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Callable

from scripts.codex_automation.audit import DEFAULT_GATE_NAMES, GateResult, build_golden_report
from scripts.codex_automation.config import CodexAutomationConfig
from scripts.codex_automation.process import ChildResult, run_owned_child
from scripts.codex_automation.word_process import (
    list_owned_word_processes,
    terminate_owned_word_processes,
)
from scripts.codex_automation.retention import prune_diagnostics, prune_stale_work
from scripts.codex_automation.state import evaluate_run, load_state, save_state
from scripts.codex_automation.trace_profile import profile_event_trace
from scripts.codex_automation.workspace import GoldenWorkspace, sha256_file
from word_replica.qa.environment import capture_environment_fingerprint


MAX_INTERACTIVE_COM_RECOVERIES = 3


def default_git_info(repo_root: Path) -> tuple[str, str, bool]:
    repo_root = Path(repo_root)
    branch = subprocess.run(
        ["git", "branch", "--show-current"], cwd=repo_root, capture_output=True, text=True, check=True
    ).stdout.strip()
    sha = subprocess.run(
        ["git", "rev-parse", "HEAD"], cwd=repo_root, capture_output=True, text=True, check=True
    ).stdout.strip()
    status = subprocess.run(
        ["git", "status", "--porcelain", "--untracked-files=all"], cwd=repo_root, capture_output=True, text=True, check=True
    ).stdout.strip()
    return branch, sha, not bool(status)


def _read_json(path: Path) -> dict:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_json(path: Path, payload: dict) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    temp.replace(path)


def _trace_summary(path: Path) -> dict:
    path = Path(path)
    if not path.exists():
        return {"event_records": 0, "first": None, "last": None}
    first = None
    last = None
    count = 0
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            if not line.strip():
                continue
            try:
                row = json.loads(line)
            except Exception:
                continue
            count += 1
            if first is None:
                first = row
            last = row
    return {"event_records": count, "first": first, "last": last}


def _performance_profile(path: Path) -> dict:
    try:
        return profile_event_trace(path)
    except (OSError, ValueError) as exc:
        return {"available": False, "error": str(exc)}


def _is_resumable_interactive_failure(payload: dict) -> bool:
    if payload.get("status") == "COM_FAIL":
        return True
    if payload.get("status") != "ENGINE_FAIL":
        return False
    return any(
        "[Errno 28]" in str(reason) or "No space left on device" in str(reason)
        for reason in payload.get("reasons", [])
    )


def _failure_report(run_id: str, source_hash: str, commit_sha: str, reconstruction_status: str, reason: str) -> dict:
    gates = {
        f"G{i}": GateResult(
            name=f"G{i}",
            passed=False,
            summary=f"not evaluable: {reason}",
            first_divergence={"error": reason} if i == 0 else None,
        )
        for i in range(len(DEFAULT_GATE_NAMES))
    }
    return build_golden_report(
        run_id=run_id,
        source_sha256=source_hash,
        commit_sha=commit_sha,
        reconstruction_status=reconstruction_status,
        gates=gates,
    )


def _archive_run_dir(source: Path, final: Path, *, expected_run_id: str) -> Path:
    try:
        shutil.move(str(source), str(final))
    except PermissionError:
        # shutil.move across directories copies first and then removes the
        # source. A Word-locked redundant rotating checkpoint can make only
        # that final removal fail. Accept the archive solely when its report
        # is complete and bound to this exact run; otherwise fail closed.
        report_path = final / "golden_report.json"
        if not report_path.is_file():
            raise
        report = _read_json(report_path)
        if report.get("run_id") != expected_run_id:
            raise
        report["archive_cleanup_pending"] = True
        _write_json(report_path, report)
        return report_path
    return final / "golden_report.json"


class GoldenRunner:
    def __init__(
        self,
        *,
        repo_root: Path,
        config: CodexAutomationConfig,
        python_executable: str | None = None,
        child_executor: Callable[..., ChildResult] = run_owned_child,
        git_info: Callable[[Path], tuple[str, str] | tuple[str, str, bool]] = default_git_info,
        environment_capture: Callable[[], dict] = capture_environment_fingerprint,
        visible_word: bool = False,
        disk_usage: Callable[[Path], object] | None = None,
    ) -> None:
        self.repo_root = Path(repo_root).resolve()
        self.config = config
        self.python_executable = python_executable or sys.executable
        self.child_executor = child_executor
        self.git_info = git_info
        self.environment_capture = environment_capture
        self.visible_word = bool(visible_word)
        self.disk_usage = disk_usage or shutil.disk_usage
        self.workspace = GoldenWorkspace(config)

    def _exec(self, command: list[str], *, timeout: int, run_dir: Path, stem: str, ownership_file: Path) -> ChildResult:
        return self.child_executor(
            command,
            timeout_seconds=timeout,
            stdout_path=run_dir / f"{stem}_stdout.log",
            stderr_path=run_dir / f"{stem}_stderr.log",
            ownership_file=ownership_file,
            cwd=self.repo_root,
            env={"PYTHONPATH": os.pathsep.join([str(self.repo_root / "src"), str(self.repo_root)])},
        )

    def _ensure_minimum_free_space(self) -> None:
        minimum_mb = int(self.config.minimum_free_space_mb)
        if minimum_mb <= 0:
            return
        free_mb = int(self.disk_usage(self.config.local_root).free // (1024 * 1024))
        if free_mb < minimum_mb:
            raise RuntimeError(
                f"Insufficient free disk space: {free_mb} MB available; {minimum_mb} MB required"
            )

    def run(self, golden_id: str = "golden_1") -> Path:
        git_info = self.git_info(self.repo_root)
        branch, commit_sha = git_info[:2]
        # Two-item injected test adapters predate the clean-worktree contract.
        worktree_clean = bool(git_info[2]) if len(git_info) >= 3 else True
        if branch != "automation-dev":
            raise RuntimeError(f"Autonomous Golden runs require branch automation-dev; current branch is {branch or '<detached>'}")

        self.workspace.acquire_lock()
        run = None
        try:
            prune_stale_work(self.config.work_dir)
            self._ensure_minimum_free_space()
            run = self.workspace.create_run(commit_sha, golden_id)
            static_dir = run.run_dir / "static"
            interactive_dir = run.run_dir / "interactive"
            child_runner = self.repo_root / "scripts" / "remote_harness" / "child_runner.py"
            audit_cli = self.repo_root / "scripts" / "codex_automation" / "audit_cli.py"

            static_result = self._exec(
                [
                    self.python_executable, str(child_runner), "--stage", "static",
                    "--source", str(run.source_copy), "--run-dir", str(static_dir),
                ],
                timeout=min(600, self.config.reconstruction_timeout_seconds),
                run_dir=run.run_dir,
                stem="static",
                ownership_file=run.ownership_file,
            )

            interactive_command = [
                self.python_executable, str(child_runner), "--stage", "interactive_maximum",
                "--source", str(run.source_copy), "--run-dir", str(interactive_dir),
                "--defer-l4-qa",
            ]
            if self.visible_word:
                interactive_command.append("--visible-word")
            interactive_result = self._exec(
                interactive_command,
                timeout=self.config.reconstruction_timeout_seconds,
                run_dir=run.run_dir,
                stem="interactive",
                ownership_file=run.ownership_file,
            )

            interactive_payload = _read_json(interactive_dir / "result.json")
            interactive_processes = [interactive_result]
            resume_project_id = interactive_payload.get("project_id")
            for resume_attempt in range(1, MAX_INTERACTIVE_COM_RECOVERIES + 1):
                if not (
                    not interactive_result.timed_out
                    and _is_resumable_interactive_failure(interactive_payload)
                    and resume_project_id
                ):
                    break
                interactive_result = self._exec(
                    [*interactive_command, "--resume-project-id", str(resume_project_id)],
                    timeout=self.config.reconstruction_timeout_seconds,
                    run_dir=run.run_dir,
                    stem=f"interactive_resume_{resume_attempt}",
                    ownership_file=run.ownership_file,
                )
                interactive_processes.append(interactive_result)
                interactive_payload = _read_json(interactive_dir / "result.json")
                resume_project_id = interactive_payload.get("project_id") or resume_project_id
            reconstruction_status = (
                "TIMEOUT" if interactive_result.timed_out
                else str(interactive_payload.get("run_status") or interactive_payload.get("status") or f"EXIT_{interactive_result.exit_code}")
            )
            output_docx = interactive_dir / "output.docx"
            report_path = run.run_dir / "golden_report.json"

            if output_docx.exists():
                audit_result = self._exec(
                    [
                        self.python_executable, str(audit_cli),
                        "--source", str(run.source_copy),
                        "--output", str(output_docx),
                        "--qa-dir", str(run.run_dir / "qa"),
                        "--report", str(report_path),
                        "--run-id", run.run_id,
                        "--source-sha256", run.source_sha256_before,
                        "--commit-sha", commit_sha,
                        "--reconstruction-status", reconstruction_status,
                        "--visual-dpi", str(self.config.visual_dpi),
                        "--changed-pixel-tolerance", str(self.config.changed_pixel_tolerance),
                        "--mae-tolerance", str(self.config.mae_tolerance),
                    ],
                    timeout=self.config.audit_timeout_seconds,
                    run_dir=run.run_dir,
                    stem="audit",
                    ownership_file=run.ownership_file,
                )
                if audit_result.timed_out or not report_path.exists():
                    report = _failure_report(
                        run.run_id, run.source_sha256_before, commit_sha, reconstruction_status,
                        "Golden audit timed out" if audit_result.timed_out else "Golden audit did not produce golden_report.json",
                    )
                else:
                    report = _read_json(report_path)
            else:
                audit_result = None
                reason = "; ".join(str(x) for x in interactive_payload.get("reasons", [])[:5]) or "interactive reconstruction produced no output.docx"
                report = _failure_report(run.run_id, run.source_sha256_before, commit_sha, reconstruction_status, reason)

            source_unchanged = self.workspace.verify_golden_unchanged(run)
            try:
                environment = self.environment_capture()
            except Exception as exc:
                environment = {"error": str(exc), "exception_type": type(exc).__name__}
            completed_processes = [static_result, *interactive_processes]
            if audit_result is not None:
                completed_processes.append(audit_result)
            verified_owner_pids = sorted({
                int(owner_pid)
                for process_result in completed_processes
                for owner_pid in process_result.owner_process_pids
            })
            final_terminated_word_pids = (
                terminate_owned_word_processes(
                    list_owned_word_processes(run.ownership_file),
                    expected_owner_pids=verified_owner_pids,
                )
                if verified_owner_pids else []
            )
            report.update({
                "run_id": run.run_id,
                "golden_id": golden_id,
                "golden_filename": self.config.golden_filename_for(golden_id),
                "gates_main_promotion": golden_id == "golden_1",
                "commit_sha": commit_sha,
                "worktree_clean": worktree_clean,
                "branch": branch,
                "source_sha256": run.source_sha256_before,
                "source_unchanged": source_unchanged,
                "source_sha256_after": sha256_file(self.config.golden_path) if self.config.golden_path.exists() else None,
                "static_process": asdict(static_result),
                "interactive_process": asdict(interactive_result),
                "interactive_processes": [asdict(item) for item in interactive_processes],
                "audit_process": asdict(audit_result) if audit_result is not None else None,
                "interactive_result": interactive_payload,
                "trace": _trace_summary(interactive_dir / "event_trace.jsonl"),
                "performance_profile": _performance_profile(interactive_dir / "event_trace.jsonl"),
                "environment": environment,
                "final_terminated_word_pids": final_terminated_word_pids,
            })
            if not source_unchanged:
                report["full_pass"] = False
                report["source_integrity_error"] = "Golden source SHA-256 changed during run"

            # golden_1 keeps the legacy filename so existing accumulated
            # production state (best_score, consecutive-pass streak) is not
            # silently abandoned by introducing per-document state files.
            state_filename = "automation_state.json" if golden_id == "golden_1" else f"automation_state_{golden_id}.json"
            state_path = self.config.state_dir / state_filename
            state = load_state(state_path)
            decision = evaluate_run(state, report)
            if not source_unchanged:
                decision.stop_required = True
                decision.reason = "Golden source integrity violation"
                decision.promotion_ready = False
            report["automation_decision"] = asdict(decision)
            save_state(state_path, state)
            _write_json(report_path, report)

            # golden_1 keeps the flat diagnostics_dir layout: AGENTS.md tells
            # operators to read the latest run straight from diagnostics_dir,
            # and it already holds real accumulated run history there.
            golden_diagnostics_dir = self.config.diagnostics_dir if golden_id == "golden_1" else self.config.diagnostics_dir / golden_id
            final_dir = golden_diagnostics_dir / run.run_id
            final_dir.parent.mkdir(parents=True, exist_ok=True)
            final_report = _archive_run_dir(
                run.run_dir, final_dir, expected_run_id=run.run_id
            )
            prune_diagnostics(
                golden_diagnostics_dir,
                keep_success=self.config.keep_success_runs,
                keep_failures=self.config.keep_failure_runs,
            )
            return final_report
        finally:
            self.workspace.release_lock()
