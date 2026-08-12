from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
from pathlib import Path
import subprocess
import sys
from typing import Callable

from scripts.codex_automation.config import load_config
from scripts.codex_automation.runner import GoldenRunner
from scripts.codex_automation.state import load_state, save_state


@dataclass(frozen=True, slots=True)
class LoopDecision:
    action: str
    reason: str
    exit_code: int


def decide_next_step(report: dict) -> LoopDecision:
    decision = report.get("automation_decision") or {}
    if bool(decision.get("promotion_ready")):
        return LoopDecision("COMPLETE", "G0-G9 passed twice on the same commit", 0)
    if bool(decision.get("stop_required")):
        return LoopDecision("STOP_FAIL_SAFE", str(decision.get("reason") or "Golden fail-safe activated"), 3)
    if bool(report.get("full_pass")):
        return LoopDecision("REPEAT_FULL_PASS", "First FULL PASS; one consecutive same-commit confirmation remains", 0)
    return LoopDecision(
        "PRODUCTION_FIX_REQUIRED",
        str(report.get("first_divergence") or report.get("first_divergent_gate") or "Golden did not pass"),
        2,
    )


class AutonomousLoopLock:
    def __init__(self, path: Path, *, pid: int | None = None) -> None:
        self.path = Path(path)
        self.pid = int(pid or 0)
        self._handle = None

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        try:
            self._handle = self.path.open("x", encoding="utf-8")
        except FileExistsError as exc:
            raise RuntimeError(f"autonomous Golden loop is already active: {self.path}") from exc
        self._handle.write(json.dumps({"pid": self.pid or None}, indent=2))
        self._handle.flush()
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._handle is not None:
            self._handle.close()
        self.path.unlink(missing_ok=True)


def _run_full_pytest(repo_root: Path, python_executable: str) -> int:
    completed = subprocess.run(
        [python_executable, "-m", "pytest", "-q"],
        cwd=repo_root,
        check=False,
    )
    return int(completed.returncode)


def run_loop(
    *,
    repo_root: Path,
    config_path: Path,
    local_root: Path,
    python_executable: str | None = None,
    pytest_runner: Callable[[Path, str], int] = _run_full_pytest,
    golden_runner_factory: Callable[..., GoldenRunner] = GoldenRunner,
    resume_after_review: bool = False,
) -> int:
    repo_root = Path(repo_root).resolve()
    local_root = Path(local_root).resolve()
    python_executable = python_executable or sys.executable
    config = load_config(Path(config_path).resolve(), local_root_override=local_root)
    lock_path = local_root / "state" / "golden_autonomous_loop.lock"
    state = load_state(config.state_dir / "automation_state.json")
    if state.no_improvement_count >= 3:
        if not resume_after_review:
            print("AUTOMATION STOP: existing Golden fail-safe requires a new reviewed production fix")
            return 3
        state.no_improvement_count = 0
        save_state(config.state_dir / "automation_state.json", state)
        print("AUTOMATION RESUME: reviewed production fix acknowledged")

    with AutonomousLoopLock(lock_path, pid=__import__("os").getpid()):
        for iteration in (1, 2):
            print(f"AUTONOMOUS GOLDEN ITERATION {iteration}/2: full pytest")
            pytest_exit = pytest_runner(repo_root, python_executable)
            if pytest_exit != 0:
                print(f"AUTOMATION STOP: full pytest failed with exit code {pytest_exit}")
                return 2

            print(f"AUTONOMOUS GOLDEN ITERATION {iteration}/2: real Microsoft Word Golden")
            report_path = golden_runner_factory(
                repo_root=repo_root,
                config=config,
                python_executable=python_executable,
            ).run()
            report = json.loads(Path(report_path).read_text(encoding="utf-8"))
            step = decide_next_step(report)
            print(f"AUTOMATION DECISION: {step.action} - {step.reason}")
            if step.action == "REPEAT_FULL_PASS":
                continue
            return step.exit_code

        return 3


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run the safe WordReplica autonomous Golden supervisor")
    parser.add_argument("--config", default="codex_automation.json")
    parser.add_argument("--local-root", default="C:\\WordReplica-Automation")
    parser.add_argument("--resume-after-review", action="store_true")
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    repo_root = Path(__file__).resolve().parents[2]
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = repo_root / config_path
    return run_loop(
        repo_root=repo_root,
        config_path=config_path,
        local_root=Path(args.local_root),
        resume_after_review=args.resume_after_review,
    )


if __name__ == "__main__":
    raise SystemExit(main())
