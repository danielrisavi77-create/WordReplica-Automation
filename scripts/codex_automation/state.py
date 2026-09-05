from __future__ import annotations

from dataclasses import asdict, dataclass, field
import json
from pathlib import Path


DEFAULT_REQUIRED_GATES = tuple(f"G{i}" for i in range(10))


@dataclass(slots=True)
class AutomationState:
    last_gate_status: dict[str, bool] = field(default_factory=dict)
    best_score: int = 0
    no_improvement_count: int = 0
    last_full_pass_commit: str | None = None
    last_full_pass_word_build: str | None = None
    consecutive_full_pass_same_commit: int = 0


@dataclass(slots=True)
class AutomationDecision:
    stop_required: bool
    reason: str | None
    score: int
    promotion_ready: bool = False


def evaluate_run(state: AutomationState, report: dict) -> AutomationDecision:
    current = {str(k): bool(v) for k, v in (report.get("gates") or {}).items()}
    regressed = sorted(name for name, was_pass in state.last_gate_status.items() if was_pass and not current.get(name, False))
    score = sum(1 for value in current.values() if value)
    commit_sha = str(report.get("commit_sha") or "")
    word_build = ((report.get("environment") or {}).get("word") or {}).get("build")
    # Which gates a FULL PASS requires is declared by the report itself, so a
    # pipeline that runs extra gates (the fidelity lab adds G10) cannot silently
    # change what gates promotion to main. Reports without the key are Golden's
    # original ten-gate shape and must keep behaving exactly as they always did.
    required = [str(name) for name in (report.get("required_gates") or DEFAULT_REQUIRED_GATES)]
    full_pass = (
        bool(report.get("full_pass"))
        and len(current) == len(required)
        and all(current.get(name, False) for name in required)
        and bool(report.get("worktree_clean", True))
    )

    if full_pass:
        # Two consecutive FULL PASS results must land on the same commit AND the
        # same Word build - a build change is an environment change, and a pass
        # verified on a different Word build is not evidence the golden commit
        # is stable on the build it was previously verified against.
        if state.last_full_pass_commit == commit_sha and state.last_full_pass_word_build == word_build:
            state.consecutive_full_pass_same_commit += 1
        else:
            state.last_full_pass_commit = commit_sha
            state.last_full_pass_word_build = word_build
            state.consecutive_full_pass_same_commit = 1
    else:
        state.last_full_pass_commit = None
        state.last_full_pass_word_build = None
        state.consecutive_full_pass_same_commit = 0

    if regressed:
        state.last_gate_status = current
        return AutomationDecision(
            True,
            f"previously passing gate regressed: {', '.join(regressed)}",
            score,
            promotion_ready=False,
        )

    if full_pass:
        state.no_improvement_count = 0
    elif score > state.best_score:
        state.best_score = score
        state.no_improvement_count = 0
    else:
        state.no_improvement_count += 1
    state.last_gate_status = current

    promotion_ready = full_pass and state.consecutive_full_pass_same_commit >= 2
    if state.no_improvement_count >= 3 and not promotion_ready:
        return AutomationDecision(True, "three consecutive iterations without gate improvement", score, False)
    return AutomationDecision(False, None, score, promotion_ready)


def load_state(path: Path) -> AutomationState:
    path = Path(path)
    if not path.exists():
        return AutomationState()
    payload = json.loads(path.read_text(encoding="utf-8"))
    allowed = set(AutomationState.__dataclass_fields__)
    return AutomationState(**{key: value for key, value in payload.items() if key in allowed})


def save_state(path: Path, state: AutomationState) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(asdict(state), indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(path)
