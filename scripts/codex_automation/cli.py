from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
for item in (str(ROOT), str(SRC)):
    if item not in sys.path:
        sys.path.insert(0, item)

from scripts.codex_automation.config import load_config
from scripts.codex_automation.runner import GoldenRunner


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run a WordReplica Golden document locally for Codex automation")
    parser.add_argument("--config", default="codex_automation.json")
    parser.add_argument("--local-root", default=None)
    parser.add_argument("--visible-word", action="store_true")
    parser.add_argument(
        "--golden", default="golden_1",
        help="Golden document id to run (default: golden_1, the sole promotion gate). "
             "See config.golden_documents for the full list.",
    )
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = ROOT / config_path
    config = load_config(
        config_path,
        local_root_override=Path(args.local_root) if args.local_root else None,
    )
    runner = GoldenRunner(repo_root=ROOT, config=config, visible_word=args.visible_word)
    report_path = runner.run(golden_id=args.golden)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    label = report.get("golden_id", args.golden)
    gates_promotion = bool(report.get("gates_main_promotion"))
    print(f"GOLDEN REPORT ({label}): {report_path}")
    print("GATES: " + " ".join(f"{name}={'PASS' if passed else 'FAIL'}" for name, passed in report.get("gates", {}).items()))
    decision = report.get("automation_decision") or {}
    if decision.get("promotion_ready"):
        if gates_promotion:
            print(f"{label.upper()}: FULL PASS x2 - PROMOTION READY")
        else:
            print(f"{label.upper()}: FULL PASS x2 - regression-stable (does not gate main promotion; only golden_1 does)")
        return 0
    if report.get("full_pass"):
        print(f"{label.upper()}: FULL PASS - repeat same commit once more")
        return 0
    if decision.get("stop_required"):
        print(f"AUTOMATION STOP REQUIRED ({label}): {decision.get('reason')}")
        return 3
    print(f"FIRST DIVERGENCE ({label}): {report.get('first_divergent_gate')} {report.get('first_divergence')}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
