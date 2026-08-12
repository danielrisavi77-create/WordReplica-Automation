from __future__ import annotations

import argparse
from dataclasses import replace
from datetime import datetime, timezone
import json
from pathlib import Path
import sys
import traceback

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
for item in (str(ROOT), str(SRC)):
    if item not in sys.path:
        sys.path.insert(0, item)

from scripts.remote_harness import HARNESS_VERSION
from scripts.remote_harness.aggregate import aggregate_results
from scripts.remote_harness.config import HarnessConfig
from scripts.remote_harness.environment import capture_environment
from scripts.remote_harness.package_results import package_results
from scripts.remote_harness.prerequisites import check_prerequisites
from scripts.remote_harness.runner import RemoteHarnessRunner


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Word Replica Remote Windows/Word Test Harness")
    parser.add_argument("--input-dir", default=str(ROOT / "realworld_input"))
    parser.add_argument("--config", default=str(ROOT / "harness_config.json"))
    parser.add_argument("--logs-only", action="store_true")
    parser.add_argument("--result-parent", default=str(ROOT / "remote_results"))
    parser.add_argument("--harness-version", default=None)
    return parser



def _safe_version(value: str) -> str:
    safe = "".join(ch if ch.isalnum() or ch in ".-_" else "_" for ch in str(value or "unknown"))
    return safe.strip("._-") or "unknown"


def build_run_name(version: str, timestamp: str) -> str:
    date, clock = timestamp.split("-", 1)
    return f"{date[:4]}-{date[4:6]}-{date[6:8]}_{clock}_v{_safe_version(version)}"


def build_result_zip_name(version: str, timestamp: str) -> str:
    return f"WordReplica-Remote-Results-{timestamp}-v{_safe_version(version)}.zip"

def _print_summary(summary: dict) -> None:
    print("\n=== Remote Word Harness Summary ===")
    print(f"Documents: {summary.get('documents', 0)}")
    print(f"Runs:      {summary.get('runs', 0)}")
    counts = summary.get("status_counts", {})
    if counts:
        print("Statuses:  " + ", ".join(f"{key}={value}" for key, value in sorted(counts.items())))
    groups = summary.get("failure_groups", {})
    if groups:
        print("Root-cause groups:")
        for key, value in list(groups.items())[:12]:
            print(f"  {key}: {value['runs']} runs / {value['documents']} documents")


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    input_dir = Path(args.input_dir).resolve()
    config = HarnessConfig.load(Path(args.config).resolve())
    if args.logs_only:
        config = replace(config, logs_only=True)

    prereq = check_prerequisites(
        input_dir,
        ROOT,
        minimum_free_space_mb=config.minimum_free_space_mb,
    )
    if not prereq.ok:
        print("REMOTE WORD HARNESS: PREREQUISITE FAIL")
        for error in prereq.errors:
            print(f"ERROR: {error}")
        return 2

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    harness_version = args.harness_version or HARNESS_VERSION
    result_parent = Path(args.result_parent).resolve()
    result_root = result_parent / build_run_name(harness_version, timestamp)
    result_root.mkdir(parents=True, exist_ok=True)
    destination_zip = result_root / build_result_zip_name(harness_version, timestamp)
    harness_log = result_root / "harness.log"
    harness_log.write_text(
        "Word Replica Remote Harness\n"
        f"Input: {input_dir}\n"
        f"Logs only: {config.logs_only}\n"
        f"Harness version: {harness_version}\n",
        encoding="utf-8",
    )
    environment = capture_environment(ROOT)
    environment["harness_version"] = harness_version
    # Reuse the prerequisite Word probe evidence so consumers see the exact
    # successful setup check even if later Word sessions fail.
    environment["prerequisite_word"] = prereq.word
    environment["disk"]["prerequisite_free_mb"] = prereq.free_space_mb
    (result_root / "environment.json").write_text(
        json.dumps(environment, indent=2, ensure_ascii=False, default=str), encoding="utf-8"
    )

    harness_error = None
    summary = None
    try:
        runner = RemoteHarnessRunner(
            source_root=ROOT,
            input_dir=input_dir,
            result_root=result_root,
            config=config,
        )
        runner.run()
    except Exception as exc:
        harness_error = exc
        (result_root / "harness_exception.json").write_text(
            json.dumps({
                "exception_type": type(exc).__name__,
                "message": str(exc),
                "traceback": traceback.format_exc(),
            }, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    finally:
        try:
            summary = aggregate_results(result_root)
            summary["harness_version"] = harness_version
            (result_root / "summary.json").write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as exc:
            (result_root / "aggregate_error.txt").write_text(traceback.format_exc(), encoding="utf-8")
            if harness_error is None:
                harness_error = exc
        try:
            package_results(result_root, destination_zip, config.logs_only)
        except Exception:
            print("REMOTE WORD HARNESS: ZIP CREATION FAIL")
            traceback.print_exc()
            return 3

    if summary is not None:
        _print_summary(summary)
    print(f"\nRESULT ZIP: {destination_zip}")
    if harness_error is not None:
        print(f"REMOTE WORD HARNESS: COMPLETE WITH HARNESS ERROR: {harness_error}")
        return 1
    print("REMOTE WORD HARNESS: COMPLETE")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
