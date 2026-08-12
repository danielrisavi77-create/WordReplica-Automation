from __future__ import annotations

import argparse
from dataclasses import asdict
import json
from pathlib import Path
import shutil
import sys
import time
import traceback

# When invoked as a script, ensure project root and src are importable before
# importing the production package.
ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
for item in (str(ROOT), str(SRC)):
    if item not in sys.path:
        sys.path.insert(0, item)

from scripts.remote_harness.contracts import HarnessStatus, normalize_failure
from scripts.remote_harness.event_trace import JsonlEventTraceObserver
from scripts.remote_harness.static_analysis import analyze_document
from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import (
    InteractiveFidelity, InteractiveSpeedMode, ReconstructionMode, RendererChoice, RunStatus, VisibilityMode,
)
from word_replica.domain.results import RunResult
from word_replica.services.rebuild import RebuildService


def serialize_run_result(document: str, stage: str, result: RunResult, elapsed_seconds: float) -> dict:
    return {
        "document": str(document),
        "stage": stage,
        "run_status": result.status.value,
        "project_id": result.project_id,
        "save_count": result.save_count,
        "warnings": [asdict(item) for item in result.warnings],
        "reasons": list(result.reasons),
        "output_path": str(result.output_path) if result.output_path else None,
        "qa_report_path": str(result.qa_report_path) if result.qa_report_path else None,
        "elapsed_seconds": float(elapsed_seconds),
    }


def classify_run_result(payload: dict) -> HarnessStatus:
    stage = payload.get("stage", "")
    if stage == "interactive_maximum" and payload.get("preflight_can_proceed") is False:
        return HarnessStatus.EXPECTED_BLOCK
    run_status = payload.get("run_status")
    warnings = payload.get("warnings") or []
    reasons = payload.get("reasons") or []
    text = "\n".join([str(x) for x in reasons] + [str(x) for x in warnings])
    if run_status == RunStatus.PASS.value:
        return HarnessStatus.PASS
    if run_status == RunStatus.WARN.value:
        return HarnessStatus.WARN
    fingerprint = normalize_failure(text)
    if fingerprint.startswith("RPC_") or fingerprint == "INVALID_ACTIVE_RANGE_TYPE":
        return HarnessStatus.COM_FAIL
    return HarnessStatus.ENGINE_FAIL


def _copy_artifact(source: str | None, destination: Path) -> str | None:
    if not source:
        return None
    src = Path(source)
    if not src.exists():
        return None
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, destination)
    return str(destination)


def _options_for(stage: str) -> RebuildOptions:
    if stage == "instant":
        return RebuildOptions(
            renderer=RendererChoice.WORD,
            visibility=VisibilityMode.BACKGROUND,
            reconstruction_mode=ReconstructionMode.INSTANT,
        )
    fidelity = InteractiveFidelity.MAXIMUM if stage == "interactive_maximum" else InteractiveFidelity.STANDARD
    return RebuildOptions(
        reconstruction_mode=ReconstructionMode.INTERACTIVE,
        renderer=RendererChoice.WORD,
        visibility=VisibilityMode.BACKGROUND,
        interactive=InteractiveOptions(
            speed_mode=InteractiveSpeedMode.MAXIMUM,
            object_step_delay_ms=0,
            fidelity=fidelity,
            checkpoint_event_interval=100000,
            verify_during_run=True,
            block_on_unsupported=True,
        ),
    )


def run_stage(stage: str, source: Path, run_dir: Path, logs_only: bool = False) -> dict:
    run_dir.mkdir(parents=True, exist_ok=True)
    if stage == "static":
        started = time.perf_counter()
        analysis = analyze_document(source)
        (run_dir / "analysis.json").write_text(json.dumps(analysis, indent=2, ensure_ascii=False), encoding="utf-8")
        ok = analysis.get("parser", {}).get("ok", False) and analysis.get("package", {}).get("zip_integrity", False)
        return {
            "document": source.name,
            "stage": stage,
            "run_status": "PASS" if ok else "FAIL",
            "status": HarnessStatus.PASS.value if ok else HarnessStatus.ENGINE_FAIL.value,
            "elapsed_seconds": time.perf_counter() - started,
            "analysis_path": str(run_dir / "analysis.json"),
            "preflight_can_proceed": analysis.get("preflight", {}).get("can_proceed"),
            "reasons": [] if ok else [analysis.get("parser", {}).get("error", "static analysis failed")],
        }

    static = analyze_document(source)
    if stage == "interactive_maximum" and static.get("preflight", {}).get("can_proceed") is False:
        return {
            "document": source.name,
            "stage": stage,
            "run_status": "FAIL",
            "status": HarnessStatus.EXPECTED_BLOCK.value,
            "elapsed_seconds": 0.0,
            "preflight_can_proceed": False,
            "reasons": list(static.get("preflight", {}).get("blocking_reasons", [])),
        }

    trace = JsonlEventTraceObserver(run_dir / "event_trace.jsonl") if stage.startswith("interactive_") else None
    service = RebuildService(app_root=run_dir / "projects")
    started = time.perf_counter()
    result = service.rebuild(source, _options_for(stage), interactive_observer=trace)
    payload = serialize_run_result(source.name, stage, result, time.perf_counter() - started)
    payload["preflight_can_proceed"] = static.get("preflight", {}).get("can_proceed")
    payload["status"] = classify_run_result(payload).value
    if not logs_only:
        copied_output = _copy_artifact(payload.get("output_path"), run_dir / "output.docx")
        copied_qa = _copy_artifact(payload.get("qa_report_path"), run_dir / "qa_report.html")
        payload["artifacts"] = {"output": copied_output, "qa_report": copied_qa}
    else:
        payload["artifacts"] = {}
    return payload


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--stage", required=True, choices=["static", "instant", "interactive_maximum", "interactive_standard"])
    parser.add_argument("--source", required=True)
    parser.add_argument("--run-dir", required=True)
    parser.add_argument("--logs-only", action="store_true")
    args = parser.parse_args(argv)
    run_dir = Path(args.run_dir).resolve()
    run_dir.mkdir(parents=True, exist_ok=True)
    result_path = run_dir / "result.json"
    try:
        payload = run_stage(args.stage, Path(args.source).resolve(), run_dir, args.logs_only)
        result_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
        return 0
    except Exception as exc:
        payload = {
            "document": Path(args.source).name,
            "stage": args.stage,
            "run_status": "FAIL",
            "status": HarnessStatus.COM_FAIL.value if normalize_failure(str(exc)).startswith("RPC_") else HarnessStatus.HARNESS_FAIL.value,
            "reasons": [str(exc)],
            "traceback": traceback.format_exc(),
        }
        result_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
