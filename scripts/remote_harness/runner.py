from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import re
import sys

from scripts.remote_harness.config import HarnessConfig
from scripts.remote_harness.contracts import DocumentSummary, HarnessStatus, StageResult, normalize_failure
from scripts.remote_harness.process_runner import run_child


def sha256_file(path: Path) -> str:
    h = sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def safe_name(name: str) -> str:
    stem = Path(name).stem
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._") or "document"
    return cleaned[:80]


class RemoteHarnessRunner:
    def __init__(self, *, source_root: Path, input_dir: Path, result_root: Path, config: HarnessConfig,
                 python_executable: str | None = None, child_runner_path: Path | None = None) -> None:
        self.source_root = Path(source_root).resolve()
        self.input_dir = Path(input_dir).resolve()
        self.result_root = Path(result_root).resolve()
        self.config = config
        self.python_executable = python_executable or sys.executable
        self.child_runner_path = Path(child_runner_path or (self.source_root / "scripts/remote_harness/child_runner.py")).resolve()

    def _run_stage(self, source: Path, stage: str, stage_dir: Path) -> StageResult:
        stage_dir.mkdir(parents=True, exist_ok=True)
        command = [
            self.python_executable,
            str(self.child_runner_path),
            "--stage", stage,
            "--source", str(source),
            "--run-dir", str(stage_dir),
        ]
        if self.config.logs_only:
            command.append("--logs-only")
        proc = run_child(
            command,
            self.config.timeout_for_stage(stage),
            stage_dir / "stdout.log",
            stage_dir / "stderr.log",
        )
        if proc.timed_out:
            payload = {
                "document": source.name, "stage": stage, "status": HarnessStatus.TIMEOUT.value,
                "reason": f"stage exceeded {self.config.timeout_for_stage(stage)} seconds",
                "elapsed_seconds": proc.elapsed_seconds, "exit_code": proc.exit_code,
            }
            (stage_dir / "result.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
            return StageResult(source.name, stage, HarnessStatus.TIMEOUT, payload["reason"], proc.elapsed_seconds, proc.exit_code)

        result_path = stage_dir / "result.json"
        if not result_path.exists():
            reason = f"child exited with code {proc.exit_code} without result.json"
            return StageResult(source.name, stage, HarnessStatus.PROCESS_CRASH, reason, proc.elapsed_seconds, proc.exit_code)
        try:
            payload = json.loads(result_path.read_text(encoding="utf-8"))
        except Exception as exc:
            return StageResult(source.name, stage, HarnessStatus.HARNESS_FAIL, f"result.json unreadable: {exc}", proc.elapsed_seconds, proc.exit_code)
        try:
            status = HarnessStatus(payload.get("status", HarnessStatus.PROCESS_CRASH.value))
        except ValueError:
            status = HarnessStatus.HARNESS_FAIL
        reasons = payload.get("reasons") or []
        reason = "; ".join(str(item) for item in reasons[:5])
        artifacts = payload.get("artifacts") or {}
        details = dict(payload)
        if reason:
            details["failure_fingerprint"] = normalize_failure(reason)
        return StageResult(source.name, stage, status, reason, float(payload.get("elapsed_seconds", proc.elapsed_seconds)), proc.exit_code, artifacts, details)

    def run(self) -> Path:
        documents_root = self.result_root / "documents"
        documents_root.mkdir(parents=True, exist_ok=True)
        sources = sorted((p for p in self.input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".docx"), key=lambda p: p.name.lower())
        for index, source in enumerate(sources, start=1):
            doc_dir = documents_root / f"{index:03d}_{safe_name(source.name)}"
            doc_dir.mkdir(parents=True, exist_ok=True)
            before = sha256_file(source)
            summary = DocumentSummary(source.name, before)
            (doc_dir / "source_manifest.json").write_text(json.dumps({
                "filename": source.name,
                "sha256_before": before,
                "size": source.stat().st_size,
            }, indent=2, ensure_ascii=False), encoding="utf-8")
            for stage in ("static", "instant", "interactive_maximum"):
                summary.stages[stage] = self._run_stage(source, stage, doc_dir / stage)
            if summary.stages["interactive_maximum"].status is HarnessStatus.EXPECTED_BLOCK:
                summary.stages["interactive_standard"] = self._run_stage(source, "interactive_standard", doc_dir / "interactive_standard")
            after = sha256_file(source)
            summary.source_sha256_after = after
            if before != after:
                summary.overall_status = HarnessStatus.SOURCE_MUTATED
                summary.reason = "source SHA-256 changed during harness run"
            else:
                statuses = {item.status for item in summary.stages.values()}
                if HarnessStatus.HARNESS_FAIL in statuses or HarnessStatus.PROCESS_CRASH in statuses:
                    summary.overall_status = HarnessStatus.HARNESS_FAIL
                elif HarnessStatus.TIMEOUT in statuses:
                    summary.overall_status = HarnessStatus.TIMEOUT
                elif HarnessStatus.COM_FAIL in statuses:
                    summary.overall_status = HarnessStatus.COM_FAIL
                elif HarnessStatus.ENGINE_FAIL in statuses:
                    summary.overall_status = HarnessStatus.ENGINE_FAIL
                elif HarnessStatus.WARN in statuses or HarnessStatus.EXPECTED_BLOCK in statuses:
                    summary.overall_status = HarnessStatus.WARN
                else:
                    summary.overall_status = HarnessStatus.PASS
            (doc_dir / "document_summary.json").write_text(
                json.dumps(summary.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8"
            )
        return self.result_root
