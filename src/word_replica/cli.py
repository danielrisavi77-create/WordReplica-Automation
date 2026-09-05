from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
from pathlib import Path

from word_replica import __version__
from word_replica.config import RebuildOptions
from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, RunStatus, VisibilityMode
from word_replica.domain.errors import RepairPackageError
from word_replica.parser.parser import DocxParser
from word_replica.qa.policy import classify_run, run_l0_l3
from word_replica.repair_contract.contract import RepairContractSchemaError
from word_replica.repair_contract.package import RepairPackageRequest
from word_replica.repair_contract.signature import RepairContractSignatureError
from word_replica.services.output_handoff import open_output_for_user
from word_replica.services.project_store import ProjectStore
from word_replica.services.rebuild import RebuildService

_EXIT = {RunStatus.PASS: 0, RunStatus.WARN: 1, RunStatus.FAIL: 2}
_REPAIR_POC_EXIT = {"FULL_PASS": 0, "RETRYABLE": 1, "FAILED": 2}
# The repair capability contract is versioned with the installed engine.
# A stale independent literal can make a valid signed package fail preflight.
REPAIR_ENGINE_VERSION = __version__
LEKTA_REPAIR_CLAIM_ENDPOINT = (
    "https://zrrjttizjyfcxmcpgzml.supabase.co/functions/v1/repair-local-claim"
)
LEKTA_REPAIR_STATUS_ENDPOINT = (
    "https://zrrjttizjyfcxmcpgzml.supabase.co/functions/v1/repair-local-status"
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="word-replica")
    sub = parser.add_subparsers(dest="command", required=True)
    rebuild = sub.add_parser("rebuild")
    rebuild.add_argument("source", type=Path)
    rebuild.add_argument("--renderer", choices=[e.value for e in RendererChoice], default="auto")
    rebuild.add_argument("--visibility", choices=[e.value for e in VisibilityMode], default="background")
    rebuild.add_argument("--fidelity", choices=[e.value for e in FidelityMode], default="clean")
    rebuild.add_argument("--metadata", choices=[e.value for e in MetadataMode], default="fresh")
    rebuild.add_argument("--preserve-author-fields", action="store_true")
    rebuild.add_argument("--preserve-custom-property", action="append", default=[])
    rebuild.add_argument("--allow-source-overwrite", action="store_true")
    inspect = sub.add_parser("inspect"); inspect.add_argument("source", type=Path)
    qa = sub.add_parser("qa"); qa.add_argument("source", type=Path); qa.add_argument("rebuilt", type=Path)
    projects = sub.add_parser("projects"); projects.add_argument("action", choices=["list"])
    project = sub.add_parser("project"); project.add_argument("action", choices=["show"]); project.add_argument("id")

    repair_poc = sub.add_parser("repair-poc")
    _add_repair_poc_package_arguments(repair_poc)
    repair_poc.add_argument(
        "--stop-after-event", type=int, default=None,
        help="Development only: request a safe stop once this many interactive events have completed.",
    )

    repair_poc_resume = sub.add_parser("repair-poc-resume")
    _add_repair_poc_package_arguments(repair_poc_resume)
    repair_poc_resume.add_argument("--job-id", required=True)

    repair_runner = sub.add_parser("repair-runner")
    repair_runner.add_argument("--launch", type=Path, required=True)
    repair_runner.add_argument("--output-dir", type=Path, required=True)

    return parser


def _add_repair_poc_package_arguments(subparser: argparse.ArgumentParser) -> None:
    subparser.add_argument("--original", type=Path, required=True)
    subparser.add_argument("--target", type=Path, required=True)
    subparser.add_argument("--contract", type=Path, required=True)
    subparser.add_argument("--public-key", type=Path, required=True)
    subparser.add_argument("--output-dir", type=Path, required=True)
    subparser.add_argument(
        "--renderer", choices=["word", "pure-docx"], default="word",
        help="'word' needs a licensed, activated Microsoft Word. 'pure-docx' "
             "reconstructs without any Word dependency, showing the visible "
             "typing preview in this console instead of a real Word window.",
    )


def require_docx(path: Path) -> Path:
    path = Path(path)
    if path.suffix.lower() != ".docx":
        raise SystemExit(f"Only .docx input is supported: {path}")
    return path


def run_rebuild(args) -> int:
    options = RebuildOptions(
        renderer=RendererChoice(args.renderer),
        visibility=VisibilityMode(args.visibility),
        fidelity=FidelityMode(args.fidelity),
        metadata=MetadataMode(args.metadata),
        allow_source_overwrite=args.allow_source_overwrite,
        preserve_author_fields=args.preserve_author_fields,
        custom_metadata_allowlist=tuple(args.preserve_custom_property),
    )
    result = RebuildService.default().rebuild(require_docx(args.source), options)
    print(f"Status: {result.status.value}")
    print(f"Output: {result.output_path or '-'}")
    print(f"QA report: {result.qa_report_path or '-'}")
    print(f"Saves: {result.save_count}")
    print(f"Warnings: {len(result.warnings)}")
    return _EXIT[result.status]


def run_inspect(path: Path) -> int:
    model = DocxParser().parse(require_docx(path))
    print(json.dumps({
        "paragraphs": sum(1 for _ in model.iter_paragraphs()),
        "sections": len(model.sections),
        "assets": len(model.assets),
    }, indent=2))
    return 0


def run_projects(store: ProjectStore) -> int:
    print(json.dumps(store.list_projects(), indent=2))
    return 0


def run_project_show(store: ProjectStore, project_id: str) -> int:
    print(json.dumps(store.get_project(project_id), indent=2))
    return 0


def run_qa(source: Path, rebuilt: Path) -> int:
    bundle = run_l0_l3(DocxParser().parse(require_docx(source)), DocxParser().parse(require_docx(rebuilt)))
    status = classify_run(bundle, warnings=[])
    print(status.value)
    return _EXIT[status]


def _repair_package_request(args) -> RepairPackageRequest:
    return RepairPackageRequest(
        original_path=args.original,
        target_path=args.target,
        contract_path=args.contract,
        public_key_path=args.public_key,
        output_dir=args.output_dir,
    )


def _console_preview_sink(kind: str, payload: dict) -> None:
    """Prints the visible typing preview directly to this console, so
    --renderer pure-docx gives real-time feedback without any Word window."""
    if kind == "text":
        print(payload.get("text", ""), end="", flush=True)
    elif kind == "newline":
        print()
    elif kind == "placeholder":
        print(f" {payload.get('label', '')} ", end="", flush=True)
    elif kind in ("pagebreak", "sectionbreak"):
        print()


def _default_repair_package_service(renderer: str = "word"):
    from datetime import datetime, timezone

    from word_replica.qa.golden_audit import audit_docx_pair
    from word_replica.repair_contract.binding import RepairRunBindingStore
    from word_replica.services.repair_package import RepairPackageService

    kwargs: dict = {}
    if renderer == "pure-docx":
        kwargs["preview_sink"] = _console_preview_sink

    return RepairPackageService(
        rebuild_service=RebuildService.default(),
        gate_auditor=audit_docx_pair,
        binding_store=RepairRunBindingStore(),
        clock=lambda: datetime.now(timezone.utc),
        engine_version=REPAIR_ENGINE_VERSION,
        **kwargs,
    )


def _write_repair_report(report, output_dir: Path) -> Path:
    report_path = Path(output_dir) / "repair_poc_report.json"
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report.to_json(), indent=2, sort_keys=True), encoding="utf-8")
    return report_path


def _stop_after_event_control(stop_after_event: int | None):
    """Development-only safe stop: request InteractiveRunControl.stop() only
    after the given event index has completed, never mid-event."""
    if stop_after_event is None:
        return None, None
    from word_replica.interactive.control import InteractiveRunControl

    control = InteractiveRunControl()
    control.start()

    class _StopAfterEventObserver:
        def event_completed(self, index, event) -> None:
            if index >= stop_after_event:
                control.stop()

    return control, _StopAfterEventObserver()


def run_repair_poc(args, *, service=None, handoff=None) -> int:
    service = service or _default_repair_package_service(getattr(args, "renderer", "word"))
    handoff = handoff or open_output_for_user
    request = _repair_package_request(args)
    control, observer = _stop_after_event_control(getattr(args, "stop_after_event", None))
    try:
        report = service.run(request, interactive_control=control, interactive_observer=observer)
    except (RepairPackageError, RepairContractSignatureError, RepairContractSchemaError) as exc:
        print(f"repair-poc invalid: {exc}")
        return 2
    report_path = _write_repair_report(report, args.output_dir)
    print(f"Report: {report_path}")
    print(f"Status: {report.status}")
    if report.full_pass:
        handoff(Path(report.output_path))
    return _REPAIR_POC_EXIT.get(report.status, 2)


def run_repair_poc_resume(args, *, service=None, handoff=None) -> int:
    service = service or _default_repair_package_service(getattr(args, "renderer", "word"))
    handoff = handoff or open_output_for_user
    request = _repair_package_request(args)
    try:
        report = service.resume(request, args.job_id)
    except (RepairPackageError, RepairContractSignatureError, RepairContractSchemaError) as exc:
        print(f"repair-poc-resume invalid: {exc}")
        return 2
    report_path = _write_repair_report(report, args.output_dir)
    print(f"Report: {report_path}")
    print(f"Status: {report.status}")
    if report.full_pass:
        handoff(Path(report.output_path))
    return _REPAIR_POC_EXIT.get(report.status, 2)


def _default_one_shot_runner():
    from word_replica.runner.http_transport import HttpsTransport
    from word_replica.runner.one_shot import OneShotRunner
    from word_replica.runner.trust_store import load_trust_keys
    from word_replica.runner.secure_retry import SecureRetryStore
    from word_replica.runner.word_preflight import run_word_preflight

    trust_path = Path(__file__).resolve().parent / "runner" / "trusted_keys.json"
    return OneShotRunner(
        preflight=run_word_preflight,
        transport=HttpsTransport(),
        package_service=_default_repair_package_service("word"),
        trust_keys=load_trust_keys(trust_path),
        retry_store=SecureRetryStore(),
    )


_RUNNER_JOB_ID = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$",
    re.I,
)


def _resume_job_id(value) -> str | None:
    if not isinstance(value, dict) or set(value) != {"version", "jobId", "resume"}:
        return None
    job_id = value.get("jobId")
    if value.get("version") != 1 or value.get("resume") is not True:
        raise ValueError("invalid local repair resume pointer")
    if not isinstance(job_id, str) or not _RUNNER_JOB_ID.fullmatch(job_id):
        raise ValueError("invalid local repair resume pointer")
    return job_id


def _replace_launch_with_resume_pointer(path: Path, job_id: str) -> None:
    payload = json.dumps(
        {"version": 1, "jobId": job_id, "resume": True},
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    descriptor, temporary_name = tempfile.mkstemp(
        prefix="lekta-resume-", suffix=".tmp", dir=path.parent
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as target:
            target.write(payload)
            target.flush()
            os.fsync(target.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def execute_repair_runner_ticket(
    ticket,
    output_dir: Path,
    *,
    runner=None,
    handoff=None,
) -> int:
    """Execute one in-memory launch ticket without persisting its bearer token.

    The portable EXE reads the ticket from its Authenticode-preserving filename.
    If the same job already has DPAPI state, resume wins before any second claim.
    """
    from word_replica.runner.one_shot import OneShotRunnerConfig

    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    active_runner = runner or _default_one_shot_runner()
    handoff = handoff or open_output_for_user
    config = OneShotRunnerConfig(
        claim_endpoint=LEKTA_REPAIR_CLAIM_ENDPOINT,
        status_endpoint=LEKTA_REPAIR_STATUS_ENDPOINT,
        output_dir=output_dir,
    )
    has_interrupted_job = getattr(active_runner, "has_interrupted_job", None)
    if callable(has_interrupted_job) and has_interrupted_job(ticket.job_id):
        result = active_runner.resume_interrupted(ticket.job_id, config)
    else:
        result = active_runner.run(ticket, config, on_claimed=None)
    handoff(Path(result.output_path))
    return 0


def run_repair_runner(args, *, runner=None, handoff=None) -> int:
    from word_replica.runner.lekta_claim import LaunchTicket
    from word_replica.runner.one_shot import OneShotRunnerConfig

    launch_path = Path(args.launch).resolve()
    output_dir = Path(args.output_dir).resolve()
    handoff = handoff or open_output_for_user
    try:
        if launch_path.stat().st_size > 4096:
            raise ValueError("launch ticket is too large")
        raw_launch = json.loads(launch_path.read_text(encoding="utf-8"))
        resume_job_id = _resume_job_id(raw_launch)
        output_dir.mkdir(parents=True, exist_ok=True)
        active_runner = runner or _default_one_shot_runner()
        config = OneShotRunnerConfig(
            claim_endpoint=LEKTA_REPAIR_CLAIM_ENDPOINT,
            status_endpoint=LEKTA_REPAIR_STATUS_ENDPOINT,
            output_dir=output_dir,
        )
        if resume_job_id is not None:
            result = active_runner.resume_interrupted(resume_job_id, config)
        else:
            ticket = LaunchTicket.parse(raw_launch)
            has_interrupted_job = getattr(active_runner, "has_interrupted_job", None)
            if callable(has_interrupted_job) and has_interrupted_job(ticket.job_id):
                _replace_launch_with_resume_pointer(launch_path, ticket.job_id)
                result = active_runner.resume_interrupted(ticket.job_id, config)
            else:
                result = active_runner.run(
                    ticket,
                    config,
                    on_claimed=lambda job_id: _replace_launch_with_resume_pointer(
                        launch_path, job_id
                    ),
                )
        launch_path.unlink()
        handoff(Path(result.output_path))
        return 0
    except Exception:
        print("Lokalni Word popravak nije dovršen. Serverska verzija ostaje dostupna u Lekti.")
        return 2


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "rebuild":
        return run_rebuild(args)
    if args.command == "repair-poc":
        return run_repair_poc(args)
    if args.command == "repair-poc-resume":
        return run_repair_poc_resume(args)
    if args.command == "repair-runner":
        return run_repair_runner(args)
    if args.command == "inspect":
        return run_inspect(args.source)
    if args.command == "qa":
        return run_qa(args.source, args.rebuilt)
    store = ProjectStore()
    if args.command == "projects":
        return run_projects(store)
    if args.command == "project":
        return run_project_show(store, args.id)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
