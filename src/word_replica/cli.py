from __future__ import annotations

import argparse
import json
from pathlib import Path

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
# Must stay within the signed contract's engineMinVersion..engineMaxVersion
# range (currently "1.0.0" in Lekta's published Repair Contract v1 fixture).
REPAIR_ENGINE_VERSION = "1.0.0"


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


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "rebuild":
        return run_rebuild(args)
    if args.command == "repair-poc":
        return run_repair_poc(args)
    if args.command == "repair-poc-resume":
        return run_repair_poc_resume(args)
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
