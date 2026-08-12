from __future__ import annotations

import argparse
import json
from pathlib import Path

from word_replica.config import RebuildOptions
from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, RunStatus, VisibilityMode
from word_replica.parser.parser import DocxParser
from word_replica.qa.policy import classify_run, run_l0_l3
from word_replica.services.project_store import ProjectStore
from word_replica.services.rebuild import RebuildService

_EXIT = {RunStatus.PASS: 0, RunStatus.WARN: 1, RunStatus.FAIL: 2}


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
    return parser


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


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "rebuild":
        return run_rebuild(args)
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
