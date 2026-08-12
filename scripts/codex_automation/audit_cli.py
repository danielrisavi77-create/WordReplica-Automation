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

from scripts.codex_automation.audit import audit_docx_pair


def run_audit_to_file(
    *,
    source: Path,
    output: Path,
    qa_dir: Path,
    report_path: Path,
    run_id: str,
    source_sha256: str,
    commit_sha: str,
    reconstruction_status: str,
    visual_dpi: int = 144,
    changed_pixel_tolerance: float = 0.001,
    mae_tolerance: float = 0.25,
    audit_fn=audit_docx_pair,
) -> Path:
    report = audit_fn(
        source,
        output,
        qa_dir,
        run_id=run_id,
        source_sha256=source_sha256,
        commit_sha=commit_sha,
        reconstruction_status=reconstruction_status,
        visual_dpi=visual_dpi,
        changed_pixel_tolerance=changed_pixel_tolerance,
        mae_tolerance=mae_tolerance,
    )
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    temp = report_path.with_suffix(report_path.suffix + ".tmp")
    temp.write_text(json.dumps(report, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    temp.replace(report_path)
    return report_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Audit one WordReplica Golden source/output pair")
    parser.add_argument("--source", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--qa-dir", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--source-sha256", required=True)
    parser.add_argument("--commit-sha", required=True)
    parser.add_argument("--reconstruction-status", required=True)
    parser.add_argument("--visual-dpi", type=int, default=144)
    parser.add_argument("--changed-pixel-tolerance", type=float, default=0.001)
    parser.add_argument("--mae-tolerance", type=float, default=0.25)
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    run_audit_to_file(
        source=Path(args.source).resolve(),
        output=Path(args.output).resolve(),
        qa_dir=Path(args.qa_dir).resolve(),
        report_path=Path(args.report).resolve(),
        run_id=args.run_id,
        source_sha256=args.source_sha256,
        commit_sha=args.commit_sha,
        reconstruction_status=args.reconstruction_status,
        visual_dpi=args.visual_dpi,
        changed_pixel_tolerance=args.changed_pixel_tolerance,
        mae_tolerance=args.mae_tolerance,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
