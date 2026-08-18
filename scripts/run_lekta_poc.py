"""Foreground, visible harness for the Lekta repair-poc local proof of concept.

Reads only the fixed local package directory below — never infers files
from a broad folder search or picks the newest file by timestamp. Records
commit SHA and source hash before running, calls the repair-poc CLI
synchronously (no background process), and writes diagnostics under
C:\\WordReplica-Automation\\diagnostics\\repair-poc-<UTC>. Retains explicit
original/target/output paths without ever copying the user's DOCX into
Git — the diagnostics root lives outside the repo entirely.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path

DEFAULT_PACKAGE_DIR = Path(r"C:\WordReplica-Automation\poc\kalogjera\package")
DEFAULT_DIAGNOSTICS_ROOT = Path(r"C:\WordReplica-Automation\diagnostics")
REQUIRED_BRANCH = "automation-dev"
# Design/plan documentation may legitimately be mid-edit during development;
# every other tracked file must be clean before a harness run is trusted.
ALLOWED_DIRTY_PREFIXES = (
    "docs/superpowers/plans/",
    "docs/superpowers/specs/",
)


class HarnessPreflightError(RuntimeError):
    pass


def _run(run_subprocess, args: list[str], *, cwd: Path):
    return run_subprocess(args, cwd=str(cwd), capture_output=True, text=True, check=True)


def _current_branch(repo_root: Path, run_subprocess) -> str:
    result = _run(run_subprocess, ["git", "rev-parse", "--abbrev-ref", "HEAD"], cwd=repo_root)
    return result.stdout.strip()


def _commit_sha(repo_root: Path, run_subprocess) -> str:
    result = _run(run_subprocess, ["git", "rev-parse", "HEAD"], cwd=repo_root)
    return result.stdout.strip()


def _dirty_production_files(repo_root: Path, run_subprocess) -> list[str]:
    result = _run(run_subprocess, ["git", "status", "--porcelain"], cwd=repo_root)
    dirty = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue
        path = line[3:].strip().replace("\\", "/")
        if any(path.startswith(prefix) for prefix in ALLOWED_DIRTY_PREFIXES):
            continue
        dirty.append(path)
    return dirty


def _sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def required_package_paths(package_dir: Path) -> dict[str, Path]:
    return {
        "original": package_dir / "original.docx",
        "target": package_dir / "target.docx",
        "contract": package_dir / "contract.json",
        "public_key": package_dir / "public-key.spki.b64url",
    }


def preflight(*, repo_root: Path, package_dir: Path, run_subprocess) -> dict[str, str]:
    """Raises HarnessPreflightError before anything else runs (in particular,
    before the CLI is invoked and long before Word could open) if any gate
    fails: wrong branch, dirty production files, or a missing package file.
    """
    branch = _current_branch(repo_root, run_subprocess)
    if branch != REQUIRED_BRANCH:
        raise HarnessPreflightError(f"Must run on branch {REQUIRED_BRANCH!r}, not {branch!r}")

    dirty = _dirty_production_files(repo_root, run_subprocess)
    if dirty:
        raise HarnessPreflightError(f"Refusing to run against a dirty working tree: {dirty}")

    paths = required_package_paths(package_dir)
    missing = [str(path) for path in paths.values() if not path.is_file()]
    if missing:
        raise HarnessPreflightError(f"Missing required package file(s): {missing}")

    return {"branch": branch, "commit_sha": _commit_sha(repo_root, run_subprocess)}


def run_harness(
    *,
    repo_root: Path,
    python_executable: Path,
    package_dir: Path,
    diagnostics_root: Path,
    renderer: str,
    stop_after_event: int | None = None,
    now: datetime | None = None,
    run_subprocess=subprocess.run,
) -> int:
    now = now or datetime.now(timezone.utc)
    preflight_info = preflight(repo_root=repo_root, package_dir=package_dir, run_subprocess=run_subprocess)
    paths = required_package_paths(package_dir)
    source_sha256 = _sha256_file(paths["original"])

    diagnostics_dir = diagnostics_root / f"repair-poc-{now.strftime('%Y%m%dT%H%M%SZ')}"
    diagnostics_dir.mkdir(parents=True, exist_ok=True)
    output_dir = diagnostics_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    command = [
        str(python_executable), "-m", "word_replica.cli", "repair-poc",
        "--original", str(paths["original"]),
        "--target", str(paths["target"]),
        "--contract", str(paths["contract"]),
        "--public-key", str(paths["public_key"]),
        "--output-dir", str(output_dir),
        "--renderer", renderer,
    ]
    if stop_after_event is not None:
        command += ["--stop-after-event", str(stop_after_event)]

    process = run_subprocess(command, cwd=str(repo_root))
    exit_code = int(getattr(process, "returncode", 2))

    report_path = output_dir / "repair_poc_report.json"
    diagnostics = {
        "branch": preflight_info["branch"],
        "commit_sha": preflight_info["commit_sha"],
        "source_sha256": source_sha256,
        "renderer": renderer,
        "started_at": now.isoformat(),
        "exit_code": exit_code,
        "report_path": str(report_path) if report_path.is_file() else None,
    }
    (diagnostics_dir / "harness_diagnostics.json").write_text(
        json.dumps(diagnostics, indent=2, sort_keys=True), encoding="utf-8"
    )
    return exit_code


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--package-dir", type=Path, default=DEFAULT_PACKAGE_DIR)
    parser.add_argument("--diagnostics-root", type=Path, default=DEFAULT_DIAGNOSTICS_ROOT)
    parser.add_argument("--renderer", choices=["word", "pure-docx"], default="word")
    parser.add_argument("--stop-after-event", type=int, default=None)
    parser.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--python", type=Path, default=Path(sys.executable))
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run_harness(
            repo_root=args.repo_root,
            python_executable=args.python,
            package_dir=args.package_dir,
            diagnostics_root=args.diagnostics_root,
            renderer=args.renderer,
            stop_after_event=args.stop_after_event,
        )
    except HarnessPreflightError as exc:
        print(f"Lekta POC harness preflight failed: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
