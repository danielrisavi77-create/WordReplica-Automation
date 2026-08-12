from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil


class MigrationCollisionError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class MigrationResult:
    copied: int
    skipped_identical: int


def _sha(path: Path) -> str:
    h = sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def import_docx_corpus(source_dir: Path, destination_dir: Path) -> MigrationResult:
    source_dir = Path(source_dir).resolve()
    destination_dir = Path(destination_dir).resolve()
    if not source_dir.is_dir():
        raise FileNotFoundError(source_dir)
    destination_dir.mkdir(parents=True, exist_ok=True)
    copied = 0
    skipped = 0
    for source in sorted((p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() == ".docx"), key=lambda p: p.name.lower()):
        target = destination_dir / source.name
        if target.exists():
            if _sha(source) != _sha(target):
                raise MigrationCollisionError(f"refusing to overwrite {target.name}: different SHA-256")
            skipped += 1
            continue
        shutil.copy2(source, target)
        copied += 1
    return MigrationResult(copied, skipped)
