from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path

from word_replica.domain.errors import SourceIntegrityError


@dataclass(frozen=True, slots=True)
class SourceSnapshot:
    path: Path
    sha256: str
    size: int


def sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def capture_source(path: Path) -> SourceSnapshot:
    resolved = path.resolve()
    return SourceSnapshot(resolved, sha256_file(resolved), resolved.stat().st_size)


def assert_source_unchanged(snapshot: SourceSnapshot) -> None:
    if (
        not snapshot.path.exists()
        or snapshot.path.stat().st_size != snapshot.size
        or sha256_file(snapshot.path) != snapshot.sha256
    ):
        raise SourceIntegrityError(f"Source changed during reconstruction: {snapshot.path}")
