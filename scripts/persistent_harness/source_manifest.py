from __future__ import annotations

from hashlib import sha256
from typing import Iterable


SOURCE_MANIFEST_NAME = "REMOTE_HARNESS_SOURCE_SHA256.txt"


def source_manifest_bytes(entries: Iterable[tuple[str, bytes]]) -> bytes:
    lines: list[str] = []
    for relative, data in sorted(entries, key=lambda item: item[0]):
        if relative == SOURCE_MANIFEST_NAME or relative == ".gitignore":
            continue
        lines.append(f"{sha256(data).hexdigest()}  {relative}")
    return ("\n".join(lines) + "\n").encode("utf-8")
