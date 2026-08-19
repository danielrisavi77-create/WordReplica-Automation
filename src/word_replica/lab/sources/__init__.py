"""Where corpus documents come from, and what we are allowed to keep.

Three adapters, one interface. The interface exists because the retention rule
differs per source and getting it wrong is a licensing problem, not a bug:

* **LibreOffice** ``sw/qa/extras/*/data`` -- MPL-2.0/LGPL project test data,
  redistributable. ~1,644 files in ooxmlexport alone, ~16 KB each, and by far
  the highest density of OOXML edge cases per megabyte anywhere.
* **Apache POI** ``test-data/document`` -- Apache-2.0, redistributable.
  Documents attached to specific historical crashes and bugs.
* **docx-corpus** -- 736k real documents from the public web. ODC-BY covers the
  *metadata*; each document's copyright stays with its author. Bytes are never
  retained, only measurements and pseudonymised structure clones.

LibreOffice is read through gitiles rather than the GitHub API: gitiles has no
rate limit (api.github.com allows 60 requests an hour unauthenticated and was
already exhausted when this was written), it serves both listings and blobs
from one origin, and it was measured at roughly half the latency.
"""
from __future__ import annotations

import base64
from dataclasses import dataclass, field
from enum import StrEnum
import json
from pathlib import Path
import random
import time
from typing import Callable, Iterator, Protocol, Sequence
import urllib.parse
import urllib.request

__all__ = [
    "CircuitOpen",
    "CorpusSource",
    "DocRef",
    "PoliteFetcher",
    "RetentionPolicy",
    "apache_poi_source",
    "libreoffice_source",
    "local_tree_source",
]

USER_AGENT = "WordReplicaLab/1.0 (+compatibility research)"

_WORD_SUFFIXES = (".docx", ".docm", ".dotx", ".dotm")

# Directories under sw/qa/extras that carry Word documents. ooxmlexport is the
# bulk of it; the others are smaller but cover import, layout and legacy paths.
LIBREOFFICE_DIRECTORIES: tuple[str, ...] = (
    "ooxmlexport",
    "ooxmlimport",
    "ooxmlfieldexport",
    "ww8export",
    "ww8import",
    "layout",
    "uiwriter",
    "cjk",
    "embedded_fonts",
    "indexing",
)


class RetentionPolicy(StrEnum):
    """What may be kept on disk once a document has been measured."""

    KEEP_VERBATIM = "KEEP_VERBATIM"   # redistributable project test data
    DISCARD_BYTES = "DISCARD_BYTES"   # third-party content: measure, then delete


class CircuitOpen(RuntimeError):
    """Raised instead of continuing to hammer a host that is refusing us."""


@dataclass(frozen=True, slots=True)
class DocRef:
    """One candidate document, before anything has been downloaded."""

    source: str
    source_ref: str          # stable path or id within the source
    name: str
    url: str | None = None
    revision: str | None = None   # upstream blob id, for cheap change detection
    size_hint: int | None = None


def _urlopen(url: str, *, timeout: float) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


@dataclass(slots=True)
class PoliteFetcher:
    """Retry, backoff and a circuit breaker around one HTTP transport.

    The circuit breaker is the part that matters at corpus scale: a crawl of
    tens of thousands of documents against a host that has started refusing
    should stop and say so, not spend hours recording the same error and
    marking every document as failed.
    """

    transport: Callable[..., bytes] = _urlopen
    sleep: Callable[[float], None] = time.sleep
    timeout: float = 45.0
    max_attempts: int = 4
    base_delay: float = 1.0
    jitter: float = 0.15
    circuit_threshold: int = 20
    _consecutive_failures: int = field(default=0, init=False)

    def get(self, url: str) -> bytes:
        if self._consecutive_failures >= self.circuit_threshold:
            raise CircuitOpen(
                f"{self._consecutive_failures} consecutive failures; refusing to keep requesting {url}"
            )
        last: BaseException | None = None
        for attempt in range(self.max_attempts):
            try:
                payload = self.transport(url, timeout=self.timeout)
            except Exception as exc:  # noqa: BLE001 - any transport failure is retryable
                last = exc
                if attempt + 1 < self.max_attempts:
                    delay = self.base_delay * (2 ** attempt)
                    if self.jitter:
                        delay += random.uniform(0, self.jitter * delay)
                    self.sleep(delay)
            else:
                self._consecutive_failures = 0
                return payload
        self._consecutive_failures += 1
        assert last is not None
        raise last


class CorpusSource(Protocol):
    name: str
    licence: str
    retention: RetentionPolicy

    def list_candidates(self, *, limit: int | None = None) -> Iterator[DocRef]: ...
    def fetch(self, ref: DocRef) -> bytes: ...


# --- LibreOffice -------------------------------------------------------------

_GITILES_ROOT = "https://git.libreoffice.org/core/+/refs/heads/master"


def _strip_xssi(payload: bytes) -> bytes:
    """Gitiles prefixes JSON with )]}' to defeat cross-site script inclusion."""
    return payload.lstrip(b")]}'").lstrip()


@dataclass(slots=True)
class _LibreOfficeSource:
    fetcher: PoliteFetcher
    directories: Sequence[str]
    name: str = "libreoffice"
    licence: str = "MPL-2.0 / LGPL-3.0 (LibreOffice project test data)"
    retention: RetentionPolicy = RetentionPolicy.KEEP_VERBATIM

    def list_candidates(self, *, limit: int | None = None) -> Iterator[DocRef]:
        seen = 0
        for directory in self.directories:
            path = f"sw/qa/extras/{directory}/data"
            try:
                payload = self.fetcher.get(f"{_GITILES_ROOT}/{path}?format=JSON")
            except CircuitOpen:
                raise
            except Exception:
                continue  # a directory that does not exist in this checkout
            entries = json.loads(_strip_xssi(payload)).get("entries", [])
            for entry in entries:
                if entry.get("type") != "blob":
                    continue
                entry_name = entry.get("name", "")
                if not entry_name.lower().endswith(_WORD_SUFFIXES):
                    continue
                yield DocRef(
                    source=self.name,
                    source_ref=f"{directory}/data/{entry_name}",
                    name=entry_name,
                    url=f"{_GITILES_ROOT}/{path}/{urllib.parse.quote(entry_name)}?format=TEXT",
                    revision=entry.get("id"),
                )
                seen += 1
                if limit and seen >= limit:
                    return

    def fetch(self, ref: DocRef) -> bytes:
        return base64.b64decode(self.fetcher.get(ref.url or ""))


def libreoffice_source(
    *,
    directories: Sequence[str] = LIBREOFFICE_DIRECTORIES,
    fetcher: PoliteFetcher | None = None,
) -> _LibreOfficeSource:
    return _LibreOfficeSource(fetcher=fetcher or PoliteFetcher(), directories=tuple(directories))


# --- Apache POI --------------------------------------------------------------

_POI_TREE = "https://api.github.com/repos/apache/poi/git/trees/trunk?recursive=1"
_POI_RAW = "https://raw.githubusercontent.com/apache/poi/trunk/"


@dataclass(slots=True)
class _ApachePoiSource:
    fetcher: PoliteFetcher
    prefix: str = "test-data/document/"
    name: str = "apache_poi"
    licence: str = "Apache-2.0 (Apache POI test data)"
    retention: RetentionPolicy = RetentionPolicy.KEEP_VERBATIM

    def list_candidates(self, *, limit: int | None = None) -> Iterator[DocRef]:
        tree = json.loads(self.fetcher.get(_POI_TREE))
        if tree.get("truncated"):
            # Listing part of a tree and calling it the corpus is exactly the
            # silent-truncation failure this lab is meant not to have.
            raise RuntimeError("apache/poi tree listing was truncated; cannot enumerate the corpus")
        seen = 0
        for entry in tree.get("tree", []):
            path = entry.get("path", "")
            if entry.get("type") != "blob" or not path.startswith(self.prefix):
                continue
            if not path.lower().endswith(_WORD_SUFFIXES):
                continue
            yield DocRef(
                source=self.name,
                source_ref=path,
                name=path.rsplit("/", 1)[-1],
                url=_POI_RAW + urllib.parse.quote(path),
                revision=entry.get("sha"),
                size_hint=entry.get("size"),
            )
            seen += 1
            if limit and seen >= limit:
                return

    def fetch(self, ref: DocRef) -> bytes:
        return self.fetcher.get(ref.url or "")


def apache_poi_source(*, fetcher: PoliteFetcher | None = None) -> _ApachePoiSource:
    return _ApachePoiSource(fetcher=fetcher or PoliteFetcher())


# --- local tree --------------------------------------------------------------

@dataclass(slots=True)
class _LocalTreeSource:
    root: Path
    name: str = "local"
    licence: str = "local"
    retention: RetentionPolicy = RetentionPolicy.KEEP_VERBATIM

    def list_candidates(self, *, limit: int | None = None) -> Iterator[DocRef]:
        for index, path in enumerate(sorted(self.root.rglob("*.docx"))):
            if limit and index >= limit:
                return
            yield DocRef(
                source=self.name,
                source_ref=path.relative_to(self.root).as_posix(),
                name=path.name,
                url=str(path),
                size_hint=path.stat().st_size,
            )

    def fetch(self, ref: DocRef) -> bytes:
        return (self.root / ref.source_ref).read_bytes()


def local_tree_source(root: Path, *, fetcher: PoliteFetcher | None = None) -> _LocalTreeSource:
    """A directory already on disk. ``fetcher`` is accepted and ignored so the
    caller can treat every source identically."""
    return _LocalTreeSource(root=Path(root))
