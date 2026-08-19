"""Hostile-input triage and single-pass package scanning.

Every document the lab ingests comes from the public web, so the first thing
that happens to it is a decision about whether it is safe to work with at all.
Three things make that decision non-optional here:

* ``opc.package_reader`` parses with lxml's *default* settings, which resolve
  entities and allow network access. A billion-laughs or XXE part in any corpus
  document would hang the worker or exfiltrate a file.
* A zip bomb is a few kilobytes on disk and hundreds of megabytes once read, on
  a machine with single-digit gigabytes free.
* Macro projects, ActiveX controls, remote templates and embedded executables
  must never reach Microsoft Word. Not opening them is the entire isolation
  strategy on a machine that cannot host a disposable VM.

The envelope check reads the ZIP *central directory* only -- nothing is ever
extracted to a temporary file, no external reference is ever resolved, and
nothing here raises: the ingestor feeds it whatever the web returned, and a
crash would stop a twenty-thousand-document batch.

``read_package`` is the shared primitive: one read, one parse, one tree walk,
consumed by both triage and fingerprinting. Doing that work twice was
measurably the dominant cost of the whole pipeline.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path, PurePosixPath
import re
from typing import Any
from zipfile import BadZipFile, ZipFile

from lxml import etree

__all__ = [
    "DEFAULT_LIMITS",
    "PackageLimits",
    "PackageScan",
    "RiskClass",
    "RiskTriage",
    "SAFE_XML_PARSER_KWARGS",
    "ZipEnvelopeReport",
    "inspect_zip_envelope",
    "read_package",
    "safe_xml_parser",
    "triage_document",
    "triage_scan",
]

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# External references of these kinds pull content or code into the document at
# open time. An external hyperlink deliberately is not one of them: hyperlinks
# are ubiquitous and quarantining them would cost most of the corpus for nothing.
_DANGEROUS_EXTERNAL_REL_TYPES = frozenset({
    "oleObject",
    "package",
    "attachedTemplate",
    "frame",
    "subDocument",
    "aFChunk",
})

_DRIVE_PREFIX = re.compile(r"^[A-Za-z]:")
_EXECUTABLE_MAGIC = (b"MZ", b"\x7fELF")
_MACRO_PARTS = frozenset({"word/vbaproject.bin", "word/vbadata.xml"})
_INSTR_TEXT_TAG = f"{{{_W_NS}}}instrText"


class RiskClass(StrEnum):
    """What the lab is allowed to do with a document.

    VALID        -- work with it normally; full fidelity is the goal.
    RECOVERABLE  -- structurally broken but not dangerous. The product must say
                    so rather than silently returning a mangled reconstruction.
    HOSTILE      -- never hand to Word. Safe refusal is the success criterion,
                    not fidelity.
    """

    VALID = "VALID"
    RECOVERABLE = "RECOVERABLE"
    HOSTILE = "HOSTILE"


@dataclass(frozen=True, slots=True)
class PackageLimits:
    """Caps applied before any part is decompressed.

    The defaults sit far above any real document -- the measured corpus p99 is
    under a megabyte -- because these are bomb guards, not style rules. A
    legitimate document should never come close to one.
    """

    max_file_bytes: int = 48 * 1024 ** 2
    max_total_uncompressed_bytes: int = 512 * 1024 ** 2
    max_part_bytes: int = 128 * 1024 ** 2
    max_compression_ratio: float = 200.0
    max_member_count: int = 20_000
    max_xml_depth: int = 512


DEFAULT_LIMITS = PackageLimits()

# resolve_entities kills entity expansion (billion laughs) and external entity
# substitution (XXE); no_network blocks DTD/entity fetches; load_dtd keeps an
# internal subset from being processed at all; huge_tree keeps libxml2's own
# depth and size guards on.
SAFE_XML_PARSER_KWARGS = {
    "resolve_entities": False,
    "no_network": True,
    "load_dtd": False,
    "dtd_validation": False,
    "huge_tree": False,
    "recover": False,
}


def safe_xml_parser() -> etree.XMLParser:
    """A parser that cannot be turned into a denial-of-service or a file read."""
    return etree.XMLParser(**SAFE_XML_PARSER_KWARGS)


@dataclass(frozen=True, slots=True)
class ZipEnvelopeReport:
    ok: bool
    reasons: tuple[str, ...]
    member_count: int
    total_uncompressed_bytes: int
    total_compressed_bytes: int
    max_part_bytes: int
    compression_ratio: float
    names: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class RiskTriage:
    risk_class: RiskClass
    reasons: tuple[str, ...]
    envelope: ZipEnvelopeReport
    word_open_allowed: bool


@dataclass(frozen=True, slots=True)
class PackageScan:
    """One read, one parse, one tree walk -- shared by triage and fingerprinting.

    Element counts are keyed by lxml's own Clark-notation tag string, so the
    walk never has to build a key: the tag *is* the key.
    """

    envelope: ZipEnvelopeReport
    parts: dict[str, bytes] = field(default_factory=dict)
    roots: dict[str, Any] = field(default_factory=dict)
    malformed: dict[str, str] = field(default_factory=dict)
    entity_parts: tuple[str, ...] = ()
    element_counts: Counter = field(default_factory=Counter)
    distinct_qname_count: int = 0
    total_element_count: int = 0
    max_xml_depth: int = 0
    max_table_depth: int = 0
    ns_mask: int = 0
    unknown_namespaces: tuple[str, ...] = ()
    field_instructions: tuple[str, ...] = ()
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None and self.envelope.ok


def _unsafe_member_name(name: str) -> bool:
    """True when extracting this member could escape the package directory.

    The lab never extracts, but a name that tries to is a strong signal about
    intent, and downstream consumers (Word, a future minimizer) might.
    """
    if not name or name.startswith(("/", "\\")):
        return True
    if _DRIVE_PREFIX.match(name):
        return True
    return any(part == ".." for part in name.replace("\\", "/").split("/"))


def _empty_envelope(reason: str) -> ZipEnvelopeReport:
    return ZipEnvelopeReport(
        ok=False,
        reasons=(reason,),
        member_count=0,
        total_uncompressed_bytes=0,
        total_compressed_bytes=0,
        max_part_bytes=0,
        compression_ratio=0.0,
    )


def inspect_zip_envelope(path: Path, limits: PackageLimits = DEFAULT_LIMITS) -> ZipEnvelopeReport:
    """Judge a package from its central directory alone.

    Nothing is decompressed. That is the point: a zip bomb is only dangerous
    once you read it, and every field needed to recognise one -- declared size,
    stored size, member count, member names -- is metadata.
    """
    path = Path(path)
    try:
        file_bytes = path.stat().st_size
    except OSError as exc:
        return _empty_envelope(f"unreadable file: {exc}")

    reasons: list[str] = []
    if file_bytes > limits.max_file_bytes:
        reasons.append(f"file exceeds {limits.max_file_bytes} bytes ({file_bytes})")

    try:
        with ZipFile(path) as archive:
            infos = archive.infolist()
    except (BadZipFile, OSError, ValueError) as exc:
        return _empty_envelope(f"not a readable zip archive: {exc}")

    names = tuple(info.filename for info in infos)
    total_uncompressed = sum(info.file_size for info in infos)
    total_compressed = sum(info.compress_size for info in infos)
    largest = max((info.file_size for info in infos), default=0)
    ratio = total_uncompressed / max(1, total_compressed)

    if len(infos) > limits.max_member_count:
        reasons.append(f"members exceed {limits.max_member_count} ({len(infos)})")
    if total_uncompressed > limits.max_total_uncompressed_bytes:
        reasons.append(
            f"declared uncompressed size exceeds {limits.max_total_uncompressed_bytes} bytes ({total_uncompressed})"
        )
    if largest > limits.max_part_bytes:
        reasons.append(f"single part exceeds {limits.max_part_bytes} bytes ({largest})")
    if ratio > limits.max_compression_ratio:
        reasons.append(f"compression ratio exceeds {limits.max_compression_ratio} ({ratio:.1f})")
    unsafe = sorted(name for name in names if _unsafe_member_name(name))
    if unsafe:
        reasons.append(f"path traversal in member names: {unsafe[:5]}")

    return ZipEnvelopeReport(
        ok=not reasons,
        reasons=tuple(reasons),
        member_count=len(infos),
        total_uncompressed_bytes=total_uncompressed,
        total_compressed_bytes=total_compressed,
        max_part_bytes=largest,
        compression_ratio=ratio,
        names=names,
    )


def _declares_dtd_or_entity(data: bytes) -> bool:
    """A byte scan, run before any parse.

    The plan is not to parse a DOCTYPE safely -- it is to never parse one. No
    part of a WordprocessingML package has a legitimate reason to declare a
    doctype or an entity, so its presence is treated as intent.
    """
    return b"<!DOCTYPE" in data[:8192] or b"<!ENTITY" in data[:65536]


def _is_xml_part(name: str) -> bool:
    lowered = name.lower()
    return lowered.endswith((".xml", ".rels"))


class _TagIndex:
    """Per-tag facts, computed once per distinct tag rather than per element.

    A real document has tens of thousands of elements but only a few hundred
    distinct tags, so memoizing here is what makes the walk affordable.
    """

    __slots__ = ("_cache", "uri_to_bit", "unknown")

    def __init__(self, uri_to_bit: dict[str, int]) -> None:
        self._cache: dict[str, tuple[int, bool, bool]] = {}
        self.uri_to_bit = uri_to_bit
        self.unknown: set[str] = set()

    def get(self, tag: str) -> tuple[int, bool, bool]:
        """(namespace bit or -1, is w:tbl, is w:instrText)."""
        hit = self._cache.get(tag)
        if hit is not None:
            return hit
        if tag.startswith("{"):
            uri, _, local = tag[1:].partition("}")
        else:
            uri, local = "", tag
        bit = self.uri_to_bit.get(uri, -1)
        if bit < 0 and uri:
            self.unknown.add(uri)
        entry = (bit, uri == _W_NS and local == "tbl", tag == _INSTR_TEXT_TAG)
        self._cache[tag] = entry
        return entry


def read_package(
    path: Path,
    *,
    limits: PackageLimits = DEFAULT_LIMITS,
    uri_to_bit: dict[str, int] | None = None,
) -> PackageScan:
    """Read, parse and walk a package exactly once. Never raises.

    ``uri_to_bit`` lets the fingerprint module supply its persisted namespace
    bit assignment without this module having to know about it.
    """
    path = Path(path)
    envelope = inspect_zip_envelope(path, limits)
    if not envelope.ok:
        return PackageScan(envelope=envelope, error="; ".join(envelope.reasons) or "unreadable package")

    try:
        with ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
    except (BadZipFile, OSError, ValueError, RuntimeError) as exc:
        return PackageScan(envelope=envelope, error=f"package could not be read: {exc}")

    parser = safe_xml_parser()
    index = _TagIndex(uri_to_bit or {})
    roots: dict[str, Any] = {}
    malformed: dict[str, str] = {}
    entity_parts: list[str] = []
    counts: Counter = Counter()
    instructions: list[str] = []
    total_elements = 0
    max_depth = 0
    max_table_depth = 0
    ns_mask = 0

    for name, data in parts.items():
        if not _is_xml_part(name):
            continue
        if _declares_dtd_or_entity(data):
            entity_parts.append(name)
            continue
        try:
            root = etree.fromstring(data, parser=parser)
        except etree.XMLSyntaxError as exc:
            malformed[name] = str(exc)[:300]
            continue
        roots[name] = root

        # Table nesting only means layout inside a story part; a nested tbl in
        # styles.xml is a definition, not a document structure.
        count_tables = "document" in name.lower()
        stack: list[tuple[Any, int, int]] = [(root, 1, 0)]
        while stack:
            node, depth, table_depth = stack.pop()
            tag = node.tag
            if type(tag) is not str:
                continue  # comments and processing instructions
            total_elements += 1
            if depth > max_depth:
                max_depth = depth
            counts[tag] += 1
            bit, is_table, is_instr = index.get(tag)
            if bit >= 0:
                ns_mask |= 1 << bit
            child_table_depth = table_depth
            if is_table and count_tables:
                child_table_depth = table_depth + 1
                if child_table_depth > max_table_depth:
                    max_table_depth = child_table_depth
            if is_instr and node.text:
                instructions.append(node.text)
            for child in node:
                stack.append((child, depth + 1, child_table_depth))

    return PackageScan(
        envelope=envelope,
        parts=parts,
        roots=roots,
        malformed=malformed,
        entity_parts=tuple(entity_parts),
        element_counts=counts,
        distinct_qname_count=len(counts),
        total_element_count=total_elements,
        max_xml_depth=max_depth,
        max_table_depth=max_table_depth,
        ns_mask=ns_mask,
        unknown_namespaces=tuple(sorted(index.unknown)),
        field_instructions=tuple(instructions),
    )


def _external_relationship_hazards(root: Any) -> list[str]:
    hazards: list[str] = []
    for node in root.findall(f"{{{_REL_NS}}}Relationship"):
        if node.get("TargetMode") != "External":
            continue
        target = (node.get("Target") or "").strip()
        rel_type = (node.get("Type") or "").rsplit("/", 1)[-1]
        if target.lower().startswith("file:") or target.startswith(("\\\\", "//")) or _DRIVE_PREFIX.match(target):
            hazards.append(f"external reference to a local or UNC path: {target[:120]}")
        elif rel_type in _DANGEROUS_EXTERNAL_REL_TYPES:
            hazards.append(f"external {rel_type} reference: {target[:120]}")
    return hazards


def content_type_gaps(parts: dict[str, bytes], root: Any) -> list[str]:
    """Parts the package never declares a content type for."""
    defaults = {node.get("Extension", "").lower() for node in root.findall(f"{{{_CT_NS}}}Default")}
    overrides = {(node.get("PartName") or "").lstrip("/") for node in root.findall(f"{{{_CT_NS}}}Override")}
    missing: list[str] = []
    for name in parts:
        if name.endswith("/") or name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        if name not in overrides and PurePosixPath(name).suffix.lstrip(".").lower() not in defaults:
            missing.append(name)
    return sorted(missing)


def resolve_relationship_target(owner_part: str, target: str) -> str:
    if target.startswith("/"):
        candidate = target.lstrip("/")
    else:
        owner_dir = PurePosixPath(owner_part).parent
        if owner_dir.name == "_rels":
            owner_dir = owner_dir.parent
        candidate = str(PurePosixPath(owner_dir, target))
    resolved: list[str] = []
    for part in PurePosixPath(candidate).parts:
        if part == "..":
            if resolved:
                resolved.pop()
        elif part not in (".", ""):
            resolved.append(part)
    return "/".join(resolved)


def dangling_relationship_targets(parts: dict[str, bytes], roots: dict[str, Any]) -> list[str]:
    """Internal relationships pointing at parts the package does not contain."""
    missing: list[str] = []
    for name, root in roots.items():
        if not name.endswith(".rels"):
            continue
        for node in root.findall(f"{{{_REL_NS}}}Relationship"):
            if node.get("TargetMode") == "External":
                continue
            resolved = resolve_relationship_target(name, node.get("Target") or "")
            if resolved and resolved not in parts:
                missing.append(f"{name}:{node.get('Id')}->{resolved}")
    return sorted(missing)


def triage_scan(scan: PackageScan, *, limits: PackageLimits = DEFAULT_LIMITS) -> RiskTriage:
    """Classify an already-read package. Pure: no I/O, no parsing.

    Hostility always wins: a package that is both malformed and macro-enabled
    is quarantined, not merely flagged. Anything that cannot be established
    counts against the document, never for it.
    """
    if not scan.ok:
        reasons = scan.envelope.reasons or ((scan.error,) if scan.error else ("unreadable package",))
        return RiskTriage(RiskClass.HOSTILE, tuple(reasons), scan.envelope, word_open_allowed=False)

    parts, roots = scan.parts, scan.roots
    hostile: list[str] = []
    recoverable: list[str] = []

    lowered_names = {name.lower(): name for name in parts}
    for macro_part in sorted(_MACRO_PARTS & lowered_names.keys()):
        hostile.append(f"macro project part present: {lowered_names[macro_part]}")
    for lowered, name in lowered_names.items():
        if lowered.startswith("word/activex/"):
            hostile.append(f"ActiveX control part present: {name}")
        if lowered.startswith("word/embeddings/") and parts[name].startswith(_EXECUTABLE_MAGIC):
            hostile.append(f"embedded executable payload: {name}")

    if b"macroEnabled" in parts.get("[Content_Types].xml", b""):
        hostile.append("macro-enabled content type declared")

    for name in scan.entity_parts:
        hostile.append(f"DOCTYPE or ENTITY declaration in {name}")
    if scan.max_xml_depth > limits.max_xml_depth:
        hostile.append(f"XML nesting depth exceeds {limits.max_xml_depth} ({scan.max_xml_depth})")

    for name, root in roots.items():
        if name.endswith(".rels"):
            hostile.extend(f"{name}: {hazard}" for hazard in _external_relationship_hazards(root))

    for name, error in scan.malformed.items():
        recoverable.append(f"malformed XML in {name}: {error}")

    ct_root = roots.get("[Content_Types].xml")
    if ct_root is None:
        recoverable.append("[Content_Types].xml is missing or unreadable")
    else:
        gaps = content_type_gaps(parts, ct_root)
        if gaps:
            recoverable.append(f"parts with no declared content type: {gaps[:5]}")

    dangling = dangling_relationship_targets(parts, roots)
    if dangling:
        recoverable.append(f"relationships pointing at absent parts: {dangling[:5]}")

    if hostile:
        return RiskTriage(RiskClass.HOSTILE, tuple(hostile), scan.envelope, word_open_allowed=False)
    if recoverable:
        return RiskTriage(RiskClass.RECOVERABLE, tuple(recoverable), scan.envelope, word_open_allowed=True)
    return RiskTriage(RiskClass.VALID, (), scan.envelope, word_open_allowed=True)


def triage_document(path: Path, *, limits: PackageLimits = DEFAULT_LIMITS) -> RiskTriage:
    """Classify a package as VALID, RECOVERABLE or HOSTILE without opening Word."""
    return triage_scan(read_package(path, limits=limits), limits=limits)
