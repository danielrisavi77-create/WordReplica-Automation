"""Hostile-input triage for corpus documents.

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

Everything here reads the ZIP *central directory* and raw part bytes only. It
never extracts to a temporary file, never resolves an external reference, and
never raises: the ingestor feeds it whatever the web returned, and a crash here
would stop a twenty-thousand-document batch.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path, PurePosixPath
import re
from zipfile import BadZipFile, ZipFile

from lxml import etree

__all__ = [
    "DEFAULT_LIMITS",
    "PackageLimits",
    "RiskClass",
    "RiskTriage",
    "SAFE_XML_PARSER_KWARGS",
    "ZipEnvelopeReport",
    "inspect_zip_envelope",
    "safe_xml_parser",
    "triage_document",
]

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_TYPE_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

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

    The defaults are deliberately far above any real document -- the measured
    corpus p99 is under a megabyte -- because these are bomb guards, not
    style rules. A legitimate document should never come close.
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


def _unsafe_member_name(name: str) -> bool:
    """True when extracting this member could escape the package directory.

    The lab never extracts, but a name that tries to is a strong signal about
    intent, and downstream consumers (Word, a future minimizer) might.
    """
    if not name or name.startswith(("/", "\\")):
        return True
    if _DRIVE_PREFIX.match(name):
        return True
    normalized = name.replace("\\", "/")
    return any(part == ".." for part in normalized.split("/"))


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
    head = data[:8192].lstrip()
    if b"<!DOCTYPE" in head or b"<!ENTITY" in data[:65536]:
        return True
    return False


def _max_depth(root) -> int:
    depth = 0
    stack = [(root, 1)]
    while stack:
        node, level = stack.pop()
        if level > depth:
            depth = level
        for child in node:
            stack.append((child, level + 1))
    return depth


def _is_xml_part(name: str) -> bool:
    lowered = name.lower()
    return lowered.endswith((".xml", ".rels"))


def _external_relationship_hazards(root) -> list[str]:
    hazards: list[str] = []
    for node in root.findall(f"{{{_REL_NS}}}Relationship"):
        if node.get("TargetMode") != "External":
            continue
        target = (node.get("Target") or "").strip()
        rel_type = (node.get("Type") or "").rsplit("/", 1)[-1]
        lowered = target.lower()
        if lowered.startswith("file:") or target.startswith(("\\\\", "//")) or _DRIVE_PREFIX.match(target):
            hazards.append(f"external reference to a local or UNC path: {target[:120]}")
        elif rel_type in _DANGEROUS_EXTERNAL_REL_TYPES:
            hazards.append(f"external {rel_type} reference: {target[:120]}")
    return hazards


def _content_type_gaps(parts: dict[str, bytes], root) -> list[str]:
    defaults = {node.get("Extension", "").lower() for node in root.findall(f"{{{_CT_NS}}}Default")}
    overrides = {(node.get("PartName") or "").lstrip("/") for node in root.findall(f"{{{_CT_NS}}}Override")}
    missing: list[str] = []
    for name in parts:
        if name.endswith("/") or name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        extension = PurePosixPath(name).suffix.lstrip(".").lower()
        if name not in overrides and extension not in defaults:
            missing.append(name)
    return sorted(missing)


def _resolve_target(owner_part: str, target: str) -> str:
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


def _dangling_relationship_targets(parts: dict[str, bytes], roots: dict[str, object]) -> list[str]:
    missing: list[str] = []
    for name, root in roots.items():
        if not name.endswith(".rels"):
            continue
        for node in root.findall(f"{{{_REL_NS}}}Relationship"):
            if node.get("TargetMode") == "External":
                continue
            resolved = _resolve_target(name, node.get("Target") or "")
            if resolved and resolved not in parts:
                missing.append(f"{name}:{node.get('Id')}->{resolved}")
    return sorted(missing)


def triage_document(path: Path, *, limits: PackageLimits = DEFAULT_LIMITS) -> RiskTriage:
    """Classify a package as VALID, RECOVERABLE or HOSTILE without opening Word.

    Hostility always wins: a package that is both malformed and macro-enabled is
    quarantined, not merely flagged. Anything that cannot be established counts
    against the document, never for it.
    """
    path = Path(path)
    envelope = inspect_zip_envelope(path, limits)
    if not envelope.ok:
        return RiskTriage(RiskClass.HOSTILE, envelope.reasons, envelope, word_open_allowed=False)

    try:
        with ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
    except (BadZipFile, OSError, ValueError, RuntimeError) as exc:
        report = _empty_envelope(f"package could not be read: {exc}")
        return RiskTriage(RiskClass.HOSTILE, report.reasons, report, word_open_allowed=False)

    hostile: list[str] = []
    recoverable: list[str] = []

    lowered_names = {name.lower(): name for name in parts}
    for macro_part in _MACRO_PARTS & lowered_names.keys():
        hostile.append(f"macro project part present: {lowered_names[macro_part]}")
    for lowered, name in lowered_names.items():
        if lowered.startswith("word/activex/"):
            hostile.append(f"ActiveX control part present: {name}")
        if lowered.startswith("word/embeddings/") and parts[name].startswith(_EXECUTABLE_MAGIC):
            hostile.append(f"embedded executable payload: {name}")

    content_types = parts.get("[Content_Types].xml", b"")
    if b"macroEnabled" in content_types:
        hostile.append("macro-enabled content type declared")

    parser = safe_xml_parser()
    roots: dict[str, object] = {}
    for name, data in parts.items():
        if not _is_xml_part(name):
            continue
        if _declares_dtd_or_entity(data):
            hostile.append(f"DOCTYPE or ENTITY declaration in {name}")
            continue
        try:
            root = etree.fromstring(data, parser=parser)
        except etree.XMLSyntaxError as exc:
            recoverable.append(f"malformed XML in {name}: {exc}")
            continue
        if _max_depth(root) > limits.max_xml_depth:
            hostile.append(f"XML nesting depth exceeds {limits.max_xml_depth} in {name}")
            continue
        roots[name] = root

    for name, root in roots.items():
        if name.endswith(".rels"):
            hostile.extend(f"{name}: {hazard}" for hazard in _external_relationship_hazards(root))

    ct_root = roots.get("[Content_Types].xml")
    if ct_root is None:
        recoverable.append("[Content_Types].xml is missing or unreadable")
    else:
        gaps = _content_type_gaps(parts, ct_root)
        if gaps:
            recoverable.append(f"parts with no declared content type: {gaps[:5]}")

    dangling = _dangling_relationship_targets(parts, roots)
    if dangling:
        recoverable.append(f"relationships pointing at absent parts: {dangling[:5]}")

    if hostile:
        return RiskTriage(RiskClass.HOSTILE, tuple(hostile), envelope, word_open_allowed=False)
    if recoverable:
        return RiskTriage(RiskClass.RECOVERABLE, tuple(recoverable), envelope, word_open_allowed=True)
    return RiskTriage(RiskClass.VALID, (), envelope, word_open_allowed=True)
