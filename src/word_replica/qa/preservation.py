"""G10 -- package preservation, judged without Microsoft Word.

Every gate from G0 to G9 compares *parsed models*, and ``parser/parser.py``
models only what WordReplica reconstructs. Content outside that model is not
merely uncompared -- it is not represented at all, so a reconstruction that
loses it reports a clean pass. Word 2010 hides the same class of loss a second
way: it silently ignores w15/w16 in the source *and* the output, so rendering
both cancels the difference out.

Measured on the first real corpus the fidelity lab ingested, 602 of 1,830
documents (32.9 %) carry content in that blind spot -- w15 comment extensions,
w16 comment ids, ISO-29500 Strict, SVG, chartex. G10 reads raw package bytes
and is blind to none of it.

**Being quiet matters as much as being sensitive.** A gate that fires on
revision-save ids or a renumbered header part gets switched off, and then the
blind spot reopens. So the projection deliberately discards:

* every ``w:rsid*`` attribute and ``w:rsids`` block, and ``w15:paraId``
* ``docProps/core.xml`` timestamps and ``cp:revision``, ``app.xml`` ``TotalTime``
* concrete part *names* -- only content types and their counts are compared, so
  ``header3.xml`` standing in for ``header1.xml`` is not a defect
* ZIP member order and compression method

and compares exactly:

* which namespaces appear anywhere in the package
* how many parts of each content type there are
* the relationship graph by type, and whether anything dangles
* the bytes of every *opaque* part -- OLE payloads, custom XML, embeddings,
  diagrams. These have no model projection whatsoever, so byte identity is the
  only assertion available about them.

The comparison is symmetric on purpose. An output that *gained* a namespace or
an opaque part is as much a defect as one that lost it: spurious injection is
one of the ways a tool silently returns an incorrect document.
"""
from __future__ import annotations

from hashlib import sha256
from pathlib import Path
from typing import Any

from word_replica.lab.safety import (
    DEFAULT_LIMITS,
    PackageLimits,
    read_package,
    resolve_relationship_target,
)
from word_replica.qa.golden_audit import GateResult

__all__ = ["G10_GATE_NAMES", "build_preservation_gate", "g10_projection"]

# Convenience for callers that want the eleven-gate contract; build_golden_report
# takes this as gate_names. Golden #1 keeps the original ten until it is green
# on eleven, so this is opt-in.
G10_GATE_NAMES: tuple[str, ...] = tuple(f"G{index}" for index in range(11))

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Parts the document model has no representation for. Whatever is in them
# survives verbatim or it is lost, and nothing else in the pipeline would know.
_OPAQUE_PREFIXES = (
    "word/embeddings/",
    "word/activex/",
    "customxml/",
    "word/diagrams/",
    "word/vbaproject.bin",
    "word/vbadata.xml",
    "word/fonts/",
)

# Content types whose parts are opaque wherever they live.
_OPAQUE_CONTENT_TYPE_MARKERS = (
    "oleobject",
    "customxmlproperties",
    "vnd.ms-office.activex",
    "ms-word.vbaproject",
    "diagramdata",
    "diagramlayout",
    "obfuscatedfont",
)


def _content_type_map(parts: dict[str, bytes], roots: dict[str, Any]) -> dict[str, str]:
    """Part name -> declared content type, resolved through Default and Override."""
    root = roots.get("[Content_Types].xml")
    if root is None:
        return {}
    defaults = {
        (node.get("Extension") or "").lower(): (node.get("ContentType") or "")
        for node in root.findall(f"{{{_CT_NS}}}Default")
    }
    overrides = {
        (node.get("PartName") or "").lstrip("/"): (node.get("ContentType") or "")
        for node in root.findall(f"{{{_CT_NS}}}Override")
    }
    resolved: dict[str, str] = {}
    for name in parts:
        if name == "[Content_Types].xml":
            continue
        if name in overrides:
            resolved[name] = overrides[name]
            continue
        extension = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        resolved[name] = defaults.get(extension, "")
    return resolved


def _is_opaque(name: str, content_type: str) -> bool:
    lowered = name.lower()
    if lowered.startswith(_OPAQUE_PREFIXES):
        return True
    ct = content_type.lower()
    return any(marker in ct for marker in _OPAQUE_CONTENT_TYPE_MARKERS)


def _namespaces(roots: dict[str, Any]) -> list[str]:
    """Every namespace URI that actually appears on an element in the package.

    Presence, not element counts. A faithful reconstruction legitimately shifts
    counts in the main wordprocessingml namespace -- runs merge, redundant run
    properties drop -- and gating on that would make G10 noise. Presence is what
    answers the question the gate is asked: did a whole feature disappear?
    """
    seen: set[str] = set()
    for root in roots.values():
        for element in root.iter():
            tag = element.tag
            if type(tag) is str and tag.startswith("{"):
                seen.add(tag[1:].partition("}")[0])
    return sorted(seen)


def _markup_compatibility(roots: dict[str, Any]) -> dict[str, Any]:
    ignorable: set[str] = set()
    requires: set[str] = set()
    alternate = 0
    for root in roots.values():
        for element in root.iter():
            tag = element.tag
            if type(tag) is not str:
                continue
            if tag == f"{{{_MC_NS}}}AlternateContent":
                alternate += 1
            elif tag == f"{{{_MC_NS}}}Choice":
                requires.update((element.get("Requires") or "").split())
            value = element.get(f"{{{_MC_NS}}}Ignorable")
            if value:
                ignorable.update(value.split())
    return {
        "alternate_content_count": alternate,
        "ignorable": sorted(ignorable),
        "choice_requires": sorted(requires),
    }


def _relationship_graph(parts: dict[str, bytes], roots: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    graph: dict[str, dict[str, int]] = {}
    dangling: list[str] = []
    for name, root in roots.items():
        if not name.endswith(".rels"):
            continue
        for node in root.findall(f"{{{_REL_NS}}}Relationship"):
            rel_type = (node.get("Type") or "").rsplit("/", 1)[-1] or "(untyped)"
            entry = graph.setdefault(rel_type, {"count": 0, "external": 0})
            entry["count"] += 1
            if node.get("TargetMode") == "External":
                entry["external"] += 1
                continue
            resolved = resolve_relationship_target(name, node.get("Target") or "")
            if resolved and resolved not in parts:
                # Recorded by type and owner, never by rId: relationship ids are
                # renumbered freely by any writer and are not a fidelity signal.
                dangling.append(f"{name}->{rel_type}")
    return graph, sorted(dangling)


def _settings(roots: dict[str, Any]) -> dict[str, Any]:
    root = roots.get("word/settings.xml")
    if root is None:
        return {}
    compat_mode = None
    compat_settings = 0
    for node in root.iter(f"{{{_W_NS}}}compatSetting"):
        compat_settings += 1
        if node.get(f"{{{_W_NS}}}name") == "compatibilityMode":
            try:
                compat_mode = int(node.get(f"{{{_W_NS}}}val") or "")
            except ValueError:
                compat_mode = None
    return {
        "compat_setting_count": compat_settings,
        "compat_mode": compat_mode,
        "document_protection": root.find(f"{{{_W_NS}}}documentProtection") is not None,
        "track_changes": root.find(f"{{{_W_NS}}}trackChanges") is not None,
        "even_odd_headers": root.find(f"{{{_W_NS}}}evenAndOddHeaders") is not None,
        "mirror_margins": root.find(f"{{{_W_NS}}}mirrorMargins") is not None,
    }


def g10_projection(path: Path, *, limits: PackageLimits = DEFAULT_LIMITS) -> dict[str, Any]:
    """Everything about a package that survives normalization. Never raises."""
    scan = read_package(Path(path), limits=limits)
    if not scan.ok:
        return {"unreadable": scan.error or "package could not be read"}

    content_types = _content_type_map(scan.parts, scan.roots)

    part_kinds: dict[str, int] = {}
    opaque: dict[str, list[str]] = {}
    for name, data in scan.parts.items():
        if name.endswith("/"):
            continue
        content_type = content_types.get(name, "")
        # Part names are deliberately dropped here: only the kind and how many
        # of it there are. header3.xml standing in for header1.xml is not a
        # defect, and gating on it would make the whole gate unusable.
        key = content_type or f"(undeclared:{name.rsplit('.', 1)[-1].lower()})"
        part_kinds[key] = part_kinds.get(key, 0) + 1
        if _is_opaque(name, content_type):
            opaque.setdefault(key, []).append(sha256(data).hexdigest())

    graph, dangling = _relationship_graph(scan.parts, scan.roots)

    return {
        "namespaces": _namespaces(scan.roots),
        "unknown_namespaces": list(scan.unknown_namespaces),
        "markup_compatibility": _markup_compatibility(scan.roots),
        "part_kinds": dict(sorted(part_kinds.items())),
        "opaque_parts": {key: sorted(values) for key, values in sorted(opaque.items())},
        "relationship_graph": dict(sorted(graph.items())),
        "dangling_relationships": dangling,
        "settings": _settings(scan.roots),
        "malformed_parts": sorted(scan.malformed),
        "entity_parts": sorted(scan.entity_parts),
    }


def build_preservation_gate(
    source: Path,
    output: Path,
    *,
    limits: PackageLimits = DEFAULT_LIMITS,
) -> GateResult:
    """Compare two packages at the level the document model cannot see."""
    from word_replica.qa.policy import compare_projection

    expected = g10_projection(source, limits=limits)
    actual = g10_projection(output, limits=limits)
    findings = compare_projection("G10", expected, actual)

    if not findings:
        return GateResult(
            name="G10",
            passed=True,
            summary="package parts, relationships and namespaces preserved",
            details={"finding_count": 0, "part_kinds": len(expected.get("part_kinds", {}))},
        )

    first = findings[0]
    return GateResult(
        name="G10",
        passed=False,
        summary=f"package preservation: {len(findings)} mismatch(es) at {getattr(first, 'path', '/')}",
        details={"finding_count": len(findings)},
        first_divergence={
            "code": getattr(first, "code", "G10_MISMATCH"),
            "path": getattr(first, "path", "/"),
            "expected": getattr(first, "expected", None),
            "actual": getattr(first, "actual", None),
        },
    )
