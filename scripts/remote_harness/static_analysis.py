from __future__ import annotations

from hashlib import sha256
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from lxml import etree

from word_replica.config import InteractiveOptions
from word_replica.domain.enums import InteractiveFidelity
from word_replica.domain.model import Paragraph, Table
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.interactive.preflight import analyze_preflight
from word_replica.parser.parser import DocxParser

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _sha256(path: Path) -> str:
    h = sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _content_type_coverage(zf: ZipFile) -> list[str]:
    root = etree.fromstring(zf.read("[Content_Types].xml"))
    defaults = {n.get("Extension", "").lower() for n in root.findall(f"{{{_CT_NS}}}Default")}
    overrides = {n.get("PartName", "").lstrip("/") for n in root.findall(f"{{{_CT_NS}}}Override")}
    missing: list[str] = []
    for name in zf.namelist():
        if name.endswith("/") or name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        ext = PurePosixPath(name).suffix.lstrip(".").lower()
        if name not in overrides and ext not in defaults:
            missing.append(name)
    return sorted(missing)


def _relationship_target_missing(zf: ZipFile) -> list[str]:
    names = set(zf.namelist())
    missing: list[str] = []
    for rel_name in sorted(n for n in names if n.endswith(".rels")):
        root = etree.fromstring(zf.read(rel_name))
        rel_dir = PurePosixPath(rel_name).parent
        if rel_dir.name == "_rels":
            owner_dir = rel_dir.parent
        else:
            owner_dir = rel_dir
        for node in root.findall(f"{{{_REL_NS}}}Relationship"):
            if node.get("TargetMode") == "External":
                continue
            target = node.get("Target") or ""
            if target.startswith("/"):
                normalized = target.lstrip("/")
            else:
                normalized = str(PurePosixPath(owner_dir, target))
            parts: list[str] = []
            for part in PurePosixPath(normalized).parts:
                if part == "..":
                    if parts:
                        parts.pop()
                elif part not in (".", ""):
                    parts.append(part)
            normalized = "/".join(parts)
            if normalized not in names:
                missing.append(f"{rel_name}:{node.get('Id')}->{normalized}")
    return sorted(missing)


def _walk_blocks(blocks):
    for block in blocks:
        yield block
        if isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    yield from _walk_blocks(cell.blocks)


def _model_counts(model) -> dict:
    paragraphs = runs = tables = merged = 0
    for block in _walk_blocks(model.body):
        if isinstance(block, Paragraph):
            paragraphs += 1
            runs += len(block.runs)
        elif isinstance(block, Table):
            tables += 1
            for row in block.rows:
                for cell in row.cells:
                    props = cell.properties
                    if int(props.get("grid_span", 1) or 1) > 1 or props.get("v_merge") in {"restart", "continue"}:
                        merged += 1
    return {
        "paragraphs": paragraphs,
        "runs": runs,
        "tables": tables,
        "merged_cells": merged,
        "images": len(model.drawings),
        "assets": len(model.assets),
        "sections": len(model.sections),
        "headers": sum(len(v) for v in model.headers.values()),
        "footers": sum(len(v) for v in model.footers.values()),
        "footnotes": len(model.footnotes),
        "endnotes": len(model.endnotes),
        "fields": len(model.fields),
        "bookmarks": len(model.bookmarks),
        "comments": len(model.comments),
        "revisions": len(model.revisions),
        "complex_parts": len(model.preserved_parts),
    }


def analyze_document(source: Path) -> dict:
    source = Path(source).resolve()
    report = {
        "source": {"filename": source.name, "sha256": _sha256(source), "size": source.stat().st_size},
        "package": {},
        "parser": {"ok": False},
    }
    try:
        with ZipFile(source) as zf:
            bad = zf.testzip()
            report["package"] = {
                "zip_integrity": bad is None,
                "bad_member": bad,
                "missing_content_types": _content_type_coverage(zf),
                "missing_relationship_targets": _relationship_target_missing(zf),
                "parts": len(zf.namelist()),
            }
    except Exception as exc:
        report["package"] = {"zip_integrity": False, "error": repr(exc)}
        return report

    try:
        model = DocxParser().parse(source)
        report["parser"] = {"ok": True, "fingerprint": model.fingerprint()}
        report["model"] = _model_counts(model)
        blueprint = BlueprintCompiler().compile(model)
        report["blueprint"] = {
            "schema_version": blueprint.schema_version,
            "fingerprint": blueprint.fingerprint,
            "total_events": blueprint.total_events,
            "total_visible_characters": blueprint.total_visible_characters,
            "semantic_counts": blueprint.semantic_counts,
        }
        options = InteractiveOptions(fidelity=InteractiveFidelity.MAXIMUM)
        preflight = analyze_preflight(model, blueprint, options, word_probe=lambda: True)
        report["preflight"] = {
            "maximum_fidelity_ready": preflight.maximum_fidelity_ready,
            "can_proceed": preflight.can_proceed,
            "blocking_reasons": list(preflight.blocking_reasons),
            "warnings": list(preflight.warnings),
            "capabilities": [
                {
                    "classification": item.classification.value,
                    "reason": item.reason,
                    "source_element_id": item.source_element_id,
                }
                for item in preflight.capability_items
            ],
        }
    except Exception as exc:
        report["parser"] = {"ok": False, "error": repr(exc)}
    return report
