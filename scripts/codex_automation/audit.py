from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(slots=True)
class GateResult:
    name: str
    passed: bool
    summary: str
    details: dict[str, Any] = field(default_factory=dict)
    first_divergence: dict[str, Any] | None = None


def compare_page_text_partitions(source_pages: list[str], output_pages: list[str]) -> GateResult:
    same_count = len(source_pages) == len(output_pages)
    first = None
    for index, (left, right) in enumerate(zip(source_pages, output_pages), start=1):
        if left != right:
            first = {"page": index, "expected": left[:500], "actual": right[:500]}
            break
    if first is None and not same_count:
        first = {"page": min(len(source_pages), len(output_pages)) + 1, "expected": "<page>", "actual": "<missing or extra page>"}
    passed = same_count and first is None
    return GateResult(
        name="G8",
        passed=passed,
        summary="pagination and page text partitions match" if passed else "pagination or page text partition mismatch",
        details={"source_page_count": len(source_pages), "output_page_count": len(output_pages)},
        first_divergence=first,
    )


def build_golden_report(*, run_id: str, source_sha256: str, commit_sha: str, reconstruction_status: str,
                        gates: dict[str, GateResult], source_page_count: int | None = None,
                        output_page_count: int | None = None, **extra: Any) -> dict[str, Any]:
    ordered = [f"G{i}" for i in range(10)]
    first_name = next((name for name in ordered if name in gates and not gates[name].passed), None)
    report = {
        "schema_version": 1,
        "run_id": run_id,
        "source_sha256": source_sha256,
        "commit_sha": commit_sha,
        "reconstruction_status": reconstruction_status,
        "gates": {name: gates[name].passed for name in ordered if name in gates},
        "gate_details": {name: asdict(gates[name]) for name in ordered if name in gates},
        "full_pass": bool(gates) and all(gates.get(name) is not None and gates[name].passed for name in ordered),
        "first_divergent_gate": first_name,
        "first_divergence": gates[first_name].first_divergence if first_name else None,
        "source_page_count": source_page_count,
        "output_page_count": output_page_count,
    }
    report.update(extra)
    return report


def _first_finding(finding) -> dict[str, Any]:
    return {
        "code": getattr(finding, "code", "MISMATCH"),
        "path": getattr(finding, "path", "/"),
        "expected": getattr(finding, "expected", None),
        "actual": getattr(finding, "actual", None),
    }


def _gate_from_projection(name: str, summary: str, expected: Any, actual: Any) -> GateResult:
    from word_replica.qa.policy import compare_projection

    findings = compare_projection(name, expected, actual)
    return GateResult(
        name=name,
        passed=not findings,
        summary=summary if not findings else f"{summary}: {len(findings)} mismatch(es)",
        details={"finding_count": len(findings)},
        first_divergence=_first_finding(findings[0]) if findings else None,
    )


def _walk_tables(blocks) -> list:
    from word_replica.domain.model import Table

    result = []
    for block in blocks:
        if isinstance(block, Table):
            result.append(block)
            for row in block.rows:
                for cell in row.cells:
                    result.extend(_walk_tables(cell.blocks))
    return result


def _clean_table_properties(properties: dict) -> dict:
    return {key: value for key, value in properties.items() if key != "borders_xml"}


def _table_projection(model) -> list[dict]:
    tables = _walk_tables(model.body)
    projection: list[dict] = []
    for table in tables:
        projection.append({
            "properties": _clean_table_properties(table.properties),
            "rows": [
                {
                    "properties": dict(row.properties),
                    "cells": [dict(cell.properties) for cell in row.cells],
                }
                for row in table.rows
            ],
        })
    return projection


def _drawing_projection(model) -> list[dict]:
    asset_hashes = {asset_id: asset.sha256 for asset_id, asset in model.assets.items()}
    keys = (
        "representation", "width_emu", "height_emu", "lock_aspect_ratio", "wrap_type",
        "horizontal_relative_from", "horizontal_position_emu", "vertical_relative_from",
        "vertical_position_emu", "distance_top_emu", "distance_bottom_emu", "distance_left_emu",
        "distance_right_emu", "crop", "rotation_degrees", "behind_text", "z_order",
    )
    return [
        {
            "asset_sha256": asset_hashes.get(drawing.asset_id),
            **{key: getattr(drawing, key) for key in keys},
        }
        for drawing in model.drawings
    ]


def _story_blocks_projection(model, blocks) -> list[dict]:
    from word_replica.domain.model import Paragraph, Table
    from word_replica.qa.formatting import normalize_formatting

    result: list[dict] = []
    for block in blocks:
        if isinstance(block, Paragraph):
            result.append({
                "kind": "Paragraph",
                "text": block.text(),
                "formatting": normalize_formatting(model, block),
            })
        elif isinstance(block, Table):
            result.append({
                "kind": "Table",
                "properties": _clean_table_properties(block.properties),
                "rows": [
                    [
                        {
                            "properties": dict(cell.properties),
                            "blocks": _story_blocks_projection(model, cell.blocks),
                        }
                        for cell in row.cells
                    ]
                    for row in block.rows
                ],
            })
    return result


def _header_footer_projection(model) -> dict:
    return {
        "headers": {
            str(key): _story_blocks_projection(model, blocks)
            for key, blocks in sorted(model.headers.items(), key=lambda item: str(item[0]))
        },
        "footers": {
            str(key): _story_blocks_projection(model, blocks)
            for key, blocks in sorted(model.footers.items(), key=lambda item: str(item[0]))
        },
    }


def _semantic_projection(model) -> dict:
    from word_replica.qa.content import _walk_blocks

    return {
        "fields": [
            {"instruction": field.instruction, "result_text": field.result_text, "locked": field.locked}
            for field in model.fields
        ],
        "bookmarks": [
            {"name": bookmark.name, "start_path": bookmark.start_path, "end_path": bookmark.end_path}
            for bookmark in model.bookmarks
        ],
        "footnotes": {
            str(key): list(_walk_blocks(blocks, f"footnote/{key}"))
            for key, blocks in sorted(model.footnotes.items(), key=lambda item: str(item[0]))
        },
        "endnotes": {
            str(key): list(_walk_blocks(blocks, f"endnote/{key}"))
            for key, blocks in sorted(model.endnotes.items(), key=lambda item: str(item[0]))
        },
    }


def build_model_gates(source_model, output_model) -> dict[str, GateResult]:
    from word_replica.qa.content import l0_projection
    from word_replica.qa.formatting import l2_projection
    from word_replica.qa.layout import l3_projection
    from word_replica.qa.structure import l1_projection

    return {
        "G0": _gate_from_projection("G0", "content fidelity", l0_projection(source_model), l0_projection(output_model)),
        "G1": _gate_from_projection("G1", "document structure fidelity", l1_projection(source_model), l1_projection(output_model)),
        "G2": _gate_from_projection("G2", "typography and paragraph formatting fidelity", l2_projection(source_model), l2_projection(output_model)),
        "G3": _gate_from_projection("G3", "table geometry and cell property fidelity", _table_projection(source_model), _table_projection(output_model)),
        "G4": _gate_from_projection("G4", "image asset and drawing geometry fidelity", _drawing_projection(source_model), _drawing_projection(output_model)),
        "G5": _gate_from_projection("G5", "page setup and section fidelity", l3_projection(source_model), l3_projection(output_model)),
        "G6": _gate_from_projection("G6", "header and footer fidelity", _header_footer_projection(source_model), _header_footer_projection(output_model)),
        "G7": _gate_from_projection("G7", "fields bookmarks and note fidelity", _semantic_projection(source_model), _semantic_projection(output_model)),
    }


def build_visual_gate(render_result, *, changed_pixel_tolerance: float, mae_tolerance: float) -> GateResult:
    first = None
    for index, metric in enumerate(getattr(render_result, "metrics", []) or [], start=1):
        if (
            not metric.same_dimensions
            or metric.changed_pixel_ratio > changed_pixel_tolerance
            or metric.mean_absolute_error > mae_tolerance
        ):
            first = {
                "page": index,
                "same_dimensions": metric.same_dimensions,
                "changed_pixel_ratio": metric.changed_pixel_ratio,
                "mean_absolute_error": metric.mean_absolute_error,
                "source_size": list(metric.source_size),
                "output_size": list(metric.rebuilt_size),
            }
            break
    if first is None and not getattr(render_result, "page_count_match", False):
        first = {
            "page": min(getattr(render_result, "source_page_count", 0), getattr(render_result, "rebuilt_page_count", 0)) + 1,
            "reason": "page count mismatch",
        }
    passed = bool(getattr(render_result, "available", False)) and bool(getattr(render_result, "within_tolerance", False))
    return GateResult(
        name="G9",
        passed=passed,
        summary="visual fidelity within tolerance" if passed else "visual fidelity outside tolerance",
        details={
            "source_page_count": getattr(render_result, "source_page_count", None),
            "output_page_count": getattr(render_result, "rebuilt_page_count", None),
            "changed_pixel_tolerance": changed_pixel_tolerance,
            "mae_tolerance": mae_tolerance,
        },
        first_divergence=first,
    )


def extract_pdf_page_texts(pdf_path) -> list[str]:
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(pdf_path))
    pages: list[str] = []
    try:
        for index in range(len(pdf)):
            page = pdf[index]
            text_page = page.get_textpage()
            try:
                text = text_page.get_text_range()
            finally:
                text_page.close()
            normalized = "\n".join(line.rstrip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")).strip()
            pages.append(normalized)
    finally:
        pdf.close()
    return pages


def audit_docx_pair(
    source_docx,
    output_docx,
    qa_dir,
    *,
    run_id: str,
    source_sha256: str,
    commit_sha: str,
    reconstruction_status: str,
    parser=None,
    pdf_exporter=None,
    pdf_comparer=None,
    page_text_extractor=None,
    visual_dpi: int = 144,
    changed_pixel_tolerance: float = 0.001,
    mae_tolerance: float = 0.25,
) -> dict[str, Any]:
    from pathlib import Path
    from word_replica.parser.parser import DocxParser
    from word_replica.qa.render import compare_pdfs
    from word_replica.qa.word_render import export_docx_to_pdf_with_word

    qa_dir = Path(qa_dir)
    qa_dir.mkdir(parents=True, exist_ok=True)
    parser = parser or DocxParser()
    pdf_exporter = pdf_exporter or export_docx_to_pdf_with_word
    pdf_comparer = pdf_comparer or compare_pdfs
    page_text_extractor = page_text_extractor or extract_pdf_page_texts

    source_model = parser.parse(Path(source_docx))
    output_model = parser.parse(Path(output_docx))
    gates = build_model_gates(source_model, output_model)

    source_pdf = qa_dir / "source.pdf"
    output_pdf = qa_dir / "output.pdf"
    source_page_texts: list[str] = []
    output_page_texts: list[str] = []
    try:
        pdf_exporter(Path(source_docx), source_pdf, visible=False)
        pdf_exporter(Path(output_docx), output_pdf, visible=False)
        source_page_texts = page_text_extractor(source_pdf)
        output_page_texts = page_text_extractor(output_pdf)
        gates["G8"] = compare_page_text_partitions(source_page_texts, output_page_texts)
        render_result = pdf_comparer(
            source_pdf,
            output_pdf,
            qa_dir,
            dpi=visual_dpi,
            changed_pixel_tolerance=changed_pixel_tolerance,
            mae_tolerance=mae_tolerance,
        )
        gates["G9"] = build_visual_gate(
            render_result,
            changed_pixel_tolerance=changed_pixel_tolerance,
            mae_tolerance=mae_tolerance,
        )
    except Exception as exc:
        message = f"PDF/visual audit unavailable: {exc}"
        gates["G8"] = GateResult("G8", False, message, first_divergence={"error": str(exc)})
        gates["G9"] = GateResult("G9", False, message, first_divergence={"error": str(exc)})

    return build_golden_report(
        run_id=run_id,
        source_sha256=source_sha256,
        commit_sha=commit_sha,
        reconstruction_status=reconstruction_status,
        gates=gates,
        source_page_count=len(source_page_texts) if source_page_texts else None,
        output_page_count=len(output_page_texts) if output_page_texts else None,
    )
