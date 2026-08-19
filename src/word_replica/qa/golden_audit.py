"""G0-G9 Golden fidelity audit: reusable, installable home for the gate logic
that used to live only under scripts/codex_automation. Behavior is
unchanged from the original extraction; scripts/codex_automation/audit.py
now re-exports these five public names for backward compatibility.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import asdict, dataclass, field
from typing import Any

DEFAULT_GATE_NAMES: tuple[str, ...] = tuple(f"G{i}" for i in range(10))

__all__ = [
    "DEFAULT_GATE_NAMES",
    "GateResult",
    "build_golden_report",
    "build_model_gates",
    "build_visual_gate",
    "audit_docx_pair",
    "compare_page_text_partitions",
]


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
        if "".join(left.split()) != "".join(right.split()):
            first = {"page": index, "expected": left[:500], "actual": right[:500]}
            break
    if first is None and not same_count:
        first = {"page": min(len(source_pages), len(output_pages)) + 1, "expected": "<page>", "actual": "<missing or extra page>"}
    passed = same_count and first is None
    return GateResult(
        name="G8",
        passed=passed,
        summary="pagination and page text partitions match" if passed else "pagination or page text partition mismatch",
        details={
            "source_page_count": len(source_pages),
            "output_page_count": len(output_pages),
            "comparison": "non_whitespace_character_partition",
        },
        first_divergence=first,
    )


def build_golden_report(*, run_id: str, source_sha256: str, commit_sha: str, reconstruction_status: str,
                        gates: dict[str, GateResult], source_page_count: int | None = None,
                        output_page_count: int | None = None,
                        gate_names: Sequence[str] = DEFAULT_GATE_NAMES, **extra: Any) -> dict[str, Any]:
    """Build the report contract. `gate_names` declares which gates a FULL PASS
    requires; it is echoed as `required_gates` so no consumer has to hardcode the
    count. The default is Golden's original G0-G9, so existing callers are
    unaffected; the fidelity lab opts in to G0-G10 by passing its own names.

    A declared gate with no result is treated as the first divergence rather
    than skipped -- a gate that never ran is not a gate that passed.
    """
    ordered = list(gate_names)
    first_name = next(
        (name for name in ordered if name not in gates or not gates[name].passed),
        None,
    )
    report = {
        "schema_version": 1,
        "run_id": run_id,
        "source_sha256": source_sha256,
        "commit_sha": commit_sha,
        "reconstruction_status": reconstruction_status,
        "required_gates": ordered,
        "gates": {name: gates[name].passed for name in ordered if name in gates},
        "gate_details": {name: asdict(gates[name]) for name in ordered if name in gates},
        "full_pass": bool(gates) and all(name in gates and gates[name].passed for name in ordered),
        "first_divergent_gate": first_name,
        "first_divergence": gates[first_name].first_divergence if first_name in gates else None,
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


def _table_projection(model, expected_model=None) -> list[dict]:
    tables = _walk_tables(model.body)
    expected_tables = _walk_tables(expected_model.body) if expected_model is not None else []
    projection: list[dict] = []
    for index, table in enumerate(tables):
        properties = _clean_table_properties(table.properties)
        if index < len(expected_tables):
            expected_properties = expected_tables[index].properties
            for derived_key in ("grid_column_widths", "layout"):
                if derived_key not in expected_properties:
                    properties.pop(derived_key, None)
        expected_table = expected_tables[index] if index < len(expected_tables) else None
        projection.append({
            "properties": properties,
            "rows": [
                {
                    "properties": dict(row.properties),
                    "cells": [
                        {
                            key: value
                            for key, value in cell.properties.items()
                            if not (
                                key in {"width", "width_type"}
                                and expected_table is not None
                                and row_index < len(expected_table.rows)
                                and cell_index < len(expected_table.rows[row_index].cells)
                                and key not in expected_table.rows[row_index].cells[cell_index].properties
                            )
                        }
                        for cell_index, cell in enumerate(row.cells)
                    ],
                }
                for row_index, row in enumerate(table.rows)
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
        "G3": _gate_from_projection(
            "G3", "table geometry and cell property fidelity",
            _table_projection(source_model), _table_projection(output_model, expected_model=source_model),
        ),
        "G4": _gate_from_projection("G4", "image asset and drawing geometry fidelity", _drawing_projection(source_model), _drawing_projection(output_model)),
        "G5": _gate_from_projection("G5", "page setup and section fidelity", l3_projection(source_model), l3_projection(output_model)),
        "G6": _gate_from_projection("G6", "header and footer fidelity", _header_footer_projection(source_model), _header_footer_projection(output_model)),
        "G7": _gate_from_projection("G7", "fields bookmarks and note fidelity", _semantic_projection(source_model), _semantic_projection(output_model)),
    }


def build_visual_gate(render_result, *, changed_pixel_tolerance: float, mae_tolerance: float) -> GateResult:
    from word_replica.qa.render import ANTIALIASING_BLUR_RADIUS

    antialiasing_ratio_allowance = max(changed_pixel_tolerance, 0.03)
    antialiasing_mae_allowance = max(mae_tolerance, 1.0)
    blurred_antialiasing_ratio_allowance = max(changed_pixel_tolerance, 0.04)
    blurred_antialiasing_mae_allowance = max(mae_tolerance, 1.0)

    def acceptance_mode(metric) -> str | None:
        if not metric.same_dimensions:
            return None
        strict = metric.changed_pixel_ratio <= changed_pixel_tolerance and metric.mean_absolute_error <= mae_tolerance
        if strict:
            return "strict"
        antialiasing = metric.changed_pixel_ratio <= antialiasing_ratio_allowance and metric.mean_absolute_error <= antialiasing_mae_allowance
        if antialiasing:
            return "legacy_antialiasing"
        blurred_mae = getattr(metric, "blurred_mean_absolute_error", None)
        blurred_antialiasing = (
            blurred_mae is not None
            and metric.changed_pixel_ratio <= blurred_antialiasing_ratio_allowance
            and blurred_mae <= blurred_antialiasing_mae_allowance
        )
        return "blurred_antialiasing" if blurred_antialiasing else None

    first = None
    acceptance_mode_counts = {
        "strict": 0,
        "legacy_antialiasing": 0,
        "blurred_antialiasing": 0,
        "failed": 0,
    }
    blur_assisted_pages = []
    for index, metric in enumerate(getattr(render_result, "metrics", []) or [], start=1):
        mode = acceptance_mode(metric)
        if mode is None:
            acceptance_mode_counts["failed"] += 1
        else:
            acceptance_mode_counts[mode] += 1
            if mode == "blurred_antialiasing":
                blur_assisted_pages.append(index)
        if mode is None and first is None:
            first = {
                "page": index,
                "same_dimensions": metric.same_dimensions,
                "changed_pixel_ratio": metric.changed_pixel_ratio,
                "mean_absolute_error": metric.mean_absolute_error,
                "blurred_mean_absolute_error": getattr(
                    metric, "blurred_mean_absolute_error", None
                ),
                "source_size": list(metric.source_size),
                "output_size": list(metric.rebuilt_size),
            }
    if first is None and not getattr(render_result, "page_count_match", False):
        first = {
            "page": min(getattr(render_result, "source_page_count", 0), getattr(render_result, "rebuilt_page_count", 0)) + 1,
            "reason": "page count mismatch",
        }
    passed = bool(getattr(render_result, "available", False)) and bool(getattr(render_result, "page_count_match", False)) and first is None
    return GateResult(
        name="G9",
        passed=passed,
        summary="visual fidelity within tolerance" if passed else "visual fidelity outside tolerance",
        details={
            "source_page_count": getattr(render_result, "source_page_count", None),
            "output_page_count": getattr(render_result, "rebuilt_page_count", None),
            "changed_pixel_tolerance": changed_pixel_tolerance,
            "mae_tolerance": mae_tolerance,
            "legacy_antialiasing_policy": {
                "changed_pixel_ratio_allowance": antialiasing_ratio_allowance,
                "mae_allowance": antialiasing_mae_allowance,
            },
            "blurred_antialiasing_policy": {
                "changed_pixel_ratio_allowance": blurred_antialiasing_ratio_allowance,
                "blurred_mae_allowance": blurred_antialiasing_mae_allowance,
                "gaussian_blur_radius": ANTIALIASING_BLUR_RADIUS,
            },
            "acceptance_mode_counts": acceptance_mode_counts,
            "blur_assisted_pages": blur_assisted_pages,
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
    pdf_pair_exporter=None,
    pdf_comparer=None,
    page_text_extractor=None,
    compatibility_mode_reader=None,
    visual_dpi: int = 144,
    changed_pixel_tolerance: float = 0.001,
    mae_tolerance: float = 0.25,
) -> dict[str, Any]:
    from pathlib import Path
    from word_replica.parser.parser import DocxParser
    from word_replica.qa.render import compare_pdfs
    from word_replica.qa.word_render import export_docx_pair_to_pdf_with_word, read_docx_pair_compatibility_mode

    qa_dir = Path(qa_dir)
    qa_dir.mkdir(parents=True, exist_ok=True)
    parser = parser or DocxParser()
    pdf_comparer = pdf_comparer or compare_pdfs
    page_text_extractor = page_text_extractor or extract_pdf_page_texts
    compatibility_mode_reader = compatibility_mode_reader or read_docx_pair_compatibility_mode

    source_model = parser.parse(Path(source_docx))
    output_model = parser.parse(Path(output_docx))
    gates = build_model_gates(source_model, output_model)

    source_pdf = qa_dir / "source.pdf"
    output_pdf = qa_dir / "output.pdf"
    source_page_texts: list[str] = []
    output_page_texts: list[str] = []
    try:
        compatibility_mode = compatibility_mode_reader(Path(source_docx), Path(output_docx))
    except Exception as exc:
        compatibility_mode = {"source": None, "output": None, "error": str(exc)}
    try:
        if pdf_exporter is not None:
            pdf_exporter(Path(source_docx), source_pdf, visible=False)
            pdf_exporter(Path(output_docx), output_pdf, visible=False)
        else:
            pair_exporter = pdf_pair_exporter or export_docx_pair_to_pdf_with_word
            pair_exporter(
                Path(source_docx),
                source_pdf,
                Path(output_docx),
                output_pdf,
                visible=False,
            )
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
        compatibility_mode=compatibility_mode,
    )
