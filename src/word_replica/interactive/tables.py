from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from word_replica.domain.model import Table


@dataclass(frozen=True, slots=True)
class MergePlan:
    start_row: int
    start_column: int
    end_row: int
    end_column: int


@dataclass(frozen=True, slots=True)
class CellPlan:
    cell_id: str
    row: int
    column: int
    column_span: int
    properties: dict[str, Any]
    blocks: tuple[object, ...]
    continuation: bool = False


@dataclass(frozen=True, slots=True)
class TablePlan:
    table_id: str
    row_count: int
    column_count: int
    properties: dict[str, Any]
    row_properties: tuple[dict[str, Any], ...]
    cells: tuple[CellPlan, ...]
    merges: tuple[MergePlan, ...]


def build_table_plan(table: Table) -> TablePlan:
    cells: list[CellPlan] = []
    row_maps: list[dict[int, CellPlan]] = []
    max_columns = 0
    for r_idx, row in enumerate(table.rows, start=1):
        col = 1
        row_map: dict[int, CellPlan] = {}
        for cell in row.cells:
            span = max(1, int(cell.properties.get("grid_span", 1) or 1))
            continuation = cell.properties.get("v_merge") == "continue"
            cp = CellPlan(
                cell_id=cell.element_id,
                row=r_idx,
                column=col,
                column_span=span,
                properties=dict(cell.properties),
                blocks=tuple(cell.blocks),
                continuation=continuation,
            )
            cells.append(cp)
            for grid_col in range(col, col + span):
                row_map[grid_col] = cp
            col += span
        max_columns = max(max_columns, col - 1)
        row_maps.append(row_map)

    merges: list[MergePlan] = []
    covered_horizontal: set[str] = set()
    # Collapse a vertical-merge chain (including horizontal span) into one rectangle.
    for cp in cells:
        if cp.properties.get("v_merge") != "restart":
            continue
        end_row = cp.row
        for next_row in range(cp.row + 1, len(row_maps) + 1):
            candidate = row_maps[next_row - 1].get(cp.column)
            if candidate is None or candidate.properties.get("v_merge") != "continue" or candidate.column_span != cp.column_span:
                break
            end_row = next_row
            covered_horizontal.add(candidate.cell_id)
        if end_row > cp.row:
            merges.append(MergePlan(cp.row, cp.column, end_row, cp.column + cp.column_span - 1))
            covered_horizontal.add(cp.cell_id)

    for cp in cells:
        if cp.column_span > 1 and cp.cell_id not in covered_horizontal:
            merges.append(MergePlan(cp.row, cp.column, cp.row, cp.column + cp.column_span - 1))

    merges.sort(key=lambda m: (m.start_row, m.start_column, m.end_row, m.end_column))
    row_properties = tuple(dict(getattr(row, "properties", {}) or {}) for row in table.rows)
    return TablePlan(
        table_id=table.element_id,
        row_count=len(table.rows),
        column_count=max_columns,
        properties=dict(table.properties),
        row_properties=row_properties,
        cells=tuple(cells),
        merges=tuple(merges),
    )
