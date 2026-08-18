from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from word_replica.domain.reconstruction import ReconstructionEvent


ALLOWED_STRUCTURE_EVENTS = {
    "BeginTable",
    "SetTableProperties",
    "SetColumnWidth",
    "SetRowProperties",
    "SetCellProperties",
    "EnterCell",
    "BeginParagraph",
    "ApplyParagraphProperties",
    "ApplyRunProperties",
    "InsertText",
    "EndParagraph",
    "LeaveCell",
    "EndTable",
}
FORBIDDEN_DELIMITERS = {"\t", "\r", "\x07"}


def _coordinates(payload: dict[str, Any]) -> tuple[int, int]:
    return int(payload["row"]), int(payload["column"])


def build_table_batch_event(
    events: Sequence[ReconstructionEvent],
) -> ReconstructionEvent | None:
    """Build one atomic text-only table event, or preserve the legacy path."""
    if len(events) < 2:
        return None
    if events[0].event_type != "BeginTable" or events[-1].event_type != "EndTable":
        return None
    if any(event.event_type not in ALLOWED_STRUCTURE_EVENTS for event in events):
        return None

    table_id = events[0].source_element_id
    if events[-1].source_element_id != table_id:
        return None

    try:
        rows = int(events[0].payload["rows"])
        columns = int(events[0].payload["columns"])
    except (KeyError, TypeError, ValueError):
        return None
    if rows <= 0 or columns <= 0:
        return None

    table_properties: dict[str, Any] | None = None
    column_widths: list[dict[str, Any]] = []
    row_properties: list[dict[str, Any]] = []
    cell_properties: dict[tuple[int, int], tuple[str, dict[str, Any]]] = {}
    index = 1

    try:
        while index < len(events) - 1 and events[index].event_type != "EnterCell":
            event = events[index]
            if event.event_type == "SetTableProperties":
                if table_properties is not None:
                    return None
                table_properties = dict(event.payload)
            elif event.event_type == "SetColumnWidth":
                column = int(event.payload["column"])
                if column < 1 or column > columns:
                    return None
                column_widths.append(dict(event.payload))
            elif event.event_type == "SetRowProperties":
                row = int(event.payload["row"])
                if row < 1 or row > rows:
                    return None
                row_properties.append(dict(event.payload))
            elif event.event_type == "SetCellProperties":
                key = _coordinates(event.payload)
                if key in cell_properties:
                    return None
                cell_properties[key] = (
                    event.source_element_id,
                    dict(event.payload.get("properties", {})),
                )
            else:
                return None
            index += 1
    except (KeyError, TypeError, ValueError):
        return None

    if table_properties is None:
        return None

    cells: list[dict[str, Any]] = []
    try:
        while index < len(events) - 1:
            enter = events[index]
            if enter.event_type != "EnterCell":
                return None
            row, column = _coordinates(enter.payload)
            key = (row, column)
            expected_cell = cell_properties.get(key)
            if expected_cell is None or expected_cell[0] != enter.source_element_id:
                return None
            index += 1

            begin = events[index]
            if begin.event_type != "BeginParagraph":
                return None
            paragraph_id = begin.source_element_id
            index += 1

            paragraph = events[index]
            if (
                paragraph.event_type != "ApplyParagraphProperties"
                or paragraph.source_element_id != paragraph_id
            ):
                return None
            paragraph_properties = dict(paragraph.payload)
            index += 1

            text_parts: list[str] = []
            runs: list[dict[str, Any]] = []
            while index < len(events) - 1 and events[index].event_type == "ApplyRunProperties":
                run_event = events[index]
                index += 1
                text_event = events[index]
                if (
                    text_event.event_type != "InsertText"
                    or text_event.source_element_id != run_event.source_element_id
                ):
                    return None
                text = str(text_event.payload["text"])
                if any(delimiter in text for delimiter in FORBIDDEN_DELIMITERS):
                    return None
                start = sum(len(part) for part in text_parts)
                text_parts.append(text)
                runs.append({
                    "source_element_id": run_event.source_element_id,
                    "start": start,
                    "end": start + len(text),
                    "properties": dict(run_event.payload),
                })
                index += 1
            if not runs:
                return None

            end = events[index]
            if end.event_type != "EndParagraph" or end.source_element_id != paragraph_id:
                return None
            index += 1
            leave = events[index]
            if (
                leave.event_type != "LeaveCell"
                or leave.source_element_id != enter.source_element_id
                or _coordinates(leave.payload) != key
            ):
                return None
            index += 1

            cells.append({
                "source_element_id": enter.source_element_id,
                "row": row,
                "column": column,
                "properties": expected_cell[1],
                "text": "".join(text_parts),
                "paragraph_properties": paragraph_properties,
                "runs": runs,
            })
    except (IndexError, KeyError, TypeError, ValueError):
        return None

    expected_coordinates = [
        (row, column)
        for row in range(1, rows + 1)
        for column in range(1, columns + 1)
    ]
    actual_coordinates = [(int(cell["row"]), int(cell["column"])) for cell in cells]
    if actual_coordinates != expected_coordinates:
        return None
    if set(cell_properties) != set(expected_coordinates):
        return None

    return ReconstructionEvent(
        "InsertTableBatch",
        table_id,
        {
            "rows": rows,
            "columns": columns,
            "table_properties": table_properties,
            "column_widths": column_widths,
            "row_properties": row_properties,
            "cells": cells,
            "paragraph_count": len(cells),
            "text_projection": "".join(str(cell["text"]) for cell in cells),
            "legacy_event_count": len(events),
        },
    )
