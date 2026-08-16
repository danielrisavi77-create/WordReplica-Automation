from pathlib import Path
import pytest

from word_replica.domain.reconstruction import ReconstructionBlueprint, ReconstructionEvent
from word_replica.renderers.interactive_word import InteractiveWordController


class FakeWordRange:
    def __init__(self):
        self.insert_after_calls = []
        self.collapse_calls = []
        self.break_calls = []
        self.paragraph_calls = 0

    def InsertAfter(self, value):
        self.insert_after_calls.append(value)

    def Collapse(self, direction):
        self.collapse_calls.append(direction)

    def InsertBreak(self, Type):
        self.break_calls.append(Type)

    def InsertParagraphAfter(self):
        self.paragraph_calls += 1


class FormattingSensitivePageBreakRange(FakeWordRange):
    """Model Word's observed formatting mutation for InsertBreak(page)."""

    def __init__(self):
        super().__init__()
        self.paragraph_alignment = "both"

    def InsertBreak(self, Type):
        super().InsertBreak(Type)
        if Type == 7:
            self.paragraph_alignment = "left"


def test_insert_character_calls_word_once_with_exactly_one_character():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "ž"}))
    assert fake.insert_after_calls == ["ž"]
    assert fake.collapse_calls == [0]


def test_insert_character_rejects_multi_character_payload():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    with pytest.raises(ValueError, match="exactly one"):
        controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "AB"}))


def test_insert_text_uses_one_word_call_for_a_contiguous_span():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("InsertText", "r1", {"text": "ABC"}))
    assert fake.insert_after_calls == ["ABC"]
    assert fake.collapse_calls == [0]


def test_semantic_tab_and_breaks_are_dedicated_atomic_calls():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("InsertTab", "r1", {}))
    controller.execute_event(ReconstructionEvent("InsertLineBreak", "r1", {}))
    controller.execute_event(ReconstructionEvent("InsertPageBreak", "r1", {}))
    assert fake.insert_after_calls == ["\t", "\f"]
    assert fake.break_calls == [6]


def test_tab_reapplies_active_run_formatting_after_word_resets_it():
    from types import SimpleNamespace

    class WordLikeTabRange(FakeWordRange):
        def __init__(self):
            super().__init__()
            self.Font = SimpleNamespace(Name=None, NameBi=None)

        def InsertAfter(self, value):
            super().InsertAfter(value)
            if value == "\t":
                self.Font.NameBi = "Times New Roman"

    fake = WordLikeTabRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r1", {
        "font_ascii": "Times New Roman",
        "font_cs": "Cambria",
    }))

    controller.execute_event(ReconstructionEvent("InsertTab", "r1", {}))

    assert fake.insert_after_calls == ["\t"]
    assert fake.Font.NameBi == "Cambria"


def test_page_break_reapplies_complex_script_font_after_word_resets_it():
    class WordLikePageBreakFont:
        def __init__(self):
            self.Name = None
            self.NameBi = None

    class WordLikePageBreakRange(FakeWordRange):
        def __init__(self):
            super().__init__()
            self.Font = WordLikePageBreakFont()

        def InsertAfter(self, value):
            super().InsertAfter(value)
            if value == "\f":
                self.Font.NameBi = "Times New Roman"

    fake = WordLikePageBreakRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r1", {
        "font_ascii": "Times New Roman",
        "font_cs": "Cambria",
    }))
    controller.execute_event(ReconstructionEvent("InsertPageBreak", "r1", {}))

    assert fake.Font.NameBi == "Cambria"


def test_interactive_renderer_has_no_forbidden_input_automation_path():
    source = Path("src/word_replica/renderers/interactive_word.py").read_text(encoding="utf-8")
    forbidden = ("Send" + "Keys", "win32" + "clipboard", "pyauto" + "gui", "keyboard" + ".write", "Selection" + ".TypeText")
    assert all(token not in source for token in forbidden)


class RecordingObject:
    def __init__(self, log, prefix):
        object.__setattr__(self, "_log", log)
        object.__setattr__(self, "_prefix", prefix)

    def __setattr__(self, name, value):
        if name.startswith("_"):
            object.__setattr__(self, name, value)
        else:
            self._log.append((f"{self._prefix}.{name}", value))
            object.__setattr__(self, name, value)


class ResettableRecordingObject(RecordingObject):
    def Reset(self):
        self._log.append((f"{self._prefix}.Reset", None))


class FormattingRange(FakeWordRange):
    def __init__(self):
        super().__init__()
        self.log = []
        self.Font = RecordingObject(self.log, "font")
        self.ParagraphFormat = RecordingObject(self.log, "paragraph")
        self.Style = None

    def InsertAfter(self, value):
        self.log.append(("insert", value))
        super().InsertAfter(value)


def test_repeated_run_properties_reuse_the_current_formatting_context():
    fake = FormattingRange()
    fake.Font = ResettableRecordingObject(fake.log, "font")
    controller = InteractiveWordController.for_testing(active_range=fake)
    properties = {"bold": True, "font_ascii": "Aptos", "size_half_points": "24"}
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r1", properties))
    controller.execute_event(ReconstructionEvent("InsertText", "r1", {"text": "A"}))
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r2", properties))

    assert [name for name, _ in fake.log if name == "font.Reset"] == ["font.Reset", "font.Reset"]
    assert [name for name, _ in fake.log if name == "font.Bold"] == ["font.Bold", "font.Bold", "font.Bold"]


def test_repeated_paragraph_style_definition_is_configured_once():
    from types import SimpleNamespace

    log = []

    class FakeStyle:
        def __init__(self, name):
            self.NameLocal = name
            self.Font = ResettableRecordingObject(log, f"style.{name}.font")
            self.ParagraphFormat = RecordingObject(log, f"style.{name}.paragraph")

    style = FakeStyle("ReplicaBody")

    class Styles:
        def __call__(self, key):
            if key != "ReplicaBody":
                raise RuntimeError("missing style")
            return style

    document = SimpleNamespace(Styles=Styles())
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = document
    definition = {
        "style_id": "ReplicaBody", "name": "ReplicaBody", "type": "paragraph",
        "run_properties": {"font_ascii": "Arial", "size_half_points": "22"},
        "paragraph_properties": {},
    }
    payload = {"style_id": "ReplicaBody", "style_definition": definition}
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p1", payload))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p2", payload))

    assert [name for name, _ in log if name == "style.ReplicaBody.font.Name"] == ["style.ReplicaBody.font.Name"]


def test_paragraph_style_definition_applies_complex_script_font():
    from types import SimpleNamespace

    log = []

    class FakeStyle:
        def __init__(self):
            self.Font = RecordingObject(log, "style.font")
            self.ParagraphFormat = RecordingObject(log, "style.paragraph")

    style = FakeStyle()

    class Styles:
        def __call__(self, key):
            if key != "Footer":
                raise RuntimeError("missing style")
            return style

    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = SimpleNamespace(Styles=Styles())
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "footer-p", {
        "style_id": "Footer",
        "style_definition": {
            "style_id": "Footer",
            "name": "footer",
            "type": "paragraph",
            "run_properties": {"font_ascii": "Times New Roman", "font_cs": "Cambria"},
            "paragraph_properties": {},
        },
    }))

    assert ("style.font.NameBi", "Cambria") in log


def test_repeated_paragraph_properties_are_not_reset_after_new_paragraph_inherits_them():
    fake = FormattingRange()
    fake.ParagraphFormat = ResettableRecordingObject(fake.log, "paragraph")
    controller = InteractiveWordController.for_testing(active_range=fake)
    properties = {"alignment": "center", "spacing_after": "120"}
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p1", properties))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p2", properties))

    assert [name for name, _ in fake.log if name == "paragraph.Reset"] == ["paragraph.Reset"]
    assert [name for name, _ in fake.log if name == "paragraph.Alignment"] == ["paragraph.Alignment"]


def test_run_formatting_is_applied_before_first_character():
    fake = FormattingRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r1", {
        "bold": True,
        "font_ascii": "Aptos",
        "size_half_points": "24",
        "italic": True,
        "color": "FF0000",
    }))
    controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "A"}))
    names = [item[0] for item in fake.log]
    assert names.index("font.Bold") < names.index("insert")
    assert names.index("font.Name") < names.index("insert")
    assert names.index("font.Size") < names.index("insert")
    assert fake.Font.Bold == -1
    assert fake.Font.Name == "Aptos"
    assert fake.Font.Size == 12.0


def test_inserted_text_receives_run_formatting_after_word_expands_the_range():
    class WordLikeFont:
        def __init__(self, owner):
            object.__setattr__(self, "owner", owner)
            object.__setattr__(self, "Bold", 0)

        def Reset(self):
            self.Bold = 0

        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == "Bold" and self.owner.inserted_text:
                self.owner.inserted_bold = value

    class WordLikeExpandedRange(FakeWordRange):
        def __init__(self):
            super().__init__()
            self.inserted_text = ""
            self.inserted_bold = 0
            self.Font = WordLikeFont(self)

        def InsertAfter(self, value):
            self.inserted_text += value
            super().InsertAfter(value)

    target = WordLikeExpandedRange()
    controller = InteractiveWordController.for_testing(active_range=target)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r1", {"bold": True}))
    controller.execute_event(ReconstructionEvent("InsertText", "r1", {"text": "Bold"}))

    assert target.inserted_bold == -1


def test_post_insert_run_repair_does_not_reset_the_whole_font():
    fake = FormattingRange()
    fake.Font = ResettableRecordingObject(fake.log, "font")
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r1", {
        "bold": True,
        "italic": False,
        "font_ascii": "Aptos",
        "size_half_points": "24",
    }))
    fake.log.clear()
    controller.execute_event(ReconstructionEvent("InsertText", "r1", {"text": "text"}))
    names = [name for name, _ in fake.log]
    assert "font.Reset" not in names
    assert "font.Bold" in names


def test_uncolored_run_clears_explicit_color_inherited_from_previous_run_after_insert():
    class WordLikeInheritedColorFont:
        def __init__(self, owner):
            object.__setattr__(self, "owner", owner)
            object.__setattr__(self, "Bold", 0)
            object.__setattr__(self, "Color", None)

        def Reset(self):
            if self.owner.range_expanded:
                self.Color = None

        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == "Color" and self.owner.range_expanded:
                self.owner.inserted_colors[-1] = value

    class WordLikeInheritedColorRange(FakeWordRange):
        def __init__(self):
            super().__init__()
            self.range_expanded = False
            self.inserted_colors = []
            self.Font = WordLikeInheritedColorFont(self)

        def InsertAfter(self, value):
            self.range_expanded = True
            self.inserted_colors.append(self.Font.Color)
            super().InsertAfter(value)

        def Collapse(self, direction):
            super().Collapse(direction)
            self.range_expanded = False

    target = WordLikeInheritedColorRange()
    controller = InteractiveWordController.for_testing(active_range=target)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "prefix", {
        "bold": True,
        "color": "1B2F4B",
    }))
    controller.execute_event(ReconstructionEvent("InsertText", "prefix", {"text": "Tablica 1."}))
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "suffix", {"bold": True}))
    controller.execute_event(ReconstructionEvent("InsertText", "suffix", {"text": " Zakonski okvir"}))

    assert target.inserted_colors == [0x4B2F1B, None]


def test_paragraph_formatting_is_applied_before_text():
    fake = FormattingRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p1", {
        "alignment": "center",
        "spacing_before": "120",
        "indent_left": "360",
        "keepNext": True,
        "pageBreakBefore": True,
    }))
    controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "A"}))
    names = [item[0] for item in fake.log]
    assert names.index("paragraph.Alignment") < names.index("insert")
    assert names.index("paragraph.SpaceBefore") < names.index("insert")
    assert names.index("paragraph.LeftIndent") < names.index("insert")
    assert fake.ParagraphFormat.Alignment == 1
    assert fake.ParagraphFormat.SpaceBefore == 6.0
    assert fake.ParagraphFormat.LeftIndent == 18.0
    assert fake.ParagraphFormat.KeepWithNext == -1
    assert fake.ParagraphFormat.PageBreakBefore == -1


class FakeCell:
    def __init__(self, row, col, log):
        self.row, self.col, self.log = row, col, log
        self.Range = FormattingRange()
        self.VerticalAlignment = None
        self.Width = None
        self.PreferredWidthType = None
        self.PreferredWidth = None
    def Merge(self, other):
        self.log.append(("merge", self.row, self.col, other.row, other.col))


class FakeTableRange:
    def __init__(self):
        self.End = 100
    @property
    def Duplicate(self):
        return FormattingRange()


class FakeRows:
    def __init__(self, count, log):
        self._rows = {i: RecordingObject(log, f"row{i}") for i in range(1, count+1)}
    def __call__(self, index): return self._rows[index]


class FakeColumns:
    def __init__(self, count, log):
        self._cols = {i: RecordingObject(log, f"column{i}") for i in range(1, count+1)}
    def __call__(self, index): return self._cols[index]


class FakeTable:
    def __init__(self, rows, cols, log):
        self.log = log
        self.Range = FakeTableRange()
        self._cells = {(r,c): FakeCell(r,c,log) for r in range(1,rows+1) for c in range(1,cols+1)}
        self.Rows = FakeRows(rows, log)
        self.Columns = FakeColumns(cols, log)
    def Cell(self, r, c):
        self.log.append(("lookup", r, c))
        return self._cells[(r,c)]


class FakeTables:
    def __init__(self, log): self.log=log
    def Add(self, Range, NumRows, NumColumns):
        self.log.append(("create_table", NumRows, NumColumns))
        return FakeTable(NumRows, NumColumns, self.log)


class FakeDocument:
    def __init__(self, log): self.Tables = FakeTables(log)


class NativeBatchRootRange:
    def __init__(self, start):
        self.Start = start
        self.End = start
        self.insert_after_calls = []

    def InsertAfter(self, value):
        self.insert_after_calls.append(value)
        self.End += len(value)


class NativeBatchInsertedRange:
    def __init__(self, start, end, table):
        self.Start = start
        self.End = end
        self.table = table
        self.convert_calls = []

    def ConvertToTable(self, **kwargs):
        self.convert_calls.append(kwargs)
        return self.table


class NativeBatchFormattingRange(FormattingRange):
    def __init__(self, start, end):
        super().__init__()
        self.Start = start
        self.End = end


class NativeBatchCellRange(NativeBatchFormattingRange):
    def __init__(self, start, end):
        super().__init__(start, end)
        self.duplicates = []

    @property
    def Duplicate(self):
        duplicate = NativeBatchFormattingRange(self.Start, self.End)
        self.duplicates.append(duplicate)
        return duplicate


class NativeBatchTableRange:
    def __init__(self, after_range):
        self.after_range = after_range
        self.duplicates = []
        self._returned_after_range = False

    @property
    def Duplicate(self):
        if not self._returned_after_range:
            self._returned_after_range = True
            return self.after_range
        duplicate = NativeBatchFormattingRange(0, 100)
        self.duplicates.append(duplicate)
        return duplicate


class NativeBatchTable(FakeTable):
    def __init__(self, rows, columns, log):
        super().__init__(rows, columns, log)
        for index, cell in enumerate(self._cells.values()):
            cell.Range = NativeBatchCellRange(start=index * 2, end=index * 2 + 2)
        self.after_range = FormattingRange()
        self.Range = NativeBatchTableRange(self.after_range)


class NativeBatchDocument(FakeDocument):
    def __init__(self, log, table):
        super().__init__(log)
        self.table = table
        self.range_calls = []
        self.inserted_range = None
        self.format_ranges = []

    def Range(self, start, end):
        self.range_calls.append((start, end))
        if self.inserted_range is None:
            self.inserted_range = NativeBatchInsertedRange(start, end, self.table)
            return self.inserted_range
        target = NativeBatchFormattingRange(start, end)
        self.format_ranges.append(target)
        return target


def _native_batch_payload():
    cells = []
    for row, texts in enumerate((("A", "B"), ("C", "D")), start=1):
        for column, value in enumerate(texts, start=1):
            cells.append({
                "source_element_id": f"c{row}{column}",
                "row": row,
                "column": column,
                "properties": {},
                "text": value,
                "paragraph_properties": {},
                "runs": [{
                    "source_element_id": f"r{row}{column}",
                    "start": 0,
                    "end": 1,
                    "properties": {},
                }],
            })
    return {
        "rows": 2,
        "columns": 2,
        "table_properties": {},
        "column_widths": [],
        "row_properties": [],
        "cells": cells,
    }


def test_native_table_batch_inserts_once_and_converts_without_tables_add():
    log = []
    root = NativeBatchRootRange(start=7)
    table = NativeBatchTable(2, 2, log)
    document = NativeBatchDocument(log, table)
    controller = InteractiveWordController.for_testing(active_range=root)
    controller.document = document

    controller.execute_event(ReconstructionEvent("InsertTableBatch", "t", _native_batch_payload()))

    assert root.insert_after_calls == ["A\tB\rC\tD"]
    assert document.range_calls[0] == (7, 14)
    assert document.inserted_range.convert_calls == [
        {"Separator": 1, "NumRows": 2, "NumColumns": 2}
    ]
    assert all(record[0] != "create_table" for record in log)
    assert controller.active_range is table.after_range


def test_native_table_batch_formats_exact_cell_and_run_ranges():
    log = []
    root = NativeBatchRootRange(start=20)
    table = NativeBatchTable(1, 1, log)
    cell_range = NativeBatchCellRange(start=20, end=31)
    table._cells[(1, 1)].Range = cell_range
    document = NativeBatchDocument(log, table)
    controller = InteractiveWordController.for_testing(active_range=root)
    controller.document = document
    payload = {
        "rows": 1,
        "columns": 1,
        "table_properties": {},
        "column_widths": [],
        "row_properties": [],
        "cells": [{
            "source_element_id": "c1",
            "row": 1,
            "column": 1,
            "properties": {},
            "text": "Alpha Beta",
            "paragraph_properties": {"alignment": "left"},
            "runs": [
                {"source_element_id": "r1", "start": 0, "end": 6, "properties": {"bold": False}},
                {"source_element_id": "r2", "start": 6, "end": 10, "properties": {"bold": True}},
            ],
        }],
    }

    controller.execute_event(ReconstructionEvent("InsertTableBatch", "t", payload))

    assert cell_range.duplicates[0].Start == 20
    assert cell_range.duplicates[0].End == 30
    assert cell_range.duplicates[0].ParagraphFormat.Alignment == 0
    assert document.range_calls == [(20, 30), (20, 26), (26, 30)]
    assert [target.Font.Bold for target in document.format_ranges] == [0, -1]


def test_native_table_batch_reports_and_consumes_phase_metrics():
    log = []
    root = NativeBatchRootRange(start=7)
    table = NativeBatchTable(2, 2, log)
    document = NativeBatchDocument(log, table)
    controller = InteractiveWordController.for_testing(active_range=root)
    controller.document = document

    controller.execute_event(ReconstructionEvent("InsertTableBatch", "t", _native_batch_payload()))
    metrics = controller.consume_table_batch_metrics()

    assert metrics["table_id"] == "t"
    assert metrics["row_count"] == 2
    assert metrics["cell_count"] == 4
    assert metrics["run_count"] == 4
    assert metrics["failed_phase"] is None
    assert metrics["success"] is True
    for name in (
        "insert_seconds", "convert_seconds", "geometry_seconds",
        "formatting_seconds", "verification_seconds", "total_seconds",
    ):
        assert metrics[name] >= 0.0
    assert controller.consume_table_batch_metrics() is None


def test_native_table_batch_applies_common_paragraph_and_run_format_once():
    log = []
    root = NativeBatchRootRange(start=7)
    table = NativeBatchTable(2, 2, log)
    document = NativeBatchDocument(log, table)
    controller = InteractiveWordController.for_testing(active_range=root)
    controller.document = document
    payload = _native_batch_payload()
    for cell in payload["cells"]:
        cell["paragraph_properties"] = {"alignment": "left"}
        cell["runs"][0]["properties"] = {"bold": False}

    controller.execute_event(ReconstructionEvent("InsertTableBatch", "t", payload))

    assert document.range_calls == [(7, 14)]
    assert len(table.Range.duplicates) == 1
    common_range = table.Range.duplicates[0]
    assert ("paragraph.Alignment", 0) in common_range.log
    assert ("font.Bold", 0) in common_range.log


def test_native_table_batch_applies_common_run_after_per_cell_paragraph_styles():
    class OrderingController(InteractiveWordController):
        def __init__(self, *, active_range):
            super().__init__()
            self.active_range = active_range
            self.formatting_order = []

        def _event_ApplyParagraphProperties(self, event):
            self.formatting_order.append(("paragraph", event.source_element_id))
            super()._event_ApplyParagraphProperties(event)

        def _event_ApplyRunProperties(self, event):
            self.formatting_order.append(("run", event.source_element_id))
            super()._event_ApplyRunProperties(event)

    log = []
    root = NativeBatchRootRange(start=7)
    table = NativeBatchTable(2, 2, log)
    document = NativeBatchDocument(log, table)
    controller = OrderingController(active_range=root)
    controller.document = document
    payload = _native_batch_payload()
    for index, cell in enumerate(payload["cells"]):
        cell["paragraph_properties"] = {
            "alignment": "left" if index % 2 == 0 else "right",
        }
        cell["runs"][0]["properties"] = {
            "font_ascii": "Times New Roman",
            "font_cs": "Cambria",
            "size_half_points": "20",
        }

    controller.execute_event(ReconstructionEvent("InsertTableBatch", "t", payload))

    baseline_index = controller.formatting_order.index(("run", "t"))
    paragraph_indices = [
        index
        for index, (kind, _) in enumerate(controller.formatting_order)
        if kind == "paragraph"
    ]
    assert paragraph_indices
    assert max(paragraph_indices) < baseline_index


def test_native_table_batch_records_failed_convert_phase_before_reraising():
    class ConversionFailureRange:
        def ConvertToTable(self, **kwargs):
            raise RuntimeError("convert failed")

    class ConversionFailureDocument(FakeDocument):
        def Range(self, start, end):
            return ConversionFailureRange()

    controller = InteractiveWordController.for_testing(active_range=NativeBatchRootRange(start=3))
    controller.document = ConversionFailureDocument([])

    with pytest.raises(RuntimeError, match="convert failed"):
        controller.execute_event(ReconstructionEvent("InsertTableBatch", "t", _native_batch_payload()))

    metrics = controller.consume_table_batch_metrics()
    assert metrics["success"] is False
    assert metrics["failed_phase"] == "convert"
    assert metrics["total_seconds"] >= 0.0


def test_table_structure_and_merges_happen_before_cell_typing_without_postmerge_lookup():
    log = []
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = FakeDocument(log)
    events = [
        ReconstructionEvent("BeginTable", "t", {"rows": 1, "columns": 2}),
        ReconstructionEvent("SetColumnWidth", "t", {"column": 1, "width_twips": 1440}),
        ReconstructionEvent("MergeCells", "t", {"start_row":1,"start_column":1,"end_row":1,"end_column":2}),
        ReconstructionEvent("EnterCell", "c", {"row":1,"column":1}),
        ReconstructionEvent("ApplyRunProperties", "r", {"bold": True}),
        ReconstructionEvent("InsertCharacter", "r", {"character":"A"}),
        ReconstructionEvent("InsertCharacter", "r", {"character":"1"}),
    ]
    for event in events: controller.execute_event(event)
    assert log[0] == ("create_table", 1, 2)
    merge_index = log.index(("merge",1,1,1,2))
    # Both cell lookups are captured before merge; entering the cell does not call Table.Cell again.
    assert log.index(("lookup",1,1)) < merge_index
    assert log.index(("lookup",1,2)) < merge_index
    assert [x for x in log[merge_index+1:] if x and x[0] == "lookup"] == []
    assert controller.active_range.insert_after_calls == ["A", "1"]


def test_cell_dxa_width_is_preferred_width_and_does_not_move_table_grid():
    log = []
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = FakeDocument(log)
    controller.execute_event(ReconstructionEvent("BeginTable", "t", {"rows": 1, "columns": 2}))
    controller.execute_event(ReconstructionEvent(
        "SetCellProperties", "c2",
        {"row": 1, "column": 2, "properties": {"width": "6236", "width_type": "dxa"}},
    ))

    cell = controller._table_stack[-1]["cells"][(1, 2)]
    assert cell.PreferredWidthType == 3
    assert cell.PreferredWidth == 311.8
    assert cell.Width is None


class FakeInlineImage:
    def __init__(self, log):
        self.log = log
        self.Width = None; self.Height = None; self.LockAspectRatio = None
        self.Range = FormattingRange()
        object.__setattr__(self.Range.Font, "NameBi", "Times New Roman")
    def ConvertToShape(self):
        self.log.append(("convert_to_shape",))
        return FakeShape(self.log)


class FakeWrap:
    def __init__(self, log): object.__setattr__(self,"log",log)
    def __setattr__(self,name,value):
        if name == "log": object.__setattr__(self,name,value)
        else: self.log.append((f"wrap.{name}",value)); object.__setattr__(self,name,value)


class FakePictureFormat:
    def __init__(self, log): object.__setattr__(self,"log",log)
    def __setattr__(self,name,value):
        if name == "log": object.__setattr__(self,name,value)
        else: self.log.append((f"crop.{name}",value)); object.__setattr__(self,name,value)


class FakeShape:
    def __init__(self, log):
        object.__setattr__(self,"log",log)
        object.__setattr__(self,"WrapFormat",FakeWrap(log))
        object.__setattr__(self,"PictureFormat",FakePictureFormat(log))
    def __setattr__(self,name,value):
        if name in {"log","WrapFormat","PictureFormat"}: object.__setattr__(self,name,value)
        else: self.log.append((f"shape.{name}",value)); object.__setattr__(self,name,value)
    def ZOrder(self, command): self.log.append(("zorder",command))


class FakeInlineShapes:
    def __init__(self, log):
        self.log = log
        self.last_inline = None
    def AddPicture(self, **kwargs):
        self.log.append(("add_picture", kwargs["FileName"]))
        self.last_inline = FakeInlineImage(self.log)
        return self.last_inline


class FakeImageDocument(FakeDocument):
    def __init__(self, log):
        super().__init__(log)
        self.InlineShapes = FakeInlineShapes(log)


def test_insert_image_preserves_active_complex_script_font_on_image_range():
    log = []
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = FakeImageDocument(log)
    controller.set_asset_resolver(lambda asset_id: Path("C:/tmp/image.png"))

    controller.execute_event(ReconstructionEvent(
        "ApplyRunProperties",
        "d1",
        {"font_ascii": "Times New Roman", "font_cs": "Cambria"},
    ))
    controller.execute_event(ReconstructionEvent(
        "InsertImage",
        "d1",
        {"asset_id": "a1", "representation": "inline", "source_path": "word/media/image1.png"},
    ))

    assert controller.document.InlineShapes.last_inline.Range.Font.NameBi == "Cambria"


def test_image_is_created_before_stepwise_geometry_mutations():
    log=[]
    controller=InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document=FakeImageDocument(log)
    controller.set_asset_resolver(lambda asset_id: Path("C:/tmp/image.png"))
    events=[
        ReconstructionEvent("InsertImage","d1",{"asset_id":"a1","representation":"floating","source_path":"word/media/image1.png"}),
        ReconstructionEvent("SetImageSize","d1",{"width_emu":914400,"height_emu":457200,"lock_aspect_ratio":True}),
        ReconstructionEvent("SetImageWrap","d1",{"wrap_type":"square","behind_text":False,"distance_top_emu":12700,"distance_bottom_emu":0,"distance_left_emu":0,"distance_right_emu":0}),
        ReconstructionEvent("SetImagePosition","d1",{"horizontal_relative_from":"column","horizontal_position_emu":12700,"vertical_relative_from":"paragraph","vertical_position_emu":25400}),
        ReconstructionEvent("SetImageRotation","d1",{"rotation_degrees":15.0}),
    ]
    for event in events: controller.execute_event(event)
    assert log[0][0] == "add_picture"
    assert Path(log[0][1]) == Path("C:/tmp/image.png")
    assert log[1] == ("convert_to_shape",)
    assert log.index(("shape.Width",72.0)) > 1
    assert log.index(("wrap.Type",0)) > 1
    assert log.index(("shape.Left",1.0)) > 1
    assert log.index(("shape.Rotation",15.0)) > 1


def test_image_dimensions_are_set_while_aspect_ratio_is_temporarily_unlocked():
    log = []
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = FakeImageDocument(log)
    controller.set_asset_resolver(lambda asset_id: Path("C:/tmp/image.png"))
    controller.execute_event(ReconstructionEvent(
        "InsertImage",
        "d1",
        {"asset_id": "a1", "representation": "floating", "source_path": "word/media/image1.png"},
    ))
    log.clear()

    controller.execute_event(ReconstructionEvent(
        "SetImageSize",
        "d1",
        {"width_emu": 914400, "height_emu": 457200, "lock_aspect_ratio": True},
    ))

    assert log == [
        ("shape.LockAspectRatio", 0),
        ("shape.Width", 72.0),
        ("shape.Height", 36.0),
        ("shape.LockAspectRatio", -1),
    ]


def test_image_dimensions_preserve_existing_aspect_lock_when_source_lock_is_unspecified():
    log = []
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = FakeImageDocument(log)
    controller.set_asset_resolver(lambda asset_id: Path("C:/tmp/image.png"))
    controller.execute_event(ReconstructionEvent(
        "InsertImage",
        "d1",
        {"asset_id": "a1", "representation": "floating", "source_path": "word/media/image1.png"},
    ))
    controller._active_image.LockAspectRatio = -1
    log.clear()

    controller.execute_event(ReconstructionEvent(
        "SetImageSize",
        "d1",
        {"width_emu": 914400, "height_emu": 457200, "lock_aspect_ratio": None},
    ))

    assert log == [
        ("shape.LockAspectRatio", 0),
        ("shape.Width", 72.0),
        ("shape.Height", 36.0),
        ("shape.LockAspectRatio", -1),
    ]


@pytest.mark.parametrize("failure_on_assignment", [1, 2])
def test_image_size_propagates_required_aspect_lock_failures(failure_on_assignment):
    class FailingLockShape:
        def __init__(self):
            object.__setattr__(self, "lock_assignments", 0)

        def __setattr__(self, name, value):
            if name == "LockAspectRatio":
                assignment = self.lock_assignments + 1
                object.__setattr__(self, "lock_assignments", assignment)
                if assignment == failure_on_assignment:
                    raise RuntimeError("lock assignment failed")
            object.__setattr__(self, name, value)

    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller._active_image = FailingLockShape()

    with pytest.raises(RuntimeError, match="lock assignment failed"):
        controller.execute_event(ReconstructionEvent(
            "SetImageSize",
            "d1",
            {"width_emu": 914400, "height_emu": 457200, "lock_aspect_ratio": True},
        ))


def test_renderer_can_resume_from_first_uncompleted_event():
    from word_replica.config import InteractiveOptions
    from word_replica.interactive.control import InteractiveRunControl
    from word_replica.renderers.interactive_word import InteractiveWordRenderer

    class RecordingController:
        def __init__(self): self.events=[]
        def execute_event(self,event): self.events.append(event.payload.get("character", event.event_type))
    controller=RecordingController()
    renderer=InteractiveWordRenderer(controller=controller, options=InteractiveOptions())
    # Avoid intentional timing in this unit test.
    renderer.speed._sleep=lambda seconds: None
    bp=ReconstructionBlueprint.build(source_sha256="a"*64, source_model_fingerprint="m", events=(
        ReconstructionEvent("InsertCharacter","r",{"character":"A"}),
        ReconstructionEvent("InsertCharacter","r",{"character":"B"}),
        ReconstructionEvent("InsertCharacter","r",{"character":"C"}),
    ))
    control=InteractiveRunControl(); control.start()
    outcome=renderer.execute_blueprint(bp, control, start_index=2)
    assert controller.events == ["C"]
    assert outcome.status == "COMPLETED"
    assert outcome.last_completed_index == 2


def test_restore_checkpoint_state_recreates_exact_body_range():
    from types import SimpleNamespace
    calls=[]
    class Doc:
        def Range(self,start,end):
            calls.append((start,end))
            return FakeCheckpointRange(start,end)
    class FakeCheckpointRange(FakeWordRange):
        def __init__(self,start,end):
            super().__init__(); self.Start=start; self.End=end
    controller=InteractiveWordController()
    controller.document=Doc()
    checkpoint=SimpleNamespace(story="body", range_start=42, range_end=42, paragraph_started=True)
    controller.restore_checkpoint_state(checkpoint)
    assert calls == [(42,42)]
    assert controller.active_range.Start == 42
    assert controller._paragraph_started is True


def test_restore_checkpoint_state_rehydrates_header_story_and_body_return_range():
    from types import SimpleNamespace

    class Range:
        def __init__(self,start,end,label): self.Start=start; self.End=end; self.label=label; self.Tables=Collection([])
        @property
        def Duplicate(self): return Range(self.Start,self.End,self.label)
        def SetRange(self,start,end): self.Start=start; self.End=end
    class Collection:
        def __init__(self,items): self.items=items
        def __call__(self,index): return self.items[index-1]
        @property
        def Count(self): return len(self.items)
    class Story:
        def __init__(self,label): self.Range=Range(0,100,label)
    class Section:
        def __init__(self):
            self.Headers=Collection([Story("header:default"),Story("header:first"),Story("header:even")])
            self.Footers=Collection([Story("footer:default"),Story("footer:first"),Story("footer:even")])
    class Doc:
        def __init__(self): self.Sections=Collection([Section()])
        def Range(self,start,end): return Range(start,end,"body")

    checkpoint=SimpleNamespace(
        story="header:default", range_start=7, range_end=7, paragraph_started=True,
        table_element_id=None, cell_element_id=None, section_element_id="s0",
        resume_state={"section_index":0,"section_started":True,"return_story":"body",
                      "return_range_start":42,"return_range_end":42,"return_paragraph_started":True},
    )
    controller=InteractiveWordController(); controller.document=Doc()
    controller.restore_checkpoint_state(checkpoint)
    assert controller._active_story == "header:default"
    assert controller.active_range.label == "header:default"
    assert (controller.active_range.Start,controller.active_range.End) == (7,7)
    assert controller._active_section_index == 0
    assert controller._section_started is True
    assert controller._story_stack
    return_range, paragraph_started, return_story = controller._story_stack[-1]
    assert return_range.label == "body" and (return_range.Start,return_range.End)==(42,42)
    assert paragraph_started is True and return_story == "body"


def test_restore_checkpoint_state_rehydrates_active_table_for_next_cell():
    from types import SimpleNamespace

    class Collection:
        def __init__(self,items): self.items=items
        def __call__(self,index): return self.items[index-1]
        @property
        def Count(self): return len(self.items)
    class Range:
        def __init__(self,start,end,label="body",tables=None):
            self.Start=start; self.End=end; self.label=label; self.Tables=Collection(tables or [])
        @property
        def Duplicate(self): return Range(self.Start,self.End,self.label,self.Tables.items)
        def Collapse(self,where): self.Start=self.End
    class Cell:
        def __init__(self): self.Range=Range(20,30,"cell")
    class Table:
        def __init__(self): self.Range=Range(10,40,"table",[self]); self.calls=[]
        def Cell(self,row,col): self.calls.append((row,col)); return Cell()
    table=Table()
    class Sections:
        def __call__(self,index): return object()
    class Doc:
        def Range(self,start,end): return Range(start,end,"body",[table])
    Doc.Sections = Sections()
    checkpoint=SimpleNamespace(
        story="body", range_start=22, range_end=22, paragraph_started=True,
        table_element_id="t1", cell_element_id="c1",
        resume_state={"section_index":0,"section_started":True,"table_structure_complete":True},
    )
    controller=InteractiveWordController(); controller.document=Doc()
    controller.restore_checkpoint_state(checkpoint)
    assert controller._table_stack and controller._table_stack[-1]["table"] is table
    controller.execute_event(ReconstructionEvent("EnterCell","c2",{"row":1,"column":2}))
    assert table.calls == [(1,2)]
    assert controller._active_cell_element_id == "c2"


def test_custom_paragraph_style_is_created_from_blueprint_before_assignment():
    from types import SimpleNamespace

    log = []

    class FakeFont:
        def __setattr__(self, name, value):
            if name != "_ready": log.append((f"style.Font.{name}", value))
            object.__setattr__(self, name, value)

    class FakeParagraphFormat:
        pass

    class FakeStyle:
        def __init__(self, name):
            self.NameLocal = name
            self.Font = FakeFont()
            self.ParagraphFormat = FakeParagraphFormat()

    class FakeStyles:
        def __init__(self): self.items = {"Normal": FakeStyle("Normal")}
        def __call__(self, key):
            if key not in self.items: raise RuntimeError("missing style")
            return self.items[key]
        def Add(self, Name, Type):
            log.append(("Styles.Add", Name, Type))
            style = FakeStyle(Name); self.items[Name] = style; return style

    fake = FormattingRange()
    styles = FakeStyles()
    document = SimpleNamespace(Styles=styles)
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.document = document
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p1", {
        "style_id": "ReplicaBody",
        "style_definition": {
            "style_id": "ReplicaBody", "name": "ReplicaBody", "type": "paragraph",
            "run_properties": {"font_ascii": "Arial", "size_half_points": "22"},
            "paragraph_properties": {},
        },
    }))
    assert ("Styles.Add", "ReplicaBody", 1) in log
    assert styles.items["ReplicaBody"].Font.Name == "Arial"
    assert styles.items["ReplicaBody"].Font.Size == 11.0
    assert fake.Style is styles.items["ReplicaBody"]


def test_controller_applies_descriptive_metadata_without_touching_lifecycle_fields():
    from types import SimpleNamespace

    class Prop:
        def __init__(self): self.Value = None
    class BuiltIns:
        def __init__(self): self.props = {name: Prop() for name in ("Title","Subject","Keywords","Category","Author")}
        def __call__(self, name): return self.props[name]
    class Custom:
        def __init__(self): self.values = {}
        def __call__(self, name): raise RuntimeError("missing")
        def Add(self, Name, LinkToContent, Type, Value): self.values[Name] = Value
    built = BuiltIns(); custom = Custom()
    controller = InteractiveWordController.for_testing(active_range=SimpleNamespace())
    controller.document = SimpleNamespace(BuiltInDocumentProperties=built, CustomDocumentProperties=custom)
    controller.apply_metadata({"title":"T","subject":"S","creator":"A","custom:StudyId":"42"})
    assert built.props["Title"].Value == "T"
    assert built.props["Subject"].Value == "S"
    assert built.props["Author"].Value == "A"
    assert custom.values["StudyId"] == "42"


def test_custom_document_property_add_uses_word_compatible_positional_arguments():
    from types import SimpleNamespace

    class Custom:
        def __init__(self):
            self.calls = []

        def __call__(self, name):
            raise RuntimeError("missing")

        def Add(self, *args):
            self.calls.append(args)

    custom = Custom()
    controller = InteractiveWordController.for_testing(active_range=SimpleNamespace())
    controller.document = SimpleNamespace(CustomDocumentProperties=custom)

    controller.set_custom_property("WordReplicaActualSaveCount", 1)

    assert custom.calls == [("WordReplicaActualSaveCount", False, 4, "1")]


def test_create_field_moves_active_range_past_field_end_marker():
    from types import SimpleNamespace

    class WordLikeFieldResult:
        def __init__(self, start, end):
            self.Start = start
            self.End = end
            self.Text = ""
            self.collapse_calls = []
            self.move_calls = []

        @property
        def Duplicate(self):
            return self

        def Collapse(self, direction):
            self.collapse_calls.append(direction)
            self.Start = self.End

        def Move(self, unit, count):
            self.move_calls.append((unit, count))
            self.Start += count
            self.End += count
            return count

    class Fields:
        def __init__(self, result):
            self.result = result

        def Add(self, Range, Type, Text, PreserveFormatting):
            return SimpleNamespace(Result=self.result)

    result = WordLikeFieldResult(start=10, end=14)
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = SimpleNamespace(Fields=Fields(result))

    controller.execute_event(ReconstructionEvent("CreateField", "field", {
        "instruction": "REF ref_tab_1 \\h",
        "cached_result": "1",
    }))

    assert result.collapse_calls == [0]
    assert result.move_calls == [(1, 1)]
    assert (controller.active_range.Start, controller.active_range.End) == (15, 15)


def test_create_field_seeds_cached_result_and_reapplies_active_run_formatting_without_refresh():
    from types import SimpleNamespace

    class ResultRange:
        def __init__(self):
            self.Text = ""; self.Start=4; self.End=4; self.collapse=[]; self.move=[]
            self.Font = SimpleNamespace(Bold=-1, Name=None, NameBi="Times New Roman")
        @property
        def Duplicate(self): return self
        def Collapse(self, direction): self.collapse.append(direction)
        def Move(self, unit, count): self.move.append((unit, count)); return count
    class Field:
        def __init__(self): self.Result=ResultRange(); self.updated=0
        def Update(self): self.updated += 1
    class Fields:
        def __init__(self): self.created=None
        def Add(self, Range, Type, Text, PreserveFormatting):
            self.created=Field(); self.args=(Range,Type,Text,PreserveFormatting); return self.created
    fields=Fields(); document=SimpleNamespace(Fields=fields)
    active=FormattingRange(); active.Start=4; active.End=4
    controller=InteractiveWordController.for_testing(active_range=active); controller.document=document
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "f1", {
        "bold": False,
        "font_ascii": "Times New Roman",
        "font_cs": "Cambria",
    }))
    controller.execute_event(ReconstructionEvent("CreateField", "f1", {
        "instruction": 'TOC \\o "1-3"', "cached_result": "Chapter One .... 1"
    }))
    assert fields.created.Result.Text == "Chapter One .... 1"
    assert fields.created.updated == 0
    assert fields.created.Result.Font.Bold == 0
    assert fields.created.Result.Font.NameBi == "Cambria"
    assert controller.active_range is fields.created.Result

class CallableComRange(FakeWordRange):
    """Mimics pywin32 COM dispatch: Range.Duplicate is itself callable via a default property."""
    def __init__(self):
        super().__init__()
        self._oleobj_ = object()
        self.Start = 7
        self.End = 7

    @property
    def Duplicate(self):
        return self

    def __call__(self):
        return "default-property-text"


def test_duplicate_range_does_not_invoke_callable_com_dispatch_default_property():
    source = CallableComRange()
    duplicate = InteractiveWordController._duplicate_range(source)
    assert duplicate is source
    assert not isinstance(duplicate, str)


class RejectedCall(Exception):
    hresult = -2147418111


class RetryCollapseRange(FakeWordRange):
    def __init__(self):
        super().__init__()
        self._collapse_attempts = 0

    def Collapse(self, direction):
        self._collapse_attempts += 1
        if self._collapse_attempts == 1:
            raise RejectedCall("Call was rejected by callee")
        super().Collapse(direction)


def test_insert_character_does_not_duplicate_character_when_only_collapse_is_temporarily_rejected():
    fake = RetryCollapseRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "A"}))
    assert fake.insert_after_calls == ["A"]
    assert fake.collapse_calls == [0]
    assert fake._collapse_attempts == 2


class RangeWithoutCollapse:
    """Word can leave a usable story range whose Collapse member is unavailable."""

    def __init__(self, position=156534):
        self.Start = position
        self.End = position
        self.insert_after_calls = []

    def InsertAfter(self, value):
        self.insert_after_calls.append(value)
        self.End += len(value)


def test_insert_text_collapses_by_coordinates_when_word_range_loses_collapse_member():
    target = RangeWithoutCollapse()
    controller = InteractiveWordController.for_testing(active_range=target)

    controller.execute_event(ReconstructionEvent("InsertText", "run", {"text": "4"}))

    assert target.insert_after_calls == ["4"]
    assert target.Start == target.End == 156535

class RejectOncePropertyRange(FormattingRange):
    def __init__(self):
        super().__init__()
        self._start_reads = 0
        self._end_reads = 0
        self._start = 7
        self._end = 7
    @property
    def Start(self):
        self._start_reads += 1
        if self._start_reads == 1:
            raise RejectedCall()
        return self._start
    @Start.setter
    def Start(self, value): self._start = value
    @property
    def End(self):
        self._end_reads += 1
        if self._end_reads == 1:
            raise RejectedCall()
        return self._end
    @End.setter
    def End(self, value): self._end = value


def test_current_state_snapshot_retries_temporarily_rejected_range_coordinates():
    target = RejectOncePropertyRange()
    controller = InteractiveWordController.for_testing(active_range=target)
    snapshot = controller.current_state_snapshot()
    assert snapshot.range_start == 7
    assert snapshot.range_end == 7
    assert target._start_reads == 2
    assert target._end_reads == 2


class RejectInsertOnceRange(FormattingRange):
    def __init__(self):
        super().__init__()
        self.attempts = 0
    def InsertAfter(self, value):
        self.attempts += 1
        if self.attempts == 1:
            raise RejectedCall()
        super().InsertAfter(value)


def test_insert_character_retries_rejected_insert_without_duplicate_character():
    target = RejectInsertOnceRange()
    controller = InteractiveWordController.for_testing(active_range=target)
    controller.execute_event(ReconstructionEvent("InsertCharacter", "r", {"character": "Ž"}))
    assert target.attempts == 2
    assert target.insert_after_calls == ["Ž"]


class RejectParagraphOnceRange(FormattingRange):
    def __init__(self):
        super().__init__()
        self.paragraph_attempts = 0
    def InsertParagraphAfter(self):
        self.paragraph_attempts += 1
        if self.paragraph_attempts == 1:
            raise RejectedCall()
        return super().InsertParagraphAfter()


def test_begin_paragraph_retries_rejected_word_call_without_double_paragraph():
    target = RejectParagraphOnceRange()
    controller = InteractiveWordController.for_testing(active_range=target)
    controller._paragraph_started = True
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p", {}))
    assert target.paragraph_attempts == 2


class RejectFontGetterOnceRange(FormattingRange):
    def __init__(self):
        super().__init__()
        self._font_value = self.Font
        self.font_reads = 0
    @property
    def Font(self):
        self.font_reads += 1
        if self.font_reads == 1:
            raise RejectedCall()
        return self._font_value
    @Font.setter
    def Font(self, value): self._font_value = value


def test_run_formatting_retries_rejected_font_getter_before_setting_format():
    # Build manually because the subclass property intentionally rejects its first COM-style read.
    target = object.__new__(RejectFontGetterOnceRange)
    FormattingRange.__init__(target)
    target._font_value = target.__dict__.get("Font", None) or RecordingObject([], "font")
    target.font_reads = 0
    controller = InteractiveWordController.for_testing(active_range=target)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r", {"bold": True}))
    assert target.font_reads == 2

class RejectingFont:
    def __init__(self):
        object.__setattr__(self, "bold_attempts", 0)
        object.__setattr__(self, "Bold", None)
    def __setattr__(self, name, value):
        if name == "Bold":
            attempts = object.__getattribute__(self, "bold_attempts") + 1
            object.__setattr__(self, "bold_attempts", attempts)
            if attempts == 1:
                raise RejectedCall()
        object.__setattr__(self, name, value)


class RangeWithRejectingFont(FormattingRange):
    def __init__(self):
        super().__init__()
        self.Font = RejectingFont()


def test_run_formatting_retries_rejected_property_setter():
    target = RangeWithRejectingFont()
    controller = InteractiveWordController.for_testing(active_range=target)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r", {"bold": True}))
    assert target.Font.bold_attempts == 2
    assert target.Font.Bold == -1


class RejectingPageSetup:
    def __init__(self):
        object.__setattr__(self, "orientation_attempts", 0)
        object.__setattr__(self, "Orientation", None)
        object.__setattr__(self, "TextColumns", RecordingObject([], "columns"))
    def __setattr__(self, name, value):
        if name == "Orientation":
            attempts = object.__getattribute__(self, "orientation_attempts") + 1
            object.__setattr__(self, "orientation_attempts", attempts)
            if attempts == 1:
                raise RejectedCall()
        object.__setattr__(self, name, value)


def test_section_properties_retry_rejected_orientation_setter():
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    setup = RejectingPageSetup()
    controller._active_section = type("S", (), {"PageSetup": setup})()
    controller.execute_event(ReconstructionEvent("ApplySectionProperties", "s", {"orientation": "landscape"}))
    assert setup.orientation_attempts == 2
    assert setup.Orientation == 1


def test_auto_line_spacing_uses_multiple_rule_without_word_application_round_trip():
    class ApplicationThatMustNotBeCalled:
        def LinesToPoints(self, lines):
            raise AssertionError("line conversion must be local, not a Word COM call")

    fake = FormattingRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.application = ApplicationThatMustNotBeCalled()
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p1", {
        "spacing_before": "0", "spacing_after": "0", "spacing_line": "300", "spacing_line_rule": "auto",
    }))
    assert fake.ParagraphFormat.SpaceBefore == 0.0
    assert fake.ParagraphFormat.SpaceAfter == 0.0
    assert fake.ParagraphFormat.LineSpacingRule == 5
    assert fake.ParagraphFormat.LineSpacing == 15.0


def test_empty_paragraph_formatting_expands_zero_length_range():
    class RangeWithExpand(FormattingRange):
        def __init__(self):
            super().__init__()
            self.Start = 10
            self.End = 10

        def Expand(self, unit):
            self.expanded_unit = unit
            return 1

    class ExpandedRange(FormattingRange):
        def __init__(self):
            super().__init__()
            self.Start = 10
            self.End = 10
            self.expanded = RangeWithExpand()

        def Duplicate(self):
            return self.expanded

    fake = ExpandedRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p1", {
        "spacing_after": "0", "spacing_line": "360", "spacing_line_rule": "auto",
    }))

    assert not hasattr(fake.ParagraphFormat, "SpaceAfter")
    assert fake.expanded.ParagraphFormat.SpaceAfter == 0.0
    assert fake.expanded.End == 11


def test_end_table_reuses_words_existing_post_table_paragraph():
    after = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=after)
    controller._table_stack = [{"table": object(), "after_range": after, "element_id": "t1"}]
    controller._active_table_element_id = "t1"
    controller._paragraph_started = False
    controller.execute_event(ReconstructionEvent("EndTable", "t1", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))
    assert after.paragraph_calls == 0


def test_page_break_after_table_stays_in_words_existing_post_table_paragraph():
    after = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=after)
    controller._table_stack = [{"table": object(), "after_range": after, "element_id": "t1"}]

    controller.execute_event(ReconstructionEvent("EndTable", "t1", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "page-break", {}))
    controller.execute_event(ReconstructionEvent("InsertPageBreak", "run", {}))
    controller.execute_event(ReconstructionEvent("EndParagraph", "page-break", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "following", {}))

    assert after.insert_after_calls == ["\f"]
    assert after.break_calls == []
    assert after.paragraph_calls == 1


def test_word2010_style_add_uses_positional_arguments():
    from types import SimpleNamespace
    class Style:
        def __init__(self):
            self.Font = RecordingObject([], "font")
            self.ParagraphFormat = RecordingObject([], "paragraph")
    class Styles:
        def __init__(self): self.items = {}
        def __call__(self, key):
            if key not in self.items: raise RuntimeError("missing")
            return self.items[key]
        def Add(self, *args, **kwargs):
            if kwargs:
                raise TypeError("Word 2010 wrapper rejects keyword arguments")
            assert args == ("ReplicaBody", 1)
            self.items["ReplicaBody"] = Style()
            return self.items["ReplicaBody"]
    fake = FormattingRange(); styles = Styles()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.document = SimpleNamespace(Styles=styles)
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p1", {
        "style_id": "ReplicaBody",
        "style_definition": {"style_id": "ReplicaBody", "name": "ReplicaBody", "type": "paragraph", "run_properties": {"font_ascii": "Arial"}, "paragraph_properties": {}},
    }))
    assert fake.Style is styles.items["ReplicaBody"]


def test_source_style_application_retries_word2010_rejected_style_font_setter():
    from types import SimpleNamespace
    class RejectingSizeFont:
        def __init__(self):
            object.__setattr__(self, "attempts", 0)
        def __setattr__(self, name, value):
            if name == "Size":
                attempts = object.__getattribute__(self, "attempts") + 1
                object.__setattr__(self, "attempts", attempts)
                if attempts == 1:
                    raise RejectedCall("busy")
            object.__setattr__(self, name, value)
    class Style:
        def __init__(self):
            self.Font = RejectingSizeFont()
            self.ParagraphFormat = RecordingObject([], "paragraph")
    style = Style()
    class Styles:
        def __call__(self, key): return style
    fake = FormattingRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.document = SimpleNamespace(Styles=Styles())
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p", {
        "style_id": "Heading1",
        "style_definition": {"style_id":"Heading1","name":"heading 1","type":"paragraph","run_properties":{"size_half_points":"28"},"paragraph_properties":{}},
    }))
    assert style.Font.attempts == 2
    assert style.Font.Size == 14.0


def test_default_word2010_retry_budget_survives_more_than_one_second_busy_burst(monkeypatch):
    from word_replica.renderers import interactive_word as module
    monkeypatch.setattr(module.time, "sleep", lambda _: None)
    attempts = {"count": 0}
    def operation():
        attempts["count"] += 1
        if attempts["count"] <= 61:
            raise RejectedCall("Word 2010 busy")
        return "ok"
    assert module._retry_rejected_com_call(operation) == "ok"
    assert attempts["count"] == 62


def test_visible_word_is_shown_only_after_blank_document_is_ready(monkeypatch):
    import sys
    from types import SimpleNamespace

    calls = []

    class FakeDocument:
        def Range(self, start, end):
            calls.append(("Range", start, end))
            return FormattingRange()

        def Activate(self):
            calls.append("document.Activate")

    class FakeDocuments:
        def Add(self):
            calls.append("Documents.Add")
            return FakeDocument()

    class FakeApplication:
        def __init__(self):
            object.__setattr__(self, "Documents", FakeDocuments())
            object.__setattr__(self, "Hwnd", 1)

        def __setattr__(self, name, value):
            if name in {"DisplayAlerts", "Visible"}:
                calls.append((name, value))
            object.__setattr__(self, name, value)

        def Activate(self):
            calls.append("application.Activate")

    application = FakeApplication()
    monkeypatch.setitem(sys.modules, "pythoncom", SimpleNamespace(CoInitialize=lambda: calls.append("CoInitialize")))
    monkeypatch.setitem(
        sys.modules,
        "win32com.client",
        SimpleNamespace(DispatchEx=lambda name: application),
    )
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=sys.modules["win32com.client"]))
    monkeypatch.setattr("word_replica.renderers.interactive_word.word_process_pids", lambda: set())
    monkeypatch.setattr("word_replica.renderers.interactive_word.record_owned_word", lambda *a, **k: 123)

    controller = InteractiveWordController(visible=True)
    controller.open_blank()

    assert calls.index("Documents.Add") < calls.index(("Visible", True))
    assert calls.index(("DisplayAlerts", 0)) < calls.index("Documents.Add")
    assert calls[-2:] == ["document.Activate", "application.Activate"]


class RejectOnceSections:
    def __init__(self):
        self.attempts = 0
        self.section = object()

    def __call__(self, index):
        self.attempts += 1
        if self.attempts == 1:
            raise RejectedCall("Word is busy")
        assert index == 1
        return self.section


def test_begin_first_section_retries_rejected_sections_collection_call():
    from types import SimpleNamespace

    sections = RejectOnceSections()
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller.document = SimpleNamespace(Sections=sections)
    controller.execute_event(ReconstructionEvent("BeginSection", "s1", {"section_index": 0}))
    assert sections.attempts == 2
    assert controller._active_section is sections.section

class ResettableRecordingObject(RecordingObject):
    def Reset(self):
        self._log.append((f"{self._prefix}.Reset", None))


class ResettableFormattingRange(FakeWordRange):
    def __init__(self):
        super().__init__()
        self.log = []
        self.Font = ResettableRecordingObject(self.log, "font")
        self.ParagraphFormat = ResettableRecordingObject(self.log, "paragraph")
        self.Style = None


def test_paragraph_properties_assign_style_then_reset_direct_formatting_before_effective_values():
    fake = ResettableFormattingRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "p2", {
        "style_id": "Normal",
        "alignment": "both",
        "spacing_after": "120",
        "spacing_line": "360",
        "spacing_line_rule": "auto",
    }))
    names = [name for name, _ in fake.log]
    assert fake.Style == "Normal"
    assert "paragraph.Reset" in names
    assert names.index("paragraph.Reset") < names.index("paragraph.Alignment")
    assert names.index("paragraph.Reset") < names.index("paragraph.SpaceAfter")


def test_run_properties_reset_direct_font_formatting_before_effective_values():
    fake = ResettableFormattingRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("ApplyRunProperties", "r2", {
        "bold": False,
        "italic": False,
        "underline": False,
        "strike": False,
        "font_ascii": "Times New Roman",
        "size_half_points": "22",
        "color": "000000",
        "vert_align": "baseline",
        "character_spacing": "0",
        "character_position": "0",
    }))
    names = [name for name, _ in fake.log]
    assert "font.Reset" in names
    assert names.index("font.Reset") < names.index("font.Bold")
    assert names.index("font.Reset") < names.index("font.Name")


def test_enter_cell_then_normal_paragraph_resets_heading_formatting_context():
    cell_range = ResettableFormattingRange()
    cell = type("Cell", (), {"Range": cell_range})()
    controller = InteractiveWordController.for_testing(active_range=ResettableFormattingRange())
    controller._table_stack = [{"table": object(), "cells": {(1, 1): cell}, "after_range": ResettableFormattingRange(), "element_id": "t", "structure_complete": True}]
    controller._active_table_element_id = "t"
    controller.execute_event(ReconstructionEvent("EnterCell", "c", {"row": 1, "column": 1}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "cp", {}))
    controller.execute_event(ReconstructionEvent("ApplyParagraphProperties", "cp", {
        "style_id": "Normal", "alignment": "both", "spacing_after": "0", "spacing_line": "240", "spacing_line_rule": "auto",
    }))
    assert controller.active_range.Style == "Normal"
    assert ("paragraph.Reset", None) in controller.active_range.log
    assert controller.active_range.ParagraphFormat.Alignment == 3


def test_page_break_at_end_of_source_paragraph_keeps_next_source_paragraph_separate():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("InsertPageBreak", "r1", {}))
    controller.execute_event(ReconstructionEvent("EndParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))
    assert fake.insert_after_calls == ["\f"]
    assert fake.break_calls == []
    assert fake.paragraph_calls == 1


def test_standalone_page_break_preserves_source_paragraph_formatting():
    fake = FormattingSensitivePageBreakRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("InsertPageBreak", "r1", {}))
    controller.execute_event(ReconstructionEvent("EndParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))

    assert fake.paragraph_alignment == "both"
    assert fake.paragraph_calls == 1


def test_content_after_page_break_clears_continuation_reuse_for_next_paragraph():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("InsertPageBreak", "r1", {}))
    controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "X"}))
    controller.execute_event(ReconstructionEvent("EndParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))
    assert fake.paragraph_calls == 1


def test_section_page_number_start_restarts_primary_footer_numbering():
    class PageNumbers:
        RestartNumberingAtSection = False
        StartingNumber = 0
    class HeaderFooter:
        def __init__(self): self.PageNumbers = PageNumbers()
    class Collection:
        def __init__(self): self.primary = HeaderFooter()
        def __call__(self, index):
            assert index == 1
            return self.primary
    class PageSetup:
        def __init__(self): self.TextColumns = RecordingObject([], "columns")
    section = type("Section", (), {})()
    section.PageSetup = PageSetup()
    section.Footers = Collection()
    section.Headers = Collection()
    controller = InteractiveWordController.for_testing(active_range=FormattingRange())
    controller._active_section = section
    controller.execute_event(ReconstructionEvent("ApplySectionProperties", "s2", {"page_number_start": "1"}))
    nums = section.Footers.primary.PageNumbers
    assert nums.RestartNumberingAtSection == -1
    assert nums.StartingNumber == 1
