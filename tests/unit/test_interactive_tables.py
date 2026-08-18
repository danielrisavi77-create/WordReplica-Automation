from word_replica.domain.model import Paragraph, Run, Table, TableCell, TableRow
from word_replica.interactive.tables import build_table_plan
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.domain.model import DocumentModel


def cell(cid, text, props=None, blocks=None):
    return TableCell(cid, blocks if blocks is not None else [Paragraph(f"p_{cid}", [Run(f"r_{cid}", text)])], props or {})


def test_rectangular_table_plan_is_deterministic():
    table = Table("t", [
        TableRow("r0", [cell("c00", "00"), cell("c01", "01"), cell("c02", "02")]),
        TableRow("r1", [cell("c10", "10"), cell("c11", "11"), cell("c12", "12")]),
    ])
    plan = build_table_plan(table)
    assert (plan.row_count, plan.column_count) == (2, 3)
    assert [(c.row, c.column, c.cell_id) for c in plan.cells] == [
        (1,1,"c00"),(1,2,"c01"),(1,3,"c02"),(2,1,"c10"),(2,2,"c11"),(2,3,"c12")
    ]
    assert plan.merges == ()


def test_horizontal_and_vertical_merges_are_stable_premerge_coordinates():
    table = Table("t", [
        TableRow("r0", [cell("a", "A", {"grid_span": 2, "v_merge": "restart"}), cell("b", "B")]),
        TableRow("r1", [cell("a2", "", {"grid_span": 2, "v_merge": "continue"}), cell("b2", "C")]),
    ])
    plan = build_table_plan(table)
    assert plan.column_count == 3
    assert [(m.start_row,m.start_column,m.end_row,m.end_column) for m in plan.merges] == [(1,1,2,2)]
    assert [c.cell_id for c in plan.cells if not c.continuation] == ["a", "b", "b2"]


def test_nested_table_is_retained_in_logical_cell_blocks():
    nested = Table("nested", [TableRow("nr", [cell("nc", "N")])])
    outer_cell = cell("outer", "", blocks=[nested])
    plan = build_table_plan(Table("outerT", [TableRow("or", [outer_cell])]))
    assert plan.cells[0].blocks[0] is nested


def test_blueprint_builds_structure_and_merges_before_cell_characters():
    table = Table("t", [TableRow("r", [cell("c", "A1", {"grid_span": 2})])])
    model = DocumentModel(source_sha256="a"*64, body=[table])
    events = BlueprintCompiler().compile(model).events
    types = [e.event_type for e in events]
    first_char = types.index("InsertText")
    assert types.index("BeginTable") < types.index("SetTableProperties") < types.index("MergeCells") < types.index("EnterCell") < first_char
    chars = [e.payload["text"] for e in events if e.event_type == "InsertText"]
    assert chars == ["A1"]


def test_word_table_property_executor_applies_modeled_width_margins_shading_and_borders():
    from types import SimpleNamespace
    from word_replica.domain.reconstruction import ReconstructionEvent
    from word_replica.renderers.interactive_word import InteractiveWordController

    class Border:
        def __init__(self): self.LineStyle=None; self.LineWidth=None; self.Color=None
    class Borders:
        def __init__(self): self.items={i:Border() for i in (-1,-2,-3,-4,-5,-6)}
        def __call__(self,index): return self.items[index]
    class Shading:
        def __init__(self): self.BackgroundPatternColor=None
    class Table:
        def __init__(self):
            self.Rows=SimpleNamespace(Alignment=None)
            self.AllowAutoFit=True
            self.PreferredWidthType=None; self.PreferredWidth=None
            self.TopPadding=self.BottomPadding=self.LeftPadding=self.RightPadding=None
            self.Shading=Shading(); self.Borders=Borders()
    table=Table()
    controller=InteractiveWordController.for_testing(active_range=SimpleNamespace())
    controller._table_stack=[{"table":table,"cells":{},"structure_complete":False}]
    controller.execute_event(ReconstructionEvent("SetTableProperties","t",{
        "alignment":"center","layout":"fixed","width":"5000","width_type":"dxa",
        "cell_margins":{"top":{"w":"40"},"bottom":{"w":"60"},"left":{"w":"80"},"right":{"w":"100"}},
        "shading_fill":"D9EAF7",
        "borders":{"top":{"val":"single","sz":"8","color":"112233"},"insideH":{"val":"double","sz":"4","color":"445566"}},
    }))
    assert table.Rows.Alignment == 1
    assert table.AllowAutoFit is False
    assert table.PreferredWidthType == 3 and table.PreferredWidth == 250.0
    assert (table.TopPadding,table.BottomPadding,table.LeftPadding,table.RightPadding) == (2.0,3.0,4.0,5.0)
    assert table.Shading.BackgroundPatternColor == controller._word_color("D9EAF7")
    assert table.Borders(-1).LineStyle == 1
    assert table.Borders(-1).LineWidth == 8
    assert table.Borders(-1).Color == controller._word_color("112233")
    assert table.Borders(-5).LineStyle == 7


def test_word_cell_property_executor_applies_shading_borders_margins_and_direction_before_typing():
    from types import SimpleNamespace
    from word_replica.domain.reconstruction import ReconstructionEvent
    from word_replica.renderers.interactive_word import InteractiveWordController

    class Border:
        def __init__(self): self.LineStyle=None; self.LineWidth=None; self.Color=None
    class Borders:
        def __init__(self): self.items={i:Border() for i in (-1,-2,-3,-4)}
        def __call__(self,index): return self.items[index]
    class Shading:
        def __init__(self): self.BackgroundPatternColor=None
    class Cell:
        def __init__(self):
            self.VerticalAlignment=None; self.Width=None
            self.PreferredWidthType=None; self.PreferredWidth=None
            self.TopPadding=self.BottomPadding=self.LeftPadding=self.RightPadding=None
            self.Shading=Shading(); self.Borders=Borders(); self.Range=SimpleNamespace(Orientation=None)
    cell=Cell(); table=SimpleNamespace()
    controller=InteractiveWordController.for_testing(active_range=SimpleNamespace())
    controller._table_stack=[{"table":table,"cells":{(1,1):cell},"structure_complete":False}]
    controller.execute_event(ReconstructionEvent("SetCellProperties","c",{"row":1,"column":1,"properties":{
        "vertical_alignment":"center","width":"2400","width_type":"dxa","shading_fill":"ABCDEF",
        "cell_margins":{"top":{"w":"20"},"bottom":{"w":"40"},"left":{"w":"60"},"right":{"w":"80"}},
        "borders":{"left":{"val":"single","sz":"8","color":"010203"}},
        "text_direction":"tbRl",
    }}))
    assert cell.VerticalAlignment == 1
    assert cell.PreferredWidthType == 3 and cell.PreferredWidth == 120.0
    assert cell.Width is None
    assert (cell.TopPadding,cell.BottomPadding,cell.LeftPadding,cell.RightPadding) == (1.0,2.0,3.0,4.0)
    assert cell.Shading.BackgroundPatternColor == controller._word_color("ABCDEF")
    assert cell.Borders(-2).LineStyle == 1 and cell.Borders(-2).LineWidth == 8
    assert cell.Range.Orientation == 3


def test_blueprint_and_word_executor_apply_source_table_style(corpus_dir):
    from types import SimpleNamespace
    from word_replica.parser.parser import DocxParser
    from word_replica.interactive.blueprint import BlueprintCompiler
    from word_replica.renderers.interactive_word import InteractiveWordController

    model = DocxParser().parse(corpus_dir / "04_tables_merged.docx")
    event = next(e for e in BlueprintCompiler().compile(model).events if e.event_type == "SetTableProperties")
    assert event.payload["style_id"] == "TableGrid"

    table = SimpleNamespace(
        Style=None, Rows=SimpleNamespace(Alignment=None), AllowAutoFit=True,
        PreferredWidthType=None, PreferredWidth=None,
        TopPadding=None, BottomPadding=None, LeftPadding=None, RightPadding=None,
        Shading=SimpleNamespace(BackgroundPatternColor=None),
        Borders=SimpleNamespace(),
    )
    controller = InteractiveWordController.for_testing(active_range=SimpleNamespace())
    controller._table_stack=[{"table":table,"cells":{},"structure_complete":False}]
    controller.execute_event(event)
    assert table.Style == "TableGrid"
