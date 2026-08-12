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
    assert fake.insert_after_calls == ["\t"]
    assert fake.break_calls == [6, 7]


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

    assert [name for name, _ in fake.log if name == "font.Reset"] == ["font.Reset"]
    assert [name for name, _ in fake.log if name == "font.Bold"] == ["font.Bold"]


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


class FakeInlineImage:
    def __init__(self, log):
        self.log = log
        self.Width = None; self.Height = None; self.LockAspectRatio = None
        self.Range = FormattingRange()
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
    def __init__(self, log): self.log=log
    def AddPicture(self, **kwargs):
        self.log.append(("add_picture", kwargs["FileName"]))
        return FakeInlineImage(self.log)


class FakeImageDocument(FakeDocument):
    def __init__(self, log):
        super().__init__(log)
        self.InlineShapes = FakeInlineShapes(log)


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


def test_create_field_seeds_cached_result_without_refreshing_mid_reconstruction():
    from types import SimpleNamespace

    class ResultRange:
        def __init__(self): self.Text = ""; self.Start=4; self.End=4; self.collapse=[]
        @property
        def Duplicate(self): return self
        def Collapse(self, direction): self.collapse.append(direction)
    class Field:
        def __init__(self): self.Result=ResultRange(); self.updated=0
        def Update(self): self.updated += 1
    class Fields:
        def __init__(self): self.created=None
        def Add(self, Range, Type, Text, PreserveFormatting):
            self.created=Field(); self.args=(Range,Type,Text,PreserveFormatting); return self.created
    fields=Fields(); document=SimpleNamespace(Fields=fields)
    active=SimpleNamespace(Start=4, End=4)
    controller=InteractiveWordController.for_testing(active_range=active); controller.document=document
    controller.execute_event(ReconstructionEvent("CreateField", "f1", {
        "instruction": 'TOC \\o "1-3"', "cached_result": "Chapter One .... 1"
    }))
    assert fields.created.Result.Text == "Chapter One .... 1"
    assert fields.created.updated == 0
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


def test_end_table_reuses_words_existing_post_table_paragraph():
    after = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=after)
    controller._table_stack = [{"table": object(), "after_range": after, "element_id": "t1"}]
    controller._active_table_element_id = "t1"
    controller._paragraph_started = False
    controller.execute_event(ReconstructionEvent("EndTable", "t1", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))
    assert after.paragraph_calls == 0


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


def test_page_break_at_end_of_source_paragraph_reuses_word_continuation_paragraph():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("InsertPageBreak", "r1", {}))
    controller.execute_event(ReconstructionEvent("EndParagraph", "p1", {}))
    controller.execute_event(ReconstructionEvent("BeginParagraph", "p2", {}))
    assert fake.break_calls == [7]
    assert fake.paragraph_calls == 0


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
