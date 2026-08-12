from word_replica.domain.model import PreservedPart
from word_replica.renderers.word_com import WordComRenderer


def test_preserved_part_without_safe_relationship_becomes_warning():
    renderer = WordComRenderer(visible=False)
    part = PreservedPart(
        "word/embeddings/object.bin",
        "application/octet-stream",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
        "abc",
        b"x",
    )
    renderer._handle_preserved_part(part)
    assert [warning.code for warning in renderer._warnings] == ["OLE_TRANSFER_UNSUPPORTED"]
    assert renderer._post_save_transfers == []


def test_word_paragraph_applies_canonical_run_formatting_without_flattening_style_semantics():
    from word_replica.domain.model import Paragraph, Run
    from word_replica.renderers.word_com import WordComRenderer

    class Font:
        Hidden = 0

    class Format:
        KeepWithNext = 0
        PageBreakBefore = 0
        KeepTogether = 0

    class DynamicRange:
        def __init__(self, doc, start, end):
            self.doc = doc
            self.Start = start
            self.End = end
            self.Bold = 0
            self.Italic = 0
            self.Underline = 0
            self.Font = Font()

        @property
        def Text(self):
            return self.doc.content[self.Start:self.End]

        @Text.setter
        def Text(self, value):
            self.doc.content = self.doc.content[:self.Start] + value + self.doc.content[self.End:]
            self.End = self.Start + len(value)

        def InsertAfter(self, text):
            self.doc.content = self.doc.content[:self.End] + text + self.doc.content[self.End:]
            self.End += len(text)

    class ParagraphRange:
        def __init__(self, doc):
            self.doc = doc
            self.Start = 10
            self.Style = None

        @property
        def End(self):
            return 10 + len(self.doc.content) + 1

    class ParagraphObject:
        def __init__(self, doc):
            self.Range = ParagraphRange(doc)
            self.Format = Format()

    class Paragraphs:
        Count = 1
        def __init__(self, paragraph): self.paragraph = paragraph
        def __call__(self, index): return self.paragraph
        def Add(self): return self.paragraph

    class Doc:
        def __init__(self):
            self.content = ''
            self.paragraph = ParagraphObject(self)
            self.Paragraphs = Paragraphs(self.paragraph)
            self.ranges = []

        def Range(self, start, end):
            r = DynamicRange(self, start - 10, end - 10)
            r.absolute_start = start
            r.absolute_end = end
            self.ranges.append(r)
            return r

    doc = Doc()
    block = Paragraph('p', runs=[
        Run('r1', text='Bold', properties={'bold': True}),
        Run('r2', text='It', properties={'italic': True, 'underline': True, 'hidden': True}),
    ])
    WordComRenderer(visible=False)._insert_paragraph(doc, block)

    assert doc.content == 'BoldIt'
    inserted = [r for r in doc.ranges if r.absolute_start == r.absolute_end]
    assert [(r.absolute_start, r.End - r.Start) for r in inserted] == [(10, 4), (14, 2)]
    assert inserted[0].Bold == -1
    assert inserted[1].Italic == -1
    assert inserted[1].Underline == 1
    assert inserted[1].Font.Hidden == -1


def test_word_assets_are_embedded_approximately_and_warn_about_position():
    from word_replica.domain.model import BinaryAsset, DocumentModel

    class Content: Start = 0; End = 20
    class InlineShapes:
        def __init__(self): self.payloads = []
        def AddPicture(self, path, link, save, target):
            from pathlib import Path
            self.payloads.append((Path(path).read_bytes(), link, save, target))
    class Doc:
        def __init__(self): self.Content = Content(); self.InlineShapes = InlineShapes()
        def Range(self, start, end): return (start, end)

    model = DocumentModel('abc')
    model.assets['a'] = BinaryAsset('a', 'word/media/image1.png', 'image/png', 'sha', b'PNGDATA')
    renderer = WordComRenderer(False); doc = Doc()
    renderer._insert_assets(doc, model)
    assert doc.InlineShapes.payloads[0][0] == b'PNGDATA'
    assert [w.code for w in renderer._warnings] == ['WORD_ASSET_POSITION_APPROXIMATION']


def test_word_notes_fields_and_bookmarks_are_recreated_with_position_warnings():
    from word_replica.domain.model import Bookmark, DocumentModel, Field, Paragraph, Run

    class Content: Start = 0; End = 50
    class Collection:
        def __init__(self): self.calls = []
        def Add(self, *args, **kwargs): self.calls.append((args, kwargs))
    class Doc:
        def __init__(self):
            self.Content = Content(); self.Footnotes = Collection(); self.Endnotes = Collection(); self.Fields = Collection(); self.Bookmarks = Collection()
        def Range(self, start, end): return (start, end)

    model = DocumentModel('abc')
    model.footnotes['1'] = [Paragraph('p1', [Run('r1', 'Foot')])]
    model.endnotes['1'] = [Paragraph('p2', [Run('r2', 'End')])]
    model.bookmarks.append(Bookmark('b1', 'TargetBookmark', '/a', '/b'))
    model.fields.append(Field('f1', 'REF TargetBookmark \\h', 'Target', False))
    renderer = WordComRenderer(False); doc = Doc()
    renderer._restore_notes(doc, model); renderer._restore_fields_bookmarks(doc, model)
    assert len(doc.Footnotes.calls) == 1 and len(doc.Endnotes.calls) == 1
    assert len(doc.Bookmarks.calls) == 1 and len(doc.Fields.calls) == 1
    codes = {w.code for w in renderer._warnings}
    assert 'WORD_NOTE_REFERENCE_POSITION_APPROXIMATION' in codes
    assert 'WORD_FIELD_POSITION_APPROXIMATION' in codes
    assert 'WORD_BOOKMARK_POSITION_APPROXIMATION' in codes


def test_word_full_review_structures_are_never_silently_dropped():
    from word_replica.domain.model import Comment, DocumentModel, Paragraph, RevisionSpan, Run
    model = DocumentModel('abc')
    model.comments['1'] = Comment('1', 'Reviewer', None, [Paragraph('p', [Run('r', 'Comment')])])
    model.revisions.append(RevisionSpan('1', 'insert', 'Reviewer', None, '/p', 'Inserted'))
    model.extras['tracked_changes_enabled'] = True
    renderer = WordComRenderer(False)
    renderer._warn_review_structures(model)
    assert [w.code for w in renderer._warnings] == ['WORD_REVIEW_STRUCTURES_APPROXIMATION']


def test_word_page_break_uses_one_continuation_paragraph_without_extra_break_character():
    from word_replica.domain.model import Paragraph, Run

    class Font:
        Hidden = 0

    class Format:
        def __init__(self):
            self.KeepWithNext = 0
            self.PageBreakBefore = 0
            self.KeepTogether = 0

    class DynamicRange:
        def __init__(self, paragraph, start, end):
            self.paragraph = paragraph
            self.Start = start
            self.End = end
            self.Bold = 0
            self.Italic = 0
            self.Underline = 0
            self.Font = Font()

        @property
        def Text(self):
            return self.paragraph.text[self.Start:self.End]

        @Text.setter
        def Text(self, value):
            self.paragraph.text = self.paragraph.text[:self.Start] + value + self.paragraph.text[self.End:]
            self.End = self.Start + len(value)

        def InsertAfter(self, text):
            self.paragraph.text = self.paragraph.text[:self.End] + text + self.paragraph.text[self.End:]
            self.End += len(text)

        def InsertBreak(self, Type=7):
            raise AssertionError("page breaks must not use Word InsertBreak because it creates extra paragraph boundaries")

    class ParagraphRange:
        def __init__(self, doc, paragraph):
            self.doc = doc
            self.paragraph = paragraph
            self.Start = 0
            self.Style = None

        @property
        def End(self):
            return len(self.paragraph.text) + 1

        @property
        def Text(self):
            return self.paragraph.text + "\r"

        @Text.setter
        def Text(self, value):
            self.paragraph.text = value.rstrip("\r")

        def InsertParagraphAfter(self):
            self.doc.paragraphs.append(ParagraphObject(self.doc))

    class ParagraphObject:
        def __init__(self, doc):
            self.doc = doc
            self.text = ""
            self.Format = Format()
            self.Range = ParagraphRange(doc, self)

    class Paragraphs:
        def __init__(self, doc):
            self.doc = doc

        @property
        def Count(self):
            return len(self.doc.paragraphs)

        def __call__(self, index):
            return self.doc.paragraphs[index - 1]

    class Doc:
        def __init__(self):
            self.paragraphs = [ParagraphObject(self)]
            self.Paragraphs = Paragraphs(self)

        def Range(self, start, end):
            paragraph = self.paragraphs[-1]
            assert 0 <= start <= end <= len(paragraph.text)
            return DynamicRange(paragraph, start, end)

    doc = Doc()
    renderer = WordComRenderer(visible=False)
    first = Paragraph('p1', runs=[
        Run('r1', text='Bold', properties={'bold': True}),
        Run('r2', text='\t'),
        Run('r3', text=' normal'),
        Run('r4', text='\n', properties={'break_types': ['page']}),
        Run('r5', text='after break'),
    ], properties={'pageBreakBefore': True})
    trailing = Paragraph('p2', runs=[])

    renderer._insert_paragraph(doc, first)
    renderer._insert_paragraph(doc, trailing)

    assert [p.text for p in doc.paragraphs] == ['Bold\t normal', 'after break', '']
    assert doc.paragraphs[0].Format.PageBreakBefore == -1
    assert doc.paragraphs[1].Format.PageBreakBefore == -1
    assert "\n".join(p.text for p in doc.paragraphs) == 'Bold\t normal\nafter break\n'



def test_word_sections_are_not_precreated_before_body_and_materialize_at_canonical_boundary():
    from word_replica.domain.model import DocumentModel, Paragraph, Run, Section

    class PageSetup:
        def __init__(self):
            self.Orientation = 0

    class SectionObject:
        def __init__(self):
            self.PageSetup = PageSetup()

    class Sections:
        def __init__(self):
            self.items = [SectionObject()]
            self.add_calls = []

        @property
        def Count(self):
            return len(self.items)

        def __call__(self, index):
            return self.items[index - 1]

        def Add(self, *args, **kwargs):
            self.add_calls.append((args, kwargs))
            section = SectionObject()
            self.items.append(section)
            return section

    class BoundaryRange:
        pass

    class BoundaryParagraph:
        def __init__(self):
            self.Range = BoundaryRange()

    class Doc:
        def __init__(self):
            self.Sections = Sections()

    model = DocumentModel('abc')
    model.sections = [
        Section('s0', {'orientation': 'portrait'}),
        Section('s1', {'orientation': 'landscape'}),
    ]
    model.body = [
        Paragraph('p0', [Run('r0', 'Heading')]),
        Paragraph('p1', [], properties={'section_index': 0}),
    ]

    doc = Doc()
    renderer = WordComRenderer(False)
    renderer._section_boundaries = {0: BoundaryParagraph()}

    renderer._apply_sections(doc, model)
    assert doc.Sections.Count == 1
    assert doc.Sections.add_calls == []

    renderer._materialize_sections(doc, model)
    assert doc.Sections.Count == 2
    assert len(doc.Sections.add_calls) == 1
    args, kwargs = doc.Sections.add_calls[0]
    assert kwargs['Range'] is renderer._section_boundaries[0].Range
    assert kwargs['Start'] == 2
    assert doc.Sections(1).PageSetup.Orientation == 0
    assert doc.Sections(2).PageSetup.Orientation == 1


def test_structural_section_boundary_paragraph_is_not_inserted_as_visible_empty_body_block():
    from word_replica.domain.model import DocumentModel, Paragraph, Run, Section

    class RangeObject:
        def __init__(self, start=0, end=0):
            self.Start = start
            self.End = end

    class Content:
        Start = 0
        End = 10

    class PageSetup:
        def __init__(self):
            self.Orientation = 0

    class SectionObject:
        def __init__(self):
            self.PageSetup = PageSetup()

    class Sections:
        def __init__(self):
            self.items = [SectionObject()]
            self.add_calls = []

        @property
        def Count(self):
            return len(self.items)

        def __call__(self, index):
            return self.items[index - 1]

        def Add(self, *args, **kwargs):
            self.add_calls.append((args, kwargs))
            section = SectionObject()
            self.items.append(section)
            return section

    class Doc:
        def __init__(self):
            self.Content = Content()
            self.Sections = Sections()
            self.range_calls = []

        def Range(self, start, end):
            self.range_calls.append((start, end))
            return RangeObject(start, end)

    model = DocumentModel('abc')
    model.sections = [
        Section('s0', {'orientation': 'portrait'}),
        Section('s1', {'orientation': 'landscape'}),
    ]
    structural = Paragraph('p-section', [], properties={'section_index': 0})

    renderer = WordComRenderer(False)
    doc = Doc()

    assert renderer._is_structural_section_boundary(structural) is True
    renderer._materialize_section_boundary(doc, model, 0, boundary=None)

    assert doc.range_calls == [(9, 9)]
    assert doc.Sections.Count == 2
    assert doc.Sections.add_calls[0][1]['Start'] == 2
    assert doc.Sections(2).PageSetup.Orientation == 1


def test_word_table_populates_all_logical_cells_before_any_merge_mutates_collection():
    from word_replica.domain.model import Paragraph, Run, Table, TableCell, TableRow

    events = []

    class Range:
        def __init__(self, cell=None): self.cell = cell; self.Start = 0; self.End = 0
        @property
        def Text(self): return self.cell.text if self.cell else ''
        @Text.setter
        def Text(self, value):
            if self.cell: self.cell.text = value; events.append(('text', self.cell.row, self.cell.col, value))

    class Cell:
        def __init__(self, table, row, col):
            self.table = table; self.row = row; self.col = col; self.text = ''; self.Range = Range(self)
        def Merge(self, other):
            events.append(('merge', self.row, self.col, other.row, other.col))
            self.table.merged = True

    class TableObject:
        def __init__(self, rows, cols):
            self.merged = False
            self.cells = {(r,c): Cell(self,r,c) for r in range(1,rows+1) for c in range(1,cols+1)}
        def Cell(self, r, c):
            # Simulate Word invalidating the old logical collection after a merge.
            if self.merged:
                raise AssertionError('renderer accessed table.Cell after merge mutated the collection')
            return self.cells[(r,c)]

    class Tables:
        def __init__(self): self.table = None
        def Add(self, insert_at, rows, cols): self.table = TableObject(rows, cols); return self.table

    class Content: End = 1
    class Doc:
        def __init__(self): self.Content = Content(); self.Tables = Tables()
        def Range(self, start, end): return Range()

    block = Table('t', rows=[
        TableRow('r1', [
            TableCell('c11', [Paragraph('p11', [Run('x11', 'Merged')])], {'grid_span': 2, 'v_merge': None}),
            TableCell('c13', [Paragraph('p13', [Run('x13', 'Tail')])], {'grid_span': 1, 'v_merge': None}),
        ]),
        TableRow('r2', [
            TableCell('c21', [Paragraph('p21', [Run('x21', 'Vertical')])], {'grid_span': 1, 'v_merge': 'restart'}),
            TableCell('c22', [Paragraph('p22', [Run('x22', 'A')])], {'grid_span': 1, 'v_merge': None}),
            TableCell('c23', [Paragraph('p23', [Run('x23', 'B')])], {'grid_span': 1, 'v_merge': None}),
        ]),
        TableRow('r3', [
            TableCell('c31', [Paragraph('p31', [])], {'grid_span': 1, 'v_merge': 'continue'}),
            TableCell('c32', [Paragraph('p32', [Run('x32', 'C')])], {'grid_span': 1, 'v_merge': None}),
            TableCell('c33', [Paragraph('p33', [Run('x33', 'D')])], {'grid_span': 1, 'v_merge': None}),
        ]),
    ])

    renderer = WordComRenderer(False); doc = Doc()
    renderer._insert_table(doc, block)
    first_merge_index = next(i for i, event in enumerate(events) if event[0] == 'merge')
    assert all(event[0] == 'text' for event in events[:first_merge_index])
    assert ('text', 1, 3, 'Tail') in events[:first_merge_index]
    assert ('merge', 2, 1, 3, 1) in events
    assert ('merge', 1, 1, 1, 2) in events


def test_word_render_uses_canonical_package_builder_then_word_validation(monkeypatch, tmp_path):
    import word_replica.renderers.word_com as module
    from word_replica.domain.model import DocumentModel
    from word_replica.renderers.base import RenderResult

    calls = []

    class FakePureRenderer:
        def render(self, model, output_path, context):
            calls.append(("pure", model.source_sha256, context))
            output_path.write_bytes(b"docx")
            return RenderResult(output_path=output_path, warnings=[], stages_completed=["body", "metadata"])

    monkeypatch.setattr(module, "PureDocxRenderer", FakePureRenderer, raising=False)
    renderer = module.WordComRenderer(visible=False)
    monkeypatch.setattr(renderer, "_validate_package_with_word", lambda output: calls.append(("word", output)))

    output = tmp_path / "rebuilt.docx"
    result = renderer.render(DocumentModel("source-hash"), output, context=None)

    assert calls == [("pure", "source-hash", None), ("word", output)]
    assert result.output_path == output
    assert result.stages_completed[-1] == "word_validation"
