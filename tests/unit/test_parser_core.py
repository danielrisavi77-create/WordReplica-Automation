from word_replica.parser.parser import DocxParser
from tests.fixtures.build_fixtures import build_core_fixture


def test_parser_extracts_text_styles_and_sections(tmp_path):
    source = build_core_fixture(tmp_path / "core.docx")
    model = DocxParser().parse(source)
    assert model.body[0].style_id == "Heading1"
    assert model.body[0].text() == "Heading One"
    assert model.body[0].properties["keepNext"] is True
    assert model.body[1].runs[0].properties["bold"] is True
    assert "\t" in model.body[1].text()
    assert "\n" in model.body[1].text()
    assert model.body[1].properties["pageBreakBefore"] is True
    assert len(model.sections) == 2
    assert model.body[-1].properties["section_index"] == 0
    assert model.sections[-1].properties["orientation"] == "landscape"
    assert model.styles_xml
    assert model.extras["source_properties"].creator is not None


def test_parser_preserves_page_break_type_for_word_renderer(tmp_path):
    source = build_core_fixture(tmp_path / "core-page-break.docx")
    model = DocxParser().parse(source)
    break_runs = [run for run in model.body[1].runs if "\n" in run.text]
    assert len(break_runs) == 1
    assert break_runs[0].properties["break_types"] == ["page"]


def test_parser_ignores_word_system_goback_bookmark_but_keeps_real_bookmarks(tmp_path):
    from zipfile import ZIP_DEFLATED, ZipFile
    from lxml import etree
    from docx import Document
    from word_replica.parser.parser import DocxParser

    source = tmp_path / 'bookmarks.docx'
    doc = Document(); doc.add_paragraph('Body'); doc.save(source)
    with ZipFile(source, 'r') as src:
        members = {name: src.read(name) for name in src.namelist()}
    root = etree.fromstring(members['word/document.xml'])
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    p = root.find(f'.//{{{ns}}}p')
    for bookmark_id, name in [('1', '_GoBack'), ('2', 'RealBookmark')]:
        start = etree.Element(f'{{{ns}}}bookmarkStart')
        start.set(f'{{{ns}}}id', bookmark_id); start.set(f'{{{ns}}}name', name)
        end = etree.Element(f'{{{ns}}}bookmarkEnd'); end.set(f'{{{ns}}}id', bookmark_id)
        p.insert(0, start); p.append(end)
    members['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
    tmp = source.with_suffix('.tmp.docx')
    with ZipFile(tmp, 'w', ZIP_DEFLATED) as dst:
        for name, data in members.items(): dst.writestr(name, data)
    tmp.replace(source)

    model = DocxParser().parse(source)
    assert [bookmark.name for bookmark in model.bookmarks] == ['RealBookmark']


def test_fldsimple_field_is_extracted_same_as_the_complex_form(tmp_path):
    from tests.fixtures.build_fixtures import build_field_simple_form

    source = build_field_simple_form(tmp_path / "field_simple.docx")
    model = DocxParser().parse(source)

    assert len(model.fields) == 1
    field = model.fields[0]
    assert field.instruction == "REF _Ref_tab1 \\h"
    assert field.result_text == "1"
    # The result text must still read normally as paragraph content -
    # unwrapping the fldSimple shorthand must not lose the visible text.
    assert "Table number: 1" in model.body[0].text()


def test_bookmark_paths_are_stable_across_sdt_flattening(tmp_path):
    # Regression: a source TOC wrapped in an <w:sdt> content control counts as
    # ONE direct child of <w:body>, however many paragraphs it contains inside.
    # A renderer that flattens the control's contents to plain paragraphs (as
    # this one does) makes each of those paragraphs its own direct child of
    # <w:body> instead - shifting the absolute w:p[N] position of every bookmark
    # that follows even though nothing about the actual content changed.
    # Confirmed live: this alone accounted for the bulk of a golden document's
    # G7 fidelity mismatches. Bookmark paths must be anchored to the enclosing
    # paragraph's position among ALL <w:p> elements (stable across sdt nesting),
    # not its position among its immediate parent's children.
    from zipfile import ZIP_DEFLATED, ZipFile
    from lxml import etree
    from word_replica.parser.parser import DocxParser

    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def build(sdt_wrapped: bool) -> bytes:
        toc_paragraphs = "".join(f"<w:p><w:r><w:t>TOC{i}</w:t></w:r></w:p>" for i in range(2))
        toc_block = (
            f"<w:sdt><w:sdtContent>{toc_paragraphs}</w:sdtContent></w:sdt>"
            if sdt_wrapped else toc_paragraphs
        )
        return (
            f"<w:document xmlns:w='{ns}'><w:body>"
            f"<w:p><w:r><w:t>Before</w:t></w:r></w:p>"
            f"{toc_block}"
            f"<w:p><w:bookmarkStart w:id='1' w:name='Target'/>"
            f"<w:r><w:t>Heading</w:t></w:r>"
            f"<w:bookmarkEnd w:id='1'/></w:p>"
            f"</w:body></w:document>"
        ).encode("utf-8")

    def make_docx(name: str, sdt_wrapped: bool):
        path = tmp_path / name
        with ZipFile(path, "w", ZIP_DEFLATED) as archive:
            archive.writestr("word/document.xml", build(sdt_wrapped))
            archive.writestr(
                "[Content_Types].xml",
                b"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>"
                b"<Default Extension='xml' ContentType='application/xml'/>"
                b"<Override PartName='/word/document.xml' "
                b"ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>"
                b"</Types>",
            )
            archive.writestr(
                "_rels/.rels",
                b"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
                b"<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' "
                b"Target='word/document.xml'/></Relationships>",
            )
            archive.writestr(
                "word/_rels/document.xml.rels",
                b"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>",
            )
        return path

    sdt_source = make_docx("sdt.docx", sdt_wrapped=True)
    flat_source = make_docx("flat.docx", sdt_wrapped=False)

    sdt_model = DocxParser().parse(sdt_source)
    flat_model = DocxParser().parse(flat_source)

    assert len(sdt_model.bookmarks) == 1
    assert len(flat_model.bookmarks) == 1
    assert sdt_model.bookmarks[0].start_path == flat_model.bookmarks[0].start_path
    assert sdt_model.bookmarks[0].end_path == flat_model.bookmarks[0].end_path
    # Both should resolve to the 4th paragraph overall (Before, TOC0, TOC1, Target).
    assert sdt_model.bookmarks[0].start_path == "//w:p[4]/w:bookmarkStart"


def _save_with_document_xml_edit(tmp_path, filename, edit):
    from zipfile import ZIP_DEFLATED, ZipFile
    from lxml import etree
    from docx import Document

    source = tmp_path / filename
    doc = Document(); doc.add_paragraph('Body'); doc.save(source)
    with ZipFile(source, 'r') as src:
        members = {name: src.read(name) for name in src.namelist()}
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    root = etree.fromstring(members['word/document.xml'])
    edit(root, ns)
    members['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
    tmp = source.with_suffix('.tmp.docx')
    with ZipFile(tmp, 'w', ZIP_DEFLATED) as dst:
        for name, data in members.items(): dst.writestr(name, data)
    tmp.replace(source)
    return source


def test_parser_flattens_sdt_content_control_paragraphs_into_the_body(tmp_path):
    from lxml import etree
    from word_replica.parser.parser import DocxParser

    def edit(root, ns):
        body = root.find(f'.//{{{ns}}}body')
        first_p = body.find(f'{{{ns}}}p')
        sdt = etree.Element(f'{{{ns}}}sdt')
        sdt_content = etree.SubElement(sdt, f'{{{ns}}}sdtContent')
        p = etree.SubElement(sdt_content, f'{{{ns}}}p')
        r = etree.SubElement(p, f'{{{ns}}}r')
        t = etree.SubElement(r, f'{{{ns}}}t'); t.text = 'Inside content control'
        first_p.addprevious(sdt)

    source = _save_with_document_xml_edit(tmp_path, 'sdt.docx', edit)
    model = DocxParser().parse(source)
    assert 'Inside content control' in model.plain_text()
    assert any(p.__class__.__name__ == 'Paragraph' and 'Inside content control' in p.text() for p in model.body)


def test_parser_flattens_nested_sdt_inside_sdt(tmp_path):
    from lxml import etree
    from word_replica.parser.parser import DocxParser

    def edit(root, ns):
        body = root.find(f'.//{{{ns}}}body')
        first_p = body.find(f'{{{ns}}}p')
        outer = etree.Element(f'{{{ns}}}sdt')
        outer_content = etree.SubElement(outer, f'{{{ns}}}sdtContent')
        inner = etree.SubElement(outer_content, f'{{{ns}}}sdt')
        inner_content = etree.SubElement(inner, f'{{{ns}}}sdtContent')
        p = etree.SubElement(inner_content, f'{{{ns}}}p')
        r = etree.SubElement(p, f'{{{ns}}}r')
        t = etree.SubElement(r, f'{{{ns}}}t'); t.text = 'Nested control'
        first_p.addprevious(outer)

    source = _save_with_document_xml_edit(tmp_path, 'nested-sdt.docx', edit)
    model = DocxParser().parse(source)
    assert 'Nested control' in model.plain_text()


def test_parser_extracts_runs_wrapped_in_hyperlink(tmp_path):
    from lxml import etree
    from word_replica.parser.parser import DocxParser

    def edit(root, ns):
        body = root.find(f'.//{{{ns}}}body')
        first_p = body.find(f'{{{ns}}}p')
        hyperlink = etree.SubElement(first_p, f'{{{ns}}}hyperlink')
        r = etree.SubElement(hyperlink, f'{{{ns}}}r')
        t = etree.SubElement(r, f'{{{ns}}}t'); t.text = 'Click here'

    source = _save_with_document_xml_edit(tmp_path, 'hyperlink.docx', edit)
    model = DocxParser().parse(source)
    assert 'Click here' in model.plain_text()
