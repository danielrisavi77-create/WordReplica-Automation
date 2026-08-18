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
