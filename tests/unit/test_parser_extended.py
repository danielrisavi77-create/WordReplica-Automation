from word_replica.domain.model import Table
from word_replica.parser.parser import DocxParser
from tests.fixtures.build_fixtures import build_extended_fixture


def test_parser_extracts_tables_assets_headers_and_notes(tmp_path):
    source = build_extended_fixture(tmp_path / "extended.docx")
    model = DocxParser().parse(source)
    table = next(block for block in model.body if isinstance(block, Table))
    assert len(table.rows) == 2
    assert table.rows[0].cells[0].properties["grid_span"] == 2
    assert len(model.assets) == 1
    assert model.extras["headers"]
    assert model.extras["footers"]
    assert model.extras["footnotes"]
    assert model.extras["endnotes"]
    assert model.extras["footnotes"]["1"][0].text() == "Footnote text"


def test_parser_ties_image_occurrence_to_asset_and_run_token(tmp_path):
    source = build_extended_fixture(tmp_path / "extended-image.docx")
    model = DocxParser().parse(source)
    assert len(model.drawings) == 1
    drawing = model.drawings[0]
    assert drawing.representation == "inline"
    assert drawing.asset_id in model.assets
    tokens = [token for p in model.iter_paragraphs() for r in p.runs for token in r.properties.get("content_tokens", [])]
    assert any(token.get("kind") == "drawing" for token in tokens)


def test_parser_exposes_used_custom_paragraph_style_definition(corpus_dir):
    from word_replica.parser.parser import DocxParser

    model = DocxParser().parse(corpus_dir / "02_headings_styles.docx")
    style = model.extras["style_definitions"]["ReplicaBody"]
    assert style["type"] == "paragraph"
    assert style["name"] == "ReplicaBody"
    assert style["run_properties"]["font_ascii"] == "Arial"
    assert style["run_properties"]["size_half_points"] == "22"


def test_parser_retains_table_style_id(corpus_dir):
    from word_replica.domain.model import Table
    from word_replica.parser.parser import DocxParser

    model = DocxParser().parse(corpus_dir / "04_tables_merged.docx")
    table = next(block for block in model.body if isinstance(block, Table))
    assert table.properties["style_id"] == "TableGrid"


def test_parser_exposes_document_defaults_font_size_and_paragraph_defaults(tmp_path):
    from zipfile import ZipFile, ZIP_DEFLATED
    source = tmp_path / "defaults.docx"
    ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
    doc = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'''
    styles = '''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
    with ZipFile(source, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>''')
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
    model = DocxParser().parse(source)
    defaults = model.extras["document_defaults"]
    assert defaults["run_properties"]["font_ascii"] == "Times New Roman"
    assert defaults["run_properties"]["size_half_points"] == "24"
    assert defaults["paragraph_properties"]["spacing_line"] == "300"
    assert defaults["paragraph_properties"]["spacing_line_rule"] == "auto"
    assert model.extras["default_paragraph_style_id"] == "Normal"


def test_theme_font_scheme_extracts_major_and_minor_latin_east_asia_and_complex_script():
    from word_replica.parser.parser import _parse_theme_font_scheme

    xml = b"""<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:fontScheme name="Office">
          <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="MS Gothic"/><a:cs typeface="Arial"/></a:majorFont>
          <a:minorFont><a:latin typeface="Cambria"/><a:ea typeface="MS Mincho"/><a:cs typeface="Times New Roman"/></a:minorFont>
        </a:fontScheme>
      </a:themeElements>
    </a:theme>"""
    scheme = _parse_theme_font_scheme({"word/theme/theme1.xml": xml})
    assert scheme["majorHAnsi"] == "Calibri"
    assert scheme["minorHAnsi"] == "Cambria"
    assert scheme["majorEastAsia"] == "MS Gothic"
    assert scheme["minorEastAsia"] == "MS Mincho"
    assert scheme["majorBidi"] == "Arial"
    assert scheme["minorBidi"] == "Times New Roman"
