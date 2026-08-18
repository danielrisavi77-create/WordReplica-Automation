from zipfile import ZipFile

from word_replica.domain.model import DocumentModel, Paragraph, PreservedPart, Run
from word_replica.renderers.pure_docx import PureDocxRenderer


def test_renderer_creates_new_docx_with_document_xml(tmp_path):
    model = DocumentModel(
        source_sha256="abc",
        body=[Paragraph("p1", runs=[Run("r1", text="Hello")])],
    )
    output = tmp_path / "out.docx"
    PureDocxRenderer().render(model, output, context=None)
    with ZipFile(output) as z:
        assert b"Hello" in z.read("word/document.xml")


def test_renderer_writes_run_font_color_size_and_language_not_only_bold_italic(tmp_path):
    model = DocumentModel(
        source_sha256="abc",
        body=[Paragraph("p1", runs=[Run("r1", text="Naslov", properties={
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_cs": "Times New Roman",
            "color": "000000",
            "size_half_points": "28",
            "strike": True,
            "highlight": "yellow",
            "vert_align": "superscript",
            "language": "en-US",
            "language_east_asia": "en-US",
            "language_bidi": "ar-SA",
            "character_spacing": "10",
            "character_position": "4",
        })])],
    )
    output = tmp_path / "out.docx"
    PureDocxRenderer().render(model, output, context=None)
    with ZipFile(output) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    assert 'w:ascii="Times New Roman"' in xml
    assert 'w:val="000000"' in xml
    assert '<w:sz w:val="28"/>' in xml
    assert '<w:szCs w:val="28"/>' in xml
    assert "<w:strike/>" in xml
    assert 'w:highlight w:val="yellow"' in xml
    assert 'w:vertAlign w:val="superscript"' in xml
    assert 'w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"' in xml
    assert '<w:spacing w:val="10"/>' in xml
    assert '<w:position w:val="4"/>' in xml


def test_renderer_prefers_theme_font_over_literal_name_when_both_present(tmp_path):
    model = DocumentModel(
        source_sha256="abc",
        body=[Paragraph("p1", runs=[Run("r1", text="x", properties={
            "font_ascii_theme": "majorHAnsi",
        })])],
    )
    output = tmp_path / "out.docx"
    PureDocxRenderer().render(model, output, context=None)
    with ZipFile(output) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    assert 'w:asciiTheme="majorHAnsi"' in xml


def test_renderer_warns_when_preserved_part_cannot_be_safely_related(tmp_path):
    model = DocumentModel(source_sha256="abc")
    model.preserved_parts["word/embeddings/object.bin"] = PreservedPart(
        "word/embeddings/object.bin", "application/octet-stream", None, "deadbeef", b"x"
    )
    output = tmp_path / "out.docx"
    result = PureDocxRenderer().render(model, output, context=None)
    assert [warning.code for warning in result.warnings] == ["UNSUPPORTED_TRANSFER_PART"]


def test_renderer_roundtrips_headers_footers_and_notes(tmp_path):
    from word_replica.parser.parser import DocxParser
    from tests.fixtures.build_fixtures import build_extended_fixture

    source = build_extended_fixture(tmp_path / 'extended.docx')
    expected = DocxParser().parse(source)
    output = tmp_path / 'rebuilt.docx'
    PureDocxRenderer().render(expected, output, context=None)
    actual = DocxParser().parse(output)

    def text_map(collection):
        return {key: [block.text() for block in blocks if hasattr(block, 'text')] for key, blocks in collection.items()}

    assert text_map(actual.headers) == text_map(expected.headers)
    assert text_map(actual.footers) == text_map(expected.footers)
    assert text_map(actual.footnotes) == text_map(expected.footnotes)
    assert text_map(actual.endnotes) == text_map(expected.endnotes)


def test_fresh_shell_uses_current_reconstruction_timestamp(tmp_path):
    from datetime import datetime, timezone
    from word_replica.opc.package_reader import DocxPackage
    from word_replica.opc.properties import read_properties

    output = tmp_path / 'fresh.docx'
    before = datetime.now(timezone.utc)
    PureDocxRenderer().render(DocumentModel(source_sha256='abc'), output, context=None)
    with DocxPackage.open(output) as package:
        props = read_properties(package)
    created = datetime.fromisoformat(props.created.replace('Z', '+00:00'))
    modified = datetime.fromisoformat(props.modified.replace('Z', '+00:00'))
    assert created >= before
    assert modified >= before
    assert props.revision == '1'
    assert props.total_editing_time == '0'


def test_renderer_warns_when_image_bytes_exist_without_reconstructed_position(tmp_path):
    from word_replica.domain.model import BinaryAsset
    model = DocumentModel(source_sha256='abc')
    model.assets['image-sha'] = BinaryAsset(
        asset_id='image-sha',
        part_name='word/media/image1.png',
        content_type='image/png',
        sha256='image-sha',
        bytes_data=b'png-bytes',
    )
    result = PureDocxRenderer().render(model, tmp_path / 'with-image.docx', context=None)
    codes = [warning.code for warning in result.warnings]
    assert 'PURE_DOCX_ASSET_POSITION_UNAVAILABLE' in codes
    assert any(w.affects_status for w in result.warnings if w.code == 'PURE_DOCX_ASSET_POSITION_UNAVAILABLE')


def test_atomic_writer_closes_mkstemp_descriptor_before_unlink(monkeypatch, tmp_path):
    """Regression: Windows refuses to unlink a mkstemp path while its fd is open."""
    import os
    from pathlib import Path
    import word_replica.renderers.pure_docx as pure_docx

    original_mkstemp = pure_docx.tempfile.mkstemp
    original_unlink = Path.unlink
    tracked = {"fd": None, "path": None}

    def tracking_mkstemp(*args, **kwargs):
        fd, name = original_mkstemp(*args, **kwargs)
        tracked["fd"] = fd
        tracked["path"] = Path(name)
        return fd, name

    def windows_like_unlink(self, *args, **kwargs):
        if tracked["path"] is not None and self == tracked["path"]:
            try:
                os.fstat(tracked["fd"])
            except OSError:
                pass
            else:
                raise PermissionError("simulated Windows sharing violation: temp fd still open")
        return original_unlink(self, *args, **kwargs)

    monkeypatch.setattr(pure_docx.tempfile, "mkstemp", tracking_mkstemp)
    monkeypatch.setattr(Path, "unlink", windows_like_unlink)

    model = DocumentModel(
        source_sha256="abc",
        body=[Paragraph("p1", runs=[Run("r1", text="Hello")])],
    )
    output = tmp_path / "atomic.docx"
    PureDocxRenderer().render(model, output, context=None)
    assert output.exists()


def test_renderer_roundtrips_bookmarks_and_fields_without_adding_visible_paragraphs(tmp_path):
    from word_replica.domain.model import Bookmark, DocumentModel, Field, Paragraph, Run
    from word_replica.parser.parser import DocxParser
    from word_replica.renderers.pure_docx import PureDocxRenderer

    model = DocumentModel('abc')
    model.body = [Paragraph('p1', [Run('r1', 'Target paragraph')]), Paragraph('p2', [Run('r2', 'See target: ')])]
    model.bookmarks = [Bookmark('5', 'TargetBookmark', '/body/0', '/body/0')]
    model.fields = [Field('f1', 'REF TargetBookmark \\h', 'Target paragraph', False)]
    output = tmp_path / 'fields-bookmarks.docx'

    PureDocxRenderer().render(model, output, context=None)
    rebuilt = DocxParser().parse(output)

    assert rebuilt.plain_text() == model.plain_text()
    assert len(rebuilt.body) == len(model.body)
    assert [b.name for b in rebuilt.bookmarks] == ['TargetBookmark']
    assert len(rebuilt.fields) == 1
    assert 'REF TargetBookmark' in rebuilt.fields[0].instruction


def test_renderer_warns_that_field_codes_move_to_end_of_body(tmp_path):
    from word_replica.domain.model import DocumentModel, Field, Paragraph, Run
    from word_replica.renderers.pure_docx import PureDocxRenderer

    model = DocumentModel('abc')
    model.body = [Paragraph('p1', [Run('r1', 'Sadržaj')])]
    model.fields = [Field('f1', 'TOC \\o "1-3" \\h \\z \\u', '', False)]
    output = tmp_path / 'field-warning.docx'

    result = PureDocxRenderer().render(model, output, context=None)
    assert "PURE_DOCX_FIELD_POSITION_UNAVAILABLE" in [w.code for w in result.warnings]


def test_renderer_keeps_field_at_its_original_paragraph_when_content_tokens_are_present(tmp_path):
    from word_replica.domain.model import DocumentModel, Field, Paragraph, Run
    from word_replica.renderers.pure_docx import PureDocxRenderer

    model = DocumentModel('abc')
    model.body = [
        Paragraph('p1', [Run('r1', text='Naslov')]),
        Paragraph('toc', [
            Run('r2', properties={'content_tokens': [{'kind': 'field_begin'}]}),
            Run('r3', properties={'content_tokens': [{'kind': 'field_instruction', 'value': ' TOC \\o "1-3" \\h \\z \\u '}]}),
            Run('r4', properties={'content_tokens': [{'kind': 'field_separate'}]}),
            Run('r5', text='1. Uvod\t1'),
            Run('r6', properties={'content_tokens': [{'kind': 'field_end'}]}),
        ]),
        Paragraph('p3', [Run('r7', text='1. Uvod')]),
        Paragraph('p4', [Run('r8', text='Zaključak')]),
    ]
    model.fields = [Field('f1', 'TOC \\o "1-3" \\h \\z \\u', '1. Uvod\t1', False)]
    output = tmp_path / 'field-inline.docx'

    result = PureDocxRenderer().render(model, output, context=None)
    assert "PURE_DOCX_FIELD_POSITION_UNAVAILABLE" not in [w.code for w in result.warnings]

    with ZipFile(output) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    from lxml import etree
    root = etree.fromstring(xml.encode('utf-8'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    paragraphs = root.findall('.//w:body/w:p', ns)
    assert len(paragraphs) == 4
    # The field markers stay in the second paragraph, not glued onto the last one.
    assert paragraphs[1].find('.//w:fldChar', ns) is not None
    assert paragraphs[-1].find('.//w:fldChar', ns) is None
    assert 'Zaključak' in ''.join(t.text or '' for t in paragraphs[-1].findall('.//w:t', ns))


def test_renderer_registers_content_type_for_installed_png_asset(tmp_path):
    from lxml import etree
    from word_replica.domain.model import BinaryAsset

    model = DocumentModel(source_sha256='abc')
    model.assets['image-sha'] = BinaryAsset(
        asset_id='image-sha',
        part_name='word/media/image1.png',
        content_type='image/png',
        sha256='image-sha',
        bytes_data=b'png-bytes',
    )
    output = tmp_path / 'with-image.docx'
    PureDocxRenderer().render(model, output, context=None)

    with ZipFile(output) as archive:
        root = etree.fromstring(archive.read('[Content_Types].xml'))
    defaults = {
        node.get('Extension').lower(): node.get('ContentType')
        for node in root
        if node.tag.endswith('Default') and node.get('Extension')
    }
    overrides = {
        node.get('PartName').lstrip('/'): node.get('ContentType')
        for node in root
        if node.tag.endswith('Override') and node.get('PartName')
    }
    assert defaults.get('png') == 'image/png' or overrides.get('word/media/image1.png') == 'image/png'


def test_header_part_uses_ooxml_hdr_root_not_truncated_hea(tmp_path):
    from lxml import etree
    from word_replica.domain.model import Paragraph, Run
    from word_replica.renderers.pure_docx import W

    model = DocumentModel(source_sha256="abc")
    model.headers["word/header-source.xml"] = [Paragraph("hp1", runs=[Run("hr1", text="Header")])]
    output = tmp_path / "header.docx"
    PureDocxRenderer().render(model, output, context=None)
    with ZipFile(output) as z:
        root = etree.fromstring(z.read("word/header1.xml"))
    assert root.tag == f"{W}hdr"
