from word_replica.parser.parser import DocxParser
from word_replica.renderers.pure_docx import PureDocxRenderer
from tests.fixtures.build_fixtures import build_extended_fixture


def test_pure_docx_round_trip_preserves_core_content(tmp_path):
    source = build_extended_fixture(tmp_path / "source.docx")
    source_model = DocxParser().parse(source)
    output = tmp_path / "out.docx"
    PureDocxRenderer().render(source_model, output, context=None)
    rebuilt = DocxParser().parse(output)
    assert rebuilt.plain_text() == source_model.plain_text()
    assert len(rebuilt.sections) == len(source_model.sections)
