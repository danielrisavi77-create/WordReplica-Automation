import os
import pytest

from tests.fixtures.build_fixtures import build_core_fixture

pytestmark = pytest.mark.word


@pytest.mark.skipif(os.environ.get("WORD_REPLICA_WORD_TESTS") != "1", reason="Word integration disabled")
def test_word_renderer_rebuilds_core_fixture(tmp_path):
    from word_replica.parser.parser import DocxParser
    from word_replica.renderers.word_com import WordComRenderer

    source = build_core_fixture(tmp_path / "source.docx")
    model = DocxParser().parse(source)
    model.extras["metadata_policy"] = {}
    output = tmp_path / "rebuilt.docx"
    renderer = WordComRenderer(visible=False)
    renderer.render(model, output, context=None)
    rebuilt = DocxParser().parse(output)
    assert rebuilt.plain_text() == model.plain_text()
    assert len(rebuilt.sections) == len(model.sections)
