from zipfile import ZipFile

from lxml import etree
import pytest

from tests.fixtures.build_fixtures import build_review_fixture, build_academic_complex
from word_replica.domain.enums import FidelityMode
from word_replica.parser.fidelity import project_fidelity
from word_replica.parser.parser import DocxParser
from word_replica.renderers.pure_docx import PureDocxRenderer


@pytest.mark.parametrize("builder", [build_review_fixture, build_academic_complex])
@pytest.mark.parametrize("mode", [FidelityMode.CLEAN, FidelityMode.FULL])
def test_comment_references_match_retained_comments(tmp_path, builder, mode):
    source = builder(tmp_path / "source.docx")
    model = DocxParser().parse(source)
    fingerprint = model.fingerprint()
    projected = project_fidelity(model, mode)
    output = tmp_path / "output.docx"
    PureDocxRenderer().render(projected, output, None)

    assert model.fingerprint() == fingerprint
    with ZipFile(output) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        references = document.xpath("//w:commentReference/@w:id", namespaces=ns)
        if mode is FidelityMode.CLEAN:
            assert "word/comments.xml" not in archive.namelist()
            assert references == [], "CLEAN document references removed comments"
        else:
            comments = etree.fromstring(archive.read("word/comments.xml"))
            ids = comments.xpath("/w:comments/w:comment/@w:id", namespaces=ns)
            assert references
            assert set(references) <= set(ids)
