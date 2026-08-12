from word_replica.domain.enums import FidelityMode
from word_replica.parser.fidelity import project_fidelity
from tests.fixtures.build_fixtures import build_review_fixture
from word_replica.parser.parser import DocxParser


def test_clean_projection_omits_comments_and_deleted_revision_text(tmp_path):
    source = build_review_fixture(tmp_path / "review.docx")
    model = DocxParser().parse(source)
    clean = project_fidelity(model, FidelityMode.CLEAN)
    assert clean.extras["comments"] == {}
    assert "deleted text" not in clean.plain_text()
    assert "inserted text" in clean.plain_text()
    assert "hidden text" not in clean.plain_text()
    assert clean.extras["tracked_changes_enabled"] is False


def test_full_projection_retains_review_structures(tmp_path):
    source = build_review_fixture(tmp_path / "review.docx")
    full = project_fidelity(DocxParser().parse(source), FidelityMode.FULL)
    assert full.extras["comments"]
    assert full.extras["revisions"]
    assert full.extras["bookmarks"]
    assert full.extras["fields"]
    assert any(r.kind == "delete" and r.text == "deleted text" for r in full.extras["revisions"])
    assert "inserted text" in full.plain_text()
    assert "deleted text" not in full.plain_text()
