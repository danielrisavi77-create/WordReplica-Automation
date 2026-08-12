import os
import pytest

pytestmark = pytest.mark.word


@pytest.mark.skipif(
    os.environ.get("WORD_REPLICA_WORD_TESTS") != "1",
    reason="set WORD_REPLICA_WORD_TESTS=1 on Windows with Word",
)
def test_word_session_creates_and_closes_document(tmp_path):
    from word_replica.renderers.word_com import WordSession

    with WordSession(visible=False) as session:
        doc = session.new_document()
        path = tmp_path / "smoke.docx"
        doc.SaveAs2(str(path))
        doc.Close(False)
        session.release_document(doc)
    assert path.exists()
