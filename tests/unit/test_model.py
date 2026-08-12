from word_replica.domain.model import DocumentModel, ElementIdFactory, Paragraph, Run


def test_element_ids_are_deterministic_for_source_positions():
    ids = ElementIdFactory("abc123")
    assert ids.make("paragraph", "body/0") == ids.make("paragraph", "body/0")
    assert ids.make("paragraph", "body/0") != ids.make("paragraph", "body/1")


def test_document_plain_text_preserves_order():
    doc = DocumentModel(
        source_sha256="abc123",
        body=[
            Paragraph("p1", runs=[Run("r1", text="Hello")]),
            Paragraph("p2", runs=[Run("r2", text=" world")]),
        ],
    )
    assert doc.plain_text() == "Hello\n world"


def test_equivalent_models_have_same_fingerprint():
    a = DocumentModel(source_sha256="x", body=[Paragraph("p", runs=[Run("r", text="A")])])
    b = DocumentModel(source_sha256="x", body=[Paragraph("p", runs=[Run("r", text="A")])])
    assert a.fingerprint() == b.fingerprint()
