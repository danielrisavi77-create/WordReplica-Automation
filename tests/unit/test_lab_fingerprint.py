"""Word-free structural fingerprinting.

The fingerprint is what lets the lab rank hundreds of thousands of documents
without opening any of them. It has to be cheap, total (never raise), and --
critically -- it has to see things the parser and Word 2010 are both blind to.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.lab.fingerprint import (
    FINGERPRINT_SCHEMA_VERSION,
    NAMESPACE_BITS,
    DocxFingerprint,
    extract_fingerprint,
    namespaces_from_mask,
)
from word_replica.lab.safety import RiskClass

W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

CT_XML = (
    b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    b'.wordprocessingml.document.main+xml"/>'
    b"</Types>"
)
RELS_XML = (
    b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    b'/officeDocument" Target="word/document.xml"/>'
    b"</Relationships>"
)
DOC_XML = (
    b'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b"<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
)
MINIMAL = {"[Content_Types].xml": CT_XML, "_rels/.rels": RELS_XML, "word/document.xml": DOC_XML}


def _docx(path, parts, *, extra=None):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, data in {**parts, **(extra or {})}.items():
            archive.writestr(name, data)
    return path


# --- the reason this module exists -------------------------------------------

def test_ns_mask_sees_a_w15_feature_that_the_document_model_cannot(tmp_path):
    """The false-pass proof, at fingerprint level.

    Word 2010 ignores w15 entirely and DocxParser only models what WordReplica
    reconstructs, so a dropped comment-reply is invisible to both. If the
    fingerprint could not see it either, the lab would have no way to even
    notice such a document is interesting.
    """
    from word_replica.parser.parser import DocxParser
    from word_replica.qa.golden_audit import build_model_gates

    plain = _docx(tmp_path / "plain.docx", MINIMAL)
    commentsex = (
        '<?xml version="1.0"?><w15:commentsEx xmlns:w15="{ns}">'
        '<w15:commentEx w15:paraId="1" w15:done="0"/>'
        "</w15:commentsEx>"
    ).format(ns=W15).encode("utf-8")
    with_w15 = _docx(tmp_path / "w15.docx", MINIMAL, extra={"word/commentsExtended.xml": commentsex})

    plain_fp = extract_fingerprint(plain)
    w15_fp = extract_fingerprint(with_w15)

    plain_model = DocxParser().parse(plain)
    w15_model = DocxParser().parse(with_w15)

    # The part is not modelled at all -- it is not even kept as a preserved part.
    assert plain_model.preserved_parts == {}
    assert w15_model.preserved_parts == {}
    # So every existing gate reports the two documents as equivalent...
    gates = build_model_gates(plain_model, w15_model)
    assert all(gate.passed for gate in gates.values()), {k: v.summary for k, v in gates.items()}
    # ...but the fingerprint can tell them apart.
    assert plain_fp.ns_mask != w15_fp.ns_mask
    assert "w15" in namespaces_from_mask(w15_fp.ns_mask)
    assert "w15" not in namespaces_from_mask(plain_fp.ns_mask)


def test_namespace_bits_are_append_only():
    # ns_mask values are persisted in the corpus database. Reordering these
    # would silently reinterpret every stored row.
    assert NAMESPACE_BITS[0] == "w"
    assert NAMESPACE_BITS.index("w14") == 1
    assert NAMESPACE_BITS.index("w15") == 2
    assert len(set(NAMESPACE_BITS)) == len(NAMESPACE_BITS)
    assert len(NAMESPACE_BITS) <= 64


# --- totality: the ingestor feeds this whatever the web returned --------------

def test_a_corrupt_package_yields_a_row_instead_of_raising(tmp_path):
    path = tmp_path / "broken.docx"
    path.write_bytes(b"PK\x03\x04 and then nothing that makes sense")

    fingerprint = extract_fingerprint(path)

    assert fingerprint.extractor_ok is False
    assert fingerprint.extract_error
    assert fingerprint.sha256  # identity is still recorded
    assert fingerprint.risk_class == RiskClass.HOSTILE


def test_an_empty_file_yields_a_row_instead_of_raising(tmp_path):
    path = tmp_path / "empty.docx"
    path.write_bytes(b"")

    assert extract_fingerprint(path).extractor_ok is False


def test_a_hostile_package_is_fingerprinted_without_a_model_parse(tmp_path):
    # HOSTILE documents stay useful -- they are scored and gate-checked -- but
    # nothing may hand them to a full parse or to Word.
    path = _docx(tmp_path / "m.docx", MINIMAL, extra={"word/vbaProject.bin": b"\x00"})

    fingerprint = extract_fingerprint(path)

    assert fingerprint.risk_class == RiskClass.HOSTILE
    assert fingerprint.word_open_allowed is False
    assert fingerprint.model_fingerprint is None
    assert fingerprint.part_count > 0


# --- identity and storage ----------------------------------------------------

def test_fingerprint_is_stable_for_the_same_bytes(tmp_path):
    path = _docx(tmp_path / "a.docx", MINIMAL)

    first, second = extract_fingerprint(path), extract_fingerprint(path)

    assert first.to_row() == second.to_row()


def test_row_is_flat_and_storable(tmp_path):
    row = extract_fingerprint(_docx(tmp_path / "a.docx", MINIMAL)).to_row()

    assert row["schema_version"] == FINGERPRINT_SCHEMA_VERSION
    for key, value in row.items():
        assert isinstance(value, (int, float, str, bytes, type(None))), (key, type(value))


def test_skipping_the_model_layer_still_produces_package_measurements(tmp_path):
    path = _docx(tmp_path / "a.docx", MINIMAL)

    cheap = extract_fingerprint(path, parse_model=False)

    assert cheap.model_fingerprint is None
    assert cheap.part_count == 3
    assert cheap.paragraph_count == 1
    assert cheap.extractor_ok is True


# --- counters over the real fixture corpus -----------------------------------

def test_plain_text_fixture_has_no_exotic_features(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "01_plain_text.docx")

    assert fingerprint.risk_class == RiskClass.VALID
    assert fingerprint.paragraph_count > 0
    assert fingerprint.table_count == 0
    assert fingerprint.drawing_inline_count == 0
    assert fingerprint.chart_count == 0
    assert fingerprint.comment_count == 0


def test_nested_table_fixture_reports_its_nesting_depth(corpus_dir):
    flat = extract_fingerprint(corpus_dir / "04_tables_merged.docx")
    nested = extract_fingerprint(corpus_dir / "15_nested_table.docx")

    assert flat.table_count >= 1
    assert flat.max_table_depth == 1
    assert nested.max_table_depth >= 2


def test_merged_cells_are_counted(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "04_tables_merged.docx")

    assert fingerprint.merged_cell_count >= 1


def test_review_fixture_reports_comments_and_revisions(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "10_comments_tracked_changes.docx")

    assert fingerprint.comment_count >= 1
    assert fingerprint.revision_count >= 1


def test_notes_fixture_reports_footnotes_and_endnotes(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "08_footnotes_endnotes.docx")

    assert fingerprint.footnote_count >= 1
    assert fingerprint.endnote_count >= 1


def test_sections_fixture_reports_more_than_one_section(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "06_sections_orientations.docx")

    assert fingerprint.section_count >= 2


def test_headers_fixture_reports_header_and_footer_parts(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "07_headers_footers_numbers.docx")

    assert fingerprint.header_count >= 1
    assert fingerprint.footer_count >= 1


def test_fields_fixture_reports_fields(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "09_toc_fields.docx")

    assert fingerprint.field_count >= 1


def test_bookmark_fixture_reports_bookmarks(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "11_bookmarks_crossrefs.docx")

    assert fingerprint.bookmark_count >= 1


def test_chart_fixture_reports_a_chart_part(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "12_charts_embedded.docx")

    assert fingerprint.chart_count >= 1


def test_image_fixture_distinguishes_inline_from_anchored(corpus_dir):
    fingerprint = extract_fingerprint(corpus_dir / "05_images_inline_floating.docx")

    assert fingerprint.media_count >= 1
    assert fingerprint.drawing_inline_count + fingerprint.drawing_anchor_count >= 1


@pytest.mark.parametrize(
    "name",
    [
        "01_plain_text.docx",
        "04_tables_merged.docx",
        "08_footnotes_endnotes.docx",
        "10_comments_tracked_changes.docx",
        "12_charts_embedded.docx",
        "13_academic_complex.docx",
        "15_nested_table.docx",
        "18_academic_citations.docx",
    ],
)
def test_every_fixture_fingerprints_cleanly(corpus_dir, name):
    fingerprint = extract_fingerprint(corpus_dir / name)

    assert fingerprint.extractor_ok is True
    assert fingerprint.risk_class in {RiskClass.VALID, RiskClass.RECOVERABLE}
    assert fingerprint.total_element_count > 0
    assert fingerprint.max_xml_depth > 0
    assert isinstance(fingerprint, DocxFingerprint)
