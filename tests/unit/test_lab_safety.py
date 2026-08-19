"""Hostile-input triage.

Every document in the corpus arrives from the public web, so the lab must
decide -- statically, before Word or a full parse is involved -- whether a
package is safe to work with at all.
"""
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import pytest
from lxml import etree

from word_replica.lab.safety import (
    DEFAULT_LIMITS,
    PackageLimits,
    RiskClass,
    SAFE_XML_PARSER_KWARGS,
    inspect_zip_envelope,
    safe_xml_parser,
    triage_document,
)

BILLION_LAUGHS = b"""<?xml version="1.0"?>
<!DOCTYPE lolz [
 <!ENTITY lol "lol">
 <!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
 <!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;">
 <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
]>
<lolz>&lol3;</lolz>"""

XXE = b"""<?xml version="1.0"?>
<!DOCTYPE r [<!ENTITY x SYSTEM "file:///c:/windows/win.ini">]>
<r>&x;</r>"""

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
    b"<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>"
)
MINIMAL = {"[Content_Types].xml": CT_XML, "_rels/.rels": RELS_XML, "word/document.xml": DOC_XML}

REL_HEAD = b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
REL_NS_TYPE = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def _docx(path, parts, *, extra=None):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, data in {**parts, **(extra or {})}.items():
            archive.writestr(name, data)
    return path


def _external_rel(rel_type: bytes, target: bytes) -> bytes:
    return (
        REL_HEAD
        + b'<Relationship Id="rId9" TargetMode="External" Type="'
        + REL_NS_TYPE
        + rel_type
        + b'" Target="'
        + target
        + b'"/></Relationships>'
    )


# --- hardened XML parsing ----------------------------------------------------

def test_safe_parser_does_not_expand_entity_bombs():
    # lxml does not error on an unresolvable entity, it leaves it as an Entity
    # node. That is the property that matters: the text never grows. The
    # default parser expands this same input to 3000 characters, and a few more
    # nesting levels would exhaust memory.
    safe = etree.fromstring(BILLION_LAUGHS, parser=safe_xml_parser())
    unsafe = etree.fromstring(BILLION_LAUGHS)

    assert "".join(safe.itertext()) == "&lol3;"
    assert len("".join(unsafe.itertext())) > 1000


def test_safe_parser_does_not_resolve_external_entities():
    safe = etree.fromstring(XXE, parser=safe_xml_parser())

    # The reference survives verbatim; no file was read.
    assert "".join(safe.itertext()) == "&x;"
    assert "[" not in "".join(safe.itertext())


def test_safe_parser_never_fetches_over_the_network():
    assert SAFE_XML_PARSER_KWARGS["no_network"] is True
    assert SAFE_XML_PARSER_KWARGS["resolve_entities"] is False
    assert SAFE_XML_PARSER_KWARGS["huge_tree"] is False


def test_safe_parser_still_reads_ordinary_wordprocessingml():
    root = etree.fromstring(DOC_XML, parser=safe_xml_parser())
    assert root.tag.endswith("}document")


# --- zip envelope, read from the central directory only ----------------------

def test_zip_bomb_is_rejected_without_extracting_anything(tmp_path):
    # A declared 8 MB member in a tiny file. The verdict must come from the
    # central directory: extracting to find out is the failure mode itself.
    path = tmp_path / "bomb.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", b"\0" * (8 * 1024 * 1024))

    report = inspect_zip_envelope(path, PackageLimits(max_total_uncompressed_bytes=1024))

    assert report.ok is False
    assert any("uncompressed" in reason for reason in report.reasons)
    assert path.stat().st_size < 1024 * 1024


def test_compression_ratio_cap_is_enforced(tmp_path):
    path = tmp_path / "ratio.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", b"\0" * (4 * 1024 * 1024))

    report = inspect_zip_envelope(path, PackageLimits(max_compression_ratio=2.0))

    assert report.ok is False
    assert any("ratio" in reason for reason in report.reasons)


def test_oversized_single_part_is_rejected(tmp_path):
    path = tmp_path / "part.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", b"\0" * (2 * 1024 * 1024))

    report = inspect_zip_envelope(path, PackageLimits(max_part_bytes=1024))

    assert report.ok is False
    assert any("part" in reason for reason in report.reasons)


def test_member_explosion_is_rejected(tmp_path):
    path = tmp_path / "many.docx"
    with ZipFile(path, "w") as archive:
        for index in range(50):
            archive.writestr("word/p{}.xml".format(index), b"<x/>")

    report = inspect_zip_envelope(path, PackageLimits(max_member_count=10))

    assert report.ok is False
    assert any("members" in reason for reason in report.reasons)


@pytest.mark.parametrize(
    "name",
    ["../../evil.xml", "/abs/evil.xml", "C:/evil.xml", "word/../../evil.xml"],
)
def test_path_traversal_member_names_are_rejected(tmp_path, name):
    path = tmp_path / "traversal.docx"
    with ZipFile(path, "w") as archive:
        archive.writestr(ZipInfo(name), b"<x/>")

    report = inspect_zip_envelope(path, DEFAULT_LIMITS)

    assert report.ok is False
    assert any("traversal" in reason for reason in report.reasons)


def test_ordinary_package_passes_the_envelope_check(tmp_path):
    report = inspect_zip_envelope(_docx(tmp_path / "ok.docx", MINIMAL), DEFAULT_LIMITS)

    assert report.ok is True
    assert report.reasons == ()
    assert report.member_count == 3


def test_envelope_records_measurements_even_when_it_passes(tmp_path):
    report = inspect_zip_envelope(_docx(tmp_path / "ok.docx", MINIMAL), DEFAULT_LIMITS)

    assert report.total_uncompressed_bytes == sum(len(v) for v in MINIMAL.values())
    assert report.max_part_bytes == max(len(v) for v in MINIMAL.values())
    assert report.compression_ratio > 0


# --- triage ------------------------------------------------------------------

def test_ordinary_document_is_valid(tmp_path):
    triage = triage_document(_docx(tmp_path / "ok.docx", MINIMAL))

    assert triage.risk_class is RiskClass.VALID
    assert triage.word_open_allowed is True
    assert triage.reasons == ()


def test_macro_project_is_hostile_and_never_opened_in_word(tmp_path):
    path = _docx(tmp_path / "m.docx", MINIMAL, extra={"word/vbaProject.bin": b"\x00\x01"})

    triage = triage_document(path)

    assert triage.risk_class is RiskClass.HOSTILE
    assert triage.word_open_allowed is False
    assert any("macro" in reason for reason in triage.reasons)


def test_activex_control_is_hostile(tmp_path):
    path = _docx(tmp_path / "a.docx", MINIMAL, extra={"word/activeX/activeX1.xml": b"<x/>"})

    assert triage_document(path).risk_class is RiskClass.HOSTILE


def test_embedded_executable_is_hostile(tmp_path):
    path = _docx(tmp_path / "x.docx", MINIMAL, extra={"word/embeddings/oleObject1.bin": b"MZ\x90\x00" + b"\x00" * 64})

    triage = triage_document(path)

    assert triage.risk_class is RiskClass.HOSTILE
    assert any("executable" in reason for reason in triage.reasons)


def test_entity_declaration_anywhere_is_hostile(tmp_path):
    path = _docx(tmp_path / "e.docx", {**MINIMAL, "word/document.xml": BILLION_LAUGHS})

    triage = triage_document(path)

    assert triage.risk_class is RiskClass.HOSTILE
    assert any("entity" in reason.lower() or "doctype" in reason.lower() for reason in triage.reasons)


def test_external_unc_relationship_target_is_hostile(tmp_path):
    rels = _external_rel(b"oleObject", b"\\\\attacker\\share\\payload.dll")
    path = _docx(tmp_path / "unc.docx", MINIMAL, extra={"word/_rels/document.xml.rels": rels})

    assert triage_document(path).risk_class is RiskClass.HOSTILE


def test_external_file_scheme_relationship_target_is_hostile(tmp_path):
    rels = _external_rel(b"attachedTemplate", b"file:///c:/evil.dotm")
    path = _docx(tmp_path / "tpl.docx", MINIMAL, extra={"word/_rels/document.xml.rels": rels})

    assert triage_document(path).risk_class is RiskClass.HOSTILE


def test_ordinary_external_hyperlink_is_not_hostile(tmp_path):
    # Hyperlinks are ubiquitous; treating them as hostile would quarantine a
    # large fraction of the corpus for no gain.
    rels = _external_rel(b"hyperlink", b"https://example.com/")
    path = _docx(tmp_path / "link.docx", MINIMAL, extra={"word/_rels/document.xml.rels": rels})

    assert triage_document(path).risk_class is RiskClass.VALID


def test_missing_content_type_override_is_recoverable_not_hostile(tmp_path):
    # A real defect the product must label -- not a threat, and not a clean file.
    broken = dict(MINIMAL)
    broken["[Content_Types].xml"] = (
        b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b"</Types>"
    )

    triage = triage_document(_docx(tmp_path / "b.docx", broken))

    assert triage.risk_class is RiskClass.RECOVERABLE
    assert triage.word_open_allowed is True


def test_dangling_relationship_target_is_recoverable(tmp_path):
    rels = (
        REL_HEAD
        + b'<Relationship Id="rId1" Type="'
        + REL_NS_TYPE
        + b'officeDocument" Target="word/missing.xml"/></Relationships>'
    )
    path = _docx(tmp_path / "d.docx", {**MINIMAL, "_rels/.rels": rels})

    assert triage_document(path).risk_class is RiskClass.RECOVERABLE


def test_a_file_that_is_not_a_zip_is_hostile(tmp_path):
    path = tmp_path / "not.docx"
    path.write_bytes(b"this is not a zip archive")

    triage = triage_document(path)

    assert triage.risk_class is RiskClass.HOSTILE
    assert triage.word_open_allowed is False


def test_hostility_wins_over_recoverability(tmp_path):
    # A package that is both malformed and macro-enabled must be quarantined,
    # not merely flagged.
    broken = dict(MINIMAL)
    broken["[Content_Types].xml"] = (
        b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
    )
    path = _docx(tmp_path / "both.docx", broken, extra={"word/vbaProject.bin": b"\x00"})

    assert triage_document(path).risk_class is RiskClass.HOSTILE


def test_triage_never_raises_on_arbitrary_bytes(tmp_path):
    # The ingestor feeds this whatever the web returned; a crash here stops a
    # 20,000-document batch.
    path = tmp_path / "junk.docx"
    path.write_bytes(bytes(range(256)) * 40)

    triage = triage_document(path)

    assert triage.risk_class is RiskClass.HOSTILE
