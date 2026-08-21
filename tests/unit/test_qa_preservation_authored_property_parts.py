"""Authored document properties are recognised by what they are, not where.

`application_properties_rewritten_by_policy` declares that the renderer authors
`docProps/app.xml` and `docProps/core.xml` rather than copying them -- page and
word counts describing the document that now exists, and the truthful-lifecycle
stamp. The carve-out named those two paths literally.

Real packages carry more than those two. LibreOffice writes a *second*
core-properties part, `docProps/core0.xml`, hung off `_rels/.rels` by a
relationship whose type URI is misspelled -- `.../officedocument/2006/...`
where the standard says `.../package/2006/...`. Its content duplicates
core.xml. The rebuild does not carry it, and because the carve-out did not
recognise it, the gate reported the loss of core-properties, dc and dcterms
namespaces on documents whose metadata policy was working exactly as designed.

That was the largest class of G10 failures in the corpus: 4 of 23.

Recognising the part by content type rather than by path keeps the rule as
narrow as it was. A part that is not a property part is still compared, so a
rebuild that quietly dropped real content still fails, and the carve-out only
applies when the caller declares the policy.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.qa.preservation import build_preservation_gate

PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CORE_TYPE = "application/vnd.openxmlformats-package.core-properties+xml"
EXTENDED_TYPE = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

CORE = (
    '<?xml version="1.0"?><cp:coreProperties '
    'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:dcterms="http://purl.org/dc/terms/">'
    "<dc:creator>someone</dc:creator><dcterms:created>2025-06-06T09:00:00Z</dcterms:created>"
    "</cp:coreProperties>"
).encode("utf-8")
APP = (
    '<?xml version="1.0"?><Properties '
    'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    "<Application>Writer</Application></Properties>"
).encode("utf-8")
DOCUMENT = (
    f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>'
    "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
).encode("utf-8")
NOTE = b'<?xml version="1.0"?><note xmlns="urn:example:note">kept</note>'


def _content_types(*, second_core: bool, note: bool) -> bytes:
    overrides = (
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.document.main+xml"/>'
        f'<Override PartName="/docProps/core.xml" ContentType="{CORE_TYPE}"/>'
        f'<Override PartName="/docProps/app.xml" ContentType="{EXTENDED_TYPE}"/>'
    )
    if second_core:
        overrides += f'<Override PartName="/docProps/core0.xml" ContentType="{CORE_TYPE}"/>'
    if note:
        overrides += '<Override PartName="/extras/note.xml" ContentType="application/xml"/>'
    return (
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{overrides}</Types>"
    ).encode("utf-8")


def _root_rels(*, second_core: bool, note: bool) -> bytes:
    entries = (
        f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
        f'<Relationship Id="rId2" Target="docProps/core.xml" '
        f'Type="{PKG_REL.replace("/relationships", "")}/relationships/metadata/core-properties"/>'
        f'<Relationship Id="rId3" Target="docProps/app.xml" Type="{REL}/extended-properties"/>'
    )
    if second_core:
        # The misspelled type LibreOffice actually writes.
        entries += (
            '<Relationship Id="rId4" Target="docProps/core0.xml" '
            'Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/'
            'metadata/core-properties"/>'
        )
    if note:
        entries += f'<Relationship Id="rId5" Target="extras/note.xml" Type="urn:example:note"/>'
    return (
        f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">{entries}</Relationships>'
    ).encode("utf-8")


def _package(path, *, second_core: bool = False, note: bool = False):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(second_core=second_core, note=note))
        archive.writestr("_rels/.rels", _root_rels(second_core=second_core, note=note))
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}"/>'.encode("utf-8"),
        )
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("docProps/core.xml", CORE)
        archive.writestr("docProps/app.xml", APP)
        if second_core:
            archive.writestr("docProps/core0.xml", CORE)
        if note:
            archive.writestr("extras/note.xml", NOTE)
    return path


@pytest.fixture
def source(tmp_path):
    return _package(tmp_path / "source.docx", second_core=True)


@pytest.fixture
def rebuilt(tmp_path):
    """What the renderer produces: the properties it authors, and no second one."""
    return _package(tmp_path / "rebuilt.docx")


def test_a_second_core_properties_part_is_covered_by_the_policy(source, rebuilt):
    gate = build_preservation_gate(
        source, rebuilt, application_properties_rewritten_by_policy=True
    )

    assert gate.passed is True, gate.first_divergence


def test_without_the_declaration_it_is_still_reported(source, rebuilt):
    # The carve-out only applies when the caller declares the policy.
    gate = build_preservation_gate(source, rebuilt)

    assert gate.passed is False


def test_a_part_that_is_not_a_property_part_is_still_compared(tmp_path):
    source = _package(tmp_path / "with-note.docx", second_core=True, note=True)
    rebuilt = _package(tmp_path / "without-note.docx", second_core=True)

    gate = build_preservation_gate(
        source, rebuilt, application_properties_rewritten_by_policy=True
    )

    assert gate.passed is False, "dropping real content must still fail"


def test_an_unchanged_package_still_passes(tmp_path):
    source = _package(tmp_path / "a.docx", second_core=True, note=True)
    same = _package(tmp_path / "b.docx", second_core=True, note=True)

    gate = build_preservation_gate(
        source, same, application_properties_rewritten_by_policy=True
    )

    assert gate.passed is True, gate.first_divergence


def test_an_invented_second_core_properties_part_is_covered_too(source, rebuilt):
    # The policy is symmetric for authored properties: the renderer writes them
    # whether or not the source had them, so gaining one is as intended as
    # losing one.
    gate = build_preservation_gate(
        rebuilt, source, application_properties_rewritten_by_policy=True
    )

    assert gate.passed is True, gate.first_divergence
