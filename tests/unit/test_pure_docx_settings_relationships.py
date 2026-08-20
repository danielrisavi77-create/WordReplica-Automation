"""The document's attached template must survive a rebuild.

Found by G10 on the corpus: two documents lost an `attachedTemplate`
relationship, which lives in `word/_rels/settings.xml.rels` and points at the
template the document was authored from -- usually the author's Normal.dotm.
The renderer writes `settings.xml` from the model but never wrote its .rels, so
the relationship simply disappeared.

Only *external* relationships are carried. An external target needs no part in
the package, so restoring it cannot leave anything dangling. An internal one
would need its part to travel too, and inventing a reference to a part that is
not there is the failure this whole line of work exists to avoid.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from lxml import etree

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
)
from word_replica.parser.parser import DocxParser
from word_replica.services.rebuild import RebuildService

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
ATTACHED_TEMPLATE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
)

CONTENT_TYPES = (
    b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    b'.wordprocessingml.document.main+xml"/>'
    b'<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument'
    b'.wordprocessingml.settings+xml"/>'
    b"</Types>"
)
ROOT_RELS = (
    b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Target="word/document.xml" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
    b"</Relationships>"
)
DOC_RELS = (
    b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId2" Target="settings.xml" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"/>'
    b"</Relationships>"
)
DOCUMENT = (
    b'<?xml version="1.0"?><w:document '
    b'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b"<w:body><w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
)
SETTINGS = (
    b'<?xml version="1.0"?><w:settings '
    b'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)


def _source(path, *, settings_rels: bytes | None):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/settings.xml", SETTINGS)
        if settings_rels is not None:
            archive.writestr("word/_rels/settings.xml.rels", settings_rels)
    return path


def _external(rel_type: str, target: str) -> bytes:
    return (
        '<?xml version="1.0"?><Relationships '
        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId9" TargetMode="External" Type="{rel_type}" Target="{target}"/>'
        "</Relationships>"
    ).encode("utf-8")


def _rebuild(tmp_path, source, suffix=""):
    result = RebuildService(app_root=tmp_path / f"app{suffix}").rebuild(
        source,
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )
    assert result.output_path is not None, result.reasons
    return result.output_path


def _settings_relationships(path):
    with ZipFile(path) as archive:
        if "word/_rels/settings.xml.rels" not in archive.namelist():
            return []
        root = etree.fromstring(archive.read("word/_rels/settings.xml.rels"))
    return [(node.get("Type"), node.get("Target"), node.get("TargetMode")) for node in root]


def test_the_parser_records_external_settings_relationships(tmp_path):
    source = _source(tmp_path / "src.docx", settings_rels=_external(ATTACHED_TEMPLATE, "Normal.dotm"))

    model = DocxParser().parse(source)

    assert model.extras["settings_relationships"] == [(ATTACHED_TEMPLATE, "Normal.dotm")]


def test_the_attached_template_survives_a_rebuild(tmp_path):
    source = _source(tmp_path / "src.docx", settings_rels=_external(ATTACHED_TEMPLATE, "Normal.dotm"))

    relationships = _settings_relationships(_rebuild(tmp_path, source))

    assert (ATTACHED_TEMPLATE, "Normal.dotm", "External") in relationships


def test_a_document_without_one_does_not_gain_one(tmp_path):
    source = _source(tmp_path / "bare.docx", settings_rels=None)

    assert _settings_relationships(_rebuild(tmp_path, source, "-bare")) == []


def test_an_internal_settings_relationship_is_not_carried(tmp_path):
    # An internal target needs its part to travel too. Restoring the reference
    # without the part is the exact failure this work exists to avoid.
    internal = (
        '<?xml version="1.0"?><Relationships '
        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId9" Target="somePart.xml" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"/>'
        "</Relationships>"
    ).encode("utf-8")
    source = _source(tmp_path / "internal.docx", settings_rels=internal)

    model = DocxParser().parse(source)

    assert model.extras["settings_relationships"] == []


def test_the_rebuild_has_no_dangling_relationships(tmp_path):
    from word_replica.qa.preservation import g10_projection

    source = _source(tmp_path / "src.docx", settings_rels=_external(ATTACHED_TEMPLATE, "Normal.dotm"))

    assert g10_projection(_rebuild(tmp_path, source, "-dangle"))["dangling_relationships"] == []
