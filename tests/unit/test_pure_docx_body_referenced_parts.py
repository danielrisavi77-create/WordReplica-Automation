"""Charts, diagrams and embedded objects can come back now.

They were captured but never restored, and the reason was sound at the time:
they are reached from inside the document body, and a renderer that rebuilds
the body cannot recreate the reference. Writing the part alone would have left
an orphan, and G10 would then have matched on a document that was actually
broken -- a false signal being worse than a reported loss.

Verbatim fragment preservation removed that constraint. The reference now
survives in the body with its relationship id remapped, so the part it points
at can be installed and the two ends meet.

What the placeholder has to carry is the *type* of the relationship. An OLE
object reached through an image relationship is not the same document: Word
uses the type to decide what a reference is for.
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
from word_replica.qa.preservation import g10_projection
from word_replica.services.rebuild import RebuildService

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Target="word/document.xml" '
    f'Type="{REL}/officeDocument"/></Relationships>'
).encode("utf-8")
DOC_RELS = (
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId7" Target="embeddings/oleObject1.bin" Type="{REL}/oleObject"/>'
    "</Relationships>"
).encode("utf-8")
DOCUMENT = (
    '<?xml version="1.0"?><w:document '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    f'xmlns:r="{REL}">'
    "<w:body><w:p><w:r>"
    '<w:object><o:OLEObject Type="Embed" ProgID="Equation.3" r:id="rId7"/></w:object>'
    "</w:r></w:p><w:sectPr/></w:body></w:document>"
).encode("utf-8")
OLE_BYTES = b"\xd0\xcf\x11\xe0 equation payload"


@pytest.fixture
def source(tmp_path):
    path = tmp_path / "ole.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/embeddings/oleObject1.bin", OLE_BYTES)
    return path


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


def test_the_embedded_object_bytes_survive(tmp_path, source):
    with ZipFile(_rebuild(tmp_path, source)) as archive:
        assert archive.read("word/embeddings/oleObject1.bin") == OLE_BYTES


def test_the_body_still_refers_to_it(tmp_path, source):
    with ZipFile(_rebuild(tmp_path, source, "-ref")) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))

    referenced = {
        node.get(f"{{{REL}}}id")
        for node in document.iter()
        if isinstance(node.tag, str) and node.get(f"{{{REL}}}id")
    }
    declared = {node.get("Id") for node in rels}
    assert referenced, "the OLE reference did not survive"
    assert referenced <= declared


def test_the_relationship_keeps_its_own_type(tmp_path, source):
    # An OLE object reached through an image relationship is not the same
    # document: Word uses the type to decide what a reference is for.
    with ZipFile(_rebuild(tmp_path, source, "-type")) as archive:
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))

    types = {node.get("Type") for node in rels if "embeddings" in (node.get("Target") or "")}
    assert types == {f"{REL}/oleObject"}


def test_nothing_is_orphaned_or_dangling(tmp_path, source):
    projection = g10_projection(_rebuild(tmp_path, source, "-clean"))

    assert projection["dangling_relationships"] == []
    assert projection["orphan_parts"] == []


def test_g10_sees_the_object_preserved(tmp_path, source):
    from word_replica.qa.preservation import build_preservation_gate

    gate = build_preservation_gate(
        source,
        _rebuild(tmp_path, source, "-gate"),
        custom_properties_dropped_by_policy=True,
        application_properties_rewritten_by_policy=True,
    )

    assert gate.passed is True, f"{gate.summary} {gate.first_divergence}"


def test_a_document_without_one_gains_nothing(tmp_path):
    path = tmp_path / "plain.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            b'<?xml version="1.0"?><Relationships '
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        )
        archive.writestr("word/document.xml", (
            '<?xml version="1.0"?><w:document '
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
        ).encode("utf-8"))

    with ZipFile(_rebuild(tmp_path, path, "-plain")) as archive:
        assert not [n for n in archive.namelist() if n.startswith("word/embeddings/")]
