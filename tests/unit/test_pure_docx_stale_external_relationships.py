"""An external relationship nothing points at is still part of the package.

A document accumulates relationships that no element references any more --
editing removes the link text and leaves the relationship behind. For an
internal target the part travels and the relationship is put back with it. An
external target has no part at all, so nothing noticed the relationship was
gone.

Measured on 090716_Studentische_Arbeit_VWS.docx: two external hyperlink
relationships, rId26 and rId31, referenced by nothing in document.xml, both
absent from the rebuild. G10 reported the hyperlink relationships disappearing
entirely.

settings.xml already had exactly this fix for its own external relationships;
this is the same rule everywhere else. Relationships are keyed by type and
target, so one the rebuild recreated for a link that *is* referenced is not
written twice.
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

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")

ONE = "https://example.invalid/one"
TWO = "https://example.invalid/two"

STALE_RELS = (
    f'<Relationship Id="rId26" Type="{REL}/hyperlink" Target="{ONE}" TargetMode="External"/>'
    f'<Relationship Id="rId31" Type="{REL}/hyperlink" Target="{TWO}" TargetMode="External"/>'
)
PLAIN = "<w:p><w:r><w:t>the link text was deleted, the relationship was not</w:t></w:r></w:p>"
LINKED = (
    f'<w:p><w:hyperlink r:id="rId26"><w:r><w:t>still linked</w:t></w:r></w:hyperlink></w:p>'
)


def _write(path, *, doc_rels: str, body: str):
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{REL}">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">{doc_rels}</Relationships>'.encode(
                "utf-8"
            ),
        )
        archive.writestr("word/document.xml", document)
    return path


@pytest.fixture
def stale(tmp_path):
    return _write(tmp_path / "stale.docx", doc_rels=STALE_RELS, body=PLAIN)


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


def _external(path):
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
    return sorted(
        node.get("Target") for node in root if node.get("TargetMode") == "External"
    )


def test_both_stale_relationships_come_back(tmp_path, stale):
    assert _external(_rebuild(tmp_path, stale)) == [ONE, TWO]


def test_g10_sees_the_hyperlink_graph_unchanged(tmp_path, stale):
    before = g10_projection(stale)
    after = g10_projection(_rebuild(tmp_path, stale, "-g10"))

    assert after["relationship_graph"].get("hyperlink") == before["relationship_graph"].get(
        "hyperlink"
    )


def test_they_are_marked_external(tmp_path, stale):
    output = _rebuild(tmp_path, stale, "-mode")
    with ZipFile(output) as archive:
        root = etree.fromstring(archive.read("word/_rels/document.xml.rels"))

    links = [node for node in root if (node.get("Type") or "").endswith("/hyperlink")]
    assert links and all(node.get("TargetMode") == "External" for node in links)


# --- what must not be disturbed ------------------------------------------------

def test_a_referenced_link_is_not_written_twice(tmp_path):
    source = _write(tmp_path / "live.docx", doc_rels=STALE_RELS, body=LINKED)

    assert _external(_rebuild(tmp_path, source, "-live")) == [ONE, TWO]


def test_the_referenced_link_still_resolves(tmp_path):
    source = _write(tmp_path / "resolve.docx", doc_rels=STALE_RELS, body=LINKED)
    output = _rebuild(tmp_path, source, "-resolve")

    with ZipFile(output) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))

    used = {
        node.get(f"{{{REL}}}id")
        for node in document.findall(f".//{{{W}}}hyperlink")
        if node.get(f"{{{REL}}}id")
    }
    declared = {node.get("Id") for node in rels}
    assert used and used <= declared


def test_a_document_with_no_external_relationships_gains_none(tmp_path):
    source = _write(tmp_path / "none.docx", doc_rels="", body=PLAIN)

    assert _external(_rebuild(tmp_path, source, "-none")) == []
