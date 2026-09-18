"""Keyboard and toolbar customisations are part of the document.

`word/customizations.xml` holds the key map and toolbar customisations saved
with a document -- `wne:tcg`, `wne:keymaps`, `wne:keymap`. It hangs off
document.xml.rels by a relationship of type `keyMapCustomizations`, which was
missing from the attachment allow-list, so the part was dropped and G10 reported
the office/word/2006/wordml namespace disappearing.

Two documents in the corpus have one, and one of them
(bnc837302.docx) also carries `word/_rels/customizations.xml.rels`, so the
sidecar has to travel with it or the part arrives with a reference to nothing.

Nothing about this is exotic: it is a part with a relationship, and the reason
it was lost is that the list of relationship types worth keeping is a list, and
this type was not on it.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
)
from word_replica.parser.parser import DocxParser
from word_replica.qa.preservation import g10_projection
from word_replica.services.rebuild import RebuildService

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WNE = "http://schemas.microsoft.com/office/word/2006/wordml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
KEYMAP = "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations"
CUSTOMIZATIONS_TYPE = (
    "application/vnd.ms-word.keyMapCustomizations+xml"
)

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    f'<Override PartName="/word/customizations.xml" ContentType="{CUSTOMIZATIONS_TYPE}"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")
DOCUMENT = (
    f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>'
    "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
).encode("utf-8")
CUSTOMIZATIONS = (
    f'<?xml version="1.0"?><wne:tcg xmlns:wne="{WNE}" xmlns:w="{W}">'
    '<wne:keymaps><wne:keymap wne:kcmPrimary="0141"/></wne:keymaps>'
    "</wne:tcg>"
).encode("utf-8")


def _write(path, *, customisations=True):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        rels = (
            f'<Relationship Id="rId1" Type="{KEYMAP}" Target="customizations.xml"/>'
            if customisations
            else ""
        )
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">{rels}</Relationships>'.encode(
                "utf-8"
            ),
        )
        archive.writestr("word/document.xml", DOCUMENT)
        if customisations:
            archive.writestr("word/customizations.xml", CUSTOMIZATIONS)
    return path


@pytest.fixture
def with_customisations(tmp_path):
    return _write(tmp_path / "keymap.docx")


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


def test_the_part_is_preserved_by_the_parser(with_customisations):
    model = DocxParser().parse(with_customisations)

    assert "word/customizations.xml" in model.preserved_parts


def test_the_part_travels_byte_for_byte(tmp_path, with_customisations):
    with ZipFile(_rebuild(tmp_path, with_customisations)) as archive:
        assert archive.read("word/customizations.xml") == CUSTOMIZATIONS


def test_the_relationship_comes_back(tmp_path, with_customisations):
    before = g10_projection(with_customisations)
    after = g10_projection(_rebuild(tmp_path, with_customisations, "-rel"))

    assert after["relationship_graph"].get("keyMapCustomizations") == before[
        "relationship_graph"
    ].get("keyMapCustomizations")


def test_g10_keeps_the_namespace(tmp_path, with_customisations):
    before = g10_projection(with_customisations)
    after = g10_projection(_rebuild(tmp_path, with_customisations, "-ns"))

    assert WNE in after["namespaces"]
    assert set(before["namespaces"]) - set(after["namespaces"]) == set()


def test_it_is_not_left_orphaned(tmp_path, with_customisations):
    projection = g10_projection(_rebuild(tmp_path, with_customisations, "-orphan"))

    assert projection["orphan_parts"] == []
    assert projection["dangling_relationships"] == []


def test_a_document_without_customisations_gains_none(tmp_path):
    source = _write(tmp_path / "plain.docx", customisations=False)

    with ZipFile(_rebuild(tmp_path, source, "-plain")) as archive:
        assert "word/customizations.xml" not in archive.namelist()
