"""A part no relationship reaches is still in the package.

Three mechanisms now carry a source's parts across: an attachment reached by an
allow-listed document relationship, a part reached from inside the body, and a
part whose relationship the rebuild puts back because nothing referenced it.
All three start from a relationship.

A package can hold a part that no relationship reaches at all.
LibreOffice's style-pane-format-filter.docx has two -- word/webSettings.xml and
word/stylesWithEffects.xml -- present as parts and absent from
document.xml.rels. Nothing looked for them, so the rebuild dropped both and G10
reported application/xml going from three parts to one.

They are inert: Word does not read a part it cannot reach. That is the argument
for carrying them, not against it. The tool's claim is that it does not change
the package, and quietly tidying away bytes the source shipped is a change --
just one whose harmlessness we would be asserting rather than checking.

The parts the renderer authors itself are left alone, or its own document.xml
would be overwritten by whatever happened to be unreachable in the source.
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
DOCUMENT = (
    f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>'
    "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
).encode("utf-8")
WEB_SETTINGS = (
    f'<?xml version="1.0"?><w:webSettings xmlns:w="{W}"><w:optimizeForBrowser/></w:webSettings>'
).encode("utf-8")
STYLES_WITH_EFFECTS = (
    f'<?xml version="1.0"?><w:styles xmlns:w="{W}"><w:style w:styleId="Kept"/></w:styles>'
).encode("utf-8")


def _write(path, *, unreachable=True):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}"/>'.encode("utf-8"),
        )
        archive.writestr("word/document.xml", DOCUMENT)
        if unreachable:
            archive.writestr("word/webSettings.xml", WEB_SETTINGS)
            archive.writestr("word/stylesWithEffects.xml", STYLES_WITH_EFFECTS)
    return path


@pytest.fixture
def unreachable(tmp_path):
    return _write(tmp_path / "unreachable.docx")


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


def _names(path):
    with ZipFile(path) as archive:
        return set(archive.namelist())


def test_both_unreachable_parts_travel(tmp_path, unreachable):
    names = _names(_rebuild(tmp_path, unreachable))

    assert "word/webSettings.xml" in names
    assert "word/stylesWithEffects.xml" in names


def test_their_bytes_are_unchanged(tmp_path, unreachable):
    with ZipFile(_rebuild(tmp_path, unreachable, "-bytes")) as archive:
        assert archive.read("word/webSettings.xml") == WEB_SETTINGS


def test_g10_counts_the_same_parts(tmp_path, unreachable):
    before = g10_projection(unreachable)
    after = g10_projection(_rebuild(tmp_path, unreachable, "-g10"))

    assert after["part_kinds"].get("application/xml") == before["part_kinds"].get(
        "application/xml"
    )


def test_they_stay_unreachable_rather_than_gaining_a_relationship(tmp_path, unreachable):
    # The source did not point at them; inventing a relationship would be its
    # own change to the package.
    before = g10_projection(unreachable)
    after = g10_projection(_rebuild(tmp_path, unreachable, "-orphan"))

    assert after["orphan_parts"] == before["orphan_parts"]
    assert after["dangling_relationships"] == []


# --- what must not be disturbed ------------------------------------------------

def test_the_renderer_still_writes_its_own_document(tmp_path, unreachable):
    from word_replica.parser.parser import DocxParser

    output = _rebuild(tmp_path, unreachable, "-doc")

    assert DocxParser().parse(output).plain_text().strip() == "text"


def test_a_package_with_nothing_unreachable_gains_nothing(tmp_path):
    source = _write(tmp_path / "clean.docx", unreachable=False)

    names = _names(_rebuild(tmp_path, source, "-clean"))

    assert "word/webSettings.xml" not in names
    assert "word/stylesWithEffects.xml" not in names
