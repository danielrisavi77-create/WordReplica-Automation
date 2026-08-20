"""The pure-docx renderer must not smuggle its template into the output.

Found by G10 on the LibreOffice corpus: reconstructions that every model gate
(G0-G7) called perfect were carrying a customXml datastore, a thumbnail and a
stylesWithEffects part that the source document never had.

`PureDocxRenderer.render` starts from `Document().save(shell)` -- python-docx's
default template -- and mutates that package. The template ships
`customXml/item1.xml`, `customXml/itemProps1.xml`, `docProps/thumbnail.jpeg`
and `word/stylesWithEffects.xml`, and no stage removed them.

The contract is symmetric, and both halves are tested here: **do not gain what
the source lacked, and do not lose what it had.** The first version of these
tests asserted only the first half, against a fixture that was itself built
from the python-docx template -- so "template debris" and "source content" were
indistinguishable in it. A hand-built package is used instead.
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
from word_replica.services.rebuild import RebuildService

CONTENT_TYPES = (
    b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    b'.wordprocessingml.document.main+xml"/>'
    b"</Types>"
)
ROOT_RELS = (
    b'<?xml version="1.0"?><Relationships '
    b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Target="word/document.xml" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
    b"</Relationships>"
)
DOC_RELS = (
    b'<?xml version="1.0"?><Relationships '
    b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)
DOCUMENT = (
    b'<?xml version="1.0"?><w:document '
    b'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b"<w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
)


@pytest.fixture
def bare_source(tmp_path):
    """A package with none of the parts the shell template happens to carry."""
    path = tmp_path / "bare.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
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
    with ZipFile(result.output_path) as archive:
        return set(archive.namelist())


# --- do not gain what the source lacked --------------------------------------

def test_the_template_custom_xml_store_is_not_added(tmp_path, bare_source):
    names = _rebuild(tmp_path, bare_source)

    smuggled = sorted(name for name in names if name.lower().startswith("customxml/"))
    assert smuggled == [], f"template customXml leaked into the rebuild: {smuggled}"


def test_the_template_thumbnail_is_not_added(tmp_path, bare_source):
    assert "docProps/thumbnail.jpeg" not in _rebuild(tmp_path, bare_source)


def test_the_template_styles_with_effects_part_is_not_added(tmp_path, bare_source):
    assert "word/stylesWithEffects.xml" not in _rebuild(tmp_path, bare_source)


def test_no_relationship_points_at_a_part_that_is_gone(tmp_path, bare_source):
    # Removing a part without removing what references it produces a package
    # Word repairs on open -- a worse defect than the one being fixed.
    from word_replica.qa.preservation import g10_projection

    result = RebuildService(app_root=tmp_path / "app-rels").rebuild(
        bare_source,
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )

    assert g10_projection(result.output_path)["dangling_relationships"] == []


def test_the_rebuild_still_produces_a_readable_package(tmp_path, bare_source):
    names = _rebuild(tmp_path, bare_source)

    assert "word/document.xml" in names
    assert "[Content_Types].xml" in names
    assert "_rels/.rels" in names


# --- and do not lose what it had ---------------------------------------------

def test_a_source_that_has_these_parts_keeps_them(tmp_path, corpus_dir):
    # The other half of the contract. A python-docx-authored document genuinely
    # owns a customXml store, a thumbnail and a stylesWithEffects part, and
    # dropping them as "template debris" would be the same defect mirrored.
    names = _rebuild(tmp_path, corpus_dir / "01_plain_text.docx", suffix="-full")

    assert "customXml/item1.xml" in names
    assert "customXml/itemProps1.xml" in names
    assert "docProps/thumbnail.jpeg" in names
    assert "word/stylesWithEffects.xml" in names


# --- optional parts the source simply does not have ---------------------------

@pytest.mark.parametrize(
    "part",
    [
        "word/numbering.xml",
        "word/styles.xml",
        "word/settings.xml",
    ],
)
def test_an_optional_part_the_source_lacks_is_not_taken_from_the_template(
    tmp_path, bare_source, part
):
    # set_numbering/set_styles/set_settings return early when the model has
    # nothing to write, which left the shell template's copy in place. For
    # numbering that was measurable on the corpus: documents with no
    # numbering.xml were handed the template's, complete with its own
    # mc:Ignorable declaration.
    names = _rebuild(tmp_path, bare_source, suffix=part.replace("/", "_"))

    assert part not in names


def test_dropping_an_optional_part_also_drops_what_referenced_it(tmp_path, bare_source):
    from word_replica.qa.preservation import g10_projection

    result = RebuildService(app_root=tmp_path / "app-optional").rebuild(
        bare_source,
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )
    projection = g10_projection(result.output_path)

    assert projection["dangling_relationships"] == []
    assert "numbering" not in projection["relationship_graph"]


def test_a_source_that_has_numbering_keeps_it(tmp_path, corpus_dir):
    names = _rebuild(tmp_path, corpus_dir / "03_lists.docx", suffix="-lists")

    assert "word/numbering.xml" in names


@pytest.mark.parametrize(
    "part",
    [
        "word/theme/theme1.xml",
        "word/webSettings.xml",
        "word/fontTable.xml",
    ],
)
def test_a_template_owned_part_the_source_lacks_is_not_added(tmp_path, bare_source, part):
    # The renderer never writes these, so whatever the shell template brought
    # stayed. Documents with no theme of their own were handed the template's,
    # and its mc:Ignorable declaration with it.
    names = _rebuild(tmp_path, bare_source, suffix=part.replace("/", "_"))

    assert part not in names


def test_a_source_that_has_a_theme_keeps_its_own(tmp_path, corpus_dir):
    # The mirror half: these are dropped from the template so the source's can
    # be restored in their place, never so the document ends up without one.
    from zipfile import ZipFile

    source = corpus_dir / "01_plain_text.docx"
    with ZipFile(source) as archive:
        expected = archive.read("word/theme/theme1.xml")

    result = RebuildService(app_root=tmp_path / "app-theme").rebuild(
        source,
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )
    with ZipFile(result.output_path) as archive:
        assert archive.read("word/theme/theme1.xml") == expected
