"""The pure-docx renderer must not smuggle its template into the output.

Found by Lane P on the LibreOffice corpus: reconstructions that every model
gate (G0-G7) called perfect were carrying a customXml datastore, custom
document properties and a thumbnail that the source document never had.

PureDocxRenderer.render starts from `Document().save(shell)` -- python-docx's
default template -- and mutates that package. The template ships
`customXml/item1.xml`, `customXml/itemProps1.xml` and `docProps/thumbnail.jpeg`,
and no stage removes them, so every rebuilt document inherited them.

This is not cosmetic. In real Word documents the customXml store carries
structured business data and content-control bindings; silently adding a
foreign one means handing back a document that is not the document supplied.
It is invisible to every gate that compares parsed models, because the parser
has no representation for custom XML at all.
"""
from zipfile import ZipFile

import pytest

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
)
from word_replica.services.rebuild import RebuildService


@pytest.fixture
def rebuilt(tmp_path, corpus_dir):
    """Rebuild a fixture that has no custom XML of its own."""
    service = RebuildService(app_root=tmp_path / "app")
    result = service.rebuild(
        corpus_dir / "01_plain_text.docx",
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


def test_the_template_custom_xml_store_is_not_in_the_output(rebuilt):
    smuggled = sorted(name for name in rebuilt if name.lower().startswith("customxml/"))

    assert smuggled == [], f"template customXml leaked into the rebuild: {smuggled}"


def test_the_template_thumbnail_is_not_in_the_output(rebuilt):
    assert "docProps/thumbnail.jpeg" not in rebuilt


def test_no_relationship_points_at_a_part_that_is_gone(rebuilt, tmp_path, corpus_dir):
    # Removing a part without removing what references it produces a package
    # Word repairs on open, which would be a worse defect than the one fixed.
    from word_replica.qa.preservation import g10_projection

    service = RebuildService(app_root=tmp_path / "app2")
    result = service.rebuild(
        corpus_dir / "01_plain_text.docx",
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )

    assert g10_projection(result.output_path)["dangling_relationships"] == []


def test_no_content_type_override_survives_for_a_removed_part(rebuilt):
    assert not any(name.lower().startswith("customxml/") for name in rebuilt)


def test_the_rebuild_still_produces_a_readable_package(rebuilt):
    assert "word/document.xml" in rebuilt
    assert "[Content_Types].xml" in rebuilt
    assert "_rels/.rels" in rebuilt


def test_the_template_styles_with_effects_part_is_not_in_the_output(rebuilt):
    # A Word 2010 compatibility part the template carries; sources that never
    # had one were being handed it back.
    assert "word/stylesWithEffects.xml" not in rebuilt
