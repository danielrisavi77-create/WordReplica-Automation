"""Pictures must come back into the document body, not just the package.

Verified directly before writing this: rebuilding
`05_images_inline_floating.docx` produced a package still containing
`word/media/image1.png` and a `word/document.xml` with no `w:drawing` at all.
Every rebuilt document lost every picture from its visible content while the
image bytes sat in the package as litter. It was the single largest remaining
G10 loss -- `wordprocessingDrawing` absent from 13 of 34 sampled documents --
and the renderer said so honestly through
`PURE_DOCX_ASSET_POSITION_UNAVAILABLE`.

The approach is to preserve the original `w:drawing` verbatim and remap its
relationship ids, rather than rebuild DrawingML by hand. Verbatim keeps crop,
rotation, wrapping, effects and anchor geometry exactly as authored; hand-built
markup would keep only what the model happens to represent.

Remapping is what makes it safe. The inline-fragment mechanism refuses any
fragment carrying an `r:` reference, because a renumbered id points at the
wrong part or none at all. Here the reference *can* be resolved -- the parser
knows which part each id addressed and the renderer owns the new package -- so
the id is rewritten rather than the fragment refused. An id that cannot be
resolved is still refused.
"""
from zipfile import ZipFile

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
from word_replica.qa.preservation import build_preservation_gate, g10_projection
from word_replica.services.rebuild import RebuildService

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


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


def _document(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/document.xml"))


def _media(path):
    with ZipFile(path) as archive:
        return sorted(n for n in archive.namelist() if n.startswith("word/media/"))


@pytest.fixture
def rebuilt(tmp_path, corpus_dir):
    source = corpus_dir / "05_images_inline_floating.docx"
    return source, _rebuild(tmp_path, source)


# --- the picture is in the document, not only in the package ------------------

def test_the_drawing_element_is_emitted(rebuilt):
    _source, output = rebuilt

    assert _document(output).findall(f".//{{{W}}}drawing"), "no w:drawing in the rebuilt body"


def test_the_media_part_still_travels_with_it(rebuilt):
    source, output = rebuilt

    assert _media(output) == _media(source)


def test_the_drawing_references_a_relationship_that_exists(rebuilt):
    # The whole risk of this change: a reference pointing at nothing produces a
    # document Word repairs on open, which is worse than a missing picture.
    _source, output = rebuilt

    assert g10_projection(output)["dangling_relationships"] == []


def test_the_image_is_no_longer_an_orphan_part(rebuilt):
    _source, output = rebuilt

    assert "image/png" not in g10_projection(output)["orphan_parts"]


def test_the_relationship_id_is_remapped_to_the_new_package(rebuilt):
    _source, output = rebuilt
    document = _document(output)

    embeds = {
        node.get(f"{{{R}}}embed")
        for node in document.iter()
        if isinstance(node.tag, str) and node.get(f"{{{R}}}embed")
    }
    assert embeds, "no r:embed survived"

    with ZipFile(output) as archive:
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
    declared = {node.get("Id") for node in rels}
    assert embeds <= declared, f"{embeds - declared} referenced but not declared"


# --- what verbatim preservation buys ------------------------------------------

def test_anchor_geometry_survives_verbatim(tmp_path, corpus_dir):
    # Hand-built DrawingML would keep only what the model represents. Verbatim
    # keeps the authored geometry, so an anchored picture stays anchored.
    source = corpus_dir / "05_images_inline_floating.docx"
    before, after = _document(source), _document(_rebuild(tmp_path, source, "-geom"))

    def _kinds(root):
        return sorted(
            node.tag.rsplit("}", 1)[-1]
            for node in root.iter()
            if isinstance(node.tag, str) and node.tag.rsplit("}", 1)[-1] in {"inline", "anchor", "extent"}
        )

    assert _kinds(after) == _kinds(before)


def test_g10_no_longer_reports_the_drawing_namespaces_as_lost(tmp_path, corpus_dir):
    source = corpus_dir / "05_images_inline_floating.docx"
    before = g10_projection(source)
    after = g10_projection(_rebuild(tmp_path, source, "-ns"))

    for uri in before["namespaces"]:
        if "wordprocessingDrawing" in uri or uri.endswith("/drawingml/2006/picture"):
            assert uri in after["namespaces"], f"{uri} lost"


def test_the_rebuild_is_still_parseable(tmp_path, corpus_dir):
    output = _rebuild(tmp_path, corpus_dir / "05_images_inline_floating.docx", "-parse")

    assert DocxParser().parse(output).drawings


def test_a_second_round_trip_is_stable(tmp_path, corpus_dir):
    # Each rebuild renumbers relationships. If remapping were not idempotent the
    # document would degrade a little on every pass.
    once = _rebuild(tmp_path, corpus_dir / "05_images_inline_floating.docx", "-1")
    twice = _rebuild(tmp_path, once, "-2")

    assert build_preservation_gate(once, twice).passed is True


def test_a_document_with_no_pictures_gains_none(tmp_path, corpus_dir):
    output = _rebuild(tmp_path, corpus_dir / "01_plain_text.docx", "-plain")

    assert not _document(output).findall(f".//{{{W}}}drawing")
    assert not _media(output)
