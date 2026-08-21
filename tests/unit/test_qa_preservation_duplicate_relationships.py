"""G10 compares what a package points at, not how many records say so.

Word writes one relationship per hyperlink *occurrence*, so a document that
links the same URL four times carries four relationship records with four ids
and one target. Any writer that allocates relationships by target -- ours does,
and so does Word on some paths -- collapses those four into one. Nothing about
the document changes: all four `w:hyperlink` elements point at the surviving
id, and every link still opens the same page.

Counting records made that collapse look like losing three links. Measured on
the corpus it was the entire `relationship_graph/hyperlink/count` class:

    tdf159897_broken_link.docx    4 records, 1 distinct target
    58618.docx                    4 records, 1 distinct target
    fdo76597.docx                39 records, 27 distinct targets

A gate that reports those as silent data loss is a gate people switch off, and
then the real losses go unnoticed with it. So the projection counts distinct
targets, which still moves the moment a link is actually dropped or invented.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.qa.preservation import build_preservation_gate, g10_projection

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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


def _package(path, links):
    """A document whose paragraphs link the given targets, in order."""
    rels = "".join(
        f'<Relationship Id="rId{index + 4}" Type="{REL}/hyperlink" '
        f'Target="{target}" TargetMode="External"/>'
        for index, target in enumerate(links)
    )
    body = "".join(
        f'<w:p><w:hyperlink r:id="rId{index + 4}"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>'
        for index in range(len(links))
    )
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{REL}">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            (
                f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">{rels}</Relationships>'
            ).encode("utf-8"),
        )
        archive.writestr("word/document.xml", document)
    return path


ONE = "https://libreoffice.org/"
TWO = "https://documentfoundation.org/"


def test_four_records_for_one_target_count_as_one(tmp_path):
    projection = g10_projection(_package(tmp_path / "four.docx", [ONE] * 4))

    assert projection["relationship_graph"]["hyperlink"]["distinct_targets"] == 1


def test_collapsing_duplicate_records_is_not_a_divergence(tmp_path):
    source = _package(tmp_path / "source.docx", [ONE] * 4)
    output = _package(tmp_path / "output.docx", [ONE])

    gate = build_preservation_gate(source, output)

    assert gate.passed is True, gate.first_divergence


def test_losing_a_distinct_target_is_still_a_divergence(tmp_path):
    source = _package(tmp_path / "both.docx", [ONE, TWO])
    output = _package(tmp_path / "one.docx", [ONE, ONE])

    gate = build_preservation_gate(source, output)

    assert gate.passed is False


def test_inventing_a_target_is_still_a_divergence(tmp_path):
    source = _package(tmp_path / "src-one.docx", [ONE, ONE])
    output = _package(tmp_path / "out-two.docx", [ONE, TWO])

    gate = build_preservation_gate(source, output)

    assert gate.passed is False


def test_external_targets_are_counted_distinctly_too(tmp_path):
    projection = g10_projection(_package(tmp_path / "ext.docx", [ONE] * 3 + [TWO]))

    entry = projection["relationship_graph"]["hyperlink"]
    assert (entry["distinct_targets"], entry["external"]) == (2, 2)


@pytest.mark.parametrize("count", [1, 2, 5])
def test_the_same_document_projects_the_same_way_however_often_it_repeats(tmp_path, count):
    # The projection must depend on the set of targets, not the repetition.
    projection = g10_projection(_package(tmp_path / f"n{count}.docx", [ONE] * count))

    assert projection["relationship_graph"]["hyperlink"]["distinct_targets"] == 1
