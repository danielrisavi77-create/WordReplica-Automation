"""G10 -- Word-free package preservation.

Every gate from G0 to G9 compares *parsed models*, and ``parser/parser.py``
only models what WordReplica reconstructs. Anything outside that model is
invisible: not compared, not even retained. Word 2010 compounds it, because it
silently ignores w15/w16 in both files, so its render cancels the difference
out too.

Measured on the first real corpus the lab ingested: **602 of 1,830 documents
(32.9 %)** carry content in that blind spot. G10 reads raw package bytes, so it
is blind to none of it.

The gate is only useful if it is quiet on things that are not defects, so the
first half of this file is about *not* failing.
"""
from hashlib import sha256
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.qa.preservation import build_preservation_gate, g10_projection

W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

CT_HEAD = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
)
CT_DOC = (
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
)
CT_CUSTOM_XML = (
    '<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>'
    '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.customXmlProperties+xml"/>'
)
CT_COMMENTS_EX = (
    '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.wordprocessingml.commentsExtended+xml"/>'
)

REL_HEAD = (
    '<?xml version="1.0"?><Relationships '
    'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
)
REL_DOC = (
    '<Relationship Id="rId1" Target="word/document.xml" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
)

CORE_XML = (
    '<?xml version="1.0"?><cp:coreProperties '
    'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dcterms="http://purl.org/dc/terms/">'
    '<cp:revision>{revision}</cp:revision>'
    '<dcterms:created>{created}</dcterms:created>'
    "</cp:coreProperties>"
)


def _document(*, rsid: str = "00AB12CD", doc_id: str = "{11111111-1111-1111-1111-111111111111}") -> str:
    return (
        '<?xml version="1.0"?><w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        f'xmlns:w15="{W15}">'
        f'<w:body><w:p w:rsidR="{rsid}" w15:paraId="{doc_id}">'
        "<w:r><w:t>hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
    )


def _package(
    *,
    custom_xml: bool = True,
    comments_ex: bool = True,
    ole: bytes | None = b"\xd0\xcf\x11\xe0 ole payload",
    rsid: str = "00AB12CD",
    revision: str = "3",
    created: str = "2020-01-01T00:00:00Z",
    header_name: str = "word/header1.xml",
    reverse_order: bool = False,
    extra_rels: str = "",
) -> dict[str, bytes]:
    content_types = CT_HEAD + CT_DOC
    rels = REL_HEAD  # targets here resolve relative to word/
    parts: dict[str, str | bytes] = {}

    content_types += (
        '<Override PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument'
        '.wordprocessingml.header+xml"/>'
    ).format(header_name)
    parts[header_name] = (
        '<?xml version="1.0"?><w:hdr '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>'
    )
    rels += (
        '<Relationship Id="rId9" Target="{}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"/>'
    ).format(header_name.removeprefix("word/"))

    if custom_xml:
        content_types += CT_CUSTOM_XML
        parts["customXml/item1.xml"] = '<?xml version="1.0"?><root xmlns="urn:acme:contract"><id>42</id></root>'
        parts["customXml/itemProps1.xml"] = (
            '<?xml version="1.0"?><ds:datastoreItem '
            'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" '
            'ds:itemID="{AAAA0000-0000-0000-0000-000000000001}"/>'
        )
        # A customXml item reaches its properties part through its own .rels.
        # Without this the props part is an orphan in the fixture itself.
        parts["customXml/_rels/item1.xml.rels"] = (
            REL_HEAD
            + '<Relationship Id="rId1" Target="itemProps1.xml" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            '/customXmlProps"/></Relationships>'
        )
        rels += (
            '<Relationship Id="rId20" Target="../customXml/item1.xml" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"/>'
        )
    if comments_ex:
        content_types += CT_COMMENTS_EX
        parts["word/commentsExtended.xml"] = (
            f'<?xml version="1.0"?><w15:commentsEx xmlns:w15="{W15}">'
            '<w15:commentEx w15:paraId="1" w15:done="0"/></w15:commentsEx>'
        )
        rels += (
            '<Relationship Id="rId21" Target="commentsExtended.xml" '
            'Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended"/>'
        )
    if ole is not None:
        content_types += '<Override PartName="/word/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>'
        parts["word/embeddings/oleObject1.bin"] = ole
        rels += (
            '<Relationship Id="rId22" Target="embeddings/oleObject1.bin" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"/>'
        )

    rels += extra_rels
    content_types += "</Types>"
    rels += "</Relationships>"

    package: dict[str, str | bytes] = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": (
            REL_HEAD
            + REL_DOC
            + '<Relationship Id="rId2" Target="docProps/core.xml" '
            'Type="http://schemas.openxmlformats.org/package/2006/relationships'
            '/metadata/core-properties"/></Relationships>'
        ),
        "word/_rels/document.xml.rels": rels,
        "word/document.xml": _document(rsid=rsid),
        "docProps/core.xml": CORE_XML.format(revision=revision, created=created),
        **parts,
    }
    encoded = {k: (v.encode("utf-8") if isinstance(v, str) else v) for k, v in package.items()}
    if reverse_order:
        encoded = {k: encoded[k] for k in reversed(list(encoded))}
    return encoded


def _write(path, parts: dict[str, bytes]):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


@pytest.fixture
def source(tmp_path):
    return _write(tmp_path / "source.docx", _package())


# --- the gate must be quiet on things that are not defects -------------------

def test_an_identical_package_passes(tmp_path, source):
    output = _write(tmp_path / "output.docx", _package())

    gate = build_preservation_gate(source, output)

    assert gate.passed is True
    assert gate.name == "G10"
    assert gate.first_divergence is None


@pytest.mark.parametrize(
    "label, kwargs",
    [
        ("revision-save ids", {"rsid": "99FF00EE"}),
        ("document revision counter", {"revision": "17"}),
        ("creation timestamp", {"created": "2026-08-19T12:00:00Z"}),
        ("zip member order", {"reverse_order": True}),
        ("part numbering", {"header_name": "word/header3.xml"}),
    ],
)
def test_incidental_differences_do_not_fail_the_gate(tmp_path, source, label, kwargs):
    # A gate that fires on rsids or a renumbered header part gets switched off,
    # and then the blind spot it exists to cover reopens.
    output = _write(tmp_path / "output.docx", _package(**kwargs))

    gate = build_preservation_gate(source, output)

    assert gate.passed is True, f"{label}: {gate.summary} {gate.first_divergence}"


def test_the_projection_itself_is_stable_across_incidental_differences(tmp_path):
    a = _write(tmp_path / "a.docx", _package())
    b = _write(tmp_path / "b.docx", _package(rsid="FFFF1111", revision="9", created="2030-01-01T00:00:00Z",
                                             header_name="word/header7.xml", reverse_order=True))

    assert g10_projection(a) == g10_projection(b)


# --- the reason the gate exists ----------------------------------------------

def test_a_dropped_custom_xml_part_fails_g10_while_every_model_gate_passes(tmp_path, source):
    """The false-pass proof.

    A reconstruction that loses the entire customXml store is reported as
    perfect by G0-G7, because the document model has no representation for it.
    """
    from word_replica.parser.parser import DocxParser
    from word_replica.qa.golden_audit import build_model_gates

    output = _write(tmp_path / "output.docx", _package(custom_xml=False))

    model_gates = build_model_gates(DocxParser().parse(source), DocxParser().parse(output))
    assert all(gate.passed for gate in model_gates.values()), {
        name: gate.summary for name, gate in model_gates.items() if not gate.passed
    }

    gate = build_preservation_gate(source, output)
    assert gate.passed is False
    assert gate.first_divergence is not None


def test_a_dropped_w15_comment_extension_fails_g10(tmp_path, source):
    output = _write(tmp_path / "output.docx", _package(comments_ex=False))

    assert build_preservation_gate(source, output).passed is False


def test_an_opaque_part_whose_bytes_changed_fails_g10(tmp_path, source):
    # OLE payloads have no model projection at all, so byte identity is the
    # only thing that can be asserted about them.
    output = _write(tmp_path / "output.docx", _package(ole=b"\xd0\xcf\x11\xe0 different payload"))

    gate = build_preservation_gate(source, output)

    assert gate.passed is False
    assert "opaque" in str(gate.first_divergence).lower() or "opaque" in gate.summary.lower()


def test_a_dropped_opaque_part_fails_g10(tmp_path, source):
    output = _write(tmp_path / "output.docx", _package(ole=None))

    assert build_preservation_gate(source, output).passed is False


def test_a_namespace_appearing_only_in_the_output_fails_g10(tmp_path, source):
    # Spurious injection is how "silently returned an incorrect document"
    # happens, so the comparison is symmetric on purpose.
    parts = _package()
    parts["word/document.xml"] = (
        '<?xml version="1.0"?><w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml">'
        "<w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p><v:shape/><w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    output = _write(tmp_path / "output.docx", parts)

    assert build_preservation_gate(source, output).passed is False


def test_a_relationship_type_that_disappeared_fails_g10(tmp_path, source):
    parts = _package()
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"',
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"',
    )
    output = _write(tmp_path / "output.docx", parts)

    assert build_preservation_gate(source, output).passed is False


def test_a_dangling_relationship_introduced_by_the_output_fails_g10(tmp_path, source):
    extra = (
        '<Relationship Id="rId99" Target="media/missing.png" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>'
    )
    output = _write(tmp_path / "output.docx", _package(extra_rels=extra))

    assert build_preservation_gate(source, output).passed is False


def test_a_lost_part_kind_fails_g10_even_when_the_count_is_all_that_changed(tmp_path, source):
    parts = _package()
    del parts["word/header1.xml"]
    output = _write(tmp_path / "output.docx", parts)

    assert build_preservation_gate(source, output).passed is False


# --- behaviour under damage ---------------------------------------------------

def test_an_unreadable_output_fails_rather_than_raising(tmp_path, source):
    output = tmp_path / "output.docx"
    output.write_bytes(b"not a zip at all")

    gate = build_preservation_gate(source, output)

    assert gate.passed is False
    assert gate.first_divergence is not None


def test_the_projection_records_what_it_compared(tmp_path, source):
    projection = g10_projection(source)

    assert "namespaces" in projection
    assert "part_kinds" in projection
    assert "opaque_parts" in projection
    assert "relationship_graph" in projection
    assert projection["dangling_relationships"] == []
    # opaque parts are compared by content hash, never by name
    digests = [d for group in projection["opaque_parts"].values() for d in group]
    assert sha256(b"\xd0\xcf\x11\xe0 ole payload").hexdigest() in digests


# --- WordReplica's own provenance marking ------------------------------------

_CUSTOM_PROPS_CT = (
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.custom-properties+xml"/>'
)


def _with_custom_properties(parts: dict[str, bytes], *names: str) -> dict[str, bytes]:
    body = "".join(
        '<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{pid}" name="{name}">'
        "<vt:lpwstr>value</vt:lpwstr></property>".format(pid=index + 2, name=name)
        for index, name in enumerate(names)
    )
    parts = dict(parts)
    parts["docProps/custom.xml"] = (
        '<?xml version="1.0"?><Properties '
        'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f"{body}</Properties>"
    ).encode("utf-8")
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>", _CUSTOM_PROPS_CT.encode("utf-8") + b"</Types>"
    )
    return parts


def test_word_replica_provenance_properties_do_not_fail_the_gate(tmp_path, source):
    # Every reconstruction stamps WordReplicaProjectId / Reconstructed /
    # ActualSaveCount. That is designed behaviour, and a gate that failed on it
    # would fail on 100% of documents and be switched off within a week.
    output = _write(
        tmp_path / "output.docx",
        _with_custom_properties(
            _package(),
            "WordReplicaProjectId",
            "WordReplicaReconstructed",
            "WordReplicaActualSaveCount",
        ),
    )

    assert build_preservation_gate(source, output).passed is True


def test_a_user_custom_property_that_disappeared_still_fails_the_gate(tmp_path):
    # The provenance carve-out is by property name, so it cannot mask a real
    # custom-property loss.
    src = _write(tmp_path / "source.docx", _with_custom_properties(_package(), "ContractNumber"))
    output = _write(tmp_path / "output.docx", _package())

    assert build_preservation_gate(src, output).passed is False


def test_a_user_custom_property_the_output_invented_fails_the_gate(tmp_path, source):
    output = _write(tmp_path / "output.docx", _with_custom_properties(_package(), "InventedByUs"))

    assert build_preservation_gate(source, output).passed is False


def test_provenance_alongside_a_real_property_is_still_compared(tmp_path):
    src = _write(tmp_path / "source.docx", _with_custom_properties(_package(), "ContractNumber"))
    output = _write(
        tmp_path / "output.docx",
        _with_custom_properties(_package(), "ContractNumber", "WordReplicaReconstructed"),
    )

    assert build_preservation_gate(src, output).passed is True


# --- differences the metadata policy deliberately makes ----------------------

def test_a_deliberately_dropped_custom_properties_part_can_be_declared(tmp_path):
    # README: "custom properties require an explicit allowlist". Under the
    # default policy a source's custom properties are dropped on purpose, so a
    # caller that used that policy can say so and G10 will not call it a defect.
    src = _write(tmp_path / "source.docx", _with_custom_properties(_package(), "ContractNumber"))
    output = _write(tmp_path / "output.docx", _package())

    assert build_preservation_gate(src, output).passed is False
    assert build_preservation_gate(src, output, custom_properties_dropped_by_policy=True).passed is True


def test_declaring_the_policy_does_not_excuse_inventing_properties(tmp_path, source):
    # The carve-out is one-directional: policy explains a loss, never a gain.
    output = _write(tmp_path / "output.docx", _with_custom_properties(_package(), "InventedByUs"))

    assert build_preservation_gate(source, output, custom_properties_dropped_by_policy=True).passed is False


def test_declaring_the_policy_does_not_excuse_other_losses(tmp_path):
    src = _write(tmp_path / "source.docx", _with_custom_properties(_package(), "ContractNumber"))
    output = _write(tmp_path / "output.docx", _package(comments_ex=False))

    assert build_preservation_gate(src, output, custom_properties_dropped_by_policy=True).passed is False


# --- parts nothing points at ---------------------------------------------------

def test_a_part_no_relationship_reaches_is_reported(tmp_path):
    # The mirror of a dangling relationship. Found by the pure-docx renderer
    # keeping an image's bytes while never emitting the w:drawing that referred
    # to it: the media part survived as litter and the picture was gone from
    # the document, and the relationship graph alone could not see it.
    parts = _package()
    parts["word/media/image1.png"] = b"\x89PNG\r\n\x1a\n orphan"
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>", b'<Default Extension="png" ContentType="image/png"/></Types>'
    )
    path = _write(tmp_path / "orphan.docx", parts)

    assert g10_projection(path)["orphan_parts"] == ["image/png"]


def test_an_ordinary_package_has_no_orphans(tmp_path, source):
    assert g10_projection(source)["orphan_parts"] == []


def test_an_output_that_orphans_a_part_fails_the_gate(tmp_path, source):
    parts = _package()
    parts["word/media/image1.png"] = b"\x89PNG\r\n\x1a\n orphan"
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>", b'<Default Extension="png" ContentType="image/png"/></Types>'
    )
    output = _write(tmp_path / "output.docx", parts)

    assert build_preservation_gate(source, output).passed is False


def test_orphan_parts_are_reported_by_kind_not_by_name(tmp_path):
    # Same normalization as everywhere else: image1.png standing in for
    # image3.png is not a defect.
    def _with_orphan(name):
        parts = _package()
        parts[f"word/media/{name}"] = b"\x89PNG\r\n\x1a\n orphan"
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
            b"</Types>", b'<Default Extension="png" ContentType="image/png"/></Types>'
        )
        return parts

    first = g10_projection(_write(tmp_path / "a.docx", _with_orphan("image1.png")))
    second = g10_projection(_write(tmp_path / "b.docx", _with_orphan("image7.png")))

    assert first["orphan_parts"] == second["orphan_parts"] == ["image/png"]
