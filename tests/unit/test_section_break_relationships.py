from zipfile import ZIP_DEFLATED, ZipFile

from word_replica.services.interactive_rebuild import InteractiveRebuildService


def test_restore_nested_section_footer_avoids_existing_relationship_id(tmp_path):
    from lxml import etree

    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    package_rels = "http://schemas.openxmlformats.org/package/2006/relationships"
    office_rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    source_xml = (
        f"<w:document xmlns:w='{word_ns}' xmlns:r='{office_rels}'><w:body>"
        "<w:sdt><w:sdtContent><w:p><w:pPr><w:sectPr>"
        "<w:footerReference w:type='default' r:id='rId9'/>"
        "</w:sectPr></w:pPr></w:p></w:sdtContent></w:sdt>"
        "</w:body></w:document>"
    ).encode()
    output_xml = (
        f"<w:document xmlns:w='{word_ns}' xmlns:r='{office_rels}'><w:body>"
        "<w:p/></w:body></w:document>"
    ).encode()
    source_relationships = (
        f"<Relationships xmlns='{package_rels}'>"
        f"<Relationship Id='rId9' Type='{office_rels}/footer' Target='footer9.xml'/>"
        "</Relationships>"
    ).encode()
    output_relationships = (
        f"<Relationships xmlns='{package_rels}'>"
        f"<Relationship Id='rId9' Type='{office_rels}/image' Target='media/image1.png'/>"
        "</Relationships>"
    ).encode()
    source_content_types = (
        f"<Types xmlns='{content_types_ns}'>"
        "<Override PartName='/word/footer9.xml' "
        "ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'/>"
        "</Types>"
    ).encode()
    output_content_types = f"<Types xmlns='{content_types_ns}'/>".encode()
    with ZipFile(source, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", source_xml)
        archive.writestr("word/_rels/document.xml.rels", source_relationships)
        archive.writestr("word/footer9.xml", b"<w:ftr xmlns:w='" + word_ns.encode() + b"'/>")
        archive.writestr("[Content_Types].xml", source_content_types)
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", output_xml)
        archive.writestr("word/_rels/document.xml.rels", output_relationships)
        archive.writestr("word/media/image1.png", b"image")
        archive.writestr("[Content_Types].xml", output_content_types)

    restored = InteractiveRebuildService._restore_source_section_breaks(output, source)

    with ZipFile(output) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        relationships = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        relationship_nodes = relationships.findall(f"{{{package_rels}}}Relationship")
        relationship_ids = [node.get("Id") for node in relationship_nodes]
        footer_relationship = next(
            node for node in relationship_nodes if (node.get("Type") or "").endswith("/footer")
        )
        footer_reference_ids = document.xpath(
            "//w:footerReference/@r:id",
            namespaces={"w": word_ns, "r": office_rels},
        )
        assert "word/footer9.xml" in archive.namelist()

    assert restored == 1
    assert len(relationship_ids) == len(set(relationship_ids))
    assert footer_relationship.get("Id") != "rId9"
    assert footer_reference_ids == [footer_relationship.get("Id")]


def test_restore_body_section_maps_header_to_existing_relationship_target(tmp_path):
    from lxml import etree

    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    package_rels = "http://schemas.openxmlformats.org/package/2006/relationships"
    office_rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    source_xml = (
        f"<w:document xmlns:w='{word_ns}' xmlns:r='{office_rels}'><w:body>"
        "<w:p/><w:sectPr><w:headerReference w:type='default' r:id='rId33'/>"
        "</w:sectPr></w:body></w:document>"
    ).encode()
    output_xml = (
        f"<w:document xmlns:w='{word_ns}' xmlns:r='{office_rels}'><w:body>"
        "<w:p/><w:sectPr><w:headerReference w:type='default' r:id='rId31'/>"
        "</w:sectPr></w:body></w:document>"
    ).encode()
    source_relationships = (
        f"<Relationships xmlns='{package_rels}'>"
        f"<Relationship Id='rId33' Type='{office_rels}/header' Target='header1.xml'/>"
        "</Relationships>"
    ).encode()
    output_relationships = (
        f"<Relationships xmlns='{package_rels}'>"
        f"<Relationship Id='rId31' Type='{office_rels}/header' Target='header1.xml'/>"
        "</Relationships>"
    ).encode()
    content_types = (
        f"<Types xmlns='{content_types_ns}'>"
        "<Override PartName='/word/header1.xml' "
        "ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'/>"
        "</Types>"
    ).encode()
    header_xml = f"<w:hdr xmlns:w='{word_ns}'><w:p/></w:hdr>".encode()
    for path, document_xml, relationships in (
        (source, source_xml, source_relationships),
        (output, output_xml, output_relationships),
    ):
        with ZipFile(path, "w", ZIP_DEFLATED) as archive:
            archive.writestr("word/document.xml", document_xml)
            archive.writestr("word/_rels/document.xml.rels", relationships)
            archive.writestr("word/header1.xml", header_xml)
            archive.writestr("[Content_Types].xml", content_types)

    restored = InteractiveRebuildService._restore_source_section_breaks(output, source)

    with ZipFile(output) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        relationships = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        header_relationships = [
            node
            for node in relationships.findall(f"{{{package_rels}}}Relationship")
            if (node.get("Type") or "").endswith("/header")
        ]
        header_reference_ids = document.xpath(
            "//w:headerReference/@r:id",
            namespaces={"w": word_ns, "r": office_rels},
        )

    assert restored == 0
    assert len(header_relationships) == 1
    assert header_relationships[0].get("Id") == "rId31"
    assert header_reference_ids == ["rId31"]
