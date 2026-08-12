from pathlib import Path
from zipfile import ZipFile

from word_replica.domain.enums import MetadataMode
from word_replica.opc.package_reader import DocxPackage
from word_replica.opc.properties import (
    DocumentProperties,
    build_output_metadata,
    read_properties,
    select_preservable_properties,
)


def test_preserve_policy_excludes_lifecycle_history():
    props = DocumentProperties(
        title="Paper",
        subject="AI",
        creator="Alice",
        created="2025-01-01",
        modified="2025-02-01",
        revision="42",
        total_editing_time="999",
    )
    selected = select_preservable_properties(props, explicit_author=False)
    assert selected == {"title": "Paper", "subject": "AI"}


def test_custom_property_is_copied_only_when_allowlisted():
    props = DocumentProperties(custom={"DatasetVersion": "v13", "Secret": "x"})
    assert select_preservable_properties(props) == {}
    assert select_preservable_properties(props, custom_allowlist=("DatasetVersion",)) == {
        "custom:DatasetVersion": "v13"
    }


def test_fresh_metadata_never_preserves_source_properties():
    props = DocumentProperties(title="Paper", creator="Alice")
    assert build_output_metadata(props, MetadataMode.FRESH, True, ("Anything",)) == {}


def test_read_properties_reads_core_extended_and_custom(tmp_path: Path):
    path = tmp_path / "props.docx"
    with ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
        z.writestr("docProps/core.xml", """<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/'><dc:title>Paper</dc:title><dc:creator>Alice</dc:creator><cp:revision>42</cp:revision><dcterms:created>2025-01-01</dcterms:created></cp:coreProperties>""")
        z.writestr("docProps/app.xml", """<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'><Company>Lab</Company><TotalTime>999</TotalTime></Properties>""")
        z.writestr("docProps/custom.xml", """<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties' xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'><property name='DatasetVersion'><vt:lpwstr>v13</vt:lpwstr></property></Properties>""")
    with DocxPackage.open(path) as package:
        props = read_properties(package)
    assert props.title == "Paper"
    assert props.creator == "Alice"
    assert props.company == "Lab"
    assert props.revision == "42"
    assert props.total_editing_time == "999"
    assert props.custom == {"DatasetVersion": "v13"}
