from pathlib import Path
from zipfile import ZipFile

from word_replica.opc.package_reader import DocxPackage


def make_package(path: Path) -> None:
    with ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
        z.writestr("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>")
        z.writestr("word/_rels/document.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='image' Target='media/image1.png'/></Relationships>")
        z.writestr("word/media/image1.png", b"png")


def test_reader_lists_parts_and_relationships(tmp_path: Path):
    path = tmp_path / "a.docx"
    make_package(path)
    with DocxPackage.open(path) as package:
        assert "word/document.xml" in package.parts
        assert package.read_bytes("word/media/image1.png") == b"png"
        assert package.relationships("word/document.xml")["rId1"].target == "media/image1.png"
        assert package.iter_parts("word/media/") == ["word/media/image1.png"]
