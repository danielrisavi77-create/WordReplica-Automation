from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree

from word_replica.domain.errors import PackageReadError

REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}


@dataclass(frozen=True, slots=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    target_mode: str | None


class DocxPackage:
    def __init__(self, path: Path, archive: ZipFile) -> None:
        self.path = path
        self.archive = archive
        self.parts = frozenset(archive.namelist())

    @classmethod
    def open(cls, path: Path) -> "DocxPackage":
        try:
            return cls(path, ZipFile(path, "r"))
        except (BadZipFile, OSError) as exc:
            raise PackageReadError(str(exc)) from exc

    def __enter__(self) -> "DocxPackage":
        return self

    def __exit__(self, *_args) -> None:
        self.archive.close()

    def read_bytes(self, part_name: str) -> bytes:
        try:
            return self.archive.read(part_name)
        except KeyError as exc:
            raise PackageReadError(f"Missing DOCX part: {part_name}") from exc

    def read_xml(self, part_name: str):
        try:
            return etree.fromstring(self.read_bytes(part_name))
        except etree.XMLSyntaxError as exc:
            raise PackageReadError(f"Malformed XML part {part_name}: {exc}") from exc

    def iter_parts(self, prefix: str) -> list[str]:
        """Parts under ``prefix``. Directory entries are not parts.

        A zip may carry zero-length entries ending in "/" to record a folder,
        and many writers emit them. Handing one back as a part means a caller
        reads a folder as content: the theme scheme gave up on the empty
        ``word/theme/`` before reaching ``word/theme/theme1.xml`` beside it, and
        ``word/media/`` would have been extracted as an image.
        """
        return sorted(
            part for part in self.parts if part.startswith(prefix) and not part.endswith("/")
        )

    def relationships(self, source_part: str) -> dict[str, Relationship]:
        source = Path(source_part)
        rel_part = str(source.parent / "_rels" / f"{source.name}.rels").replace("\\", "/")
        if rel_part not in self.parts:
            return {}
        root = self.read_xml(rel_part)
        result: dict[str, Relationship] = {}
        for node in root.xpath("//pr:Relationship", namespaces=REL_NS):
            rel = Relationship(
                node.get("Id"),
                node.get("Type"),
                node.get("Target"),
                node.get("TargetMode"),
            )
            result[rel.rel_id] = rel
        return result
