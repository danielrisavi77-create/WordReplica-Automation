from dataclasses import asdict, dataclass, field, is_dataclass
from hashlib import sha256
import json
from typing import Any, Iterator


class ElementIdFactory:
    def __init__(self, source_sha256: str) -> None:
        self.source_sha256 = source_sha256

    def make(self, kind: str, source_position: str) -> str:
        raw = f"{self.source_sha256}|{kind}|{source_position}".encode()
        return f"{kind[:3]}_{sha256(raw).hexdigest()[:16]}"


@dataclass(slots=True)
class Run:
    element_id: str
    text: str = ""
    properties: dict[str, Any] = field(default_factory=dict)
    hidden: bool = False


@dataclass(slots=True)
class Paragraph:
    element_id: str
    runs: list[Run] = field(default_factory=list)
    style_id: str | None = None
    properties: dict[str, Any] = field(default_factory=dict)

    def text(self) -> str:
        return "".join(run.text for run in self.runs)


@dataclass(slots=True)
class TableCell:
    element_id: str
    blocks: list[Any] = field(default_factory=list)
    properties: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class TableRow:
    element_id: str
    cells: list[TableCell] = field(default_factory=list)
    properties: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class Table:
    element_id: str
    rows: list[TableRow] = field(default_factory=list)
    properties: dict[str, Any] = field(default_factory=dict)




@dataclass(slots=True)
class DrawingRef:
    element_id: str
    asset_id: str
    source_path: str
    representation: str
    width_emu: int | None = None
    height_emu: int | None = None
    lock_aspect_ratio: bool | None = None
    wrap_type: str | None = None
    horizontal_relative_from: str | None = None
    horizontal_position_emu: int | None = None
    vertical_relative_from: str | None = None
    vertical_position_emu: int | None = None
    distance_top_emu: int | None = None
    distance_bottom_emu: int | None = None
    distance_left_emu: int | None = None
    distance_right_emu: int | None = None
    crop: dict[str, int] = field(default_factory=dict)
    rotation_degrees: float | None = None
    behind_text: bool | None = None
    z_order: int | None = None

@dataclass(slots=True)
class BinaryAsset:
    asset_id: str
    part_name: str
    content_type: str | None
    sha256: str
    bytes_data: bytes


@dataclass(slots=True)
class RelationshipRef:
    rel_id: str
    rel_type: str
    target: str
    external: bool = False


@dataclass(slots=True)
class Section:
    element_id: str
    properties: dict[str, Any] = field(default_factory=dict)


def _stable(value: Any) -> Any:
    if is_dataclass(value):
        return {k: _stable(v) for k, v in asdict(value).items() if k != "bytes_data"}
    if isinstance(value, dict):
        return {str(k): _stable(value[k]) for k in sorted(value, key=str)}
    if isinstance(value, (list, tuple)):
        return [_stable(v) for v in value]
    if isinstance(value, bytes):
        return sha256(value).hexdigest()
    return value


@dataclass(slots=True)
class Comment:
    comment_id: str
    author: str | None
    date: str | None
    blocks: list[object] = field(default_factory=list)


@dataclass(slots=True)
class Bookmark:
    bookmark_id: str
    name: str
    start_path: str
    end_path: str | None


@dataclass(slots=True)
class Field:
    field_id: str
    instruction: str
    result_text: str
    locked: bool = False


@dataclass(slots=True)
class RevisionSpan:
    revision_id: str
    kind: str
    author: str | None
    date: str | None
    path: str
    text: str


@dataclass(slots=True)
class PreservedPart:
    """A package part the document model does not otherwise represent.

    How it is reached decides what a renderer can do with it:

    * ``relationship_type`` set -- a document-level attachment. Part plus one
      relationship off document.xml.rels is the whole of it, so it can be
      restored exactly.
    * ``sidecar`` -- reached from the attachment's own .rels (a customXml
      item's properties part). Restored as a part; its relationship comes back
      with the .rels file itself.
    * neither -- reached from inside the document body (charts, embedded
      objects, diagrams). A renderer that rebuilds the body cannot restore the
      reference, so writing the part alone would only produce an orphan. It is
      captured so the loss can be reported rather than hidden.
    """

    part_name: str
    content_type: str | None
    relationship_type: str | None
    sha256: str
    data: bytes
    sidecar: bool = False
    # Which .rels file carries the relationship reaching this part. Almost
    # everything hangs off the document; the package thumbnail hangs off the
    # package itself.
    owner_rels: str = "word/_rels/document.xml.rels"


@dataclass(slots=True)
class DocumentModel:
    source_sha256: str
    body: list[Any] = field(default_factory=list)
    sections: list[Section] = field(default_factory=list)
    styles_xml: bytes | None = None
    numbering_xml: bytes | None = None
    settings_xml: bytes | None = None
    theme_parts: dict[str, bytes] = field(default_factory=dict)
    relationships: dict[str, RelationshipRef] = field(default_factory=dict)
    assets: dict[str, BinaryAsset] = field(default_factory=dict)
    drawings: list[DrawingRef] = field(default_factory=list)
    headers: dict[str, list[Any]] = field(default_factory=dict)
    footers: dict[str, list[Any]] = field(default_factory=dict)
    footnotes: dict[str, list[Any]] = field(default_factory=dict)
    endnotes: dict[str, list[Any]] = field(default_factory=dict)
    comments: dict[str, Comment] = field(default_factory=dict)
    bookmarks: list[Bookmark] = field(default_factory=list)
    fields: list[Field] = field(default_factory=list)
    revisions: list[RevisionSpan] = field(default_factory=list)
    preserved_parts: dict[str, PreservedPart] = field(default_factory=dict)
    extras: dict[str, Any] = field(default_factory=dict)

    def plain_text(self) -> str:
        return "\n".join(block.text() for block in self.body if isinstance(block, Paragraph))

    def fingerprint(self) -> str:
        payload = json.dumps(
            _stable(self),
            sort_keys=True,
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        return sha256(payload).hexdigest()

    def iter_paragraphs(self) -> Iterator[Paragraph]:
        def walk(blocks: list[Any]) -> Iterator[Paragraph]:
            for block in blocks:
                if isinstance(block, Paragraph):
                    yield block
                elif isinstance(block, Table):
                    for row in block.rows:
                        for cell in row.cells:
                            yield from walk(cell.blocks)

        yield from walk(self.body)
