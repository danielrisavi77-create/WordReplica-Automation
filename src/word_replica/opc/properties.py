from dataclasses import dataclass, field

from word_replica.domain.enums import MetadataMode
from word_replica.opc.package_reader import DocxPackage

CP = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
}
EP = {"ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}
CUST = {
    "c": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}


@dataclass(slots=True)
class DocumentProperties:
    title: str | None = None
    subject: str | None = None
    keywords: str | None = None
    category: str | None = None
    language: str | None = None
    creator: str | None = None
    company: str | None = None
    created: str | None = None
    modified: str | None = None
    revision: str | None = None
    total_editing_time: str | None = None
    custom: dict[str, str] = field(default_factory=dict)


def select_preservable_properties(
    props: DocumentProperties,
    explicit_author: bool = False,
    custom_allowlist: tuple[str, ...] = (),
) -> dict[str, str]:
    names = ("title", "subject", "keywords", "category", "language")
    selected = {name: value for name in names if (value := getattr(props, name))}
    if explicit_author and props.creator:
        selected["creator"] = props.creator
    if explicit_author and props.company:
        selected["company"] = props.company
    for key in custom_allowlist:
        if key in props.custom:
            selected[f"custom:{key}"] = props.custom[key]
    return selected


def build_output_metadata(
    props: DocumentProperties,
    mode: MetadataMode,
    preserve_author_fields: bool,
    custom_allowlist: tuple[str, ...],
) -> dict[str, str]:
    if mode is MetadataMode.FRESH:
        return {}
    return select_preservable_properties(props, preserve_author_fields, custom_allowlist)


def _text(root, xpath: str, ns: dict[str, str]) -> str | None:
    nodes = root.xpath(xpath, namespaces=ns)
    return nodes[0].text if nodes and nodes[0].text is not None else None


def read_properties(package: DocxPackage) -> DocumentProperties:
    props = DocumentProperties()
    if "docProps/core.xml" in package.parts:
        root = package.read_xml("docProps/core.xml")
        props.title = _text(root, "//dc:title", CP)
        props.subject = _text(root, "//dc:subject", CP)
        props.keywords = _text(root, "//cp:keywords", CP)
        props.category = _text(root, "//cp:category", CP)
        props.language = _text(root, "//dc:language", CP)
        props.creator = _text(root, "//dc:creator", CP)
        props.created = _text(root, "//dcterms:created", CP)
        props.modified = _text(root, "//dcterms:modified", CP)
        props.revision = _text(root, "//cp:revision", CP)
    if "docProps/app.xml" in package.parts:
        root = package.read_xml("docProps/app.xml")
        props.company = _text(root, "//ep:Company", EP)
        props.total_editing_time = _text(root, "//ep:TotalTime", EP)
    if "docProps/custom.xml" in package.parts:
        root = package.read_xml("docProps/custom.xml")
        for node in root.xpath("//c:property", namespaces=CUST):
            child = next(iter(node), None)
            if child is not None and child.text is not None and node.get("name"):
                props.custom[node.get("name")] = child.text
    return props
