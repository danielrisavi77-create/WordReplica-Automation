from hashlib import sha256
import posixpath

from word_replica.domain.model import BinaryAsset, RelationshipRef
from word_replica.opc.package_reader import DocxPackage

CT_NS = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}


def content_types(package: DocxPackage) -> tuple[dict[str, str], dict[str, str]]:
    if "[Content_Types].xml" not in package.parts:
        return {}, {}
    root = package.read_xml("[Content_Types].xml")
    defaults = {
        node.get("Extension", "").lower(): node.get("ContentType")
        for node in root.xpath("//ct:Default", namespaces=CT_NS)
    }
    overrides = {
        node.get("PartName", "").lstrip("/"): node.get("ContentType")
        for node in root.xpath("//ct:Override", namespaces=CT_NS)
    }
    return defaults, overrides


def content_type_for(package: DocxPackage, part_name: str) -> str | None:
    defaults, overrides = content_types(package)
    if part_name in overrides:
        return overrides[part_name]
    ext = part_name.rsplit(".", 1)[-1].lower() if "." in part_name else ""
    return defaults.get(ext)


def extract_asset(package: DocxPackage, part_name: str, content_type: str | None = None) -> BinaryAsset:
    data = package.read_bytes(part_name)
    digest = sha256(data).hexdigest()
    return BinaryAsset(
        f"asset_{digest[:16]}",
        part_name,
        content_type if content_type is not None else content_type_for(package, part_name),
        digest,
        data,
    )


def resolve_relationship_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target))


def collect_relationships(package: DocxPackage, source_part: str) -> dict[str, RelationshipRef]:
    refs: dict[str, RelationshipRef] = {}
    for rel_id, rel in package.relationships(source_part).items():
        external = (rel.target_mode or "").lower() == "external"
        target = rel.target if external else resolve_relationship_target(source_part, rel.target)
        refs[f"{source_part}:{rel_id}"] = RelationshipRef(
            rel_id=rel_id,
            rel_type=rel.rel_type,
            target=target,
            external=external,
        )
    return refs
