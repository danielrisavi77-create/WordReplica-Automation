"""Word-free structural fingerprinting of a .docx package.

This is what lets the lab rank a corpus it can never open in Word. Three
properties matter more than completeness:

* **Cheap.** The XML layer is a single pass over every part with a hardened
  parser -- roughly 90 ms on a real-world document, so a whole corpus is hours
  rather than months. A full ``DocxParser`` parse is an order of magnitude more
  expensive and is therefore optional.
* **Total.** It never raises. The ingestor feeds it whatever the web returned;
  a corrupt package must produce a row saying so, not stop a batch.
* **Blind to nothing.** ``qa/policy.py`` compares *parsed models*, and
  ``parser/parser.py`` only models what WordReplica reconstructs. A ``w15``
  comment reply is invisible to the parser and to Word 2010 alike. The
  namespace and element histograms here are taken from raw part bytes, so the
  lab can at least tell such a document apart -- which is the precondition for
  the G10 preservation gate catching a silent loss.

Nothing here needs Microsoft Word, a network, or a temporary file.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
from pathlib import Path
from typing import Any, Iterable, Iterator

from word_replica.lab.safety import (
    DEFAULT_LIMITS,
    PackageLimits,
    PackageScan,
    RiskClass,
    RiskTriage,
    content_type_gaps,
    read_package,
    resolve_relationship_target,
    triage_scan,
)

__all__ = [
    "FINGERPRINT_SCHEMA_VERSION",
    "NAMESPACE_BITS",
    "DocxFingerprint",
    "extract_fingerprint",
    "extract_fingerprint_stream",
    "namespaces_from_mask",
]

FINGERPRINT_SCHEMA_VERSION = 1

# Bit index per namespace URI. APPEND ONLY: ns_mask values are persisted in the
# corpus database, so reordering this tuple silently reinterprets every stored
# row. A new namespace goes on the end and nowhere else.
NAMESPACE_BITS: tuple[str, ...] = (
    "w",           # 0  wordprocessingml 2006 main
    "w14",         # 1  Word 2010 extensions   -- Word 2010 understands these
    "w15",         # 2  Word 2013 extensions   -- Word 2010 silently ignores
    "w16cid",      # 3  Word 2016 comment ids  -- ignored
    "w16se",       # 4  Word 2015 symbol ex    -- ignored
    "w16",         # 5  Word 2018 extensions   -- ignored
    "wp",          # 6  wordprocessingDrawing
    "wp14",        # 7
    "wps",         # 8  wordprocessingShape
    "wpg",         # 9  wordprocessingGroup
    "wpi",         # 10 wordprocessingInk
    "a",           # 11 drawingml main
    "a14",         # 12 drawing 2010
    "pic",         # 13
    "c",           # 14 chart
    "cx",          # 15 chartex             -- Word 2010 cannot render
    "m",           # 16 OMML
    "v",           # 17 VML
    "o",           # 18 office
    "w10",         # 19 legacy word
    "mc",          # 20 markup compatibility
    "dgm",         # 21 SmartArt diagram
    "sl",          # 22 schema library
    "asvg",        # 23 SVG                  -- Word 2010 cannot render
    "wne",         # 24 vba/wordml
    "strict",      # 25 ISO-29500 Strict     -- Word 2010 cannot write
    "ct",          # 26 package content types
    "pr",          # 27 package relationships
    "r",           # 28 officeDocument relationships
    "cp",          # 29 core properties
    "dc",          # 30
    "dcterms",     # 31
    "vt",          # 32 docPropsVTypes
    "ep",          # 33 extended properties
    "custom",      # 34 custom properties
    "xdr",         # 35 spreadsheet drawing (embedded workbooks)
    "wpc",         # 36 wordprocessingCanvas
    "cdr",         # 37 chart drawing
    "bib",         # 38 bibliography (standard; was being counted as unknown)
    "cxml",        # 39 customXml part schema (standard; likewise)
)

_NAMESPACE_URIS: dict[str, str] = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "strict": "http://purl.oclc.org/ooxml/wordprocessingml/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "custom": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cdr": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "bib": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
    "cxml": "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
}
_URI_TO_BIT: dict[str, int] = {
    _NAMESPACE_URIS[prefix]: index for index, prefix in enumerate(NAMESPACE_BITS)
}

_W = _NAMESPACE_URIS["w"]
_REL_NS = _NAMESPACE_URIS["pr"]

_MEDIA_SUFFIXES = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg")
_PAGINATION_FIELDS = ("PAGE", "NUMPAGES", "TOC", "PAGEREF", "SECTIONPAGES")


def namespaces_from_mask(mask: int) -> frozenset[str]:
    """Prefixes present in a stored ns_mask -- for reports and explanations."""
    return frozenset(name for index, name in enumerate(NAMESPACE_BITS) if mask >> index & 1)


@dataclass(frozen=True, slots=True)
class DocxFingerprint:
    """One row of the corpus table: everything measurable without Word."""

    # identity
    sha256: str
    byte_size: int
    schema_version: int = FINGERPRINT_SCHEMA_VERSION
    extractor_ok: bool = False
    extract_error: str | None = None

    # risk (from lab.safety)
    risk_class: str = RiskClass.HOSTILE.value
    risk_reasons: tuple[str, ...] = ()
    word_open_allowed: bool = False

    # package envelope
    part_count: int = 0
    xml_part_count: int = 0
    media_count: int = 0
    total_uncompressed_bytes: int = 0
    max_part_bytes: int = 0
    compression_ratio: float = 0.0
    document_xml_bytes: int = 0
    missing_content_type_count: int = 0
    missing_rel_target_count: int = 0
    external_rel_count: int = 0
    rel_type_count: int = 0

    # xml / namespace layer
    ns_mask: int = 0
    unknown_ns_count: int = 0
    unknown_ns_sample: str | None = None
    distinct_qname_count: int = 0
    total_element_count: int = 0
    max_xml_depth: int = 0
    mc_alternate_content_count: int = 0

    # structure
    paragraph_count: int = 0
    run_count: int = 0
    char_count: int = 0
    table_count: int = 0
    max_table_depth: int = 0
    row_count: int = 0
    cell_count: int = 0
    merged_cell_count: int = 0
    section_count: int = 0
    multicolumn_section_count: int = 0

    # drawings and embedded objects
    drawing_inline_count: int = 0
    drawing_anchor_count: int = 0
    behind_text_count: int = 0
    textbox_count: int = 0
    vml_shape_count: int = 0
    ole_object_count: int = 0
    chart_count: int = 0
    chartex_count: int = 0
    smartart_count: int = 0
    omml_count: int = 0

    # references and review
    field_count: int = 0
    pagination_field_count: int = 0
    bookmark_count: int = 0
    hyperlink_count: int = 0
    footnote_count: int = 0
    endnote_count: int = 0
    comment_count: int = 0
    comment_reply_count: int = 0
    revision_count: int = 0
    sdt_count: int = 0

    # styling and layout sensitivity
    numbering_def_count: int = 0
    numbering_max_level: int = 0
    style_def_count: int = 0
    header_count: int = 0
    footer_count: int = 0
    even_odd_headers: bool = False
    keep_together_count: int = 0
    page_break_before_count: int = 0
    frame_pr_count: int = 0
    tbl_layout_auto_count: int = 0
    tab_leader_count: int = 0
    complex_script_run_count: int = 0
    rtl_run_count: int = 0
    embedded_font_count: int = 0
    custom_xml_part_count: int = 0
    compat_setting_count: int = 0
    compat_mode: int | None = None
    doc_protection: bool = False
    track_changes_on: bool = False
    macro_present: bool = False
    activex_count: int = 0

    # model layer (optional, expensive)
    model_fingerprint: str | None = None
    model_error: str | None = None

    # populated later by lab.scoring; NULL until the second pass
    structural_score: float | None = None
    rarity_score: float | None = None
    failure_score: float | None = None
    layout_score: float | None = None
    complexity_score: float | None = None
    score_model_version: int | None = None

    element_counts: dict[str, int] = field(default_factory=dict, compare=False)

    def to_row(self) -> dict[str, Any]:
        """Flat, SQLite-storable projection. Excludes the element histogram,
        which is a diagnostic aid rather than a stored column."""
        row: dict[str, Any] = {}
        for name in self.__slots__:
            if name == "element_counts":
                continue
            value = getattr(self, name)
            if isinstance(value, bool):
                row[name] = int(value)
            elif isinstance(value, tuple):
                row[name] = "\n".join(value)[:2000] or None
            else:
                row[name] = value
        return row


def _sha256_file(path: Path) -> str:
    digest = sha256()
    try:
        with Path(path).open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError:
        return ""
    return digest.hexdigest()


def _w(name: str) -> str:
    """Clark notation, so the same string works as an lxml tag and a count key."""
    return f"{{{_W}}}{name}"


def _ns(prefix: str, name: str) -> str:
    return f"{{{_NAMESPACE_URIS[prefix]}}}{name}"


def _failed(path: Path, reason: str, triage: RiskTriage | None = None) -> DocxFingerprint:
    return DocxFingerprint(
        sha256=_sha256_file(path),
        byte_size=path.stat().st_size if path.exists() else 0,
        extractor_ok=False,
        extract_error=reason[:500],
        risk_class=(triage.risk_class.value if triage else RiskClass.HOSTILE.value),
        risk_reasons=(triage.reasons if triage else (reason,)),
        word_open_allowed=bool(triage.word_open_allowed) if triage else False,
        part_count=triage.envelope.member_count if triage else 0,
    )


def extract_fingerprint(
    path: Path,
    *,
    limits: PackageLimits = DEFAULT_LIMITS,
    scan: PackageScan | None = None,
    triage: RiskTriage | None = None,
    parse_model: bool = True,
) -> DocxFingerprint:
    """Measure one package. Never raises.

    ``scan`` lets a caller that has already read the package (the ingestor
    does, to triage it) pay for the read, parse and tree walk exactly once.

    ``parse_model=False`` skips the ``DocxParser`` layer, which is an order of
    magnitude more expensive than everything else here and is only worth paying
    for on documents that survive scoring. A HOSTILE package never gets the
    model layer regardless of the flag.
    """
    path = Path(path)
    scan = scan if scan is not None else read_package(path, limits=limits, uri_to_bit=_URI_TO_BIT)
    triage = triage or triage_scan(scan, limits=limits)
    if not scan.ok:
        return _failed(path, scan.error or "unreadable package", triage)

    try:
        return _measure(path, scan, triage, parse_model=parse_model)
    except Exception as exc:  # totality is the contract; a batch must not stop
        return _failed(path, f"{type(exc).__name__}: {exc}", triage)


def _measure(
    path: Path,
    scan: PackageScan,
    triage: RiskTriage,
    *,
    parse_model: bool,
) -> DocxFingerprint:
    parts, roots = scan.parts, scan.roots
    counts = scan.element_counts
    xml_part_count = sum(1 for name in parts if name.lower().endswith((".xml", ".rels")))
    lowered_names = {name.lower() for name in parts}

    header_parts = sum(1 for n in lowered_names if n.startswith("word/header") and n.endswith(".xml"))
    footer_parts = sum(1 for n in lowered_names if n.startswith("word/footer") and n.endswith(".xml"))
    media_count = sum(1 for n in lowered_names if n.startswith("word/media/") and n.endswith(_MEDIA_SUFFIXES))
    custom_xml_parts = sum(1 for n in lowered_names if n.startswith("customxml/") and n.endswith(".xml"))
    embedded_fonts = sum(1 for n in lowered_names if n.startswith("word/fonts/"))
    activex = sum(1 for n in lowered_names if n.startswith("word/activex/"))
    charts = sum(1 for n in lowered_names if n.startswith("word/charts/") and n.endswith(".xml"))

    external_rels = 0
    rel_types: set[str] = set()
    missing_targets = 0
    for name, root in roots.items():
        if not name.endswith(".rels"):
            continue
        for node in root.findall(f"{{{_REL_NS}}}Relationship"):
            rel_types.add((node.get("Type") or "").rsplit("/", 1)[-1])
            if node.get("TargetMode") == "External":
                external_rels += 1
                continue
            resolved = resolve_relationship_target(name, node.get("Target") or "")
            if resolved and resolved not in parts:
                missing_targets += 1

    ct_root = roots.get("[Content_Types].xml")
    missing_content_types = len(content_type_gaps(parts, ct_root)) if ct_root is not None else 0

    # w:cols appears in every sectPr, so the element alone says nothing; only a
    # w:num above one is actually a multi-column section.
    multicolumn_sections = 0
    for name, root in roots.items():
        if not name.lower().endswith(".xml"):
            continue
        for node in root.iter(_w("cols")):
            try:
                if int(node.get(_w("num")) or "1") > 1:
                    multicolumn_sections += 1
            except ValueError:
                continue

    instructions = " ".join(scan.field_instructions).upper()
    pagination_fields = sum(instructions.count(token) for token in _PAGINATION_FIELDS)

    settings_root = roots.get("word/settings.xml")
    compat_mode, compat_settings, protection, track_changes, even_odd = _read_settings(settings_root)

    numbering_root = roots.get("word/numbering.xml")
    numbering_defs, numbering_levels = _read_numbering(numbering_root)

    char_count = sum(
        len(node.text or "")
        for name, root in roots.items()
        if name.lower().endswith(".xml")
        for node in root.iter(_w("t"))
    )

    model_fingerprint, model_error = None, None
    if parse_model and triage.risk_class is not RiskClass.HOSTILE:
        model_fingerprint, model_error = _model_fingerprint(path)

    unknown = list(scan.unknown_namespaces)
    envelope = scan.envelope

    return DocxFingerprint(
        sha256=_sha256_file(path),
        byte_size=path.stat().st_size,
        extractor_ok=True,
        extract_error=None,
        risk_class=triage.risk_class.value,
        risk_reasons=triage.reasons,
        word_open_allowed=triage.word_open_allowed,
        part_count=envelope.member_count,
        xml_part_count=xml_part_count,
        media_count=media_count,
        total_uncompressed_bytes=envelope.total_uncompressed_bytes,
        max_part_bytes=envelope.max_part_bytes,
        compression_ratio=envelope.compression_ratio,
        document_xml_bytes=len(parts.get("word/document.xml", b"")),
        missing_content_type_count=missing_content_types,
        missing_rel_target_count=missing_targets,
        external_rel_count=external_rels,
        rel_type_count=len(rel_types),
        ns_mask=scan.ns_mask,
        unknown_ns_count=len(unknown),
        unknown_ns_sample=unknown[0][:200] if unknown else None,
        distinct_qname_count=scan.distinct_qname_count,
        total_element_count=scan.total_element_count,
        max_xml_depth=scan.max_xml_depth,
        mc_alternate_content_count=counts[_ns("mc", "AlternateContent")],
        paragraph_count=counts[_w("p")],
        run_count=counts[_w("r")],
        char_count=char_count,
        table_count=counts[_w("tbl")],
        max_table_depth=scan.max_table_depth,
        row_count=counts[_w("tr")],
        cell_count=counts[_w("tc")],
        merged_cell_count=counts[_w("gridSpan")] + counts[_w("vMerge")],
        section_count=counts[_w("sectPr")],
        multicolumn_section_count=multicolumn_sections,
        drawing_inline_count=counts[_ns("wp", "inline")],
        drawing_anchor_count=counts[_ns("wp", "anchor")],
        behind_text_count=counts[_ns("wp", "wrapNone")],
        textbox_count=counts[_ns("wps", "txbx")] + counts[_w("txbxContent")],
        vml_shape_count=counts[_ns("v", "shape")] + counts[_ns("v", "rect")],
        ole_object_count=counts[_ns("o", "OLEObject")] + counts[_w("object")],
        chart_count=charts + counts[_ns("c", "chart")],
        chartex_count=counts[_ns("cx", "chart")],
        smartart_count=counts[_ns("dgm", "relIds")],
        omml_count=counts[_ns("m", "oMath")],
        field_count=counts[_w("fldSimple")] + counts[_w("instrText")],
        pagination_field_count=pagination_fields,
        bookmark_count=counts[_w("bookmarkStart")],
        hyperlink_count=counts[_w("hyperlink")],
        footnote_count=counts[_w("footnoteReference")],
        endnote_count=counts[_w("endnoteReference")],
        comment_count=counts[_w("commentReference")],
        comment_reply_count=counts[_ns("w15", "commentEx")],
        revision_count=counts[_w("ins")] + counts[_w("del")],
        sdt_count=counts[_w("sdt")],
        numbering_def_count=numbering_defs,
        numbering_max_level=numbering_levels,
        style_def_count=counts[_w("style")],
        header_count=header_parts,
        footer_count=footer_parts,
        even_odd_headers=even_odd,
        keep_together_count=counts[_w("keepNext")] + counts[_w("keepLines")],
        page_break_before_count=counts[_w("pageBreakBefore")],
        frame_pr_count=counts[_w("framePr")],
        tbl_layout_auto_count=counts[_w("tblLayout")],
        tab_leader_count=counts[_w("tabs")],
        complex_script_run_count=counts[_w("cs")] + counts[_w("szCs")],
        rtl_run_count=counts[_w("rtl")] + counts[_w("bidi")],
        embedded_font_count=embedded_fonts,
        custom_xml_part_count=custom_xml_parts,
        compat_setting_count=compat_settings,
        compat_mode=compat_mode,
        doc_protection=protection,
        track_changes_on=track_changes,
        macro_present=any(n in lowered_names for n in ("word/vbaproject.bin", "word/vbadata.xml")),
        activex_count=activex,
        model_fingerprint=model_fingerprint,
        model_error=model_error,
        element_counts=dict(counts),
    )


def _read_settings(root: Any) -> tuple[int | None, int, bool, bool, bool]:
    if root is None:
        return None, 0, False, False, False
    compat_settings = 0
    compat_mode: int | None = None
    for node in root.iter(_w("compatSetting")):
        compat_settings += 1
        if node.get(_w("name")) == "compatibilityMode":
            try:
                compat_mode = int(node.get(_w("val")) or "")
            except ValueError:
                compat_mode = None
    protection = root.find(_w("documentProtection")) is not None
    track_changes = root.find(_w("trackChanges")) is not None
    even_odd = root.find(_w("evenAndOddHeaders")) is not None
    return compat_mode, compat_settings, protection, track_changes, even_odd


def _read_numbering(root: Any) -> tuple[int, int]:
    if root is None:
        return 0, 0
    definitions = sum(1 for _ in root.iter(_w("abstractNum")))
    levels = 0
    for node in root.iter(_w("lvl")):
        try:
            levels = max(levels, int(node.get(_w("ilvl")) or "0") + 1)
        except ValueError:
            continue
    return definitions, levels


def _model_fingerprint(path: Path) -> tuple[str | None, str | None]:
    """The expensive layer. Isolated so its cost and its failures are visible."""
    try:
        from word_replica.parser.parser import DocxParser

        return DocxParser().parse(path).fingerprint(), None
    except Exception as exc:
        return None, f"{type(exc).__name__}: {exc}"[:500]


def extract_fingerprint_stream(
    paths: Iterable[Path],
    *,
    limits: PackageLimits = DEFAULT_LIMITS,
    parse_model: bool = True,
) -> Iterator[DocxFingerprint]:
    """Sequential for now; the ingestor is what parallelises across cores."""
    for path in paths:
        yield extract_fingerprint(Path(path), limits=limits, parse_model=parse_model)
