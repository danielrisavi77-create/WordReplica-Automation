"""Comparing what the reader actually sees.

Every other measurement in the lab compares structure -- the document model,
then the package behind it. Neither can tell you a rebuild reflowed onto an
extra page, or that a floating image landed two centimetres to the left. Only
rendering can, which is why this is the one lane that needs real Microsoft
Word: both documents are exported to PDF by the same Word build, so its own
layout quirks appear on both sides and cancel.

Two answers are kept apart on purpose:

* **Pagination differs** -- the document is a different length, or the same
  length with different text on the pages. That is a layout failure, and it
  points at margins, section properties or a floating object, not at rendering.
* **Pixels differ** -- pagination matches and the pages still do not look the
  same. Usually smaller, sometimes only antialiasing, and a different thing to
  go and look at.

Collapsing them into one verdict would hide which happened. And a comparison
that could not complete is never a match: "not verified" and "verified the
same" are different answers.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Any, Callable

__all__ = [
    "VisualOutcome",
    "VisualVerdict",
    "classify_visual",
    "compare_visually",
]


class VisualOutcome(StrEnum):
    MATCH = "MATCH"
    PAGINATION_DIFFERS = "PAGINATION_DIFFERS"
    PIXELS_DIFFER = "PIXELS_DIFFER"
    UNVERIFIED = "UNVERIFIED"


@dataclass(frozen=True, slots=True)
class VisualVerdict:
    outcome: VisualOutcome
    source_pages: int | None = None
    output_pages: int | None = None
    changed_pixel_ratio: float | None = None
    first_divergent_page: int | None = None
    detail: str = ""


def _first_text_divergence(source_pages: list[str], output_pages: list[str]) -> int | None:
    """First page whose text differs, ignoring whitespace.

    PDF text extraction inserts and drops spaces around glyph runs, so the
    existing partition comparison comes down to non-whitespace characters and
    so does this.
    """
    for index, (left, right) in enumerate(zip(source_pages, output_pages), start=1):
        if "".join(left.split()) != "".join(right.split()):
            return index
    if len(source_pages) != len(output_pages):
        return min(len(source_pages), len(output_pages)) + 1
    return None


def classify_visual(
    source_pages: list[str],
    output_pages: list[str],
    render_result: Any,
) -> VisualVerdict:
    """Turn page texts and a render comparison into one verdict."""
    divergent_page = _first_text_divergence(source_pages, output_pages)
    paginated_the_same = (
        divergent_page is None and getattr(render_result, "page_count_match", True)
    )

    worst = None
    for metric in getattr(render_result, "metrics", ()) or ():
        if not getattr(metric, "within_tolerance", True):
            worst = metric
            break
    ratio = getattr(worst, "changed_pixel_ratio", None) if worst is not None else None

    if not paginated_the_same:
        return VisualVerdict(
            VisualOutcome.PAGINATION_DIFFERS,
            source_pages=len(source_pages),
            output_pages=len(output_pages),
            changed_pixel_ratio=ratio,
            first_divergent_page=divergent_page,
            detail=f"{len(source_pages)} page(s) became {len(output_pages)}"
            if len(source_pages) != len(output_pages)
            else f"page {divergent_page} carries different text",
        )

    if not getattr(render_result, "within_tolerance", True):
        return VisualVerdict(
            VisualOutcome.PIXELS_DIFFER,
            source_pages=len(source_pages),
            output_pages=len(output_pages),
            changed_pixel_ratio=ratio,
            first_divergent_page=getattr(worst, "page", None),
            detail="pagination matches; rendered pages differ beyond tolerance",
        )

    return VisualVerdict(
        VisualOutcome.MATCH,
        source_pages=len(source_pages),
        output_pages=len(output_pages),
        detail="pagination and rendering match",
    )


def compare_visually(
    source: Path,
    output: Path,
    qa_dir: Path,
    *,
    exporter: Callable[..., Any] | None = None,
    comparer: Callable[..., Any] | None = None,
    page_texts: Callable[[Path], list[str]] | None = None,
    dpi: int = 144,
    changed_pixel_tolerance: float = 0.001,
    mae_tolerance: float = 0.25,
) -> VisualVerdict:
    """Render both documents with Word and compare what came out.

    The pair is exported in one Word session rather than two: each round trip
    costs tens of seconds, and both documents have to meet the same Word build
    anyway for its layout quirks to cancel.
    """
    qa_dir = Path(qa_dir)
    qa_dir.mkdir(parents=True, exist_ok=True)
    source_pdf, output_pdf = qa_dir / "source.pdf", qa_dir / "output.pdf"

    if exporter is None:
        from word_replica.qa.word_render import export_docx_pair_to_pdf_with_word as exporter
    if comparer is None:
        from word_replica.qa.render import compare_pdfs as comparer
    if page_texts is None:
        from word_replica.qa.golden_audit import extract_pdf_page_texts as page_texts

    try:
        exporter(Path(source), source_pdf, Path(output), output_pdf, visible=False)
        rendered = comparer(
            source_pdf,
            output_pdf,
            qa_dir,
            dpi=dpi,
            changed_pixel_tolerance=changed_pixel_tolerance,
            mae_tolerance=mae_tolerance,
        )
        source_page_texts = page_texts(source_pdf)
        output_page_texts = page_texts(output_pdf)
    except Exception as exc:
        # A comparison that did not finish is never a match.
        return VisualVerdict(VisualOutcome.UNVERIFIED, detail=f"{type(exc).__name__}: {exc}")

    return classify_visual(source_page_texts, output_page_texts, rendered)
