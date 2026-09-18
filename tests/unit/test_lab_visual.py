"""The dimension the lab has never measured: what the reader actually sees.

Everything so far compares structure -- the document model, then the package
behind it. Neither can tell you a rebuild reflowed onto an extra page, or that a
floating image landed two centimetres to the left. Only rendering can, which is
why this is the one lane that needs real Word.

Two answers are kept apart on purpose. A page-count change is a layout failure:
the document is a different length than the one handed in. Pixel differences at
matching pagination are a rendering difference, usually smaller and sometimes
only antialiasing. Collapsing them into one verdict would hide which happened.
"""
from pathlib import Path

import pytest

from word_replica.lab.visual import (
    VisualOutcome,
    VisualVerdict,
    classify_visual,
    compare_visually,
)


class _Metric:
    def __init__(self, page, ratio, mae=0.0, within=True):
        self.page = page
        self.changed_pixel_ratio = ratio
        self.mean_absolute_error = mae
        self.within_tolerance = within


class _Render:
    def __init__(self, *, page_count_match=True, within=True, metrics=()):
        self.page_count_match = page_count_match
        self.within_tolerance = within
        self.metrics = list(metrics)


# --- what the answers mean ----------------------------------------------------

def test_matching_pages_and_pixels_is_a_match():
    verdict = classify_visual(["a", "b"], ["a", "b"], _Render())

    assert verdict.outcome is VisualOutcome.MATCH
    assert verdict.source_pages == 2
    assert verdict.output_pages == 2


def test_a_different_page_count_is_reported_as_pagination():
    # The document is a different length than the one handed in. That is a
    # layout failure, not a rendering difference, and saying so points at a
    # different cause.
    verdict = classify_visual(["a", "b"], ["a", "b", "c"], _Render(page_count_match=False))

    assert verdict.outcome is VisualOutcome.PAGINATION_DIFFERS
    assert verdict.source_pages == 2
    assert verdict.output_pages == 3


def test_text_moving_between_pages_is_pagination_even_at_equal_counts():
    # Same number of pages, different text on them: the reflow happened, it
    # just happened to land on the same total.
    verdict = classify_visual(["abc", "def"], ["ab", "cdef"], _Render())

    assert verdict.outcome is VisualOutcome.PAGINATION_DIFFERS
    assert verdict.first_divergent_page == 1


def test_pixels_differing_at_matching_pagination_is_its_own_answer():
    verdict = classify_visual(
        ["a"], ["a"], _Render(within=False, metrics=[_Metric(1, 0.02, 0.9, within=False)])
    )

    assert verdict.outcome is VisualOutcome.PIXELS_DIFFER
    assert verdict.changed_pixel_ratio == pytest.approx(0.02)
    assert verdict.first_divergent_page == 1


def test_whitespace_only_page_text_changes_do_not_count_as_reflow():
    # PDF text extraction inserts and drops spaces around glyph runs. The
    # existing partition comparison already ignores that, and so must this.
    verdict = classify_visual(["a b c"], ["abc"], _Render())

    assert verdict.outcome is VisualOutcome.MATCH


# --- never call a missing answer a pass ---------------------------------------

def test_a_failed_export_is_unverified(tmp_path):
    def _explode(*_args, **_kwargs):
        raise RuntimeError("Word went away")

    verdict = compare_visually(
        tmp_path / "src.docx", tmp_path / "out.docx", tmp_path / "qa", exporter=_explode
    )

    assert verdict.outcome is VisualOutcome.UNVERIFIED
    assert "Word went away" in verdict.detail


def test_unverified_is_neither_a_match_nor_a_difference(tmp_path):
    def _explode(*_args, **_kwargs):
        raise RuntimeError("nope")

    verdict = compare_visually(
        tmp_path / "src.docx", tmp_path / "out.docx", tmp_path / "qa", exporter=_explode
    )

    assert verdict.outcome is not VisualOutcome.MATCH
    assert verdict.outcome is not VisualOutcome.PIXELS_DIFFER
    assert verdict.outcome is not VisualOutcome.PAGINATION_DIFFERS


def test_a_comparison_that_raises_is_unverified(tmp_path):
    def _export(source, source_pdf, output, output_pdf, visible=False):
        return source_pdf, output_pdf

    def _boom(*_args, **_kwargs):
        raise ValueError("rasterizer failed")

    verdict = compare_visually(
        tmp_path / "src.docx",
        tmp_path / "out.docx",
        tmp_path / "qa",
        exporter=_export,
        comparer=_boom,
        page_texts=lambda _p: ["a"],
    )

    assert verdict.outcome is VisualOutcome.UNVERIFIED


# --- the whole path, with Word replaced ---------------------------------------

def test_the_pair_is_exported_once_not_twice(tmp_path):
    # Each Word round trip costs tens of seconds. The pair exporter opens one
    # session for both documents, and this lane must use it that way.
    calls = []

    def _export(source, source_pdf, output, output_pdf, visible=False):
        calls.append((source, output))
        return source_pdf, output_pdf

    compare_visually(
        tmp_path / "src.docx",
        tmp_path / "out.docx",
        tmp_path / "qa",
        exporter=_export,
        comparer=lambda *a, **k: _Render(),
        page_texts=lambda _p: ["page"],
    )

    assert len(calls) == 1


def test_the_verdict_carries_what_a_reader_would_ask(tmp_path):
    verdict = compare_visually(
        tmp_path / "src.docx",
        tmp_path / "out.docx",
        tmp_path / "qa",
        exporter=lambda s, sp, o, op, visible=False: (sp, op),
        comparer=lambda *a, **k: _Render(within=False, metrics=[_Metric(3, 0.05, 1.2, within=False)]),
        page_texts=lambda _p: ["one", "two", "three"],
    )

    assert isinstance(verdict, VisualVerdict)
    assert verdict.source_pages == 3
    assert verdict.first_divergent_page == 3
    assert verdict.changed_pixel_ratio == pytest.approx(0.05)
