"""Complexity scoring.

The score decides which documents earn the scarce resource in this lab: time
inside Microsoft Word. It has to be monotone, explainable, and immune to the
obvious failure mode -- a 300-page plain-text document outranking a two-page
nightmare because it has more paragraphs.
"""
import pytest

from word_replica.lab.fingerprint import DocxFingerprint
from word_replica.lab.scoring import (
    FEATURE_BITS,
    SCORE_MODEL_VERSION,
    WEIGHTS,
    ScoringContext,
    feature_mask,
    features_from_mask,
    score_document,
)


def _fp(**kwargs) -> DocxFingerprint:
    base = dict(sha256="a" * 64, byte_size=10_000, extractor_ok=True, risk_class="VALID")
    return DocxFingerprint(**{**base, **kwargs})


EMPTY = ScoringContext(corpus_n=0, document_frequency={}, failure_stats={})


def test_weights_are_the_declared_split():
    assert WEIGHTS == {"structural": 0.35, "rarity": 0.25, "failure": 0.25, "layout": 0.15}
    assert sum(WEIGHTS.values()) == pytest.approx(1.0)


def test_total_stays_in_range_for_an_empty_and_an_extreme_document():
    plain = score_document(_fp(), EMPTY)
    extreme = score_document(
        _fp(
            paragraph_count=100_000, run_count=200_000, table_count=500, max_table_depth=9,
            merged_cell_count=5_000, section_count=200, drawing_anchor_count=900,
            drawing_inline_count=900, field_count=800, bookmark_count=400,
            numbering_max_level=9, style_def_count=600, header_count=30, footer_count=30,
            footnote_count=400, endnote_count=400, sdt_count=200, chartex_count=9,
            multicolumn_section_count=40, textbox_count=90, tbl_layout_auto_count=40,
            pagination_field_count=60, frame_pr_count=20, compat_setting_count=40,
        ),
        EMPTY,
    )

    for score in (plain, extreme):
        assert 0.0 <= score.total <= 1.0
        for part in (score.structural, score.failure, score.layout):
            assert 0.0 <= part <= 1.0
    assert extreme.total > plain.total


def test_log_saturation_keeps_bulk_from_beating_complexity():
    # The single most important property. Without log saturation a long plain
    # document dominates the queue and the Word budget is spent on nothing.
    bulk = _fp(paragraph_count=10_000, run_count=20_000, char_count=500_000)
    nightmare = _fp(
        paragraph_count=40, run_count=90, table_count=6, max_table_depth=3,
        merged_cell_count=25, section_count=5, drawing_anchor_count=7,
        field_count=12, bookmark_count=9, numbering_max_level=6,
        style_def_count=40, header_count=3, footer_count=3, sdt_count=4,
    )

    assert score_document(nightmare, EMPTY).structural > score_document(bulk, EMPTY).structural


def test_structural_score_is_monotone_in_nesting():
    flat = _fp(table_count=3, max_table_depth=1)
    nested = _fp(table_count=3, max_table_depth=4)

    assert score_document(nested, EMPTY).structural > score_document(flat, EMPTY).structural


def test_layout_score_rewards_the_signals_word_is_fragile_on():
    calm = _fp(paragraph_count=200)
    fragile = _fp(
        paragraph_count=200, drawing_anchor_count=6, behind_text_count=3,
        textbox_count=4, tbl_layout_auto_count=5, multicolumn_section_count=2,
        pagination_field_count=4, frame_pr_count=2,
    )

    assert score_document(fragile, EMPTY).layout > score_document(calm, EMPTY).layout


# --- rarity is a two-pass quantity -------------------------------------------

def test_rarity_is_deferred_when_the_corpus_is_still_empty():
    # With no corpus every feature looks maximally rare, which would rank every
    # document at 1.0 and make the queue meaningless. Defer instead of lying.
    score = score_document(_fp(chartex_count=1), EMPTY)

    assert score.rarity is None
    # total renormalises over the components it actually has, so a deferred
    # rarity depresses every rank uniformly rather than reordering them.
    present = WEIGHTS["structural"] + WEIGHTS["failure"] + WEIGHTS["layout"]
    assert score.total == pytest.approx(
        (
            WEIGHTS["structural"] * score.structural
            + WEIGHTS["failure"] * score.failure
            + WEIGHTS["layout"] * score.layout
        )
        / present,
        rel=1e-9,
    )


def test_rarity_rewards_features_few_documents_have():
    common = _fp(table_count=1)
    rare = _fp(table_count=1, chartex_count=1)
    context = ScoringContext(
        corpus_n=10_000,
        document_frequency={bit: 9_000 for bit in range(len(FEATURE_BITS))} | {
            FEATURE_BITS.index("chartex"): 3
        },
        failure_stats={},
    )

    assert score_document(rare, context).rarity > score_document(common, context).rarity


def test_rarity_never_divides_by_zero_on_a_feature_the_corpus_has_never_seen():
    # The corpus has been counted, but this particular feature has df == 0.
    # That is the maximum-idf case and it must stay finite and in range.
    context = ScoringContext(
        corpus_n=5,
        document_frequency={FEATURE_BITS.index("tables"): 4},
        failure_stats={},
    )
    score = score_document(_fp(chartex_count=1), context)

    assert score.rarity is not None
    assert 0.0 <= score.rarity <= 1.0


# --- historical failure probability ------------------------------------------

def test_failure_score_reflects_recorded_history():
    bit = FEATURE_BITS.index("chartex")
    quiet = ScoringContext(corpus_n=100, document_frequency={}, failure_stats={bit: (100, 0)})
    noisy = ScoringContext(corpus_n=100, document_frequency={}, failure_stats={bit: (100, 95)})

    assert score_document(_fp(chartex_count=1), noisy).failure > score_document(_fp(chartex_count=1), quiet).failure


def test_failure_score_uses_a_seeded_prior_before_any_history_exists():
    # A cold start must not flatten every document to the same value; the seeds
    # come from features the codebase already declares unsupported or fragile.
    exotic = score_document(_fp(ole_object_count=2, vml_shape_count=3), EMPTY)
    plain = score_document(_fp(paragraph_count=10), EMPTY)

    assert exotic.failure > plain.failure


def test_failure_score_is_clipped_below_one():
    context = ScoringContext(
        corpus_n=10,
        document_frequency={},
        failure_stats={bit: (100, 100) for bit in range(len(FEATURE_BITS))},
    )
    everything = _fp(
        chartex_count=1, ole_object_count=1, vml_shape_count=1, smartart_count=1,
        omml_count=1, sdt_count=1, comment_reply_count=1, embedded_font_count=1,
    )

    assert score_document(everything, context).failure <= 0.99


# --- feature mask ------------------------------------------------------------

def test_feature_bits_are_append_only_and_fit_the_mask():
    assert len(set(FEATURE_BITS)) == len(FEATURE_BITS)
    assert len(FEATURE_BITS) <= 64


def test_feature_mask_round_trips():
    fingerprint = _fp(table_count=2, chartex_count=1, omml_count=4)

    names = features_from_mask(feature_mask(fingerprint))

    assert "tables" in names
    assert "chartex" in names
    assert "equations" in names
    assert "smartart" not in names


# --- explainability ----------------------------------------------------------

def test_top_contributors_explain_the_rank():
    score = score_document(
        _fp(table_count=8, max_table_depth=4, drawing_anchor_count=9, field_count=20),
        EMPTY,
    )

    assert len(score.top_contributors) == 5
    values = [value for _, value in score.top_contributors]
    assert values == sorted(values, reverse=True)
    assert all(isinstance(name, str) for name, _ in score.top_contributors)


def test_score_records_its_model_version():
    assert score_document(_fp(), EMPTY).model_version == SCORE_MODEL_VERSION
