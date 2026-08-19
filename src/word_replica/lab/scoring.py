"""Complexity scoring: which documents earn time inside Microsoft Word.

Word is the scarce resource in this lab -- roughly 330 to 800 documents a
night against tens of thousands the Word-free lanes can handle -- so the
ranking is what decides whether that budget buys new information or another
long report that exercises nothing.

The score is the declared 35/25/25/15 split:

    total = 0.35 * structural + 0.25 * rarity + 0.25 * failure + 0.15 * layout

Three deliberate choices:

* **Log saturation on every structural count.** Without it a 10,000-paragraph
  plain-text document outranks a two-page nightmare, and the Word budget goes
  to documents that exercise nothing.
* **Rarity is deferred, not guessed.** With no corpus, every feature has
  maximal inverse document frequency and every document scores ~1.0. Rather
  than emit a number that means nothing, ``rarity`` is ``None`` until
  ``document_frequency`` has been computed by a second pass, and ``total``
  renormalises over the components it actually has.
* **Noisy-OR for failure probability, not a fitted model.** With fewer than
  ~10k labelled runs a regression overfits. Noisy-OR is monotone, needs no
  training loop, and every term is explainable.

Everything here is a pure function of one fingerprint row plus corpus-level
statistics. No I/O, no Word, no network.
"""
from __future__ import annotations

from dataclasses import dataclass
from math import exp, log, log1p
from typing import Callable, Mapping

from word_replica.lab.fingerprint import DocxFingerprint

__all__ = [
    "FEATURE_BITS",
    "SCORE_MODEL_VERSION",
    "WEIGHTS",
    "ComplexityScore",
    "ScoringContext",
    "feature_mask",
    "features_from_mask",
    "score_document",
]

SCORE_MODEL_VERSION = 1

WEIGHTS: dict[str, float] = {
    "structural": 0.35,
    "rarity": 0.25,
    "failure": 0.25,
    "layout": 0.15,
}

# Rarity calibration. Larger K means a document needs more rare features before
# it saturates; 40 puts a document carrying three genuinely unusual features
# around 0.5 on a corpus of a few tens of thousands.
_RARITY_K = 40.0


# --- feature vocabulary ------------------------------------------------------
# APPEND ONLY: bit positions are persisted per document in feature_mask, and
# document_frequency / failure_stats are keyed by them. Reordering silently
# reinterprets every stored row and every accumulated failure statistic.
_FEATURES: tuple[tuple[str, Callable[[DocxFingerprint], bool]], ...] = (
    ("tables", lambda f: f.table_count > 0),
    ("nested_tables", lambda f: f.max_table_depth >= 2),
    ("merged_cells", lambda f: f.merged_cell_count > 0),
    ("multi_section", lambda f: f.section_count > 1),
    ("multi_column", lambda f: f.multicolumn_section_count > 0),
    ("inline_images", lambda f: f.drawing_inline_count > 0),
    ("floating_images", lambda f: f.drawing_anchor_count > 0),
    ("behind_text", lambda f: f.behind_text_count > 0),
    ("textboxes", lambda f: f.textbox_count > 0),
    ("vml", lambda f: f.vml_shape_count > 0),
    ("ole", lambda f: f.ole_object_count > 0),
    ("charts", lambda f: f.chart_count > 0),
    ("chartex", lambda f: f.chartex_count > 0),
    ("smartart", lambda f: f.smartart_count > 0),
    ("equations", lambda f: f.omml_count > 0),
    ("fields", lambda f: f.field_count > 0),
    ("pagination_fields", lambda f: f.pagination_field_count > 0),
    ("bookmarks", lambda f: f.bookmark_count > 0),
    ("hyperlinks", lambda f: f.hyperlink_count > 0),
    ("footnotes", lambda f: f.footnote_count > 0),
    ("endnotes", lambda f: f.endnote_count > 0),
    ("comments", lambda f: f.comment_count > 0),
    ("comment_replies", lambda f: f.comment_reply_count > 0),
    ("revisions", lambda f: f.revision_count > 0),
    ("content_controls", lambda f: f.sdt_count > 0),
    ("multilevel_numbering", lambda f: f.numbering_max_level >= 2),
    ("headers", lambda f: f.header_count > 0),
    ("footers", lambda f: f.footer_count > 0),
    ("even_odd_headers", lambda f: f.even_odd_headers),
    ("frames", lambda f: f.frame_pr_count > 0),
    ("auto_table_layout", lambda f: f.tbl_layout_auto_count > 0),
    ("embedded_fonts", lambda f: f.embedded_font_count > 0),
    ("custom_xml", lambda f: f.custom_xml_part_count > 0),
    ("alternate_content", lambda f: f.mc_alternate_content_count > 0),
    ("external_refs", lambda f: f.external_rel_count > 0),
    ("document_protection", lambda f: f.doc_protection),
    ("track_changes_on", lambda f: f.track_changes_on),
    ("legacy_compat", lambda f: bool(f.compat_mode) and f.compat_mode < 15),
    ("unknown_namespace", lambda f: f.unknown_ns_count > 0),
    ("broken_content_types", lambda f: f.missing_content_type_count > 0),
    ("broken_relationships", lambda f: f.missing_rel_target_count > 0),
)

FEATURE_BITS: tuple[str, ...] = tuple(name for name, _ in _FEATURES)

# Cold-start priors for the failure component, seeded from what the codebase
# already knows rather than left flat at 0.5. Features that interactive/
# capabilities.py and interactive/preflight.py classify as UNSUPPORTED get a
# high prior; the hand-written failure classes in remote_harness/contracts.py
# (INVALID_WORD_HEADER_XML, IMAGE_FORMAT_UNDEFINED, IMAGE_RELATIONSHIP_ERROR)
# name real historical failures and seed theirs.
_FAILURE_PRIOR: dict[str, float] = {
    # Read these as "probability this feature is what breaks a document that
    # has it", not as a vague fragility rating. Noisy-OR multiplies survival
    # probabilities, so a dozen features at 0.5 would saturate every document
    # at 0.999 and the component would rank nothing. The scale is calibrated so
    # a document carrying three genuinely fragile features lands near 0.5.
    "ole": 0.35,
    "smartart": 0.35,
    "custom_xml": 0.30,
    "chartex": 0.30,
    "vml": 0.25,
    "comment_replies": 0.25,
    "unknown_namespace": 0.25,
    "broken_content_types": 0.30,
    "broken_relationships": 0.30,
    "charts": 0.20,
    "textboxes": 0.20,
    "behind_text": 0.18,
    "floating_images": 0.15,
    "content_controls": 0.12,
    "frames": 0.12,
    "embedded_fonts": 0.10,
    "alternate_content": 0.10,
    "equations": 0.10,
    "revisions": 0.08,
    "comments": 0.06,
    "auto_table_layout": 0.08,
    "nested_tables": 0.08,
    "external_refs": 0.05,
    "even_odd_headers": 0.04,
    "legacy_compat": 0.04,
    "pagination_fields": 0.04,
}
_DEFAULT_FAILURE_PRIOR = 0.01

# (attribute, weight, saturation cap). Caps are the p99 the score treats as
# "as complex as this dimension gets"; a value at the cap contributes its full
# weight. Bootstrapped from the LibreOffice regression corpus and recalibrated
# once real corpus percentiles exist.
_STRUCTURAL_TERMS: tuple[tuple[str, float, float], ...] = (
    ("paragraph_count", 0.06, 400),
    ("run_count", 0.06, 1200),
    ("table_count", 0.10, 20),
    ("max_table_depth", 0.10, 4),
    ("merged_cell_count", 0.08, 60),
    ("section_count", 0.06, 8),
    ("drawing_anchor_count", 0.08, 15),
    ("drawing_inline_count", 0.04, 25),
    ("field_count", 0.08, 40),
    ("bookmark_count", 0.04, 40),
    ("numbering_max_level", 0.06, 9),
    ("style_def_count", 0.06, 120),
    ("sdt_count", 0.06, 10),
)
# Two terms combine more than one counter, so they are computed separately.
_STRUCTURAL_HEADER_WEIGHT, _STRUCTURAL_HEADER_CAP = 0.06, 6
_STRUCTURAL_NOTE_WEIGHT, _STRUCTURAL_NOTE_CAP = 0.06, 30

# (label, numerator, denominator, weight) -- densities, saturating at 1.0.
# These are the signals Word's layout engine is empirically fragile on.
_LAYOUT_TERMS: tuple[tuple[str, Callable[[DocxFingerprint], float], float, float], ...] = (
    ("floating drawings", lambda f: f.drawing_anchor_count, 5, 0.16),
    ("behind-text drawings", lambda f: f.behind_text_count, 2, 0.10),
    ("textboxes and shapes", lambda f: f.textbox_count + f.vml_shape_count, 3, 0.10),
    ("auto-layout tables", lambda f: f.tbl_layout_auto_count, 4, 0.12),
    ("multi-column sections", lambda f: f.multicolumn_section_count, 1, 0.10),
    ("pagination fields", lambda f: f.pagination_field_count, 3, 0.12),
    ("keep-together density", lambda f: f.keep_together_count / max(1, f.paragraph_count) / 0.30, 1, 0.08),
    ("tab leaders", lambda f: f.tab_leader_count, 10, 0.06),
    ("complex scripts", lambda f: f.rtl_run_count, 20, 0.08),
    ("frames", lambda f: f.frame_pr_count, 1, 0.04),
    ("compatibility settings", lambda f: f.compat_setting_count, 20, 0.04),
)


@dataclass(frozen=True, slots=True)
class ScoringContext:
    """Corpus-level statistics the per-document score needs.

    ``document_frequency`` and ``failure_stats`` are keyed by feature bit index
    and are produced by a single GROUP BY over the finished corpus table -- which
    is why they arrive on a second pass rather than during ingest.
    """

    corpus_n: int
    document_frequency: Mapping[int, int]
    failure_stats: Mapping[int, tuple[int, int]]

    @property
    def has_corpus_statistics(self) -> bool:
        return self.corpus_n > 0 and bool(self.document_frequency)


@dataclass(frozen=True, slots=True)
class ComplexityScore:
    structural: float
    rarity: float | None
    failure: float
    layout: float
    total: float
    model_version: int
    feature_mask: int
    top_contributors: tuple[tuple[str, float], ...]


def feature_mask(fingerprint: DocxFingerprint) -> int:
    mask = 0
    for index, (_, present) in enumerate(_FEATURES):
        try:
            if present(fingerprint):
                mask |= 1 << index
        except (TypeError, AttributeError):
            continue
    return mask


def features_from_mask(mask: int) -> frozenset[str]:
    return frozenset(name for index, name in enumerate(FEATURE_BITS) if mask >> index & 1)


def _saturate(value: float, cap: float) -> float:
    """Log saturation, so bulk cannot outrank complexity."""
    if value <= 0 or cap <= 0:
        return 0.0
    return min(1.0, log1p(value) / log1p(cap))


def structural_score(fingerprint: DocxFingerprint) -> tuple[float, list[tuple[str, float]]]:
    total = 0.0
    terms: list[tuple[str, float]] = []
    for attribute, weight, cap in _STRUCTURAL_TERMS:
        contribution = weight * _saturate(float(getattr(fingerprint, attribute, 0) or 0), cap)
        total += contribution
        if contribution:
            terms.append((attribute.replace("_", " "), contribution))

    headers = fingerprint.header_count + fingerprint.footer_count
    notes = fingerprint.footnote_count + fingerprint.endnote_count
    for label, value, weight, cap in (
        ("headers and footers", headers, _STRUCTURAL_HEADER_WEIGHT, _STRUCTURAL_HEADER_CAP),
        ("footnotes and endnotes", notes, _STRUCTURAL_NOTE_WEIGHT, _STRUCTURAL_NOTE_CAP),
    ):
        contribution = weight * _saturate(float(value), cap)
        total += contribution
        if contribution:
            terms.append((label, contribution))
    return min(1.0, total), terms


def rarity_score(fingerprint: DocxFingerprint, context: ScoringContext) -> tuple[float | None, list[tuple[str, float]]]:
    """Inverse document frequency over the feature set.

    Returns None until the corpus has been counted: with an empty
    ``document_frequency`` every feature is maximally rare and every document
    scores ~1.0, which is worse than admitting we do not know yet.
    """
    if not context.has_corpus_statistics:
        return None, []

    mask = feature_mask(fingerprint)
    accumulated = 0.0
    terms: list[tuple[str, float]] = []
    for index, name in enumerate(FEATURE_BITS):
        if not mask >> index & 1:
            continue
        df = context.document_frequency.get(index, 0)
        idf = log((context.corpus_n + 1) / (df + 1))
        accumulated += idf
        terms.append((f"rare: {name}", idf))
    score = 1.0 - exp(-accumulated / _RARITY_K)
    scale = WEIGHTS["rarity"] * score / accumulated if accumulated else 0.0
    return score, [(name, value * scale) for name, value in terms]


def failure_score(fingerprint: DocxFingerprint, context: ScoringContext) -> tuple[float, list[tuple[str, float]]]:
    """Noisy-OR over per-feature failure probabilities.

    Laplace smoothing means a feature seen once and failed once is treated as
    2/3, not 1.0 -- one observation is not a law.
    """
    mask = feature_mask(fingerprint)
    survival = 1.0
    terms: list[tuple[str, float]] = []
    for index, name in enumerate(FEATURE_BITS):
        if not mask >> index & 1:
            continue
        observed = context.failure_stats.get(index)
        if observed:
            runs, fails = observed
            probability = (fails + 1) / (runs + 2)
        else:
            probability = _FAILURE_PRIOR.get(name, _DEFAULT_FAILURE_PRIOR)
        survival *= 1.0 - probability
        terms.append((f"known-fragile: {name}", probability))
    score = min(0.99, 1.0 - survival)
    if not terms:
        return score, []
    scale = WEIGHTS["failure"] * score / sum(value for _, value in terms)
    return score, [(name, value * scale) for name, value in terms]


def layout_sensitivity_score(fingerprint: DocxFingerprint) -> tuple[float, list[tuple[str, float]]]:
    total = 0.0
    terms: list[tuple[str, float]] = []
    for label, numerator, denominator, weight in _LAYOUT_TERMS:
        try:
            value = float(numerator(fingerprint))
        except (TypeError, ZeroDivisionError):
            continue
        contribution = weight * min(1.0, value / denominator if denominator else 0.0)
        total += contribution
        if contribution:
            terms.append((label, contribution))
    return min(1.0, total), terms


def score_document(fingerprint: DocxFingerprint, context: ScoringContext) -> ComplexityScore:
    structural, structural_terms = structural_score(fingerprint)
    rarity, rarity_terms = rarity_score(fingerprint, context)
    failure, failure_terms = failure_score(fingerprint, context)
    layout, layout_terms = layout_sensitivity_score(fingerprint)

    components = {"structural": structural, "failure": failure, "layout": layout}
    if rarity is not None:
        components["rarity"] = rarity
    # Renormalise over the components we actually have, so a deferred rarity
    # depresses ranks uniformly instead of shifting them.
    weight_total = sum(WEIGHTS[name] for name in components)
    total = sum(WEIGHTS[name] * value for name, value in components.items()) / weight_total

    contributions = [
        *((name, value * WEIGHTS["structural"] / max(structural, 1e-9) * structural) for name, value in structural_terms),
        *((name, value * WEIGHTS["layout"] / max(layout, 1e-9) * layout) for name, value in layout_terms),
        *failure_terms,
        *rarity_terms,
    ]
    contributions.sort(key=lambda item: item[1], reverse=True)
    top = tuple(contributions[:5])
    while len(top) < 5:
        top = top + (("", 0.0),)

    return ComplexityScore(
        structural=structural,
        rarity=rarity,
        failure=failure,
        layout=layout,
        total=min(1.0, max(0.0, total)),
        model_version=SCORE_MODEL_VERSION,
        feature_mask=feature_mask(fingerprint),
        top_contributors=top,
    )
