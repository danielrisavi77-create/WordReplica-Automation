# WordReplica Fidelity Contract

## Status
Formal specification. Describes what WordReplica already measures (G0-G9 in the Golden pipeline, L0-L4 in the per-run QA pipeline) as an explicit contract, rather than an implicit consequence of gate code. This document does not introduce new gates; it names and bounds the guarantee the existing gates already provide, and states what they do not provide.

## Goal
"100% isti dokument" is not a testable claim on its own. This contract makes it testable by splitting document equivalence into four independently checkable dimensions, defining exactly which automated gate produces evidence for each dimension, and stating the scope within which a PASS result is meaningful.

A WordReplica result is reported as `FULL PASS` only when every dimension below has passing evidence **and** the run's [Environment Fingerprint](../../../src/word_replica/qa/environment.py) is recorded alongside it. A `FULL PASS` is a claim about one output document, verified against one source document, inside one recorded environment — not a claim that holds for every Word build, OS, or font set.

## The four fidelity dimensions

| Dimension | Question it answers | Evidence source (this repo) |
|---|---|---|
| Structural | Do all elements, relationships, and objects the source declared still exist in the output, with the same identity and nesting? | G1 (`l1_projection` — document structure), G3 (table geometry/cell properties), G4 (image assets and drawing geometry), G6 (headers/footers), G7 (fields, bookmarks, footnotes/endnotes — existence and linkage) |
| Visual | Do pages, breaks, positions, fonts, and spacing render identically? | G8 (`compare_page_text_partitions` — pagination and page-text partition), G9 (`build_visual_gate` — pixel-level PDF comparison at 144 DPI) |
| Semantic | Do elements carry the same meaning and relationships, not just the same surface text? | G0 (`l0_projection` — content), G7 (field *instruction* and *result_text*, not just presence; bookmark start/end paths) |
| Functional | Does the reconstructed content actually work — fields resolve, numbering is live (not hand-typed), references point at real targets? | G7 (field/bookmark/note linkage), G8 (pagination is a *consequence* of correct functional layout, not checked independently) |

Several gates contribute evidence to more than one dimension; the table above lists a gate under every dimension it bears on, not a single "owning" dimension. G2 (typography and paragraph/run formatting) underlies both Structural and Visual fidelity — it is evidence that the two dimensions must agree with, but it is not itself declared as a top-level dimension in the source strategy document, so it is folded into Structural here (formatting is a property of structural elements) with a note that G9 is the independent visual check that would catch a G2 gap G2 itself missed.

Golden gate reference: `AGENTS.md` § "Golden #1 gates"; implementation: `scripts/codex_automation/audit.py::build_model_gates`, `build_visual_gate`, `compare_page_text_partitions`.

## Relationship to L0-L4 (per-run QA)

G0-G9 (this contract) run in the Golden CI pipeline against one pinned source document and require a real Microsoft Word render. L0-L4 (`word_replica/qa/policy.py::run_l0_l3`, plus the render/L4 comparison in `interactive_rebuild.py`/`rebuild.py`) run on every ordinary rebuild or interactive run, against whatever document the user supplied, and are the fast/local subset of the same idea:

| L-level (per-run QA) | Covers | Nearest G-gate(s) |
|---|---|---|
| L0 | Content | G0 |
| L1 | Structure | G1 (not G3/G4/G6/G7 — those are Golden-only, more expensive checks) |
| L2 | Typography/paragraph formatting | G2 |
| L3 | Page setup and sections | G5 |
| L4 (render) | Visual (PDF pixel compare) | G9 |

L0-L4 is not a lighter *version* of the fidelity contract with a different bar — it is a subset of the same gates, chosen because they're cheap enough to run on every document without a pinned Golden baseline. A document that passes L0-L4 has not been checked against G3/G4/G6/G7/G8; `qa_report.html` should not be read as implying table, image, header/footer, field, or pagination fidelity unless those specific gates ran.

## Tolerances (current, from `audit.py` defaults)

- Visual (G9): rendered at 144 DPI; `changed_pixel_ratio <= 0.001` and `mean_absolute_error <= 0.25` for a strict pass. Two documented fallback acceptance modes exist (`legacy_antialiasing`, `blurred_antialiasing`) with looser bounds, used only to absorb renderer antialiasing noise — every use of a fallback mode is recorded per-page in `gate_details.G9.details.acceptance_mode_counts` and `blur_assisted_pages`, so a PASS achieved through a fallback mode is visible, not hidden.
- Pagination (G8): exact page count match plus exact non-whitespace character partition per page. No tolerance — any difference fails the gate.
- Structural/semantic gates (G0, G1, G3, G4, G6, G7): exact projection equality (`compare_projection` in `qa/policy.py`); no tolerance band.

## Scope and explicit non-goals

- **This is an empirical guarantee, not a formal proof.** Microsoft Word's layout engine is closed-source and not guaranteed deterministic across builds, even nominally identical ones. A `FULL PASS` means: verified equivalent, under this contract's gates, on the Word build/OS/font set recorded in that run's Environment Fingerprint. It is not a claim that any other environment would reproduce the same result — see `capture_environment_fingerprint()` in `qa/environment.py`.
- **Compatibility Mode is recorded, not yet gated.** `audit_docx_pair` now records `Document.CompatibilityMode` for both source and output (`compatibility_mode` key in `golden_report.json`), but a mismatch does not currently fail a gate. A source/output `CompatibilityMode` mismatch is a real fidelity risk (it changes Word's own layout rules) and should be promoted to a gate — or folded into G5/G9 — once there is a real corpus document that exercises it; tracked as follow-up, not implemented here.
- **Promotion now requires environment stability, not just commit stability.** `scripts/codex_automation/state.py::evaluate_run` requires two consecutive `FULL PASS` results on the *same commit* **and** the *same Word build* before `promotion_ready` is set. A Word update between runs resets the counter even if the commit hasn't changed.
- **Macros, custom XML, embedded OLE/Excel objects, and SmartArt are not covered by any gate today.** They are listed in the source strategy document's "Document Digital Twin" wishlist but have no projection function in `word_replica/qa/`. A `FULL PASS` makes no claim about them.
- **The contract is defined against one Golden document today** (`golden/Glavna verzija rektorova (grupno)(1).docx`). A gate passing on that document is evidence the gate logic is correct for the element types that document contains — it is not evidence the gate generalizes to element types the document doesn't exercise (e.g. multi-column sections, RTL text, embedded charts). Closing that gap is the Golden corpus expansion tracked separately in the same plan.

## How a report should be read

A `golden_report.json` or `qa_report.html` fully satisfies this contract when, and only when:
1. `full_pass` is `true` (all ten gates passed, per `build_golden_report`), and
2. `environment` is present and its `word.available` is `true` (the run was verified against a real Word render, not degraded to a fallback), and
3. the reader knows which document(s) were exercised, per the Golden corpus coverage caveat above.

A report missing any of these should be described by which dimensions it *did* verify (e.g. "L0-L3 verified, no visual/table/image/header gates ran"), not summarized as "fidelity confirmed."
