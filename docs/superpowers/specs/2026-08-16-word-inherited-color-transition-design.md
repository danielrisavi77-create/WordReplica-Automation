# Word Inherited Color Transition Design

## Problem

The latest retained Golden run (`20260816T052406Z_af539637`) has 45 G2 findings. Forty-four are explained by caption runs whose source changes from an explicit `1B2F4B` color to no direct color. The blueprint represents that transition correctly, but Word keeps the previous direct color on a collapsed insertion range. `Font.Reset()` before insertion therefore does not clear the formatting of text that does not exist yet, and the newly inserted suffix inherits the caption-prefix color.

## Evidence

- G0, G1, G5, and G7 remain passing.
- G2 has 45 findings; the first is `174/runs/0/text_len`, expected 10 and actual 63.
- The source paragraph has a colored caption prefix and an uncolored suffix.
- The blueprint emits color only for the prefix and omits color for the suffix.
- The interactive renderer resets the collapsed range before insertion and applies color after insertion only when the next run has an explicit non-automatic color.

## Considered Approaches

1. **Targeted post-insert reset on explicit-to-default color transitions (selected).** Track the transition in `ApplyRunProperties`. After the next insertion expands the Word range over real text, reset that inserted range and reapply the active run properties. This adds work only at the diagnosed transitions and preserves semantic absence of direct color.
2. **Assign Word automatic color when the next run omits color.** This is smaller but may serialize an explicit `w:color w:val="auto"`, replacing the current mismatch with a different G2 mismatch.
3. **Reset every inserted range.** This is broad, slower, and risks changing unrelated properties on hundreds of correct runs.

## Design

`InteractiveWordController` stores a one-shot boolean indicating that the next inserted range must shed inherited direct formatting. `ApplyRunProperties` computes it from the previous and new run-property dictionaries: it is true only when the previous run has an explicit color other than `auto` or `none`, and the new run does not.

`_apply_post_insert_run_properties` consumes the flag while the active range covers newly inserted text. It calls Word's `Font.Reset()` once, then uses the existing post-insert repair path to reapply every explicit property represented by the new run. The flag is cleared before the COM call so a failed operation cannot leak stale transition state into a later run.

This logic also applies to page-break and cached-field insertion because they already use the same post-insert repair method. No new Word process lifecycle behavior is introduced.

## Regression Test

A Word-like range fake models the observed asymmetry: `Font.Reset()` on a collapsed range cannot clear inherited color, while reset on an expanded post-insert range clears the newly inserted text. The test inserts a colored caption prefix and then an uncolored suffix. Before the fix, both pieces remain colored. After the fix, only the prefix is colored and the suffix has no direct color.

## Acceptance Criteria

- The new regression test is observed failing before production code changes and passing afterward.
- The complete local pytest suite passes.
- The real Microsoft Word Golden run preserves every currently passing gate.
- G2 decreases from 45 findings, with the diagnosed caption-color findings removed.
- The Golden source SHA-256 remains `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`.
- A checkpoint is committed and pushed to `automation-dev` only after measurable Golden improvement.
