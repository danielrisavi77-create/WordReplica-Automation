# Word Field-Result Formatting Preservation

## Context and evidence

The accepted GOLDEN #1 run `20260816T023116Z_85b73369` has 114 G2 findings. Its first divergence is `173/runs/0/text_len`: the source paragraph normalizes to one 1,351-character run, while Word splits the output into 1,324 normal characters, one bold REF result, and 26 normal characters.

The source paragraph contains `REF ref_tab_1 \\h` with cached result `1`. Its blueprint is correct: every field token is preceded by `ApplyRunProperties` with `bold: false`, Times New Roman, Cambria complex-script font, and 11 pt size. In the output, however, the result run is `bold: true` because `Fields.Add` derives the live REF result from the bold caption bookmark. Assigning `field.Result.Text` restores the cached text but does not restore its source formatting.

This pattern affects 17 field-containing paragraphs and accounts for 69 of the remaining 114 G2 findings.

## Goal

After creating a Word field and restoring its cached result, apply the blueprint's active run properties to the field result range so the visible result has the same formatting as the source run.

## Non-goals

- Changing field instructions, bookmark targets, cached result text, or field update policy.
- Refreshing fields during reconstruction.
- Repairing caption color inheritance or other non-field G2 findings in this iteration.
- Adding an OOXML post-processing rule.

## Chosen design

`_event_CreateField` will keep the current field creation and cached-result assignment. It will then duplicate `field.Result`. If a cached result exists and active run properties are available, the handler will make that result range active and call the existing `_apply_post_insert_run_properties` helper with the cached text. Finally, it will collapse the same result range to its end and retain it as `active_range`, preserving the existing continuation behavior.

The complete active run-property set is authoritative. Applying only `Bold` would fix the first observed result but would leave the same Word inheritance bug possible for italic, underline, color, font, size, hidden text, vertical alignment, or character spacing. Reusing the post-insert formatter also keeps field behavior aligned with text and page-break insertion.

No formatting is replayed when there is no cached result or no active run properties. Field creation remains valid in those cases.

## Error handling

Required field-result formatting follows the existing post-insert formatting policy: required Word property failures propagate through the normal COM retry/error path, while optional font properties remain best-effort inside the helper. The current best-effort cached-text assignment remains unchanged.

## Performance

The change adds one existing run-formatting pass per visible cached field result. GOLDEN #1 has 28 fields; current `CreateField` work totals about 89 seconds, while ordinary `ApplyRunProperties` averages about 0.1 seconds. The expected runtime cost is small relative to the fidelity gain and requires no additional save, field update, or verification boundary.

## Regression test

Extend the existing cached-field unit test so Word's fake result begins bold with the wrong complex-script font. Apply active run properties with `bold: false` and `font_cs: Cambria`, create the cached field, and assert:

- cached result text is still seeded;
- the field is not refreshed;
- result `Bold` becomes Word false;
- result `NameBi` becomes Cambria;
- the active range remains the collapsed field result.

The test must be observed RED before production code changes and GREEN afterward.

## Golden acceptance

After focused and full local tests pass, a clean real-Word Golden run must show:

- the first divergence moves beyond `173/runs/0/text_len`;
- the 17 diagnosed field-containing G2 groups disappear or are reduced exactly according to any remaining independently proven formatting difference;
- G2 falls materially below 114, with an expected reduction of up to 69 findings;
- G0, G1, G5, and G7 remain PASS;
- source hash remains `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`;
- runtime does not materially regress.

Only an accepted Golden improvement is eligible for an `automation-dev` checkpoint commit and push. This iteration is not overall completion; the two-consecutive-FULL-PASS and runtime gates remain in force.

## Safety

The implementation stays on `automation-dev`, does not modify the Golden source DOCX, does not globally terminate Word, and does not close unrelated Word documents.
