# Footer Font, Field Topology, and Stable Word Audit Implementation Plan

**Goal:** Preserve source-equivalent footer/header fonts and field topology, make the controlled Word audit deterministic, and advance Golden #1 to G0-G8 PASS.

## Tasks

- [x] Read all three G6 findings and isolate the footer font as the first divergence.
- [x] Prove the Footer style event contains `font_cs=Cambria` while style configuration omits `NameBi`.
- [x] Add a regression test and observe RED.
- [x] Apply `font_cs` through `Font.NameBi` and verify targeted GREEN.
- [x] Run the full suite: 424 passed, 52 skipped.
- [x] Run real Word and honor the G1/G7 fail-safe.
- [x] Prove fields were converted to `w:fldSimple`, not deleted, and that PASS/fail CreateField ranges were identical.
- [x] Add `fldSimple`→complex topology regressions and observe RED.
- [x] Harden pairing for nested/split/flagged/cross-paragraph/nested-instruction fields; preserve output result content and formatting.
- [x] Verify on the actual fail-safe artifact: 26 restored; G1 PASS; G7 PASS.
- [x] Prove separate Word PDF applications render the unchanged source inconsistently.
- [x] Add paired-export RED/GREEN coverage and render source/output through one owned Word application.
- [x] Restore relationship-free source headers after Word updates `STYLEREF`; skip direct and nested header relationship parts.
- [x] Verify G6 PASS on a copy of a real Golden artifact.
- [x] Prove a bare tab run causes the G8 leader-space mismatch; route tabs through formatted insertion under RED/GREEN.
- [x] Verify full suite: 437 passed, 52 skipped.
- [x] Run Golden `20260816T082601Z_687e3788`: prior gates preserved, G6 reduced to two, G8 PASS, 73/73 pages, source unchanged, 245.63-second runtime.
- [x] Run Golden `20260816T093134Z_687e3788`: G0-G8 PASS, only G9 remains, 73/73 pages, source unchanged, score 9.
- [x] Obtain read-only review with no Critical or Important findings.
- [ ] Commit and push only to `automation-dev`, then continue with the first G9 visual divergence.
