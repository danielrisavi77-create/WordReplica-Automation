# Word Replica Golden #1 Repair Batch Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the confirmed Golden #1 fidelity defects so the same real academic DOCX can continue past the first table without false verification pauses and with materially correct paragraph/style/table/header/footer layout.

**Architecture:** Keep the existing event pipeline and Word COM controller. Fix fidelity at the source of state leakage: resolve the effective default paragraph style in the blueprint, reset Word paragraph/run state to that style before applying effective source properties, reuse Word-created post-page-break paragraph context only when appropriate, include semantic field cached results in live expected text, and apply source section page-number restart metadata. Validate each change against focused unit tests plus Golden-source proxy assertions before building a persistent update.

**Tech Stack:** Python 3.11+, lxml, pywin32/Word COM on Windows acceptance, pytest, existing WordReplica parser/blueprint/interactive renderer/persistent updater.

## Global Constraints

- Preserve the existing Instant Reconstruction path unchanged.
- Interactive text remains one-character-per-event; no SendKeys, clipboard typing, or fabricated human behavior.
- No silent fidelity downgrade in Maximum Fidelity.
- Do not modify the Golden source DOCX.
- Do not alter `realworld_input` or `results` during persistent updates.
- Every production change requires a failing regression test first.
- The local sandbox may not claim Word-side success; final proof still requires the user's Microsoft Word 2010 run.

---

### Task 1: Parse and resolve the default paragraph style

**Files:**
- Modify: `src/word_replica/parser/parser.py`
- Modify: `src/word_replica/interactive/blueprint.py`
- Test: `tests/unit/test_parser_extended.py`
- Test: `tests/unit/test_blueprint_compiler.py`

**Interfaces:**
- Produces `model.extras["default_paragraph_style_id"]: str | None`.
- `BlueprintCompiler` uses the explicit paragraph style when present, otherwise the parsed default paragraph style.

- [ ] Write a parser test with `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">` and assert `default_paragraph_style_id == "Normal"`.
- [ ] Run the focused parser test and verify RED because the field is not currently populated.
- [ ] Implement `_parse_default_paragraph_style_id()` and store its result during `DocxParser.parse()`.
- [ ] Write a blueprint test where a paragraph has `style_id=None`, `Normal` overrides docDefaults, and assert `ApplyParagraphProperties.style_id == "Normal"` plus Normal effective paragraph/run properties.
- [ ] Run the blueprint test and verify RED on current behavior.
- [ ] Update `_effective_paragraph_properties()` / `_effective_run_properties()` to use the effective paragraph style id (`paragraph.style_id` or model default).
- [ ] Run the focused parser/blueprint tests and verify GREEN.

### Task 2: Reset Word paragraph and run formatting before applying source-effective state

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Test: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- `ApplyParagraphProperties` assigns the resolved style first, resets direct paragraph formatting, then applies effective source properties.
- `ApplyRunProperties` resets direct font formatting before applying effective source run properties.

- [ ] Write a fake-COM test where paragraph A sets `SpaceBefore=96pt`, paragraph B has effective `Normal` with no `SpaceBefore`, and assert paragraph B calls `ParagraphFormat.Reset()` before its properties so A's spacing cannot leak.
- [ ] Run and verify RED because no paragraph reset occurs.
- [ ] Implement a retry-safe `_reset_paragraph_format()` and reorder `ApplyParagraphProperties`: ensure/assign style → reset paragraph formatting → apply effective properties/tabs.
- [ ] Write a run test where run A is bold/red and run B is ordinary; assert `Font.Reset()` occurs before run B properties and bold/color do not stay sticky.
- [ ] Run and verify RED.
- [ ] Implement retry-safe font reset before each `ApplyRunProperties` event, then apply effective properties.
- [ ] Run focused executor tests and verify GREEN.

### Task 3: Isolate table-cell paragraphs from the preceding heading/style context

**Files:**
- Modify: `src/word_replica/interactive/blueprint.py`
- Test: `tests/unit/test_blueprint_compiler.py`
- Test: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- A table-cell paragraph with no explicit `w:pStyle` compiles with the document default paragraph style (Golden: `Normal`).
- Entering a cell does not inherit the pre-table heading's direct paragraph/run formatting after Task 2 resets.

- [ ] Add a blueprint test with a `Heading1` paragraph followed by a table whose cell paragraph has `style_id=None`; assert cell `ApplyParagraphProperties.style_id == "Normal"` and Normal effective properties.
- [ ] Run and verify RED before Task 1/2 behavior is complete.
- [ ] Add a controller trace test that enters a cell after a heading and asserts the cell paragraph applies `Normal`, resets paragraph format, and then applies cell spacing.
- [ ] Run and verify GREEN after the minimal implementation from Tasks 1–2; do not add table-specific hacks unless the test exposes one.

### Task 4: Reuse Word's page-break continuation paragraph without merging mid-paragraph breaks

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Test: `tests/unit/test_interactive_word_executor.py`

**Interfaces:**
- Controller tracks whether the most recent page break remains the final content in the current source paragraph.
- The next `BeginParagraph` reuses the Word continuation paragraph only when no content was inserted after that page break.

- [ ] Write a test: `BeginParagraph → InsertPageBreak → EndParagraph → BeginParagraph` must not call `InsertParagraphAfter()`.
- [ ] Run and verify RED because the current code creates an extra paragraph.
- [ ] Implement a `_page_break_continuation_pending` flag set by `InsertPageBreak`, consumed by the next `BeginParagraph`, and cleared by any character/tab/line-break/image/field insertion after the break.
- [ ] Add a second test: `InsertPageBreak → InsertCharacter("X") → next BeginParagraph` must insert a real new paragraph, proving mid-paragraph page breaks are not collapsed incorrectly.
- [ ] Run both tests and verify GREEN.

### Task 5: Make live verification semantic-field aware

**Files:**
- Modify: `src/word_replica/services/interactive_rebuild.py`
- Test: `tests/unit/test_interactive_verification.py`

**Interfaces:**
- `_InteractiveServiceObserver._track_expected_body_text()` appends `CreateField.cached_result` while in body story.
- Header/footer/note field cached results remain excluded from body expected text.

- [ ] Write a test with body events spelling `Tablica ` + `CreateField(cached_result="1")` + `.` and assert the expected prefix is `Tablica 1.`.
- [ ] Run and verify RED because `CreateField` is currently ignored.
- [ ] Implement body-story `CreateField` cached-result tracking.
- [ ] Add a header-field test proving header cached results do not enter body prefix.
- [ ] Run focused verifier/service-observer tests and verify GREEN.

### Task 6: Reconstruct section page-number restart and header/footer effective formatting

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Test: `tests/unit/test_interactive_word_executor.py`
- Test: `tests/unit/test_blueprint_compiler.py`

**Interfaces:**
- `ApplySectionProperties.page_number_start` maps to Word section page-number restart (`RestartNumberingAtSection=True`, `StartingNumber=<source>`), using the section's primary footer PageNumbers collection with a header fallback.
- Header/footer paragraph/run events use the same resolved style + reset pipeline as body text.

- [ ] Add a controller test with a fake section/footer PageNumbers collection and assert `page_number_start=1` sets restart true and starting number 1.
- [ ] Run and verify RED because `ApplySectionProperties` ignores `page_number_start`.
- [ ] Implement retry-safe page-number restart with footer primary first and header primary fallback.
- [ ] Add a Golden-source blueprint proxy test asserting the header paragraph resolves `Header` based on `Normal`, its run carries italic/8.5pt/gray, footer resolves `Footer`, and section 2 carries `page_number_start == "1"`.
- [ ] Run focused tests and verify GREEN.

### Task 7: Golden #1 structural proxy gate

**Files:**
- Create: `tests/integration/test_golden_rektorova_proxy.py`
- Use fixture source from test temp copy or environment-gated `/mnt/data` only in local verification; do not hard-code user paths into shipped runtime tests.

**Interfaces:**
- Validates source model/blueprint invariants that caused the observed defects without requiring Word COM.

- [ ] Add a local regression test helper that accepts the Golden source path from `WORD_REPLICA_GOLDEN_DOCX`; skip when absent.
- [ ] Assert: default paragraph style is Normal; Normal resolves Times New Roman/justify/6pt-after/1.5-line; first table cell resolves Normal/10pt direct run where applicable; header resolves Header with italic 8.5pt gray run; section 2 starts page numbering at 1; semantic `REF ref_tab_1` field cached result is `1`.
- [ ] Run against the uploaded Golden source and verify GREEN.
- [ ] Compile the blueprint and assert no paragraph with `style_id=None` remains unresolved when a default paragraph style exists.

### Task 8: Full regression, version 2.0.4 update package, and Windows handoff

**Files:**
- Use: `scripts/persistent_harness/build_update.py`
- Update package output: `/mnt/data/WordReplica-Remote-Harness-Update-2.0.4.zip`

**Interfaces:**
- Persistent update upgrades 2.0.3 → 2.0.4 and leaves `realworld_input` / `results` untouched.

- [ ] Run the focused Golden repair tests.
- [ ] Run the complete local pytest suite with `PYTHONPATH=src` and verify 0 failures.
- [ ] Run `python -m compileall -q src scripts` and verify success.
- [ ] Build update ZIP targeting `2.0.4`, minimum installed `2.0.3`.
- [ ] Simulate update over a 2.0.3 persistent layout containing sentinel DOCX/result files; assert version changes to 2.0.4, source manifest passes, and sentinels are byte-identical.
- [ ] Report explicitly that Microsoft Word 2010 acceptance remains unverified until the user reruns only Golden #1.
