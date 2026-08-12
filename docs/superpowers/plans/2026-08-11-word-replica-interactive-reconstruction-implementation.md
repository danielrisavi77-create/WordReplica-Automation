# Word Replica Interactive Reconstruction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Windows/Microsoft Word Interactive Reconstruction mode that visibly rebuilds a new DOCX character-by-character and object-by-object, with formatting active at insertion time, safe Pause/Resume/Stop, resumable checkpoints, Maximum Fidelity blocking, and final L0-L4 QA, while preserving the verified Instant path.

**Architecture:** Keep the existing Instant path (`PureDocxRenderer` + hybrid `WordComRenderer`) behaviorally isolated. Add an ordered `ReconstructionBlueprint` compiler and a separate `InteractiveWordRenderer`/executor subsystem that owns one visible Word COM session and executes one semantic atomic event at a time. `RebuildService` routes by a new `ReconstructionMode`; UI/worker control is queue/event based, and persisted checkpoints bind source hash + blueprint fingerprint + output hash before restart resume.

**Tech Stack:** Python 3.12+, dataclasses/StrEnum, lxml/OPC parser, pywin32 Word COM on Windows, Tkinter desktop UI, SQLite/project JSON storage, pytest/TDD, Pillow+pypdfium2 for L4 raster comparison, PyInstaller release build.

## Global Constraints

- Target platform: Windows desktop with Microsoft Word installed for Interactive Reconstruction.
- Existing Instant Reconstruction remains a separate fallback path and must not change behavior.
- Baseline Windows invariant: the previously verified Instant release path was 97/97 passing before this feature.
- Interactive visible text is inserted one Unicode character per semantic Word insertion operation.
- Formatting is applied before the first character requiring that formatting; no post-pass is used as the normal text-formatting path.
- Text inside table cells is also inserted one character at a time.
- Tables are created as real Word tables, with geometry/merge/property operations before cell typing.
- Supported images are inserted as new Word objects, then configured step by step.
- Interactive v1 always runs with Microsoft Word visible.
- No SendKeys, clipboard paste as the main reconstruction mechanism, synthetic keyboard events, or mouse automation.
- Slow/Fast/Custom/Maximum modes execute the same event sequence; only deterministic inter-operation delay changes.
- Maximum speed still emits one `InsertCharacter` event/call per visible character.
- Pause/Stop take effect only after the current atomic COM operation returns.
- Maximum Fidelity never silently simplifies an unsupported complex object.
- Source is read-only by default; resume requires unchanged source hash.
- Audit/save history records real automated operations only; no fabricated manual authorship or fake revision history.
- Tkinter main thread never performs long Word COM calls.
- Overall release PASS requires both independent sections: `INSTANT GATE` and `INTERACTIVE GATE`.
- Current packaged source tree has no `.git`; commit steps below are mandatory when executing in a real Git checkout. In this packaged tree, record the completed task and test evidence instead of inventing a commit hash.

## File Structure

### New production files

- `src/word_replica/domain/reconstruction.py` — immutable event schema, blueprint, progress/location, capability and checkpoint domain types.
- `src/word_replica/interactive/__init__.py` — Interactive subsystem package exports.
- `src/word_replica/interactive/blueprint.py` — deterministic `DocumentModel -> ReconstructionBlueprint` compiler.
- `src/word_replica/interactive/preflight.py` — Word/font/asset/capability readiness analysis.
- `src/word_replica/interactive/speed.py` — deterministic delay policy with injectable clock/sleeper.
- `src/word_replica/interactive/control.py` — thread-safe Pause/Resume/Stop/speed state machine.
- `src/word_replica/interactive/tables.py` — stable table grid/merge plan generation independent of mutable Word cell collections.
- `src/word_replica/interactive/capabilities.py` — `RECONSTRUCTED`/`PRESERVED`/`UNSUPPORTED` classification.
- `src/word_replica/interactive/checkpoints.py` — persisted interactive checkpoint schema/read/write/resume validation.
- `src/word_replica/interactive/verification.py` — safe-boundary live verifier and progress snapshot production.
- `src/word_replica/renderers/interactive_word.py` — dedicated visible Word COM session/controller + atomic event executor.
- `src/word_replica/services/interactive_rebuild.py` — orchestration of blueprint, preflight, execution, checkpoints, final QA and result construction.
- `src/word_replica/qa/word_render.py` — controlled Word-to-PDF export used by L4.

### Existing production files to modify

- `src/word_replica/domain/enums.py` — add reconstruction/speed/fidelity/capability/run-state enums.
- `src/word_replica/config.py` — add `InteractiveOptions` and `RebuildOptions.reconstruction_mode`/`interactive`.
- `src/word_replica/domain/model.py` — only add modeled properties needed by Interactive fidelity when parser tasks prove they are absent; do not restructure existing types.
- `src/word_replica/parser/parser.py` — extend run/paragraph/section/drawing metadata extraction required by approved Interactive fidelity.
- `src/word_replica/parser/tables.py` — extend table/cell property extraction needed by table blueprint operations.
- `src/word_replica/services/rebuild.py` — route Interactive mode to `InteractiveRebuildService` before Instant renderer selection; preserve Instant logic.
- `src/word_replica/services/project_store.py` — expose resumable stopped/interrupted Interactive projects and project path rehydration.
- `src/word_replica/desktop.py` — mode-first UI, interactive option panel, preflight/live controller, pause/resume/stop, resume card, transparency copy.
- `RUN_WINDOWS_RELEASE_GATE.ps1` — split and enforce independent Instant and Interactive gates.
- `BUILD_WINDOWS_APP.ps1` — include new interactive modules and require full release gate before EXE build.

### New tests

- `tests/unit/test_interactive_config.py`
- `tests/unit/test_reconstruction_events.py`
- `tests/unit/test_blueprint_compiler.py`
- `tests/unit/test_interactive_speed.py`
- `tests/unit/test_interactive_control.py`
- `tests/unit/test_interactive_preflight.py`
- `tests/unit/test_interactive_tables.py`
- `tests/unit/test_interactive_capabilities.py`
- `tests/unit/test_interactive_checkpoints.py`
- `tests/unit/test_interactive_verification.py`
- `tests/unit/test_interactive_word_executor.py`
- `tests/unit/test_interactive_rebuild_service.py`
- `tests/unit/test_desktop_interactive.py`
- `tests/integration/word/test_interactive_word_text.py`
- `tests/integration/word/test_interactive_word_tables.py`
- `tests/integration/word/test_interactive_word_images.py`
- `tests/integration/word/test_interactive_word_structures.py`
- `tests/integration/word/test_interactive_resume.py`
- `tests/acceptance/test_interactive_acceptance_matrix.py`

---

### Task 1: Configuration, enums, and immutable reconstruction event schema

**Files:**
- Modify: `src/word_replica/domain/enums.py`
- Modify: `src/word_replica/config.py`
- Create: `src/word_replica/domain/reconstruction.py`
- Create: `tests/unit/test_interactive_config.py`
- Create: `tests/unit/test_reconstruction_events.py`

**Interfaces:**
- Produces: `ReconstructionMode`, `InteractiveSpeedMode`, `InteractiveFidelity`, `CapabilityClass`, `InteractiveRunState`.
- Produces: `InteractiveOptions`, available as `RebuildOptions.interactive`.
- Produces: `ReconstructionEvent(event_type, source_element_id, payload)`, `ReconstructionBlueprint`, `SemanticLocation`, `InteractiveProgress`.
- Consumes: existing `RebuildOptions`, `DocumentModel` fingerprints, `RendererChoice` only for Instant mode.

- [ ] **Step 1: Write failing option-default and validation tests**

```python
from dataclasses import FrozenInstanceError
import pytest
from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import (
    InteractiveFidelity,
    InteractiveSpeedMode,
    ReconstructionMode,
)


def test_interactive_options_have_deterministic_release_defaults():
    options = InteractiveOptions()
    assert options.speed_mode is InteractiveSpeedMode.FAST
    assert options.characters_per_second == 25.0
    assert options.object_step_delay_ms == 150
    assert options.fidelity is InteractiveFidelity.MAXIMUM
    assert options.checkpoint_after_tables is True
    assert options.checkpoint_after_images is True
    assert options.checkpoint_after_sections is True
    assert options.checkpoint_event_interval == 500
    assert options.verify_during_run is True
    assert options.block_on_unsupported is True


def test_rebuild_options_default_to_existing_instant_behavior():
    options = RebuildOptions()
    assert options.reconstruction_mode is ReconstructionMode.INSTANT
    assert options.interactive == InteractiveOptions()


def test_interactive_options_reject_invalid_rates():
    with pytest.raises(ValueError, match="characters_per_second"):
        InteractiveOptions(characters_per_second=0)
    with pytest.raises(ValueError, match="object_step_delay_ms"):
        InteractiveOptions(object_step_delay_ms=-1)
    with pytest.raises(ValueError, match="checkpoint_event_interval"):
        InteractiveOptions(checkpoint_event_interval=0)
```

- [ ] **Step 2: Run the config tests and verify RED**

Run:
```powershell
python -m pytest tests/unit/test_interactive_config.py -vv
```
Expected: collection/import FAIL because the new enums/options do not exist.

- [ ] **Step 3: Add exact enums and `InteractiveOptions`**

Add to `domain/enums.py`:
```python
class ReconstructionMode(StrEnum):
    INSTANT = "instant"
    INTERACTIVE = "interactive"

class InteractiveSpeedMode(StrEnum):
    SLOW = "slow"
    FAST = "fast"
    CUSTOM = "custom"
    MAXIMUM = "maximum"

class InteractiveFidelity(StrEnum):
    STANDARD = "standard"
    MAXIMUM = "maximum"

class CapabilityClass(StrEnum):
    RECONSTRUCTED = "RECONSTRUCTED"
    PRESERVED = "PRESERVED"
    UNSUPPORTED = "UNSUPPORTED"

class InteractiveRunState(StrEnum):
    CREATED = "CREATED"
    PREFLIGHT = "PREFLIGHT"
    RUNNING = "RUNNING"
    PAUSED = "PAUSED"
    STOPPED = "STOPPED"
    VERIFYING = "VERIFYING"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"
```

Add a frozen `InteractiveOptions` dataclass to `config.py` with the defaults above and explicit numeric validation. Add:
```python
reconstruction_mode: ReconstructionMode = ReconstructionMode.INSTANT
interactive: InteractiveOptions = InteractiveOptions()
```
to `RebuildOptions` while leaving all existing Instant fields/defaults unchanged.

- [ ] **Step 4: Write failing immutable event/blueprint tests**

```python
from word_replica.domain.reconstruction import ReconstructionBlueprint, ReconstructionEvent


def test_blueprint_fingerprint_is_deterministic_and_payload_sensitive():
    a = ReconstructionBlueprint.build(
        source_sha256="a" * 64,
        source_model_fingerprint="model",
        events=(ReconstructionEvent("InsertCharacter", "run_1", {"character": "A"}),),
    )
    b = ReconstructionBlueprint.build(
        source_sha256="a" * 64,
        source_model_fingerprint="model",
        events=(ReconstructionEvent("InsertCharacter", "run_1", {"character": "A"}),),
    )
    c = ReconstructionBlueprint.build(
        source_sha256="a" * 64,
        source_model_fingerprint="model",
        events=(ReconstructionEvent("InsertCharacter", "run_1", {"character": "B"}),),
    )
    assert a.fingerprint == b.fingerprint
    assert a.fingerprint != c.fingerprint
    assert a.total_visible_characters == 1
```

- [ ] **Step 5: Run event tests and verify RED**

Run:
```powershell
python -m pytest tests/unit/test_reconstruction_events.py -vv
```
Expected: import FAIL because `domain.reconstruction` does not exist.

- [ ] **Step 6: Implement immutable event and blueprint types**

`domain/reconstruction.py` must use frozen/slots dataclasses. `ReconstructionBlueprint.build` must derive:
- `schema_version = 1`;
- `total_events = len(events)`;
- visible-character count from `event_type == "InsertCharacter"`;
- semantic counts passed by compiler later;
- stable SHA-256 fingerprint from canonical JSON (`sort_keys=True`, UTF-8, compact separators).

Do not put COM objects, `Path` objects, locks, or mutable UI state into the blueprint.

- [ ] **Step 7: Run new tests plus existing config regression**

Run:
```powershell
python -m pytest tests/unit/test_interactive_config.py tests/unit/test_reconstruction_events.py tests/unit/test_config.py -vv
```
Expected: PASS.

- [ ] **Step 8: Version checkpoint**

In a Git checkout:
```bash
git add src/word_replica/domain/enums.py src/word_replica/config.py src/word_replica/domain/reconstruction.py tests/unit/test_interactive_config.py tests/unit/test_reconstruction_events.py
git commit -m "feat: add interactive reconstruction configuration and events"
```
In the current packaged tree, record the passing command/output in the task log instead of claiming a commit.

---

### Task 2: Deterministic blueprint compiler for paragraphs, runs, tabs, and breaks

**Files:**
- Create: `src/word_replica/interactive/__init__.py`
- Create: `src/word_replica/interactive/blueprint.py`
- Modify: `src/word_replica/parser/parser.py`
- Create: `tests/unit/test_blueprint_compiler.py`

**Interfaces:**
- Consumes: `DocumentModel`, `Paragraph`, `Run`, `Section`, `ReconstructionEvent`.
- Produces: `BlueprintCompiler.compile(model: DocumentModel) -> ReconstructionBlueprint`.
- Event ordering contract for a paragraph: `BeginParagraph -> ApplyParagraphProperties -> [ApplyRunProperties -> InsertCharacter/InsertTab/InsertLineBreak/InsertPageBreak]* -> EndParagraph`.

- [ ] **Step 1: Write RED test proving one event per visible character and no bulk text event**

```python
from word_replica.domain.model import DocumentModel, Paragraph, Run
from word_replica.interactive.blueprint import BlueprintCompiler


def test_compiler_emits_one_insert_character_event_per_character():
    model = DocumentModel(
        source_sha256="a" * 64,
        body=[Paragraph("p1", [Run("r1", "Až B", {"bold": True})])],
    )
    blueprint = BlueprintCompiler().compile(model)
    chars = [e.payload["character"] for e in blueprint.events if e.event_type == "InsertCharacter"]
    assert chars == ["A", "ž", " ", "B"]
    assert not any("text" in e.payload and len(str(e.payload["text"])) > 1 for e in blueprint.events)
```

- [ ] **Step 2: Run and verify RED**

Run:
```powershell
python -m pytest tests/unit/test_blueprint_compiler.py::test_compiler_emits_one_insert_character_event_per_character -vv
```
Expected: import FAIL for missing compiler.

- [ ] **Step 3: Implement paragraph/run event compilation**

Create `BlueprintCompiler` with these exact method contracts:
- `compile(model: DocumentModel) -> ReconstructionBlueprint`
- `_compile_blocks(blocks: list[object], events: list[ReconstructionEvent], location: SemanticLocation) -> None`
- `_compile_paragraph(paragraph: Paragraph, events: list[ReconstructionEvent], location: SemanticLocation) -> None`
- `_compile_run(run: Run, events: list[ReconstructionEvent], location: SemanticLocation) -> None`

The compiler mutates only its local `events` list while compiling; it never mutates `DocumentModel`.

Rules:
- copy properties through stable JSON-safe payloads;
- `ApplyRunProperties` occurs before any visible character in that run;
- `\t` becomes `InsertTab`, not `InsertCharacter`;
- run `break_types` become dedicated line/page-break events in source order;
- paragraph `pageBreakBefore` stays a paragraph property, not a typed character;
- hidden run text remains in blueprint with a hidden run property so canonical content is not silently lost.

- [ ] **Step 4: Add RED ordering tests for formatting transitions, tab, line break, and page break**

```python
def test_run_properties_precede_first_character_of_each_run():
    model = DocumentModel(
        source_sha256="a" * 64,
        body=[Paragraph(
            "p1",
            runs=[
                Run("r1", "A", {"bold": False}),
                Run("r2", "B", {"bold": True}),
            ],
        )],
    )
    events = BlueprintCompiler().compile(model).events
    trace = [
        (event.event_type, event.source_element_id, event.payload)
        for event in events
        if event.event_type in {"ApplyRunProperties", "InsertCharacter"}
    ]
    assert trace == [
        ("ApplyRunProperties", "r1", {"bold": False}),
        ("InsertCharacter", "r1", {"character": "A"}),
        ("ApplyRunProperties", "r2", {"bold": True}),
        ("InsertCharacter", "r2", {"character": "B"}),
    ]


def test_tabs_and_breaks_use_semantic_events_in_source_order():
    run = Run(
        "r1",
        text="A\tB\n\n",
        properties={
            "content_tokens": [
                {"kind": "text", "value": "A"},
                {"kind": "tab"},
                {"kind": "text", "value": "B"},
                {"kind": "line_break"},
                {"kind": "page_break"},
            ]
        },
    )
    model = DocumentModel(source_sha256="a" * 64, body=[Paragraph("p1", [run])])
    events = BlueprintCompiler().compile(model).events
    trace = [
        (event.event_type, event.payload)
        for event in events
        if event.event_type.startswith("Insert")
    ]
    assert trace == [
        ("InsertCharacter", {"character": "A"}),
        ("InsertTab", {}),
        ("InsertCharacter", {"character": "B"}),
        ("InsertLineBreak", {}),
        ("InsertPageBreak", {}),
    ]
```

To make the second test possible without guessing break positions, extend `parse_run` to preserve exact child order in `run.properties["content_tokens"]`. Each token is one of `text`, `tab`, `line_break`, or `page_break`; `Run.text` remains the existing flattened text for Instant compatibility. The compiler uses `content_tokens` when present and falls back to character-scanning `Run.text` only for legacy/injected models.

- [ ] **Step 5: Extend parser properties only where blueprint tests demonstrate missing source semantics**

Add exact extraction for approved run properties currently absent from `parse_run` when present in OOXML:
- font family (`w:rFonts` ascii/hAnsi/eastAsia/cs);
- size (`w:sz` half-points);
- strike (`w:strike`);
- color (`w:color`);
- highlight (`w:highlight`);
- vertAlign (`superscript`/`subscript`);
- language (`w:lang`);
- character spacing/position (`w:spacing`, `w:position`).

Add paragraph tab-stop extraction from `w:tabs/w:tab` and paragraph borders/shading only as structured dictionaries. Do not change existing keys used by Instant renderer tests.

- [ ] **Step 6: Run compiler/parser regressions**

Run:
```powershell
python -m pytest tests/unit/test_blueprint_compiler.py tests/unit/test_parser_core.py tests/unit/test_parser_extended.py -vv
```
Expected: PASS.

- [ ] **Step 7: Version checkpoint**

Git checkout commit:
```bash
git add src/word_replica/interactive src/word_replica/parser/parser.py tests/unit/test_blueprint_compiler.py
git commit -m "feat: compile deterministic interactive text blueprint"
```

---

### Task 3: Word COM controller and character-by-character executor

**Files:**
- Create: `src/word_replica/renderers/interactive_word.py`
- Create: `tests/unit/test_interactive_word_executor.py`
- Create: `tests/integration/word/test_interactive_word_text.py`

**Interfaces:**
- Consumes: `ReconstructionEvent`, `ReconstructionBlueprint`, `InteractiveOptions`.
- Produces: `InteractiveWordController` with `open_blank()`, `execute_event(event)`, `save(path)`, `set_custom_property(name, value)`, `current_state_snapshot()`, `close()`.
- Produces: `InteractiveWordRenderer` that owns one controller/session and exposes the renderer save/property methods required by existing metadata/checkpoint infrastructure.

- [ ] **Step 1: Write RED COM-contract fake proving exactly one insertion call per character**

Use a minimal fake Range object that records `.InsertAfter(value)` calls. Test:
```python
def test_insert_character_calls_word_once_with_exactly_one_character():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "ž"}))
    assert fake.insert_after_calls == ["ž"]


def test_insert_character_rejects_multi_character_payload():
    fake = FakeWordRange()
    controller = InteractiveWordController.for_testing(active_range=fake)
    with pytest.raises(ValueError, match="exactly one"):
        controller.execute_event(ReconstructionEvent("InsertCharacter", "r1", {"character": "AB"}))
```

- [ ] **Step 2: Verify RED**

Run:
```powershell
python -m pytest tests/unit/test_interactive_word_executor.py -vv
```
Expected: import FAIL.

- [ ] **Step 3: Implement dedicated Word session and text atomic operations**

Requirements:
- COM initialization/uninitialization occurs in the executor thread;
- `DispatchEx("Word.Application")`;
- `Visible = True` unconditionally for Interactive v1;
- create a new blank document, not a prebuilt package;
- maintain one collapsed insertion Range;
- after each character insertion, collapse range to its end;
- dedicated handlers for `InsertTab`, `InsertLineBreak`, `InsertPageBreak`, `BeginParagraph`, `EndParagraph`;
- no calls to `Selection.TypeText`, clipboard, SendKeys, or bulk run/paragraph `.Text = full_text`.

- [ ] **Step 4: Add source-code guard test against forbidden input mechanisms**

```python
def test_interactive_renderer_has_no_sendkeys_or_clipboard_bulk_path():
    source = Path("src/word_replica/renderers/interactive_word.py").read_text(encoding="utf-8")
    forbidden = ("SendKeys", "win32clipboard", "pyautogui", "keyboard.write", "Selection.TypeText")
    assert all(token not in source for token in forbidden)
```

- [ ] **Step 5: Add real Word integration test for a short Unicode paragraph**

Mark with `@pytest.mark.word` and `WORD_REPLICA_WORD_TESTS=1`. Build blueprint for `"Až B"`, execute against blank visible Word, save, reparse, assert `plain_text()` equals source. Also assert the test double/unit trace proves four character calls.

- [ ] **Step 6: Run unit tests, then Windows integration when available**

Local/non-Word:
```powershell
python -m pytest tests/unit/test_interactive_word_executor.py -vv
```
Windows Word:
```powershell
$env:WORD_REPLICA_WORD_TESTS="1"
python -m pytest tests/integration/word/test_interactive_word_text.py -vv --tb=long
```
Expected: PASS.

- [ ] **Step 7: Version checkpoint**

Git checkout commit:
```bash
git add src/word_replica/renderers/interactive_word.py tests/unit/test_interactive_word_executor.py tests/integration/word/test_interactive_word_text.py
git commit -m "feat: add visible character-by-character Word executor"
```

---

### Task 4: Deterministic speed controller and atomic Pause/Resume/Stop state machine

**Files:**
- Create: `src/word_replica/interactive/speed.py`
- Create: `src/word_replica/interactive/control.py`
- Modify: `src/word_replica/domain/reconstruction.py` — add `ControlDecision` and `ExecutionOutcome`.
- Create: `tests/unit/test_interactive_speed.py`
- Create: `tests/unit/test_interactive_control.py`
- Modify: `src/word_replica/renderers/interactive_word.py`

**Interfaces:**
- Produces: `SpeedController.delay_after(event_type: str) -> None`, `set_mode(mode: InteractiveSpeedMode) -> None`, `set_custom_rate(characters_per_second: float) -> None`.
- Produces: `InteractiveRunControl.before_next_event(last_completed_index: int) -> ControlDecision`, `pause() -> None`, `resume() -> None`, `stop() -> None`, `set_speed(mode: InteractiveSpeedMode, characters_per_second: float | None = None) -> None`, and read-only `state`.
- Executor calls control only between atomic events.

- [ ] **Step 1: Write RED speed-policy tests with fake sleeper**

Test exact semantics:
- Slow and Fast call injected sleeper with deterministic nonzero delay after `InsertCharacter`;
- Custom 20 cps yields `0.05` seconds;
- Maximum yields no sleeper call;
- object event uses `object_step_delay_ms / 1000`;
- changing speed changes the next delay, not prior events.

- [ ] **Step 2: Verify RED and implement `SpeedController`**

Run then implement using an injected `sleep_fn` (default `time.sleep`) so tests never use wall-clock sleeping.

- [ ] **Step 3: Write RED control-state tests**

Use a worker thread and barriers/events, not arbitrary sleeps. Prove:
- pause request made during event N allows N to complete but blocks before N+1;
- resume releases N+1;
- stop request preserves last completed index and returns a stop decision before N+1;
- state transitions are `RUNNING -> PAUSED -> RUNNING -> STOPPED`;
- no direct Tkinter object is stored in control state.

- [ ] **Step 4: Implement thread-safe control with `threading.Condition`**

`before_next_event()` is the only wait point. The COM handler itself must never be interrupted or run while holding the control lock.

- [ ] **Step 5: Integrate executor event loop**

Add a method equivalent to:
```python
def execute_blueprint(self, blueprint, control, observer):
    for index, event in enumerate(blueprint.events):
        decision = control.before_next_event(index - 1)
        if decision.stop_requested:
            return ExecutionOutcome.stopped(index - 1)
        self.execute_event(event)
        observer.event_completed(index, event)
        self.speed.delay_after(event.event_type)
```
Use concrete dataclass `ExecutionOutcome` in `domain/reconstruction.py` rather than a dict.

- [ ] **Step 6: Run unit suite for tasks 1-4**

```powershell
python -m pytest tests/unit/test_interactive_config.py tests/unit/test_reconstruction_events.py tests/unit/test_blueprint_compiler.py tests/unit/test_interactive_word_executor.py tests/unit/test_interactive_speed.py tests/unit/test_interactive_control.py -vv
```
Expected: PASS.

- [ ] **Step 7: Version checkpoint**

Commit in Git checkout with message `feat: add interactive speed and atomic run controls`.

---

### Task 5: Run and paragraph formatting applied before typing

**Files:**
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `src/word_replica/parser/parser.py`
- Create/extend: `tests/unit/test_interactive_word_executor.py`
- Extend: `tests/integration/word/test_interactive_word_text.py`

**Interfaces:**
- Consumes: `ApplyParagraphProperties` and `ApplyRunProperties` payloads from Task 2.
- Produces: Word paragraph/run formatting mutations on the active range before `InsertCharacter`.

- [ ] **Step 1: Write RED call-order tests**

Fake controller objects must record operations and prove:
```text
set_bold(True)
set_font_name("Aptos")
set_font_size(24 half-points -> 12 pt)
insert_after("A")
```
with every formatting mutation occurring before the first affected character insertion.

Add paragraph ordering test: alignment/indents/spacing/style/list/page-break-before applied before paragraph text.

- [ ] **Step 2: Verify RED**

Run targeted executor tests and confirm ordering assertions fail.

- [ ] **Step 3: Implement exact Word property mapping**

Map approved properties only:
- run: font, size, bold, italic, underline, strike, color, highlight, superscript/subscript, hidden, language, character spacing/position;
- paragraph: style, alignment, indents, before/after/line spacing, keep-next/keep-lines, widow control, page-break-before, tab stops, borders/shading where represented.

Use helper functions in `interactive_word.py` or a focused private formatter class inside that file; do not modify `WordComRenderer` Instant behavior.

- [ ] **Step 4: Add real Word mixed-format integration fixture**

Create a source with normal + bold + italic + colored runs and paragraph spacing. Interactive rebuild, reparse, compare L0/L2 projections. The visual process must still have one character event per visible character.

- [ ] **Step 5: Run parser + formatting + Instant renderer regression tests**

```powershell
python -m pytest tests/unit/test_parser_core.py tests/unit/test_parser_extended.py tests/unit/test_interactive_word_executor.py tests/unit/test_word_com_renderer_unit.py -vv
```
Windows:
```powershell
python -m pytest tests/integration/word/test_interactive_word_text.py tests/integration/word/test_word_com_renderer.py -vv --tb=short
```
Expected: PASS.

- [ ] **Step 6: Version checkpoint**

Commit: `feat: apply Word formatting during interactive typing`.

---

### Task 6: Table blueprint, stable merge plan, and cell-by-cell character typing

**Files:**
- Create: `src/word_replica/interactive/tables.py`
- Modify: `src/word_replica/interactive/blueprint.py`
- Modify: `src/word_replica/parser/tables.py`
- Modify: `src/word_replica/renderers/interactive_word.py`
- Create: `tests/unit/test_interactive_tables.py`
- Create: `tests/integration/word/test_interactive_word_tables.py`

**Interfaces:**
- Produces: `TablePlan`/`MergePlan` dataclasses from canonical `Table`.
- Blueprint table event sequence: `BeginTable -> SetTableProperties -> width/row/cell structure -> merges -> cell traversal -> EndTable`.
- Cell text uses the same paragraph/run/character events from Tasks 2/5.

- [ ] **Step 1: Write RED table plan tests**

Concrete tests must cover:
- rectangular 2x3 grid;
- horizontal merge;
- vertical merge;
- combined merge where lookup coordinates would become invalid after the first Word merge;
- deterministic logical cell traversal excluding continuation cells;
- nested table inside a cell.

Assert merge operations are generated from stable pre-merge coordinates and are executed after initial cell handles/grid state are established.

- [ ] **Step 2: Verify RED and implement pure table plan generation**

No COM code in `interactive/tables.py`. It must be testable entirely from `Table` dataclasses.

- [ ] **Step 3: Extend parser table properties with exact modeled fidelity keys**

Only add source properties needed by the approved table requirements:
- table alignment;
- preferred width/type;
- fixed/auto layout;
- grid column widths;
- row height + exact/at-least;
- repeat-header/cant-split;
- cell margins;
- vertical alignment;
- borders style/size/color/spacing;
- shading/fill;
- text direction when present.

Preserve current keys expected by Instant tests.

- [ ] **Step 4: Compile table events and nested blocks**

`BlueprintCompiler._compile_table()` must emit structural operations before entering logical cells. Every visible cell character must still be an individual `InsertCharacter` event.

- [ ] **Step 5: Write RED executor call-order test**

Assert:
```text
create_table
set_widths
set_rows
set_cell_properties
merge_cells
enter_cell(1,1)
apply_run_properties
insert_character("A")
insert_character("1")
leave_cell(1,1)
```
No cell text call may occur before table structure/merge operations are complete; use a two-character cell value `A1` in the test so the trace proves two separate insertion calls.

- [ ] **Step 6: Implement Word table operations without post-merge collection lookup**

Before mutating merges, capture required Word cell COM references/Range bookmarks from the initial grid. Populate structural formatting and perform merges using the stable plan. Do not reproduce the previously fixed pattern of calling `Table.Cell(row, col)` after earlier merges have changed the collection.

- [ ] **Step 7: Add real Word tests**

Required cases:
- merged/shaded/bordered table;
- text in merged cells typed character-by-character;
- nested table where supported.

Reparse output and require L0/L1; require modeled L2/L3 for table properties.

- [ ] **Step 8: Run table + Instant acceptance regression subset**

```powershell
python -m pytest tests/unit/test_interactive_tables.py tests/unit/test_parser_extended.py -vv
```
Windows:
```powershell
python -m pytest tests/integration/word/test_interactive_word_tables.py -vv --tb=long
python -m pytest tests/acceptance/test_acceptance_matrix.py -m word -k "04_tables_merged" -vv --tb=short
```
Expected: PASS.

- [ ] **Step 9: Version checkpoint**

Commit: `feat: rebuild interactive tables cell by cell`.

---

### Task 7: Image/drawing extraction, capability classification, and stepwise Word placement

**Files:**
- Create: `src/word_replica/interactive/capabilities.py`
- Modify: `src/word_replica/domain/reconstruction.py` — add `CapabilityDecision`.
- Modify: `src/word_replica/parser/parser.py`
- Modify: `src/word_replica/domain/model.py` — add a focused `DrawingRef` dataclass and `DocumentModel.drawings` occurrence list so assets are tied to source positions/geometry.
- Modify: `src/word_replica/interactive/blueprint.py`
- Modify: `src/word_replica/renderers/interactive_word.py`
- Create: `tests/unit/test_interactive_capabilities.py`
- Extend/create: `tests/unit/test_blueprint_compiler.py`
- Create: `tests/integration/word/test_interactive_word_images.py`

**Interfaces:**
- Produces: `DrawingRef(element_id, asset_id, source_path, representation, width_emu, height_emu, lock_aspect_ratio, wrap_type, horizontal_relative_from, horizontal_position_emu, vertical_relative_from, vertical_position_emu, distance_top_emu, distance_bottom_emu, distance_left_emu, distance_right_emu, crop, rotation_degrees, behind_text, z_order)` stored in `DocumentModel.drawings`.
- Produces: drawing metadata sufficient for `InsertImage`, dimensions, inline/floating, wrap, anchor/reference, X/Y position, crop, rotation, z-order.
- Produces: `classify_drawing(drawing: DrawingRef) -> CapabilityDecision`, where `CapabilityDecision` contains `classification: CapabilityClass`, `reason: str`, and `source_element_id: str`.

- [ ] **Step 1: Write RED parser/compiler tests for inline and floating image metadata**

Use fixture XML/source documents containing:
- `wp:inline` image with extent;
- `wp:anchor` image with positionH/positionV + wrap + extent;
- crop/rotation when represented.

Assert compiler emits `InsertImage` first, followed by geometry events in deterministic order.

- [ ] **Step 2: Write RED capability tests**

Examples:
- supported PNG/JPEG drawing -> `RECONSTRUCTED`;
- safely transferable but not reconstructable approved object -> `PRESERVED`;
- active/unsupported DrawingML/VML semantics -> `UNSUPPORTED`.

Each classification must carry an explicit reason, never only an enum.

- [ ] **Step 3: Implement drawing metadata extraction and classification**

Binary bytes continue to come from `DocumentModel.assets`; blueprint payload stores only asset ID/hash/geometry, never raw bytes.

- [ ] **Step 4: Write RED Word call-order tests**

Assert `AddPicture`/new image object creation occurs before size/wrap/position/crop/rotation/z-order mutations. Assert floating source is not silently left inline in Maximum Fidelity.

- [ ] **Step 5: Implement stepwise image operations**

Use `InlineShapes.AddPicture` for source inline drawings. For source floating drawings, insert the image and convert/use the Word `Shape` API before applying wrap/anchor/position properties. Configure each supported property as a separate atomic event so object delays and live progress are visible.

- [ ] **Step 6: Add real Word image integration tests**

Required:
- inline image;
- floating anchored image with wrapping and position.

Save/reparse and compare modeled image count/geometry. Run the existing image acceptance fixture to ensure the OPC content-type corruption fixed in RC2 does not regress.

- [ ] **Step 7: Run image + RC2 regression subset**

Windows:
```powershell
python -m pytest tests/integration/word/test_interactive_word_images.py -vv --tb=long
python -m pytest tests/acceptance/test_acceptance_matrix.py -m word -k "05_images_inline_floating or 13_academic_complex" -vv --tb=short
```
Expected: PASS.

- [ ] **Step 8: Version checkpoint**

Commit: `feat: reconstruct interactive images with geometry`.

---

### Task 8: Sections, headers, footers, and Word story contexts

**Files:**
- Modify: `src/word_replica/parser/parser.py`
- Modify: `src/word_replica/interactive/blueprint.py`
- Modify: `src/word_replica/renderers/interactive_word.py`
- Create: `tests/integration/word/test_interactive_word_structures.py`
- Extend: `tests/unit/test_blueprint_compiler.py`

**Interfaces:**
- Produces story-aware events with payload field `story` (`body`, header/footer type, footnote/endnote later).
- Section boundary events reference canonical `Section.element_id` and occur at body boundary, not as a pre-created section list.

- [ ] **Step 1: Write RED section-boundary ordering test**

Use the existing core fixture pattern with two sections. Assert second `BeginSection/ApplySectionProperties` occurs only after the canonical boundary paragraph/event, never before body reconstruction.

- [ ] **Step 2: Write RED section-property tests**

Assert payload models page size, orientation, margins, gutter, header/footer distances, break type, columns, first/even-page flags when present.

- [ ] **Step 3: Implement section event compilation and Word application**

Create sections at collapsed document boundary ranges only. Reuse lessons from the verified Instant section-break bug fixes: do not pre-create later sections and do not insert structural empty paragraphs as visible content.

- [ ] **Step 4: Compile and execute header/footer stories**

For each actual header/footer relationship/story, create/select correct Word header/footer and run the same paragraph/run/character pipeline there. Do not assign whole header `.Range.Text` in Interactive mode.

- [ ] **Step 5: Add real Word tests**

Required:
- portrait -> landscape sections;
- distinct header/footer text typed by character operations;
- first-page/even-odd linkage where fixture supports it.

- [ ] **Step 6: Run structure + existing section regressions**

```powershell
python -m pytest tests/integration/word/test_interactive_word_structures.py -vv --tb=long
python -m pytest tests/acceptance/test_acceptance_matrix.py -m word -k "06_sections_orientations or 07_headers_footers_numbers" -vv --tb=short
```
Expected: PASS.

- [ ] **Step 7: Version checkpoint**

Commit: `feat: reconstruct interactive sections and stories`.

---

### Task 9: Notes, lists, bookmarks, fields, and TOC semantics

**Files:**
- Modify: `src/word_replica/interactive/blueprint.py`
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `src/word_replica/parser/parser.py` only where exact source location metadata is missing.
- Extend: `tests/integration/word/test_interactive_word_structures.py`
- Extend: `tests/unit/test_blueprint_compiler.py`

**Interfaces:**
- Produces/consumes events: `CreateFootnote`, `CreateEndnote`, `CreateListBinding`, `CreateBookmark`, `CreateField` and story-local character events.
- Fields create real Word fields from `Field.instruction`; visible result is updated by Word at safe points, not typed as a fake static substitute when semantic reconstruction is supported.

- [ ] **Step 1: Write RED blueprint location tests for notes/bookmarks/fields**

Prove each event carries source element ID and a deterministic anchor location rather than the current Instant approximation of always using document start/end.

- [ ] **Step 2: Write RED list-semantic test**

For a true numbered/bulleted paragraph, assert compiler emits `CreateListBinding`/list property event and does not prepend a literal bullet/number character to `InsertCharacter` events.

- [ ] **Step 3: Implement anchor-aware compilation**

If current `DocumentModel` lacks required exact field/bookmark/note anchors, extend parser/model minimally with stable source paths/range anchors. Preserve existing QA projections unless new fields are explicitly added to L1/L2/L3 later.

- [ ] **Step 4: Implement real Word semantics**

- create note reference at current output Range then type note story char-by-char;
- apply Word list template/numbering semantics;
- create bookmark at compiled range;
- create Word field with instruction at compiled range;
- update fields only at controlled safe boundaries/final verification.

- [ ] **Step 5: Add real Word integration tests**

Cover fixture groups 03, 08, 09, 11. Require L0/L1 PASS and no approximation warnings for supported anchored semantics.

- [ ] **Step 6: Version checkpoint**

Commit: `feat: rebuild notes lists bookmarks and fields interactively`.

---

### Task 10: Preflight and Maximum Fidelity capability gate

**Files:**
- Create: `src/word_replica/interactive/preflight.py`
- Modify: `src/word_replica/domain/reconstruction.py` — add `PreflightReport` and `PreparedInteractiveRun`.
- Extend: `src/word_replica/interactive/capabilities.py`
- Create: `tests/unit/test_interactive_preflight.py`
- Modify: `src/word_replica/services/interactive_rebuild.py` (created in this task)
- Create: `tests/unit/test_interactive_rebuild_service.py`

**Interfaces:**
- Produces: `PreflightReport` with source hash, blueprint fingerprint/status, Word availability, asset/font checks, counts, capability items, `maximum_fidelity_ready`, blocking reasons.
- Produces: initial `InteractiveRebuildService.prepare(source, options, paths, audit) -> PreparedInteractiveRun`.

- [ ] **Step 1: Write RED preflight tests**

Use injectable probes:
- Word unavailable -> blocking;
- missing image asset -> blocking;
- unsupported object + Maximum + `block_on_unsupported=True` -> blocking;
- preserved object -> allowed only with explicit preservation policy;
- Standard fidelity may proceed with explicit warning when policy permits;
- source/blueprint counts are reported exactly.

- [ ] **Step 2: Implement pure preflight with injectable environment probes**

Do not launch and quit an extra Word process just to probe when actual session startup can be the definitive gate; on Windows, environment probe should be lightweight and test-injectable. Font discovery may use Windows font registry/directories but must not mutate the system.

- [ ] **Step 3: Write RED orchestration test proving preflight happens before blank Word creation**

Inject a fake controller factory and blocked preflight. Assert controller factory is never called and result/state is `PREFLIGHT`/blocked with audit reason.

- [ ] **Step 4: Implement `InteractiveRebuildService.prepare` and blocked result path**

Persist blueprint JSON in project `working/blueprint.json`, write preflight JSON to `logs/preflight.json`, audit `BLUEPRINT_COMPILED` and `PREFLIGHT_COMPLETED`.

- [ ] **Step 5: Run unit preflight/orchestration tests**

```powershell
python -m pytest tests/unit/test_interactive_preflight.py tests/unit/test_interactive_rebuild_service.py -vv
```
Expected: PASS.

- [ ] **Step 6: Version checkpoint**

Commit: `feat: add interactive preflight and maximum fidelity gate`.

---

### Task 11: Persisted interactive checkpoints and restart resume validation

**Files:**
- Create: `src/word_replica/interactive/checkpoints.py`
- Modify: `src/word_replica/services/project_store.py`
- Modify: `src/word_replica/services/interactive_rebuild.py`
- Create: `tests/unit/test_interactive_checkpoints.py`
- Create: `tests/integration/word/test_interactive_resume.py`

**Interfaces:**
- Produces: `InteractiveCheckpoint` with all fields in design §16.2.
- Produces: `InteractiveCheckpointStore.write/load_latest/validate_resume`.
- `ProjectStore.list_resumable_projects()` returns STOPPED/interrupted Interactive projects with checkpoint path/status.

- [ ] **Step 1: Write RED checkpoint round-trip test**

Create source/output files, hashes, blueprint fingerprint and settings; write checkpoint; reload and assert every field survives JSON serialization exactly.

- [ ] **Step 2: Write RED tamper/source-change validation tests**

Assert resume fails separately for:
- changed source hash;
- changed blueprint fingerprint/schema;
- modified partial DOCX hash;
- missing output;
- invalid last event index.

Error result must name the failed validation dimension.

- [ ] **Step 3: Implement atomic checkpoint persistence**

Write JSON to temp sibling then `os.replace` to final checkpoint metadata file. Save the Word document first, compute its actual SHA-256, then persist metadata. Never record a checkpoint hash before save returns successfully.

- [ ] **Step 4: Integrate checkpoint triggers**

Observer triggers after table/image/section boundaries, every configured event interval, Stop, and major milestones. Each save increments the existing truthful save history/audit path.

- [ ] **Step 5: Implement project rehydration/resume entry point**

`InteractiveRebuildService.resume(project_id, control, observer)` must:
1. rehydrate `ProjectPaths` from DB/root;
2. load source/project JSON and blueprint;
3. validate source + blueprint + output hashes;
4. open checkpoint DOCX in the dedicated visible Word session;
5. re-establish expected semantic state;
6. continue from `last_completed_event_index + 1`.

- [ ] **Step 6: Add real Word stop/restart/resume tests**

Required:
- stop mid-paragraph, close session, resume, final L0 PASS;
- stop after/mid table at atomic boundary, resume, final table L1 PASS.

- [ ] **Step 7: Run checkpoint unit tests + Word resume tests**

Expected: PASS with real saved partial outputs and actual hashes.

- [ ] **Step 8: Version checkpoint**

Commit: `feat: add crash-safe interactive resume checkpoints`.

---

### Task 12: Word state guard, live progress, and safe-boundary verification

**Files:**
- Create: `src/word_replica/interactive/verification.py`
- Modify: `src/word_replica/domain/reconstruction.py` — add `WordStateSnapshot` and `LiveVerificationResult`.
- Modify: `src/word_replica/renderers/interactive_word.py`
- Modify: `src/word_replica/services/interactive_rebuild.py`
- Create: `tests/unit/test_interactive_verification.py`
- Extend: `tests/integration/word/test_interactive_resume.py`

**Interfaces:**
- Produces: `WordStateSnapshot` and `LiveVerificationResult` domain types.
- Produces: `LiveVerifier.verify_boundary(source_model, output_path, boundary) -> LiveVerificationResult`.
- Observer emits `InteractiveProgress` snapshots after events without accessing Tkinter directly.

- [ ] **Step 1: Write RED state mismatch test**

Inject a controller whose current document/story/table context differs from expected state before the next event. Assert executor:
- does not execute the next event;
- transitions to PAUSED;
- emits/audits `WORD_STATE_MISMATCH` with event/source ID;
- attempts deterministic recovery only from last validated checkpoint.

- [ ] **Step 2: Implement state snapshot/guard**

Track document identity, story type, collapsed insertion position/range signature, section index, table/cell logical context, last completed event. Do not rely on foreground keyboard focus.

- [ ] **Step 3: Write RED progress snapshot tests**

Given event index N, assert progress reports:
- total/completed events;
- visible characters completed/total;
- current source element/location;
- completed tables/images/sections;
- last checkpoint timestamp/index;
- current run state and verification status.

- [ ] **Step 4: Implement live verifier at safe boundaries**

At completed table/image/section/configured milestone, save or inspect a stable output snapshot, reparse it, and compare the relevant canonical projection/subtree. A mismatch returns `PAUSED_FIDELITY_MISMATCH`; it never edits projections to hide the difference.

- [ ] **Step 5: Add real Word mismatch/recovery integration test**

Use controlled test hook to move expected range/state between events, prove next character is not inserted into wrong location, restore from checkpoint, continue or stop safely.

- [ ] **Step 6: Version checkpoint**

Commit: `feat: guard Word state and verify interactive progress`.

---

### Task 13: Controlled L4 Word-to-PDF visual QA and final Interactive QA contract

**Files:**
- Create: `src/word_replica/qa/word_render.py`
- Modify: `src/word_replica/qa/render.py`
- Modify: `src/word_replica/services/interactive_rebuild.py`
- Extend: `tests/unit/test_render_diff.py`
- Create/extend: `tests/integration/word/test_interactive_word_structures.py`

**Interfaces:**
- Produces: `export_docx_to_pdf_with_word(docx_path, pdf_path, visible=False)` using a dedicated controlled Word session.
- Consumes existing `compare_pdfs` and `run_l0_l3`.
- Final Interactive `PASS` requires required L0-L3 PASS plus capability policy; L4 is attached separately and classified by calibrated tolerance.

- [ ] **Step 1: Write RED export contract test with fake Word document**

Assert export opens source read-only, calls Word fixed-format PDF export, closes without saving/mutating source, and cleanup errors do not overwrite an earlier real export failure reason.

- [ ] **Step 2: Implement Word PDF export**

Use Word object model `ExportAsFixedFormat`/equivalent PDF format constant. Source and reconstructed documents must be rendered by the same helper/environment for L4 comparison.

- [ ] **Step 3: Extend render QA tests for aggregate policy**

Assert page-count mismatch or page metric beyond tolerance sets `within_tolerance=False`; exact/near-identical test images pass.

- [ ] **Step 4: Integrate final Interactive QA**

After final Word save:
1. reparse output;
2. run L0-L3;
3. enforce Maximum Fidelity capability policy;
4. export source/output to PDF under same Word environment;
5. compare pages;
6. write existing HTML QA report with L4 data;
7. only then mark project COMPLETED/PASS/WARN/FAIL.

Do not mark successful completion merely because executor reached `EndDocument`.

- [ ] **Step 5: Windows integration verification**

Run one plain text and one academic fixture through Interactive final QA and assert PDF artifacts + QA report exist.

- [ ] **Step 6: Version checkpoint**

Commit: `feat: add controlled Word visual QA for interactive runs`.

---

### Task 14: RebuildService routing and desktop mode-first live controller UX

**Files:**
- Modify: `src/word_replica/services/rebuild.py`
- Modify: `src/word_replica/desktop.py`
- Create: `tests/unit/test_desktop_interactive.py`
- Extend: `tests/unit/test_rebuild_service.py`
- Extend: `tests/unit/test_desktop_app.py`

**Interfaces:**
- `RebuildService.rebuild()` routes `ReconstructionMode.INTERACTIVE` to `InteractiveRebuildService`; Instant path remains byte/behavior compatible.
- Desktop worker emits queue events: `preflight`, `progress`, `state`, `checkpoint`, `verification`, `warning`, `error`, `result`.
- Desktop calls control methods thread-safely; worker never calls Tkinter widgets.

- [ ] **Step 1: Write RED RebuildService isolation test**

Inject fake Instant renderer and fake Interactive service. Assert:
- default `RebuildOptions()` uses existing Instant renderer;
- `reconstruction_mode=INTERACTIVE` never calls `_select_renderer()`/Instant renderer;
- Interactive on non-Windows/Word unavailable returns explicit failure/preflight result rather than silently falling back to Pure DOCX.

- [ ] **Step 2: Implement routing with minimal Instant diff**

At the start of `rebuild`, after extension/source/project prerequisites and before Instant renderer selection, branch to the Interactive orchestration using a focused helper. Do not refactor the existing Instant renderer selection/render/QA flow beyond the minimal branch required.

- [ ] **Step 3: Write RED desktop state mapping tests**

Assert new `DesktopState` defaults to Instant and maps Interactive options exactly. When Interactive is selected:
- visible Word is fixed;
- speed preset/custom rate/object delay/fidelity/checkpoints/preflight map into `InteractiveOptions`;
- renderer selection moves to Advanced and cannot choose Pure DOCX as an Interactive fallback.

- [ ] **Step 4: Implement mode-first UI**

Main panel order:
1. source;
2. reconstruction mode (`Instant`, `Interactive`);
3. fidelity/metadata;
4. Interactive options shown only in Interactive mode;
5. Advanced Instant renderer options.

Replace footer copy with transparent automation wording from approved spec §21.5.

- [ ] **Step 5: Implement preflight panel and live compact controller**

Live controls/state:
- determinate overall progress;
- semantic location;
- char/table/image/section counters;
- speed selector/slider;
- Pause/Resume single stateful button;
- Stop;
- last checkpoint;
- live verification status.

`_poll_events` handles only queue payloads and updates Tkinter on main thread.

- [ ] **Step 6: Implement resumable project card**

On app startup, call `ProjectStore.list_resumable_projects()`. Display source name, progress from checkpoint, last checkpoint time, Resume. Validate before enabling Resume and show exact disabled reason if source/checkpoint changed.

- [ ] **Step 7: Add desktop tests without launching real Tk window where avoidable**

Unit-test state mapping, control callback routing, event-to-viewmodel updates and transparency copy. Keep existing build/installer tests passing.

- [ ] **Step 8: Run full non-Word suite**

```powershell
python -m pytest -q
```
On non-Windows, expected: all non-Word tests PASS and Word-marked tests SKIP only. No existing Instant unit/acceptance failure is allowed.

- [ ] **Step 9: Version checkpoint**

Commit: `feat: integrate interactive reconstruction into desktop app`.

---

### Task 15: Full Windows + Microsoft Word Interactive acceptance corpus and dual release gate

**Files:**
- Create: `tests/acceptance/test_interactive_acceptance_matrix.py`
- Add/extend fixture builders in: `tests/fixtures/build_fixtures.py`
- Add any new corpus DOCX fixtures under: `tests/fixtures/corpus/`
- Modify: `RUN_WINDOWS_RELEASE_GATE.ps1`
- Modify: `BUILD_WINDOWS_APP.ps1`
- Extend: `tests/unit/test_windows_builder_scripts.py`
- Create: `docs/INTERACTIVE_WINDOWS_RELEASE.md`

**Interfaces:**
- Release script emits explicit section headings and exits nonzero if either gate fails.
- Interactive acceptance always uses real Word (`WORD_REPLICA_WORD_TESTS=1`) and never substitutes the Instant renderer.

- [ ] **Step 1: Write RED release-script tests**

Assert script text contains and enforces:
```text
=== INSTANT GATE ===
=== INTERACTIVE GATE ===
WORD REPLICA WINDOWS RELEASE GATE: PASS
```
and uses PowerShell failure handling (`$ErrorActionPreference = "Stop"` plus explicit `$LASTEXITCODE` checks) so a failed subcommand cannot fall through to PASS.

- [ ] **Step 2: Define Interactive acceptance matrix**

Create parametrized real-Word tests for at least these 16 scenarios from the approved spec:
1. plain mixed formatting;
2. headings/styles;
3. lists/tab stops/page breaks;
4. merged/shaded/bordered table;
5. merged-cell text typing;
6. inline image;
7. floating image geometry;
8. multiple sections/orientations;
9. headers/footers;
10. notes/fields/bookmarks;
11. nested table where supported;
12. academic complex fixture;
13. pause/resume mid-paragraph;
14. pause/resume mid-table;
15. stop + restart + resume;
16. Word state mismatch recovery.

For fidelity variants where behavior differs, parametrize `STANDARD` and `MAXIMUM`.

- [ ] **Step 3: Add invariant assertions beyond final status**

Every applicable acceptance test must assert:
- source SHA-256 unchanged;
- output exists and opens in Word;
- L0/L1 required PASS;
- Maximum policy has zero silent unsupported downgrades;
- audit contains real Interactive events;
- per-character executor metrics/counters equal blueprint visible-character count;
- table/image counts match supported source model counts;
- final custom save count equals save history rows.

- [ ] **Step 4: Update release gate**

`INSTANT GATE` runs the previously verified complete suite/Instant acceptance path first. `INTERACTIVE GATE` runs new Interactive unit + real Word integration + acceptance tests. Overall PASS prints only after both commands return zero.

- [ ] **Step 5: Run full Windows gate**

On target Windows machine:
```powershell
powershell -ExecutionPolicy Bypass -File .\RUN_WINDOWS_RELEASE_GATE.ps1
```
Required final evidence:
```text
=== INSTANT GATE ===
INSTANT RESULT: 0 failed
=== INTERACTIVE GATE ===
INTERACTIVE RESULT: 0 failed
WORD REPLICA WINDOWS RELEASE GATE: PASS
```
Do not claim release readiness from local non-Word skips.

- [ ] **Step 6: Build EXE only after gate PASS**

```powershell
powershell -ExecutionPolicy Bypass -File .\BUILD_WINDOWS_APP.ps1
```
Verify `dist\WordReplica.exe` exists and build script reports nonzero failure if gate fails.

- [ ] **Step 7: Manual release smoke checklist on built EXE**

Using a copy of corpus fixtures, verify through the installed GUI:
- select Interactive;
- Word opens visible;
- mixed text appears character-by-character with active formatting;
- table visibly constructs then types cells;
- image visibly inserts/configures;
- speed can change live;
- Pause holds between events and Resume continues;
- Stop preserves resumable partial output;
- restart app shows Resume card;
- final QA/report opens;
- Instant mode still rebuilds normally.

Record exact EXE SHA-256 and release-gate output in `docs/INTERACTIVE_WINDOWS_RELEASE.md`.

- [ ] **Step 8: Final version checkpoint**

In a Git checkout:
```bash
git add src tests RUN_WINDOWS_RELEASE_GATE.ps1 BUILD_WINDOWS_APP.ps1 docs/INTERACTIVE_WINDOWS_RELEASE.md
git commit -m "feat: ship interactive Word reconstruction"
```
In the packaged tree, record source ZIP/EXE hashes and the full Windows gate transcript instead.

---

## Plan Self-Review

### Spec coverage map

- Configuration/event schema: Task 1.
- Blueprint and character-by-character semantics: Task 2.
- Visible Word COM/no SendKeys: Task 3.
- Slow/Fast/Custom/Maximum + Pause/Resume/Stop: Task 4.
- Formatting while typing: Task 5.
- Tables/cell typing/nested tables: Task 6.
- Images/geometry/capability behavior: Task 7.
- Sections/headers/footers: Task 8.
- Notes/lists/fields/bookmarks/TOC semantics: Task 9.
- Preflight/Maximum Fidelity block: Task 10.
- Persisted checkpoints/restart resume/source hash validation: Task 11.
- Word state mismatch + live verification/progress: Task 12.
- L4 and final QA gate: Task 13.
- Mode-first Tkinter UX/live controller/resume card/transparency copy: Task 14.
- 16-scenario Windows corpus + independent Instant/Interactive gates + EXE: Task 15.
- Security/source integrity/no fake authorship: Global Constraints + Tasks 10/11/15.
- Truthful audit/save count: Tasks 10/11/15.

### Type consistency check

The plan uses these stable names throughout:
- `ReconstructionMode`
- `InteractiveSpeedMode`
- `InteractiveFidelity`
- `CapabilityClass`
- `InteractiveRunState`
- `InteractiveOptions`
- `ReconstructionEvent`
- `ReconstructionBlueprint`
- `SemanticLocation`
- `InteractiveProgress`
- `ControlDecision`
- `ExecutionOutcome`
- `DrawingRef`
- `CapabilityDecision`
- `BlueprintCompiler`
- `SpeedController`
- `InteractiveRunControl`
- `InteractiveWordController`
- `InteractiveWordRenderer`
- `PreflightReport`
- `PreparedInteractiveRun`
- `InteractiveCheckpoint`
- `InteractiveCheckpointStore`
- `WordStateSnapshot`
- `LiveVerificationResult`
- `LiveVerifier`
- `InteractiveRebuildService`

No later task should rename these without first updating this plan/spec and all dependent task interfaces.

### Execution discipline

For every task:
1. write the failing test first;
2. run it and confirm the expected RED failure;
3. implement only enough production behavior for GREEN;
4. run targeted tests;
5. run relevant Instant regression tests before moving to the next slice;
6. obtain real Windows Word evidence wherever the task adds COM behavior;
7. never mark a task complete from a skipped Word test.
