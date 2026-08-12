# Word Replica v1 — Interactive Reconstruction Design

**Date:** 2026-08-11  
**Status:** Design approved in conversation; implementation not started  
**Target platform:** Windows desktop with Microsoft Word installed  
**Baseline source:** WordReplica App Builder v3 (`/mnt/data/wrapp_v3`)  
**Baseline invariant:** The existing Instant Reconstruction engine is preserved as a separate fallback path. The user has already verified its Windows release gate at **97/97 tests passing**.

## 1. Purpose

Add a new **Interactive Reconstruction** mode that creates a new Microsoft Word document visibly, step by step, from a read-only source DOCX.

The defining behavior is that the new document must actually be constructed through Microsoft Word operations rather than assembled first and merely displayed afterward:

- body text is inserted **one character at a time**;
- formatting is active **while each character is inserted**;
- table geometry and formatting are created step by step;
- text inside table cells is also inserted **one character at a time**;
- images are inserted as new Word objects and then sized, positioned, anchored, wrapped, cropped, rotated, and layered according to the source blueprint;
- sections, headers, footers, notes, fields, lists, and other supported Word structures are reconstructed as real Word structures;
- the Word window remains visible during the interactive mode;
- Word Replica remains the process controller and exposes Start/Pause/Resume/Stop/Speed controls.

The goal is a **true visible reconstruction** of a new Word document, not a copy/paste replay.

## 2. Product Modes

Word Replica keeps two independent reconstruction modes.

### 2.1 Instant Reconstruction

The existing stable path remains unchanged in behavior and architecture.

Purpose:

- fast reconstruction;
- current Pure DOCX / Word compatibility workflow;
- existing QA and audit pipeline;
- fallback when Interactive Reconstruction is unavailable or unnecessary.

The new feature must not replace or silently alter this mode.

### 2.2 Interactive Reconstruction

A new Word-only renderer that:

1. parses the original DOCX into a detailed canonical blueprint;
2. runs a preflight against the blueprint and the local Word environment;
3. opens a new blank Microsoft Word document;
4. reconstructs the document through atomic Word COM operations;
5. shows the document being created live;
6. supports Pause, Resume, Stop, speed changes, checkpoints, and later resume;
7. performs final L0-L4 QA before reporting success.

Interactive Reconstruction requires Microsoft Word desktop automation. It does not fall back to Pure DOCX while claiming to remain interactive.

## 3. Explicit Non-Goals and Transparency Boundary

Interactive Reconstruction may visually resemble manual transcription because content appears character by character and objects are built step by step. It remains an automated reconstruction feature.

The implementation must **not** add:

- random intentional typos;
- fake corrections intended to imitate a human author;
- random pauses designed to fabricate a human work pattern;
- simulated keyboard history as proof of manual authorship;
- fabricated edit/save timestamps or fake revision history;
- false metadata claiming a person manually authored the output.

Slow/Fast/Custom speed is a deterministic user-selected presentation/execution rate. Audit records remain truthful and identify the process as automated Word Replica reconstruction.

## 4. Core Architectural Rule

Do not retrofit character-by-character behavior into `WordComRenderer` in a way that destabilizes the verified Instant path.

Add a separate renderer and supporting subsystem:

```text
RebuildService
├── Instant path
│   ├── PureDocxRenderer
│   └── WordComRenderer
└── Interactive path
    ├── BlueprintCompiler
    ├── PreflightAnalyzer
    ├── ReconstructionExecutor
    ├── InteractiveWordController
    ├── SpeedController
    ├── ReconstructionCheckpointManager
    └── LiveVerifier
```

The Instant and Interactive renderers may share parsing/model utilities, metadata policy, project storage, QA code, source guards, and audit primitives. They must not share mutable reconstruction state.

## 5. Configuration Model

Introduce a first-class reconstruction mode rather than overloading `RendererChoice`.

Conceptual configuration:

```python
ReconstructionMode = INSTANT | INTERACTIVE
InteractiveSpeedMode = SLOW | FAST | CUSTOM | MAXIMUM
InteractiveFidelity = STANDARD | MAXIMUM
```

`RebuildOptions` should gain an interactive configuration object rather than many unrelated top-level fields.

Conceptually:

```python
InteractiveOptions(
    speed_mode=InteractiveSpeedMode.FAST,
    characters_per_second=25.0,
    object_step_delay_ms=150,
    fidelity=InteractiveFidelity.MAXIMUM,
    checkpoint_after_tables=True,
    checkpoint_after_images=True,
    checkpoint_after_sections=True,
    checkpoint_event_interval=500,
    verify_during_run=True,
    block_on_unsupported=True,
)
```

Rules:

- `characters_per_second` is used only where the selected speed mode requires it.
- `MAXIMUM` speed still emits one character insertion operation per character; it simply introduces no intentional delay between operations.
- object delays are deterministic, user-controlled delays between visible object-construction steps.
- Interactive Reconstruction always uses visible Microsoft Word in v1 of this feature.

## 6. Blueprint Layer

### 6.1 Why a Blueprint Is Required

`DocumentModel` is currently sufficient for canonical reconstruction/QA, but Interactive Reconstruction needs an explicit, ordered execution plan.

The source DOCX is parsed once. The reconstruction executor must never continuously read the source while typing. This gives deterministic resume behavior and protects the source.

### 6.2 Blueprint Contents

A `ReconstructionBlueprint` contains at minimum:

- source SHA-256;
- source canonical fingerprint;
- ordered reconstruction events;
- total event count;
- total visible character count;
- counts of paragraphs, tables, images, sections, fields, notes, and complex objects;
- preflight requirements such as fonts and source assets;
- classification of complex elements as `RECONSTRUCTED`, `PRESERVED`, or `UNSUPPORTED`;
- blueprint schema version.

The blueprint is serialized to the project folder so a stopped/crashed run can be resumed without reparsing a potentially changed source.

### 6.3 Event Model

Events are semantic operations, not mouse/keyboard replay instructions.

Representative event types:

```text
BeginDocument
BeginSection
ApplySectionProperties
BeginParagraph
ApplyParagraphProperties
ApplyRunProperties
InsertCharacter
InsertTab
InsertLineBreak
InsertPageBreak
EndParagraph
BeginTable
SetTableProperties
SetColumnWidth
SetRowHeight
MergeCells
SetCellProperties
EnterCell
LeaveCell
EndTable
InsertImage
SetImageDimensions
SetImageWrapping
SetImageAnchor
SetImagePosition
SetImageCrop
SetImageRotation
SetImageZOrder
CreateHeader
CreateFooter
CreateFootnote
CreateEndnote
CreateField
CreateBookmark
CreateListBinding
CheckpointBoundary
EndDocument
```

The event schema must carry stable source element IDs so errors and QA findings can point back to the canonical source element.

### 6.4 Atomicity

An **atomic reconstruction operation** is the smallest operation at which pause/stop is safe.

For visible text, `InsertCharacter` is atomic.

For object operations, one COM property mutation is normally atomic. A table merge or image insertion is atomic as one Word operation even if Word internally performs multiple actions.

Pause/Stop never interrupts a COM call in progress. It takes effect immediately after the current atomic operation returns.

## 7. Microsoft Word Control Model

### 7.1 Word Session

Interactive mode owns one dedicated Word application process via COM for the lifetime of the active reconstruction.

Required behavior:

- `DispatchEx("Word.Application")`;
- Word visible;
- alerts suppressed only where safe;
- one owned output document;
- source is never opened writable;
- output document is a newly created document or a verified checkpoint document when resuming;
- COM initialization/uninitialization is owned by the executor thread;
- teardown must tolerate Word RPC disconnects without reporting a false reconstruction success.

### 7.2 No SendKeys

Character insertion must use Word object-model APIs (`Range`/equivalent semantic insertion), one character at a time.

Do not use:

- SendKeys;
- clipboard paste as the main reconstruction mechanism;
- synthetic Windows keyboard events;
- mouse automation.

This avoids focus/layout/keyboard-language fragility while still creating the document incrementally in Word.

### 7.3 Visible Cursor and Foreground Word

The active output range should track the current insertion point. Word should visibly show the current document position where technically practical.

Word Replica may request Word activation when reconstruction starts or resumes, but correctness must not depend on Windows keyboard focus.

### 7.4 User Editing During Reconstruction

The output document is considered **controller-owned** while reconstruction is active.

The user may:

- watch;
- pause;
- resume;
- stop;
- adjust speed.

The user must not edit the document while active. The controller continuously maintains an expected Word state and pauses on state mismatch rather than blindly continuing.

## 8. Text Reconstruction

### 8.1 Character-by-Character Requirement

No event may insert an entire run or paragraph text in Interactive mode.

For every visible Unicode character in a supported text story:

1. the required run formatting is established;
2. one character is inserted;
3. progress is updated;
4. speed control is honored;
5. pause/stop is checked before the next event.

Tabs, breaks, fields, and structural markers are represented by dedicated events rather than pretending they are ordinary characters when Word semantics differ.

### 8.2 Formatting During Typing

Formatting is applied **before** the first character requiring it.

Supported run-level properties should include, as modeled/available:

- font family;
- font size;
- bold;
- italic;
- underline;
- strike;
- color;
- highlight;
- superscript/subscript;
- hidden text;
- language;
- character spacing/position where modeled.

A transition from one run to another updates formatting before the next character is inserted.

### 8.3 Paragraph Formatting

Paragraph properties are established before typing the paragraph content where Word semantics permit it:

- paragraph style;
- alignment;
- left/right/first-line/hanging indent;
- spacing before/after;
- line spacing;
- keep-next/keep-lines;
- widow/orphan behavior where modeled;
- page-break-before;
- tab stops;
- borders/shading where modeled;
- list/numbering binding.

## 9. Tables

Tables are a first-class fidelity target, not a text container approximation.

### 9.1 Required Reconstruction Sequence

For each source table:

1. create the Word table with the required initial grid;
2. set table-level properties;
3. set column widths/grid behavior;
4. set row heights and row properties;
5. establish cell margins and vertical alignment;
6. apply merges using a merge plan that remains valid after Word mutates the cell collection;
7. apply borders/shading/table style behavior;
8. enter each logical cell in deterministic order;
9. reconstruct every block in the cell;
10. type all cell text character by character with formatting active during insertion;
11. reconstruct nested tables recursively;
12. verify the table structure before proceeding in Maximum Fidelity mode.

### 9.2 Table Fidelity Requirements

Where represented in the source/blueprint, preserve:

- table alignment;
- preferred width and width type;
- fixed/auto layout;
- grid column widths;
- row height + exact/at-least semantics;
- header rows;
- cant-split behavior;
- horizontal/vertical merges;
- borders (side/style/size/color/spacing);
- shading/fill;
- cell margins;
- cell vertical alignment;
- cell text direction where supported;
- paragraph/run formatting inside cells;
- nested tables.

Maximum Fidelity must not silently replace a complex table with a default Word table.

## 10. Images and Drawing Objects

### 10.1 Image Asset Source

The original image bytes are extracted from the source package as a binary asset. The output image is inserted as a **new Word image object** using those bytes; the finished source drawing object is not copied wholesale as the normal reconstruction path.

### 10.2 Image Reconstruction Steps

Where modeled and supported:

1. insert image asset;
2. set inline vs floating representation;
3. set width/height;
4. set aspect-lock behavior;
5. set wrapping type;
6. set anchor association;
7. set horizontal relative reference and position;
8. set vertical relative reference and position;
9. set distance from surrounding text;
10. set crop;
11. set rotation;
12. set z-order / behind-or-in-front-of-text;
13. verify resulting drawing geometry.

### 10.3 Fidelity Classification

If the source drawing contains semantics not yet modeled, Maximum Fidelity must classify the element before reconstruction as one of:

- `RECONSTRUCTED`: fully recreated through supported Word operations;
- `PRESERVED`: original OOXML/object must be preserved because reconstruction would lose semantics;
- `UNSUPPORTED`: cannot be faithfully represented or safely preserved.

No silent conversion to a generic inline image is permitted in Maximum Fidelity.

## 11. Sections, Headers, Footers, Notes, Lists, and Fields

### 11.1 Sections

Section boundaries are executed at canonical boundaries, not pre-created before body content.

Support where modeled:

- page size;
- portrait/landscape;
- margins;
- gutter;
- header/footer distances;
- section break type;
- columns;
- header/footer linkage;
- first-page/even-odd behavior;
- page numbering configuration where modeled.

### 11.2 Headers and Footers

Create the appropriate Word story/section header or footer, then reconstruct supported content using the same text/table/image rules as the body.

### 11.3 Footnotes and Endnotes

Create the real Word note and reconstruct note content incrementally. The visible note text is typed character by character.

### 11.4 Lists and Numbering

Use actual Word numbering/list semantics. Do not type bullet/number characters as plain text when the source is a true list.

### 11.5 Fields and TOC

When the source uses a real Word field, create a real Word field where supported rather than typing only its visible result.

Examples include:

- PAGE;
- NUMPAGES;
- TOC;
- cross-reference fields;
- other supported field instructions.

Field result text may be refreshed by Word at defined safe points. Maximum Fidelity must report fields that cannot be semantically reconstructed.

## 12. Complex Object Policy

Complex Word elements include, but are not limited to:

- text boxes;
- shapes;
- equations;
- charts;
- SmartArt;
- OLE/embedded Excel objects;
- content controls;
- comments;
- tracked revisions;
- unsupported DrawingML/VML constructs.

Every such element receives an explicit capability classification during preflight:

### RECONSTRUCTED

The feature can recreate the object with equivalent Word semantics.

### PRESERVED

The feature cannot reasonably recreate the object without semantic loss, but the original object can be transferred/preserved safely. The final report explicitly identifies it as preserved rather than rebuilt.

### UNSUPPORTED

The object cannot be faithfully recreated or safely preserved.

In **Maximum Fidelity**:

- default behavior is to block before reconstruction if any element is `UNSUPPORTED`;
- a user may explicitly choose a permitted preservation fallback only when the element is `PRESERVED`;
- Word Replica never silently substitutes a simpler object.

## 13. Preflight

Preflight runs before a blank Word document begins visible reconstruction.

It reports at minimum:

- source integrity and SHA-256;
- blueprint compilation status;
- Microsoft Word availability;
- required image assets;
- fonts referenced by the source and whether they are available;
- tables/merges that require complex reconstruction;
- sections;
- fields;
- notes;
- complex object capability classifications;
- whether Maximum Fidelity can proceed.

If Maximum Fidelity is blocked, the run does not start unless the user explicitly chooses an allowed fallback.

## 14. Speed Controller

### 14.1 Modes

Expose:

- Slow;
- Fast;
- Custom;
- Maximum.

The exact presets may be tuned during implementation; correctness must not depend on them.

### 14.2 Semantics

All modes perform the same reconstruction operations in the same order.

Only intentional inter-operation delay changes.

`Maximum` means zero intentional character delay, **not** bulk insertion.

The user may alter speed while reconstruction is active. The new rate applies before the next eligible delay.

## 15. Pause, Resume, Stop

### 15.1 Pause

Pause is cooperative and atomic:

- finish the current Word COM operation;
- record current event index and state;
- enter PAUSED;
- do not execute the next event until resumed.

### 15.2 Resume

Resume continues from the next unexecuted event using the same Word document/session when still active.

If resuming from a persisted checkpoint after application restart, the checkpoint document is reopened and state is revalidated before execution continues.

### 15.3 Stop

Stop is a safe termination, not a discard:

1. complete the current atomic operation;
2. write a real checkpoint save;
3. persist resume metadata;
4. mark the project `STOPPED`;
5. preserve the partial output.

The UI may separately offer an explicit discard action.

## 16. Checkpoint and Crash Recovery

### 16.1 Checkpoint Triggers

At minimum support configurable checkpoints:

- after each table;
- after each image;
- after section boundaries;
- after a configurable number of reconstruction events;
- on Stop;
- at defined major milestones.

### 16.2 Checkpoint Data

Persist:

- project ID;
- source SHA-256;
- source model fingerprint;
- blueprint SHA-256/fingerprint + schema version;
- output document path and SHA-256;
- last completed event index;
- current section/block/table/cell source element IDs;
- save sequence;
- current reconstruction settings relevant to correctness;
- timestamp of the actual checkpoint;
- status.

### 16.3 Resume Validation

A persisted run may resume only if:

- source SHA-256 still matches;
- blueprint fingerprint/schema matches;
- checkpoint document SHA-256 matches recorded value;
- expected structural checkpoint state can be re-established in Word.

If the source changed, the run is not resumed against the new source.

## 17. Word State Guard

The executor tracks expected state sufficient to detect dangerous drift, including:

- expected Word process/session;
- expected output document;
- expected story/range;
- expected section;
- expected table/cell context when applicable;
- last completed event;
- document fingerprint/checkpoint state at configured boundaries.

On mismatch:

1. pause execution;
2. emit `WORD_STATE_MISMATCH` to audit/UI;
3. attempt a deterministic recovery to the last validated state;
4. if recovery is unsafe, stop and preserve the latest valid checkpoint.

The renderer never keeps typing into an unknown Word range.

## 18. Live Verification

Maximum Fidelity may perform lightweight verification during reconstruction at safe boundaries.

Minimum verification boundaries:

- completed table;
- completed image placement;
- completed section;
- configurable paragraph/event milestones.

Live verification is not allowed to mutate source truth or hide differences. Its purpose is early failure detection.

A mismatch may transition to:

```text
PAUSED — fidelity mismatch
```

before the remaining document is constructed.

## 19. Final QA Contract

Interactive Reconstruction does not report successful completion immediately after the final character/object operation.

It must save and run QA.

### L0 — Content

All canonical textual content and required visible values match.

### L1 — Structure

Paragraphs, tables, rows/cells/merges, sections, headers/footers, notes, lists, fields, bookmarks, and other modeled structures match according to the canonical QA contract.

### L2 — Formatting

Compare modeled paragraph/run/table/cell/drawing formatting.

### L3 — Layout / Geometry

Compare modeled section geometry, table dimensions, image geometry/positioning, and relevant layout metadata.

### L4 — Visual

On a controlled Windows/Word environment:

1. render source through Word to PDF;
2. render reconstructed output through the same Word environment to PDF;
3. rasterize comparable pages;
4. compute visual differences;
5. report per-page and aggregate difference metrics.

No universal claim of pixel-identical layout is made across different Office/font/printer environments. L4 is meaningful only under a controlled shared environment.

### Maximum Fidelity Success Rule

Maximum Fidelity may report `PASS` only when required L0-L3 checks pass and required unsupported-element policy is satisfied. L4 is reported separately with a documented threshold/policy once calibrated against the Windows acceptance corpus.

## 20. Metadata and Audit

Reuse the truthful metadata/audit philosophy of the existing application.

Record real events such as:

- blueprint compiled;
- preflight result;
- reconstruction started;
- character/object counters;
- pause/resume/stop;
- actual checkpoint saves;
- Word state mismatch/recovery;
- object capability classification;
- final save;
- QA result.

Output custom properties may identify Word Replica reconstruction, project ID, and actual save count as already supported.

Do not fabricate manual authorship or human edit history.

## 21. Desktop UI Design

### 21.1 Main Mode Selection

Replace the current renderer-first primary decision with a reconstruction-mode-first design:

```text
NAČIN REKONSTRUKCIJE
● Instant Reconstruction
○ Interactive Reconstruction
```

Renderer details move under Advanced where appropriate.

### 21.2 Interactive Options

When Interactive is selected, show:

- Word visible (fixed/required in first release);
- speed preset;
- custom characters/sec when relevant;
- object step delay;
- Standard vs Maximum Fidelity;
- checkpoint preferences;
- preflight preferences;
- metadata mode.

### 21.3 Preflight Screen/Panel

Before reconstruction starts, show counts and capability status for:

- characters;
- paragraphs;
- tables;
- images;
- sections;
- notes;
- fields;
- complex objects;
- unsupported/preserved objects;
- Maximum Fidelity readiness.

### 21.4 Live Controller

While active, the Word Replica window becomes a compact controller with:

- overall progress;
- current semantic position (paragraph/table/cell/image/section);
- current character count;
- counts of completed tables/images/sections;
- speed slider/preset;
- Pause/Resume;
- Stop;
- last checkpoint;
- live verification status.

Microsoft Word remains visible as the document work surface.

### 21.5 Transparency Copy

The current desktop footer text that says Word Replica does not simulate human typing must be revised for Interactive mode. The product copy must instead make the distinction accurately:

- Interactive Reconstruction performs automated character-by-character and object-by-object Word operations;
- the visible process may resemble manual transcription;
- it is not presented as proof of manual human authorship;
- save/audit/metadata records describe real automated operations only.

### 21.6 Resume Card

On startup, valid stopped/interrupted projects appear as resumable work with:

- source filename;
- progress percentage;
- last checkpoint time;
- Resume action.

Resume is disabled with a clear explanation when source/checkpoint validation fails.

## 22. Threading and UI Responsiveness

Word COM execution runs on its own worker thread with COM initialized on that thread.

The Tkinter main thread never performs long Word COM calls.

Worker-to-UI communication is event/queue based, consistent with the current desktop architecture.

Required worker events include:

- preflight progress/result;
- reconstruction progress;
- current semantic location;
- status transition;
- checkpoint completed;
- verification update;
- warning/error;
- final result.

Pause/stop/speed controls use thread-safe controller state, not direct Tkinter calls from the worker.

## 23. Failure Handling

Errors are classified by stage:

- parse/blueprint failure;
- preflight blocked;
- Word startup failure;
- Word COM operation failure;
- state mismatch;
- checkpoint save failure;
- resume validation failure;
- live verification failure;
- final QA failure.

Rules:

- source remains unchanged unless the existing explicit verified-overwrite option is used;
- never mark PASS after a failed build/reconstruction/install-like substage;
- preserve last valid partial output when possible;
- reasons shown to the user must identify the source element/event when available;
- Word RPC cleanup failures must not hide an earlier real reconstruction outcome, but also must not be silently ignored when they imply output uncertainty.

## 24. Test Strategy

Implementation follows TDD. No Interactive production behavior is added without a failing test first.

### 24.1 Unit Tests

Cover at minimum:

- interactive option validation;
- blueprint event compilation and deterministic ordering;
- one `InsertCharacter` event per visible character;
- run formatting transition ordering;
- paragraph property ordering;
- table merge planning;
- table cell traversal and nested content;
- drawing event compilation;
- speed timing policy without relying on wall-clock sleeps in tests;
- pause/resume atomic boundaries;
- stop checkpoint behavior;
- source/blueprint/checkpoint resume validation;
- complex object capability classification;
- desktop state mapping and live controller state;
- audit events.

### 24.2 COM Contract Tests

Use minimal fakes/test doubles only where unavoidable to prove call ordering and state-machine behavior without requiring Word for every test.

Examples:

- formatting is set before character insertion;
- exactly one insertion call is made per character;
- no bulk paragraph insertion path is used in Interactive mode;
- table structural operations occur before cell text typing;
- image geometry operations follow image insertion;
- pause takes effect between atomic events.

### 24.3 Windows + Microsoft Word Integration Tests

Real Word tests are mandatory before release.

Minimum targeted fixtures:

1. plain text with mixed run formatting;
2. headings/styles;
3. lists/tab stops/page breaks;
4. merged/shaded/bordered tables;
5. text inside merged cells;
6. inline image;
7. floating anchored image with wrapping/position;
8. multiple sections/orientations;
9. headers/footers;
10. notes/fields/bookmarks;
11. nested table where supported;
12. complex academic document combining major features;
13. pause/resume mid-paragraph;
14. pause/resume mid-table;
15. stop + restart + resume;
16. Word state mismatch recovery.

Both Standard and Maximum Fidelity variants are tested where behavior differs.

### 24.4 Regression Rule

The existing Instant Windows release gate remains part of every release candidate. Interactive changes may not regress the previously verified Instant acceptance behavior.

The release gate therefore has two independent sections:

```text
INSTANT GATE
INTERACTIVE GATE
```

A release cannot claim overall PASS when either required gate fails.

## 25. Performance and Timing

Interactive mode intentionally trades speed for visible construction.

Requirements:

- no correctness behavior depends on delays;
- zero-delay Maximum still uses atomic character events;
- speed can change live;
- estimated remaining time is optional and must be clearly approximate if added;
- checkpoints prevent long runs from becoming all-or-nothing.

No promise is made that Slow mode matches a particular person's physical typing speed. It is a deterministic presentation rate selected by the user.

## 26. Security and Source Protection

- source is hashed before reconstruction;
- source is treated read-only by default;
- output is created in a project output location;
- resume validates the source hash;
- binary assets originate from the source package unless explicitly otherwise supported;
- no macros are executed as part of reconstruction;
- external relationships are not automatically followed/executed;
- unsupported active content is classified explicitly.

## 27. Acceptance Definition for This Feature

Interactive Reconstruction v1 is considered implementation-complete only when all of the following are true:

1. Instant mode remains independently passing its existing Windows release gate.
2. Interactive mode visibly constructs a new Word document.
3. Every supported visible text character is inserted via an individual semantic Word insertion operation.
4. Formatting is active at insertion time.
5. Table structures are created through Word and cell text is typed character by character.
6. Supported images are inserted as new objects and their modeled geometry is reconstructed.
7. Pause/Resume/Stop are safe at atomic boundaries.
8. persisted checkpoints can resume after app restart when validation passes.
9. Maximum Fidelity never silently simplifies an unsupported element.
10. final success is gated by required QA, not merely by reaching the final event.
11. audit/save history reflects only real automated operations.
12. the full Windows + Word Interactive acceptance suite passes on the release environment.

## 28. Implementation Sequencing Constraint

Implementation should proceed in vertical, testable slices rather than attempting all Word features at once.

Recommended order:

1. configuration + event model;
2. blueprint compiler for plain paragraphs/runs;
3. Word executor with one-character insertion;
4. speed/pause/resume/stop state machine;
5. paragraph formatting;
6. tables + cell typing;
7. images and geometry;
8. sections/headers/footers;
9. notes/lists/fields;
10. checkpoint persistence + restart resume;
11. live verification;
12. complex-object capability policy;
13. L4 controlled render QA;
14. desktop UX integration;
15. full Windows release gate.

Each slice must preserve the Instant path and pass its own regression tests before the next slice begins.

## 29. Decisions Locked by User Approval

The following decisions are explicit and should not be reopened during implementation unless a technical blocker requires a design revision:

- Interactive mode must create text **character by character**.
- Text inside tables must also be typed character by character.
- Formatting must be active while text is being inserted, not applied as a post-pass.
- Tables must visibly be built step by step.
- Images must visibly be inserted/configured step by step.
- Microsoft Word remains visible and shows the reconstruction.
- Word Replica controls the process; the user does not manually edit the document during an active run.
- Both slow and fast/custom visible speeds are required.
- Pause/Resume/Stop and crash-safe checkpoints are required.
- Maximum Fidelity must not silently downgrade complex objects.
- The existing stable Instant engine remains a separate fallback.

## 30. Repository/Versioning Note

The current App Builder v3 working source in this environment is a packaged source tree and does **not** contain a `.git` directory. This design is therefore written into the project under `docs/superpowers/specs/`, but no Git commit hash is claimed for it in this environment.
