# Word Replica Interactive Reconstruction - Windows Release Evidence

Date: 2026-08-11

## Scope

This release candidate adds Interactive Reconstruction alongside the previously verified Instant Reconstruction engine. Interactive Reconstruction creates a new Microsoft Word document through real Word COM operations, including per-character text insertion, in-flow formatting changes, stepwise table construction/cell typing, supported image reconstruction, sections/stories, notes, bookmarks and fields.

The automation is transparent: character-by-character reconstruction is an automated process and is not evidence of manual authorship. The application does not fabricate typing mistakes, human pauses, edit history, timestamps, or authorship provenance.

## Local sandbox verification

Fresh non-Windows verification after the implementation:

```text
186 passed, 51 skipped, 0 failed
```

The skipped tests require Windows + Microsoft Word desktop automation and therefore cannot be counted as release evidence in the sandbox.

The local Maximum Fidelity preflight matrix is:

- READY: 01-09, 11, 14-17
- BLOCKED by design: 10 (comments/tracked revisions), 12 (chart/OLE parts), 13 (academic complex with unsupported review/chart/OLE content)

Maximum Fidelity never silently downgrades these unsupported structures.

## Required target Windows gate

Extract the entire release ZIP, then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\RUN_WINDOWS_RELEASE_GATE.ps1
```

Required transcript ending:

```text
=== INSTANT GATE ===
INSTANT RESULT: 0 failed
=== INTERACTIVE GATE ===
INTERACTIVE RESULT: 0 failed
WORD REPLICA WINDOWS RELEASE GATE: PASS
```

`START_HERE.cmd` runs this gate automatically before the EXE build and installation. The build script must not produce a release EXE if either gate fails.

## Interactive real-Word acceptance coverage

The target gate covers, among other cases:

- character-by-character Unicode text;
- mixed formatting applied before the next character;
- headings and custom paragraph styles;
- lists, tabs, line breaks, and page breaks;
- merged, shaded, bordered, and nested tables;
- cell text typed character-by-character;
- inline and floating image reconstruction/geometry;
- section orientation and page setup;
- headers and footers;
- footnotes/endnotes;
- real Word fields/TOC cached results;
- bookmarks/cross-references;
- pause/resume mid-paragraph;
- pause/resume/stop/restart in table content;
- source/output/checkpoint hash validation;
- Word-state mismatch pause before the next operation;
- same-Word source/output PDF export and L4 comparison;
- truthful save count and audit invariants;
- independent Instant and Interactive release gates.

## Manual installed-EXE smoke checklist

After the automated Windows gate and EXE build pass, verify in the installed GUI:

- Interactive mode opens Microsoft Word visibly;
- text visibly appears character-by-character;
- formatting is active while characters appear;
- a table visibly constructs before its cell text is typed;
- supported images are inserted/configured as Word objects;
- live speed changes take effect;
- Pause/Resume work between atomic operations;
- Stop preserves a resumable checkpoint;
- restarting Word Replica shows the resumable project;
- final QA report opens;
- Instant Reconstruction still works normally.

## Target evidence

Pending execution on the target Windows + Microsoft Word machine:

- Full gate transcript: PENDING
- Built EXE path: PENDING
- Built EXE SHA-256: PENDING
- Installed EXE SHA-256: PENDING
- Manual smoke checklist: PENDING

Do not mark this release as fully verified until the target evidence above is recorded.
