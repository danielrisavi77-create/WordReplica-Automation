# Word Replica v1 — Design Specification

Date: 2026-08-11
Status: Approved design, pending implementation-plan approval
Primary platform: Windows 10/11
Primary input: `.docx`

## 1. Purpose

Word Replica v1 is a local desktop/CLI application that accepts an existing `.docx` document and reconstructs it into a new `.docx` with maximum practical fidelity.

The application is intended for legitimate document reconstruction, archival, normalization, migration, QA, and reproducible Word automation. It may reconstruct text, formatting, sections, tables, images, footnotes/endnotes, headers/footers, fields, numbering, bookmarks, comments, and other supported Word structures.

The system must maintain a truthful audit trail of what the software actually did. It must not fabricate human authorship, fake typing behavior, falsify save history, forge editing duration, backdate timestamps, or manipulate provenance metadata to conceal the origin of content.

## 2. v1 Scope

### Included

- `.docx` input only.
- Read-only source by default.
- Working-copy reconstruction.
- Canonical internal document model.
- Dual rendering engines:
  - Microsoft Word COM renderer for full-fidelity mode when desktop Word is installed.
  - Pure DOCX renderer for fallback use without Word.
- Visible and background execution modes.
- Clean Replica and Full Fidelity modes.
- Fresh metadata and Preserve Legitimate Metadata modes.
- Real save/checkpoint history.
- Project-local artifact folder plus central project index.
- Structural and render-level QA.
- PASS / WARN / FAIL result states.
- Desktop GUI and CLI sharing the same engine.

### Excluded from v1

- PDF, image, scan, or ZIP inputs.
- OCR.
- Cloud document editing.
- Google Docs or LibreOffice automation.
- AI text generation as part of the reconstruction engine.
- Synthetic human typing simulation.
- Fake keystroke timing, fake pauses, deliberate typos, or fake corrections.
- Forged metadata or fabricated Word revision history.
- Claims that the reconstructed document was manually typed by a person.

## 3. User Modes

### 3.1 Rendering mode

#### Visible
Microsoft Word remains visible while the reconstruction runs. The user can watch document elements being inserted and checkpoints being saved.

#### Background
The renderer runs hidden where supported and produces the result without keeping Word visibly in the foreground.

### 3.2 Fidelity mode

#### Clean Replica
Reconstructs the visible final document while omitting historical/editorial baggage that is not needed for the final appearance, unless required for rendering.

Typical omissions may include old comments, resolved tracked changes, or hidden review artifacts.

#### Full Fidelity
Attempts to preserve legitimate internal Word structure in addition to the visible result, including supported:

- comments,
- tracked changes,
- bookmarks,
- fields,
- cross-references,
- hidden text,
- citation fields,
- footnotes/endnotes,
- section-level settings,
- relationship structures.

### 3.3 Metadata mode

#### Fresh Metadata — default
The output is treated as a newly reconstructed file. New timestamps and real reconstruction events are retained. Author/company fields may be set explicitly by the user where Word normally permits this.

#### Preserve Legitimate Source Metadata
Copies only selected descriptive source metadata that remains semantically valid for the reconstructed document. Reconstruction remains recorded in the application audit trail.

The default preservation allowlist is limited to fields such as title, subject, keywords, category, language, and explicitly selected custom descriptive properties. Author/company fields are copied only when the user deliberately enables them and they remain truthful.

A newly reconstructed output does **not** inherit the source document's creation/modification timestamps, revision count, total editing time, or historical save/revision values. Those fields belong to the new file's actual lifecycle.

The application must never falsify provenance fields for the purpose of making the document appear manually authored over a fictional period.

## 4. High-Level Architecture

```text
SOURCE.docx [read-only by default]
        |
        v
DOCX Forensic Parser
        |
        v
Canonical Document Model
        |
        +--------------------+
        |                    |
        v                    v
Word COM Renderer      Pure DOCX Renderer
        |                    |
        +----------+---------+
                   |
                   v
          reconstructed.docx
                   |
                   v
             QA / Diff Engine
                   |
                   v
            PASS / WARN / FAIL
```

Parallel services:

```text
Audit Engine
Project Store
Save/Checkpoint Manager
Central Project Index
Error/Recovery Manager
```

## 5. Canonical Document Model

The parser converts the source `.docx` into a renderer-neutral internal model. Every meaningful element receives a stable element ID.

Example object classes:

- `Document`
- `Section`
- `Paragraph`
- `Run`
- `StyleRef`
- `NumberingRef`
- `Table`
- `TableRow`
- `TableCell`
- `Image`
- `Shape`
- `ChartRef`
- `EmbeddedObjectRef`
- `Footnote`
- `Endnote`
- `Comment`
- `Bookmark`
- `Field`
- `Hyperlink`
- `Header`
- `Footer`
- `RelationshipRef`

Each element records, where applicable:

- source position/order,
- text/content,
- formatting properties,
- style inheritance references,
- section membership,
- relationship IDs,
- dimensions,
- alignment,
- page/layout settings,
- references to binary assets,
- reconstruction strategy,
- QA status.

## 6. Parsing Strategy

The parser must inspect both high-level document content and the underlying DOCX Open Packaging Convention parts.

Relevant package parts include, where present:

- `word/document.xml`
- `word/styles.xml`
- `word/numbering.xml`
- `word/settings.xml`
- `word/footnotes.xml`
- `word/endnotes.xml`
- `word/comments.xml`
- `word/header*.xml`
- `word/footer*.xml`
- `word/_rels/*.rels`
- `word/media/*`
- embedded objects and chart parts
- document properties

The parser must preserve unsupported or difficult binary/embedded objects as transferable package assets when reconstruction from first principles would risk quality loss.

## 7. Rendering Strategy

### 7.1 Word COM Renderer

Used when Microsoft Word desktop is installed and automation is available.

Responsibilities:

- create/open output document,
- reproduce page/section settings,
- insert text and paragraphs,
- apply styles and direct formatting,
- reconstruct tables,
- insert images and supported shapes,
- recreate footnotes/endnotes,
- rebuild headers/footers,
- restore fields/bookmarks/references where supported,
- trigger field updates,
- update TOC where appropriate,
- save real checkpoints,
- export comparison PDF for QA where supported.

### 7.2 Pure DOCX Renderer

Used when Word is unavailable or the user explicitly selects fallback mode.

Responsibilities:

- generate a standards-compliant `.docx`,
- reconstruct supported document structures directly through OPC/XML manipulation,
- reuse preserved binary assets when required,
- surface features that cannot be rendered with sufficient fidelity as warnings.

### 7.3 Hybrid Transfer Rule

Default behavior:

- text, paragraphs, headings, tables, numbering, and standard Word structures are reconstructed;
- complex binary or embedded objects may be transferred from the source package when that produces higher fidelity and does not misrepresent provenance.

## 8. Save and Checkpoint System

All saves must correspond to actual file save operations.

Two save triggers are enabled by default:

1. periodic saves during reconstruction;
2. logical checkpoint saves after important document stages.

Example checkpoints:

- document shell created,
- cover/front matter complete,
- section complete,
- chapter complete,
- tables/figures complete,
- references complete,
- final fields/TOC update,
- final save before QA.

Each save creates an audit event with:

- event ID,
- UTC and local timestamp,
- save sequence number,
- reason,
- current document path,
- document hash after save where practical,
- current reconstruction stage,
- success/failure state.

The software must not perform meaningless save loops merely to inflate revision counts or create a misleading historical appearance.

## 9. Audit Trail

Each project produces at minimum:

```text
project.json
audit.jsonl
save_history.jsonl
warnings.json
qa_report.html
```

Recommended project structure:

```text
<source-folder>/
  original.docx
  original_rebuild/
    source_snapshot/
    working/
    output/
    backups/
    logs/
    qa/
```

A central application index is also stored under an application-controlled directory such as:

```text
%USERPROFILE%\Documents\WordReplica\Projects\
```

The central index stores project metadata and paths, not duplicate source files unless the user opts in.

## 10. Original File Protection

Default behavior:

- source document is opened read-only;
- a source hash is recorded before work starts;
- reconstruction occurs in a separate working copy/output path;
- source hash is rechecked before completion.

An explicit advanced setting may permit working directly on the original, but it must require deliberate user selection and a backup must be created first.

## 11. Error Model

### Critical error — stop

Examples:

- source package is corrupt or unreadable,
- core document XML cannot be parsed,
- output cannot be saved,
- source protection check fails,
- renderer loses a required document section,
- reconstruction would overwrite source unexpectedly.

### Non-critical fidelity issue — continue with warning

Examples:

- unsupported embedded object,
- minor shape positioning mismatch,
- field cannot be fully recreated,
- exact line/page break differs,
- theme behavior differs between rendering engines.

Non-critical issues must be recorded in `warnings.json` and the final QA report.

## 12. QA / Diff Levels

### L0 — Content

Compare textual content and semantic ordering.

Target: exact equality unless the user selected a normalization option.

### L1 — Structure

Compare:

- paragraph count/order,
- headings,
- section hierarchy,
- tables and cells,
- images,
- footnotes/endnotes,
- headers/footers,
- field/bookmark counts where applicable.

Target: 100% for supported structures.

### L2 — Formatting

Compare:

- styles,
- direct formatting,
- fonts,
- font sizes,
- bold/italic/underline,
- spacing,
- indentation,
- alignment,
- borders/shading,
- table formatting.

Target: 100% for supported properties.

### L3 — Layout

Compare:

- margins,
- page size/orientation,
- sections,
- header/footer distances,
- page numbering,
- line/page-break settings,
- column settings.

Target: 100% for supported properties.

### L4 — Render

Where Word rendering is available:

1. render/export source to PDF;
2. render/export reconstructed output to PDF;
3. rasterize pages consistently;
4. compare page count, dimensions, and visual difference metrics;
5. generate page-level visual diff artifacts for mismatches.

Target: as close to pixel-identical as Word allows on the same machine, Word version, fonts, and printer/layout environment.

The report must not claim pixel identity when environmental factors make it unverifiable.

## 13. PASS / WARN / FAIL Policy

### PASS

- source integrity preserved,
- L0 exact,
- L1-L3 within supported exactness targets,
- no critical errors,
- L4 within configured tolerance where render comparison is available.

### WARN

- core content is intact,
- no critical error,
- one or more non-critical fidelity issues remain,
- or L4 differs beyond preferred but acceptable tolerance.

### FAIL

- content loss,
- structural loss of required elements,
- output corruption,
- source integrity failure,
- unsatisfied critical reconstruction requirement.

## 14. GUI

Minimal v1 GUI:

1. Select `.docx`
2. Choose:
   - Visible / Background
   - Clean Replica / Full Fidelity
   - Fresh Metadata / Preserve Legitimate Metadata
3. Start reconstruction
4. Show progress by phase
5. Show current checkpoint/save count
6. Show warnings/errors
7. Open output folder
8. Open QA report

Advanced options remain collapsed by default.

## 15. CLI

Representative command shape:

```powershell
word-replica rebuild "C:\Docs\paper.docx" `
  --renderer auto `
  --visibility visible `
  --fidelity full `
  --metadata fresh
```

Other commands may include:

```powershell
word-replica inspect <docx>
word-replica qa <source.docx> <rebuilt.docx>
word-replica projects list
word-replica project show <id>
```

The GUI and CLI must invoke the same application/service layer rather than duplicating reconstruction logic.

## 16. Platform Behavior

### Windows with desktop Word

- full COM features available;
- Full Fidelity renderer preferred;
- source/output PDF rendering available through Word where permitted.

### Windows without desktop Word

- Pure DOCX renderer only;
- UI must clearly mark unavailable Word-specific features;
- QA report must identify fidelity limitations caused by the missing renderer.

## 17. Technology Direction

Recommended stack for implementation planning:

- Python 3.12+
- `pywin32` for Word COM automation
- `python-docx` only where its abstraction is sufficient
- `lxml` / direct OPC XML handling for unsupported low-level structures
- `zipfile` for DOCX package inspection
- SQLite for central project index
- JSONL for append-only operational logs
- Tkinter/customtkinter or PySide6 for desktop GUI, to be decided in implementation planning
- pytest for automated tests

The implementation plan must verify library choices against the specific Word structures required before coding around them.

## 18. Security and Privacy

- Local-first operation.
- No network upload required for reconstruction.
- Source file is not transmitted to external services by default.
- Logs must avoid copying full document text unless explicitly needed for debugging.
- Hashes may be used to verify source/output integrity.
- Temporary files should be project-scoped and cleaned safely.

## 19. Provenance and Integrity Boundary

The system may:

- reconstruct a document locally,
- create a genuinely new document,
- perform real saves/checkpoints,
- preserve legitimate source metadata when explicitly requested,
- record truthful author/application settings,
- maintain an application audit trail.

The system must not:

- impersonate manual human typing,
- fabricate a human writing session,
- manufacture false save counts,
- forge revision history,
- falsify creation/modification dates,
- inflate editing time to create a fictional work history,
- strip or alter metadata with the purpose of defeating provenance or AI-detection systems,
- claim the output was manually typed when it was reconstructed automatically.

This boundary is a product requirement, not an optional setting.

## 20. Testing Strategy

Testing must use a fixture corpus that grows from simple to pathological DOCX files.

Fixture classes:

1. plain text document;
2. headings and styles;
3. numbered/bulleted lists;
4. tables with merged cells;
5. inline and floating images;
6. multiple sections/orientations;
7. headers/footers/page numbers;
8. footnotes/endnotes;
9. TOC and fields;
10. comments and tracked changes;
11. bookmarks/cross-references;
12. charts and embedded objects;
13. mixed complex academic document.

Every fixture should have parser, renderer, save-log, and QA assertions.

For Word COM tests, Windows/Word integration tests are separated from pure unit tests so the core model can be tested without Word.

## 21. Acceptance Criteria for v1

v1 is complete when all of the following are true:

- A valid `.docx` can be selected in GUI and CLI.
- Source remains unchanged in default mode.
- Canonical model is produced deterministically for supported fixtures.
- Word COM renderer reconstructs all mandatory v1 fixture types on a Windows machine with Word.
- Pure DOCX renderer reconstructs the supported fallback subset without Word.
- Visible and background modes work.
- Clean and Full Fidelity modes are implemented with documented differences.
- Fresh and Preserve Legitimate Metadata modes are implemented without provenance falsification.
- Real checkpoint saves are logged.
- QA levels L0-L3 are produced for every run.
- L4 visual comparison works where Word/PDF rendering is available.
- Every run finishes as PASS, WARN, or FAIL with reasons.
- Critical errors never silently continue.
- A user can open a final HTML QA report and understand what differed.

## 22. Deferred v2+ Ideas

Not part of v1:

- PDF/image/scan ingestion,
- OCR,
- ZIP multi-file projects,
- template-rule systems for universities/faculties,
- batch reconstruction,
- cloud sync,
- Google Docs integration,
- plugin architecture,
- document repair suggestions,
- academic citation validation,
- direct integration into other applications.

## 23. Final Design Decisions

The following decisions are locked for v1:

- Input: `.docx` only.
- Fidelity target: maximum practical / pixel-level target where verifiable.
- Architecture: dual-engine with canonical document model.
- Modes: Visible + Background.
- Fidelity profiles: Clean Replica + Full Fidelity.
- Metadata profiles: Fresh + Preserve Legitimate Source Metadata.
- Unsupported elements: critical-stop / non-critical-best-effort hybrid.
- Interfaces: shared engine with GUI + CLI.
- Platform: works without Word; Full Fidelity enhanced when Word is available.
- Reconstruction strategy: rebuild standard elements, transfer complex binary assets when safer.
- Save strategy: genuine periodic + logical checkpoint saves.
- Storage: project-local artifacts + central index.
- Source protection: read-only by default, explicit override only.
- Integrity boundary: no simulated human authorship or falsified provenance/history.
