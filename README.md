# Word Replica

Local `.docx` reconstruction and QA for Windows. Word Replica rebuilds a document into a new DOCX through a canonical document model, using Microsoft Word COM when available and a Pure DOCX fallback otherwise.

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[test]"
```

For a packaged Windows application, also install the build extra:

```powershell
python -m pip install -e ".[test,build]"
```

## GUI

```powershell
word-replica-gui
```

The GUI and CLI call the same `RebuildService`. Microsoft Word desktop is optional for the application as a whole, but the Word COM renderer and its Full Fidelity capabilities require Windows with desktop Word installed.

## CLI

```powershell
word-replica rebuild "C:\Docs\paper.docx" --renderer auto --visibility visible --fidelity full --metadata fresh
```

Other commands:

```powershell
word-replica inspect "C:\Docs\paper.docx"
word-replica qa "C:\Docs\paper.docx" "C:\Docs\paper_reconstructed.docx"
word-replica projects list
word-replica project show PROJECT_ID
```

## Modes

**Visible / Background** controls whether the Microsoft Word desktop process is shown during a Word COM reconstruction. It does not change the reconstruction semantics. Pure DOCX does not need Word and therefore has no visible Word session.

**Clean Replica / Full Fidelity** controls which source structures are projected into the reconstruction. Clean Replica removes review-only structures such as comments, deleted revision content and hidden text from the clean projection. Full Fidelity preserves the canonical evidence for comments, revisions, bookmarks, fields and supported structured parts; renderer limitations are reported instead of hidden.

**Fresh / Preserve Legitimate Metadata** controls descriptive metadata only. Fresh starts from the new output document's own metadata. Preserve Legitimate may copy only explicitly allowed descriptive properties. Author/company fields require explicit opt-in, and custom properties require an explicit allowlist.

## Source safety

The selected source is read-only by default. Word Replica records a pre-run SHA-256 and verifies that the source bytes are unchanged after reconstruction. The advanced source-overwrite option is opt-in and creates a hash-verified backup before any replacement.

## Output and QA

Each run creates a project folder containing:

- `project.json` — source identity and selected options.
- `logs/audit.jsonl` — append-only phase/event audit.
- `logs/save_history.jsonl` — one row per successful real file save.
- `logs/warnings.json` — renderer/QA limitations.
- `qa/qa_report.html` — final QA status and findings.
- `output/*_reconstructed.docx` — reconstructed output in default read-only-source mode.

The custom document property `WordReplicaActualSaveCount` is set by the final metadata seal and must equal the number of rows in `save_history.jsonl`. It does **not** repurpose Word's built-in revision count or editing-time fields.

QA levels are:

- **L0 Content** — visible text and note text.
- **L1 Structure** — body/table shape, assets and supported document structures.
- **L2 Formatting** — paragraph/run formatting model.
- **L3 Layout** — section/page layout.
- **L4 Render** — PDF/page visual comparison only when a controlled render path is actually available. If not available, the report says so; it never invents a visual PASS.

`PASS` means all required checks supported in the run matched. `WARN` means a non-critical fidelity limitation or supported formatting/layout difference remains. `FAIL` means critical content/required-structure loss or another critical rebuild error occurred.

## Tests

Non-Word gate:

```powershell
python -m pytest tests/unit tests/integration/test_pure_docx_e2e.py tests/acceptance -m "not word" -v
```

Windows + Word gate:

```powershell
$env:WORD_REPLICA_WORD_TESTS="1"
python -m pytest tests/integration/word tests/acceptance -m word -v
```

The Word gate must be run on the target Windows machine with Microsoft Word desktop before a release is declared complete.

## Integrity boundary

Word Replica performs automated document reconstruction. It does not simulate manual typing and does not fabricate timestamps, editing time, revision counts, save history, or provenance. It does not add deliberate typos, corrections, random delays or keystroke timing to make automated work appear human-authored. Preserve Metadata copies only explicitly allowed legitimate descriptive fields.

## Known limitations

Pixel layout can vary with Microsoft Word version, installed fonts, printer/layout environment and unsupported embedded objects. Image position, complex charts/OLE objects, comments, tracked changes, fields and cross-references may require Word Full Fidelity support and may still produce explicit warnings when exact reconstruction is not technically verifiable. The QA report must state these limitations instead of claiming unverifiable pixel identity.

The Pure DOCX renderer is the fallback path. Its acceptance gate covers plain text, headings/styles, lists, merged tables, sections/orientations, headers/footers/page fields and footnotes/endnotes. More complex objects are retained as canonical evidence and are either reconstructed by the Word renderer or reported as fidelity limitations.

## Windows release gate

For the release-candidate build, run the complete Windows + Microsoft Word gate with one command:

```powershell
powershell -ExecutionPolicy Bypass -File .\RUN_WINDOWS_RELEASE_GATE.ps1
```

The Word renderer uses a hybrid path: deterministic canonical DOCX reconstruction followed by a real Microsoft Word read-only compatibility open. This avoids fragile COM collection/range mutation during reconstruction while keeping Word as the compatibility gate.


## Windows desktop EXE

The lightweight desktop application uses Windows/Python Tkinter rather than PySide6, so the packaged app does not require the large Qt dependency. Build the verified single-file application with:

```powershell
powershell -ExecutionPolicy Bypass -File .\BUILD_WINDOWS_APP.ps1
```

The builder runs the full Windows + Microsoft Word test gate before producing the EXE and stops if any test fails. On success the application is written to:

```text
dist\WordReplica.exe
```

See `WINDOWS_APP.md` for the desktop workflow and build details.

## Persistent Remote Harness

The Remote Word Harness also supports a persistent installation layout. See `docs/PERSISTENT_REMOTE_HARNESS.md`. In this mode the user's DOCX corpus and historical results live outside the replaceable `current` code tree, and future harness versions are applied with hash-verified rollback-safe update ZIPs.
