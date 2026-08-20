# Document Fidelity Lab — Milestone 1 findings

## Status
Measured results from the first corpus the lab ingested. Every number here was
produced on this machine, Word-free, and is reproducible with the two commands
in "How to reproduce" below.

## What Milestone 1 built

A Word-free pipeline: `fetch -> triage -> fingerprint -> score -> store -> rank`.

| Module | Home |
|---|---|
| Hostile-input triage, single-pass package scan | `src/word_replica/lab/safety.py` |
| Structural fingerprint (~80 measurements per document) | `src/word_replica/lab/fingerprint.py` |
| Complexity score (35 % structural / 25 % rarity / 25 % failure / 15 % layout) | `src/word_replica/lab/scoring.py` |
| Corpus database (SQLite) | `src/word_replica/lab/store.py` |
| Disk governor | `src/word_replica/lab/budget.py` |
| Source adapters (LibreOffice, Apache POI, local tree) | `src/word_replica/lab/sources/` |
| CLIs | `scripts/fidelity_lab/{fetch_cli,scan_cli}.py` |

No Microsoft Word, no .NET, no temporary files, and no dependency that is not
already installed.

## The corpus

| Source | Documents | Size | Licence |
|---|---|---|---|
| LibreOffice `sw/qa/extras/*/data` | 1,701 | 40.4 MB | MPL-2.0 / LGPL |
| Apache POI `test-data/document` | 129 | 7.6 MB | Apache-2.0 |
| **Total** | **1,830** | **48 MB** | redistributable |

Database cost: **1,164 KB for 1,830 documents = 651 bytes per document.**
Extrapolated to the full 736,706-document docx-corpus that is ~480 MB, which
fits the lab's ~2 GB budget with room to spare. Storing the *documents* would
be ~69 GB and does not fit — which is why the design streams and discards.

LibreOffice is read through gitiles rather than the GitHub API: no rate limit
(api.github.com allows 60 requests/hour unauthenticated and was exhausted),
one origin for both listings and blobs, and roughly half the latency.

## Finding 1 — a third of a real corpus is invisible to every existing gate

**602 of 1,830 documents (32.9 %)** carry content Word 2010 silently ignores:

| Namespace | Documents | What it is |
|---|---|---|
| `w15` | 599 | Word 2013 extensions — comment replies, repeating-section content controls |
| `w16cid` | 10 | Word 2016 comment identifiers |
| `strict` | 7 | ISO-29500 Strict — Word 2010 cannot write it |
| `asvg` | 3 | SVG in DrawingML |
| `cx` | 1 | chartex |

This matters because of how the existing gates work. `qa/policy.py` compares
*parsed models*, and `parser/parser.py` only models what WordReplica
reconstructs. Demonstrated rather than asserted, in
`tests/unit/test_lab_fingerprint.py::test_ns_mask_sees_a_w15_feature_that_the_document_model_cannot`:

> Two packages differing only by a `word/commentsExtended.xml` part.
> `build_model_gates` reports **all of G0–G7 passing**, and the part is not
> even retained as a `preserved_part` — the parser does not model it at all.

Word 2010 renders both source and output, so its layout differences cancel and
it rarely produces a false FAIL. The hazard is the opposite: **a false PASS**
for anything it ignores in both files. On this corpus that is a third of the
documents. This is the empirical case for a Word-free G10 preservation gate
over raw package parts, relationships and namespaces.

## Finding 2 — a real document crashed the parser

Apache POI's `Bug66263-table.docx` took `DocxParser` down with

```
AttributeError: '_cython_3_2_4.cython_function_or_method' object has no attribute 'rsplit'
```

Its `w:body` opens with a Word-generated conditional comment. lxml gives
comments and processing instructions a *callable* tag, and `parse_blocks` did
`child.tag.rsplit("}", 1)` on every child. Any `.docx` whose body carried a
comment crashed outright — the one outcome the product contract rules out.

Fixed in `parser/nodes.py::local_name`, applied across `parser.py` (six sites)
and `text.py` (one). Three sites remain in `parser/tables.py`, deferred only
because that file carried unrelated uncommitted work at the time.

## Finding 3 — the lab's own risk classifier was wrong

Of the first 105 documents classified HOSTILE, **57 were flagged for nothing
but their own attached template**: `file:///C:\Program Files\...\Normal.dotm`
in `word/_rels/settings.xml.rels`. Word writes that into essentially every file
it saves, and chart parts routinely link the workbook they came from.

The rule "any external reference to a local path is hostile" would have
quarantined a large share of any real corpus and permanently excluded those
documents from the Word oracle. Narrowed to what is actually dangerous: a
content-pulling reference reaching *another machine* (remote template
injection), or a reference naming an executable wherever it points.

Rescanned: **HOSTILE 105 → 47, VALID 1,653 → 1,710.**

A follow-on bug surfaced immediately: reading the file extension off the whole
target classified every link to a `.com` site as a DOS executable. The suffix
is now read off the URL *path*.

## Risk breakdown after calibration

| Class | Documents | |
|---|---|---|
| VALID | 1,710 | full fidelity is the goal |
| RECOVERABLE | 73 | structurally broken, must be labelled not silently mangled |
| HOSTILE | 47 | macros, ActiveX, remote references, entity declarations, unreadable zips — never handed to Word |

36 documents defeated the pipeline entirely: 19 unreadable zips (encrypted or
clusterfuzz-minimised), 14 `PackageReadError`, 2 CRC failures, 1 `ValueError`.
All produced a labelled row rather than an exception — the extractor's totality
contract held on every one.

## Throughput

Measured on a machine concurrently running a Golden Word reconstruction, so
these are pessimistic:

- ~515 ms per document with the full `DocxParser` layer, ~300 ms without.
- ~59k XML elements/second. This is element-bound, not document-bound: the
  fixture corpus averages 18.7k elements in 36 KB because python-docx's
  template carries a full latent-style table, while the docx-corpus median
  document is 33.6 KB.

## Holdout

89 of 1,830 documents (4.9 %) are sealed in GOLDEN-HOLDOUT and excluded from
ranking. Membership is decided by `sha256(digest + "holdout") % 100 < 5` —
content hash alone, before any score exists — so it can never be influenced by
anything the lab subsequently learns.

## How to reproduce

```powershell
python -m scripts.fidelity_lab.fetch_cli --source libreoffice
python -m scripts.fidelity_lab.fetch_cli --source poi
python -m scripts.fidelity_lab.scan_cli --root C:\WordReplica-Automation\lab\docs\libreoffice `
    --db C:\WordReplica-Automation\lab\lab.db --source libreoffice
```

Corpus bytes and the database live under `C:\WordReplica-Automation\lab\`,
outside git, per `AGENTS.md` § "Data and storage".

## G10 — built, and what it found

`src/word_replica/qa/preservation.py`, 21 tests. It compares raw package bytes:
which namespaces appear anywhere, how many parts of each content type exist,
the relationship graph by type, whether anything dangles, and the bytes of
every opaque part (OLE payloads, custom XML, embeddings, diagrams, embedded
fonts) — content with no model projection at all, where byte identity is the
only available assertion.

It normalizes away what is not a defect: `w:rsid*`, `w15:paraId`, core-property
timestamps and `cp:revision`, ZIP member order, concrete part names, and
WordReplica's own provenance properties. Half the test file asserts the gate
does *not* fire. That is deliberate — a gate that fails on every document for
an intended reason gets switched off, and then the blind spot reopens.

### Lane P: the false-pass rate, measured

`scripts/fidelity_lab/lane_p_cli.py` reconstructs corpus documents through the
pure-docx path and reports G0–G7 against G10. Over 60 documents sampled across
the complexity range:

| | Documents |
|---|---|
| G0–G7 fail (already visible today) | 41 |
| **G0–G7 pass but G10 fails** | **18** |
| G0–G7 pass and G10 passes | 1 |

**Exactly one of sixty real documents round-trips cleanly.** Eighteen are
reported as perfect reconstructions by every gate the project had before today,
and are not.

## Acting on what G10 found

Seven rounds of fix-and-measure over the same 60-document sample. Every number
below comes from `scripts/fidelity_lab/lane_p_cli.py`.

| | at G10 landing | now |
|---|---|---|
| G0–G7 pass **and** G10 pass | 1 | **29** |
| G0–G7 fail (already visible) | 41 | **14** |
| G0–G7 pass but G10 fails | 18 | 17 |

Model-gate failures fell by two thirds and clean round-trips went from one
document to twenty-nine. The middle row moved because several fixes were not
package-level at all: restoring the source's own theme and font table fixed
typography that G2 had been reporting.

### What was fixed

**Template debris.** `PureDocxRenderer` starts from python-docx's default
template, so every output carried its `customXml` store, thumbnail and
`stylesWithEffects` part. `MutableDocxPackage.drop_part` removes a part with
its content-type override and every relationship resolving to it.

**Attachments were never carried.** The parser captured only `word/charts/`,
`word/embeddings/` and `word/diagrams/`, and the renderer never installed
preserved parts at all — it warned and dropped them. So every rebuild shipped
the *template's* theme and font table rather than the source's, and lost the
source's custom XML outright. `PreservedPart` now records how a part is
reached, because that decides what can be done with it.

**Optional parts the source lacks.** `set_numbering`/`set_styles`/`set_settings`
returned early when there was nothing to write, leaving the template's copy.

**Shapes had no representation.** `parse_drawings` only produces a `DrawingRef`
when there is a `blip` — a picture. A group shape, text box, diagram or VML has
none, so it was not modelled and the rebuilt body emitted nothing where it was.
Such fragments are now preserved verbatim and re-emitted.

**Pictures were not in the document at all.** Verified directly: rebuilding a
fixture with an image produced a package still holding `word/media/image1.png`
and a `word/document.xml` with no `w:drawing`. Every rebuilt document lost every
picture from its visible content while the bytes sat in the package as litter.
The `w:drawing` is now preserved verbatim with its relationship ids remapped —
which keeps the authored anchor geometry, wrapping, crop and rotation exactly,
where hand-built DrawingML would have kept only what `DrawingRef` models.

**Two differences the product makes on purpose.** Custom document properties
are dropped unless allow-listed (README: "custom properties require an explicit
allowlist"), and `docProps/app.xml` holds statistics about the document that now
exists — page counts, editing time, application and version — which cannot match
the source by construction. Both are declared parameters rather than hidden
rules, so the gate reports them only when the caller has not said it did them
deliberately.

### The measurement that made shape preservation safe

Re-emitting a fragment verbatim is only safe if it carries no relationship id:
ids are renumbered by any writer, so a stale one points at the wrong part or
none at all, and Word repairs such a document on open. **Every
`mc:AlternateContent` block sampled from the corpus carried none**, which is
why verbatim re-emission was viable. Where an id *is* present and can be
resolved — the parser knows which part it addressed and the renderer owns the
new package — it is rewritten. Where it cannot be resolved, the fragment is
refused: a missing shape is a visible loss, a corrupted package is not.

### Five self-inflicted bugs, all caught by measuring

Worth recording, because none was caught by reading the code:

1. The first attachment rule was "anything with a document-level relationship",
   which also matched `styles.xml`, `settings.xml` and `numbering.xml` — parts
   the renderer builds itself, silently overwritten by restoring them verbatim.
2. Widening `preserved_parts` made `preflight.py` mark every new attachment
   UNSUPPORTED, which would have blocked interactive runs that have always
   proceeded — Golden #1's included. The full suite caught it as 8 failures.
3. Trimming the finished projection for the metadata carve-out stripped
   `docPropsVTypes`, which `docProps/app.xml` also uses. Lane P caught it: clean
   documents fell from 6 to 0. Fixed by re-projecting with the part ignored
   rather than subtracting from the result.
4. A new inline-capture branch was unreachable, because `w:drawing` already had
   its own branch emitting the token the interactive executor dispatches on.
   The token was extended instead of replaced.
5. `word/stylesWithEffects.xml` was put on the template-debris list, but real
   documents legitimately have one — dropping it unconditionally was the same
   defect mirrored.

Three times the fixture, not the code, turned out to be what was wrong: an
over-minimal package with relationships pointing nowhere, missing `_rels` for a
customXml item, and no package relationship to `docProps/core.xml`. Each was
corrected rather than the check loosened.

### What remains

Seventeen of sixty, each on something specific rather than one systemic cause:

| First divergence | Documents |
|---|---|
| namespaces | 5 |
| `attachedTemplate` relationship | 2 |
| theme part | 2 |
| `webSettings` part | 2 |
| `mc:Ignorable` | 2 |
| comments part | 1 |
| OLE object bytes | 1 |
| core-properties part | 1 |

Charts, diagrams and embedded objects are still refused rather than restored:
they are reached from inside the document body, so writing the part without the
reference the rebuilt body no longer carries would produce an orphan, and G10
would then match on a document that is actually broken. A false signal is worse
than a reported loss.

## What is still not done

No Word lane (G8/G9 over the corpus), no minimizer, no generators, no tier
scheduler, and no docx-corpus ingest. Three occurrences of the non-element-node
crash remain in `parser/tables.py`, deferred while that file carried
uncommitted work from the autonomous harness.

Golden #1 remains on ten gates. `G10_GATE_NAMES` exists for callers that opt in
to eleven via `build_golden_report(gate_names=...)`.
