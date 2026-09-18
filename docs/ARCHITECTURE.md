# Word Replica Architecture and Integration Boundaries

## 1. Product responsibility

Word Replica is a local Windows document execution and fidelity engine.

It accepts an explicit source or a cryptographically verified target, reconstructs a new DOCX through the canonical document model, and emits evidence about what was preserved. It does not decide academic rules, write academic arguments, or own project-completion policy.

The public product promise is deliberately narrow:

> Execute a known document transformation locally and report fidelity truthfully.

## 2. Internal layers

```text
Entry points
  ├─ CLI
  ├─ Desktop GUI
  ├─ Lekta one-shot runner
  └─ Engineering harnesses
          │
          ▼
Application services
  ├─ RebuildService
  ├─ InteractiveRebuildService
  └─ RepairPackageService
          │
          ▼
Canonical document pipeline
  DOCX → parser → canonical model → reconstruction blueprint
          │
          ▼
Renderers
  ├─ Pure DOCX
  ├─ Microsoft Word compatibility path
  └─ Interactive Microsoft Word
          │
          ▼
Verification
  ├─ L0-L4 user-facing QA
  └─ G0-G10 release/integration gates
```

### Core engine

`src/word_replica/domain`, `parser`, `opc`, `renderers`, `services` and `qa` form the reusable engine.

User interfaces and external integrations may depend on the engine. The engine must not depend on Lekta, Katedra, Academic Completion, payment systems or web application frameworks.

### Lekta runner

`src/word_replica/repair_contract` and `src/word_replica/runner` implement a narrow external execution boundary.

The runner may understand the versioned Repair Contract format, trusted public keys, claim/status transport and one-shot lifecycle. It must not contain Lekta rule semantics or decide which repairs are academically correct.

### Fidelity Lab and automation

`src/word_replica/lab`, `scripts/codex_automation`, Golden corpora and remote harnesses are engineering infrastructure.

They may test production code, but production behavior must not require a development harness or Golden dataset.

### Web and landing

The landing site is a documentation/distribution surface. It may explain the product, expose release metadata, compatibility and integrations, but it is not the document-processing runtime.

Raw user DOCX processing remains local unless a separate product explicitly owns a different, documented upload flow.

## 3. Academic Suite integration model

```text
Katedra
 writing / process support
       │
       ▼
Lekta
 deterministic document analysis
       │
       │ signed Repair Contract + verified target
       ▼
Word Replica
 local execution + fidelity proof
       │
       │ structured completion status
       ▼
Academic Completion
 project state / next action
```

### Lekta → Word Replica

This is the primary integration.

Lekta owns:

- academic profiles and rules;
- repair selection;
- entitlement/commerce decisions;
- generation of the corrected target;
- signing of the Repair Contract.

Word Replica owns:

- signature and package validation;
- source/target identity checks;
- local reconstruction;
- Microsoft Word ownership safety;
- fidelity gates;
- original-source recheck;
- completion evidence.

Word Replica must fail before opening Word when the contract, key, expiry, engine compatibility or file identity is invalid.

### Katedra → Word Replica

Katedra must not directly import Word Replica renderer or COM modules.

The preferred path is:

```text
Katedra manuscript
  → Lekta verification/repair
  → signed Word Replica execution
  → verified final DOCX
```

A future direct Katedra handoff is allowed only through a small versioned artifact contract with the same principles: explicit file identity, no arbitrary commands, bounded capability and structured result.

### Academic Completion → Word Replica

Academic Completion may consume structured state such as:

- Word Replica run ID;
- engine version;
- terminal status;
- required gate results;
- output hash;
- source-unchanged result;
- completion timestamp.

It should not receive the raw DOCX or document text merely to represent completion state.

A useful state distinction is:

```text
LEKTA_VERIFIED
  ≠ WORD_REPLICA_VERIFIED
```

Lekta verification means the document satisfies its verified academic checks. Word Replica verification means the delivered artifact matches the approved target under the configured technical/fidelity policy.

### katedra-pkg / rad-docx

`rad-docx` can generate the manuscript DOCX. Word Replica can then act as a final openability/fidelity gate.

The packages should communicate by file + report/contract boundaries rather than importing each other's internal implementation.

## 4. Safety boundaries

Word Replica must preserve these invariants:

1. The original source is read-only by default.
2. Source identity is checked before and after a run.
3. No process-wide `taskkill /IM WINWORD.EXE` behavior.
4. Only owned Word processes with matching process identity may be terminated.
5. An unsupported or unverifiable fidelity check is WARN/FAIL/unevaluable, never an invented PASS.
6. Automated execution must not fabricate editing time, typing history, revision provenance or human authorship.
7. External contracts are allowlisted, versioned and fail closed.
8. Private keys never ship in the runner.

## 5. Branch and release model

### Branches

- `automation-dev` — active integration and release-candidate branch.
- `main` — promoted stable history only after explicit gates.
- short-lived feature/hardening branches — normal change surface via pull request.

Do not infer release readiness merely from branch name or a green non-Word test.

### Required gates

Every pull request:

- repository-hygiene tests;
- compile gate;
- full non-Word pytest suite on Windows and Linux.

Renderer/fidelity/runner release changes additionally require:

- target Windows environment;
- Microsoft Word desktop;
- `RUN_WINDOWS_RELEASE_GATE.ps1`;
- relevant Golden/real-document evidence when the change affects fidelity;
- clean source tree;
- source commit recorded in the release manifest.

Production executable release additionally requires:

- valid Authenticode/public code-signing result;
- timestamp;
- artifact SHA-256;
- publisher identity;
- Repair Contract key identity;
- engine version;
- source commit and clean-tree attestation.

## 6. Version authority

`pyproject.toml [project].version` and `word_replica.__version__` are required to match and are guarded by CI.

Release scripts must derive their expected version from those checked values rather than introducing another hard-coded product version.

Public landing pages must not claim a release version before an actual release artifact exists.

## 7. Future refactor boundary

Large implementation modules should be split by capability without changing the public service boundary.

Recommended split for `interactive_word.py`:

- session/ownership;
- text and run insertion;
- tables;
- fields/bookmarks;
- images/drawings;
- headers/footers/sections;
- save/finalization.

Recommended split for `interactive_rebuild.py`:

- preflight;
- orchestration;
- checkpoint/resume;
- verification;
- completion/reporting.

The refactor should be behavior-preserving and Golden-driven, not bundled with new product features.
