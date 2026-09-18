# WordReplica Full Release Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the reviewed `automation-dev` state into a release candidate with real Microsoft Word evidence, a production-ready Lekta integration path, signed-artifact gates, controlled main promotion, and a deployable public landing site.

**Architecture:** Keep WordReplica local-first. Release evidence is produced by a self-hosted Windows + Microsoft Word runner; Lekta remains the authority for repair semantics and signed Repair Contracts; WordReplica remains the local execution/fidelity layer. Promotion and publication are automated only when all trust and signing inputs are present and verified.

**Tech Stack:** Python 3.12, pytest, PowerShell, GitHub Actions, PyInstaller, Microsoft Word COM, Authenticode / Artifact Signing, TypeScript/Vite/Vitest for Lekta, Supabase, static GitHub Pages-compatible landing.

**Spec:** `docs/ARCHITECTURE.md`, `docs/RELEASE.md`, Lekta `docs/LOCAL_REPAIR_RELEASE.md`

## Global Constraints

- Never fake a real-Word PASS.
- Never weaken Authenticode or Artifact Signing requirements for production.
- Never promote `main` unless release evidence is tied to the exact source SHA.
- Never put signing private keys, Supabase secrets, claim tokens, raw private DOCX files, Golden outputs, or diagnostics into Git.
- WordReplica remains local-first; the website is documentation/download only.
- Lekta remains the academic rule and repair-semantics authority.
- Cross-repo communication uses versioned contracts/status, not direct internal imports.
- Existing production data or Supabase state is mutated only by the established fail-closed release flow.

---

### Task 1: Release branch orchestration and real-Word evidence

**Files:**
- Modify: `.github/workflows/windows-word-release.yml`
- Create: `docs/releases/RELEASE_CANDIDATE.md`
- Test: workflow syntax + existing CI

**Interfaces:**
- Consumes: exact release branch SHA.
- Produces: self-hosted Windows release gate run and downloadable evidence artifact.

- [ ] Add `push` trigger for `release/**` while retaining manual dispatch.
- [ ] Record source SHA, Python version, Word availability and release-gate result.
- [ ] Upload non-sensitive release evidence even when the gate fails.
- [ ] Push the release branch and inspect whether a compatible self-hosted runner actually picks up the job.
- [ ] If no compatible runner exists, record the blocker instead of claiming PASS.

### Task 2: Production artifact gate

**Files:**
- Create: `BUILD_WINDOWS_RELEASE.ps1`
- Create: `tests/unit/test_windows_release_script.py`
- Modify: `docs/RELEASE.md`

**Interfaces:**
- Consumes: `dist/WordReplica.exe`, trusted signing configuration, exact source SHA.
- Produces: signed `WordReplica.exe` + release manifest with signature identity, timestamp, SHA-256, size, version and source commit.

- [ ] Write tests requiring fail-closed signing inputs and manifest fields.
- [ ] Implement CertificateStore and ArtifactSigning modes using the same trust principles as the Lekta runner.
- [ ] Re-verify signature after signing.
- [ ] Refuse dirty source trees and wrong source refs.
- [ ] Refuse public release when no trusted signing identity exists.
- [ ] Run focused tests and full CI.

### Task 3: Lekta current integration reconciliation

**Repositories:** WordReplica + Lekta.

**Files:**
- Lekta: existing `feature/repair-contract-v1-current-v2` branch and release scripts.
- WordReplica: release docs only unless a compatibility defect is proven.

**Interfaces:**
- Consumes: Repair Contract v1 fixture/policy, runner manifest v2, local claim/status protocol.
- Produces: reviewed PR path from current Lekta master to the current Repair Contract implementation, with CI evidence.

- [ ] Open a fresh Lekta PR from `feature/repair-contract-v1-current-v2` to `master`.
- [ ] Run/inspect Lekta CI and resolve merge conflicts or current-master incompatibilities without bypassing tests.
- [ ] Verify cross-language Repair Contract/status tests.
- [ ] Verify the production release preflight still rejects unsigned/development artifacts.
- [ ] Do not deploy Supabase/Netlify until production signing identity and all required release inputs exist.

### Task 4: Landing deployment

**Files:**
- Create: `.github/workflows/pages.yml`
- Modify: `landing/index.html`
- Create: `landing/release.json` only when a real release exists.

**Interfaces:**
- Consumes: static `landing/`.
- Produces: GitHub Pages-compatible static site; download metadata remains unpublished until a real signed release exists.

- [ ] Add static Pages build/deploy workflow for `main` and manual preview.
- [ ] Keep download CTA non-release until a real artifact/manifest exists.
- [ ] Verify HTML is deployable without backend dependencies.
- [ ] Attempt deployment only through supported GitHub Pages configuration; record a Pages-settings blocker if repository administration is required.

### Task 5: Main promotion

**Files:** Git refs only; no new product behavior.

**Interfaces:**
- Consumes: exact release SHA + green CI + real Word PASS + signed artifact evidence + compatible Lekta state.
- Produces: deliberate `main` promotion.

- [ ] Reconcile the two `main`-only commits explicitly.
- [ ] Re-run CI on the exact candidate.
- [ ] Verify real Word release evidence belongs to the exact candidate SHA.
- [ ] Verify signed artifact manifest belongs to the exact candidate SHA.
- [ ] Verify Lekta compatibility/release preflight.
- [ ] Promote `main` only if every previous item is green; otherwise leave `main` unchanged and record the precise blocker.

## Completion Gate

The release is complete only if:
- cross-platform CI is green on the exact candidate;
- the self-hosted Windows + Microsoft Word release gate is green on the exact candidate;
- a trusted code-signed artifact exists and its manifest verifies against the exact candidate;
- Lekta compatibility is proven against the current master path;
- the landing site can publish verified release metadata;
- `main` points to the reviewed release commit.

Anything less is a release candidate with an explicit blocker, not a production release.
