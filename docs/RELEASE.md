# Release Policy

## Purpose

A green unit-test suite is necessary but not sufficient to call Word Replica released. A public Windows artifact must be traceable to reviewed source and must pass the real Microsoft Word gate when its behavior depends on Word.

## Promotion path

```text
feature/hardening branch
        ↓ PR + non-Word CI
automation-dev
        ↓ real Windows + Word gate
release candidate
        ↓ signed artifact + manifest verification
main / tagged release
```

Do not fast-forward or force-update `main` merely because `automation-dev` is ahead. Promotion is an explicit release operation.

## Required evidence

### Every PR into automation-dev

- Linux non-Word CI green.
- Windows non-Word CI green.
- Compile gate green.
- Repository hygiene green.
- Focused regression tests for changed behavior.

### Renderer / fidelity / repair runner changes

In addition to the PR gate:

- run `RUN_WINDOWS_RELEASE_GATE.ps1` on the approved Windows machine with Microsoft Word desktop;
- run the relevant Golden or real-document acceptance scenario;
- preserve unrelated Word sessions;
- record any fidelity limitation rather than converting it to a PASS.

### Public executable

A production artifact must have:

- product/engine version;
- exact artifact SHA-256;
- exact artifact size;
- source commit;
- clean source-tree attestation for release builds;
- valid code-signing identity;
- timestamped signature;
- contract key identity where the artifact executes Lekta Repair Contracts.

The desktop builder currently writes `dist/word-replica-build-manifest.json`. That manifest proves build identity fields; it does not by itself prove Authenticode validity.

For a production desktop release, run `BUILD_WINDOWS_RELEASE.ps1` from an exact `release/**` branch. It always invokes the verified desktop builder with the real Microsoft Word gate enabled, validates the unsigned build manifest against the exact source commit and bytes, requires trusted code signing, re-verifies the signer, and writes `dist/word-replica-release-manifest.json`.

Example with a local trusted code-signing certificate:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\BUILD_WINDOWS_RELEASE.ps1 \
  -SigningMode CertificateStore \
  -SigningCertificateThumbprint "<trusted-thumbprint>" \
  -TimestampServer "<trusted-rfc3161-or-authenticode-timestamp-url>"
```

Artifact Signing is supported with `-SigningMode ArtifactSigning`, but additionally requires the exact SignTool, `Azure.CodeSigning.Dlib.dll`, metadata file, timestamp server and independently trusted publisher thumbprint. Missing signing material is a release blocker, never a reason to produce an unsigned production artifact.

The Lekta production runner additionally uses `lekta-repair-runner-manifest.json` and the stricter signing checks in `BUILD_LEKTA_REPAIR_RUNNER.ps1`.

## Version rules

`pyproject.toml [project].version` and `word_replica.__version__` must match. CI enforces this.

Do not publish a version string on the landing page before a matching release artifact exists.

A version bump should occur in the same reviewed release change that defines the release candidate.

## Landing/download publication

The landing page may be deployed independently as documentation, but a download CTA should only be treated as an actual release when the linked artifact has:

1. a release manifest;
2. a SHA-256;
3. source-commit traceability;
4. required real-Word evidence;
5. valid code signing for the production channel.

The web site is not a DOCX processing backend. Word Replica document execution remains local.

## Main branch repair

At the time this policy was introduced, active development lived on `automation-dev` and `main` lagged/diverged.

Do not resolve that divergence by an unreviewed merge. First make `automation-dev` the reviewed release candidate, then reconcile the small set of `main`-only commits explicitly, rerun all gates, and promote the resulting commit deliberately.
