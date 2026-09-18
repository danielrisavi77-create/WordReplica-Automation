# Word Replica landing site

This directory is the public documentation/download surface for Word Replica.

## What the site is for

- explain local DOCX reconstruction and fidelity verification;
- document Windows/Word compatibility;
- explain Academic Suite integrations;
- link to verified releases and their hashes/manifests;
- expose privacy and integrity guarantees.

## What the site is not for

The landing site is not a DOCX processing backend. Do not add a generic manuscript upload flow merely to make Word Replica "web based". The core product advantage is that Word/COM execution and document reconstruction remain local on the user's Windows machine.

If another Academic Suite product uploads a document for its own clearly documented purpose, that remains that product's boundary and is not a reason to move Word Replica execution into the browser.

## Release metadata

Do not hard-code a public version that does not have a matching artifact.

When a production release exists, the download surface should display metadata derived from the release artifact/manifest:

- version;
- file name;
- exact size;
- SHA-256;
- source commit;
- code-signing publisher/status;
- compatible Repair Contract version/key identity when applicable.

The local desktop builder emits `dist/word-replica-build-manifest.json` as build evidence. A public release additionally requires the signing and release gates in `docs/RELEASE.md`.

## Deployment

Deployment is intentionally separate from application execution. This static directory can be hosted by a simple static host after review, without adding backend document-processing infrastructure.
