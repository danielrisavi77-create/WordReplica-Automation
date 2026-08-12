# Word Replica v1 — Windows RC2

RC2 fixes the final Windows release-gate failures observed in image-bearing fixtures.

## Root cause fixed
The deterministic DOCX renderer retained media parts such as `word/media/image1.png` but did not register an OPC content type for those new parts in `[Content_Types].xml`. Desktop Microsoft Word rejected those packages as corrupted even though ZIP and python-docx checks passed.

RC2 registers each installed binary asset with its canonical content type using an explicit OPC Override. This covers PNG and any future asset type carried by the canonical `BinaryAsset.content_type` field.

## Release-gate usage
Run only:

```powershell
powershell -ExecutionPolicy Bypass -File .\RUN_WINDOWS_RELEASE_GATE.ps1
```

The gate creates an isolated lean `.venv_gate`, installs core + test dependencies only, enables Microsoft Word tests, compiles source/tests, and runs the complete pytest suite.
