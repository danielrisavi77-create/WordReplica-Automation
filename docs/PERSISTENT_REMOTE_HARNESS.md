# Word Replica Persistent Remote Harness

## One-time setup

1. Extract the persistent harness ZIP once to a permanent folder, for example `C:\Users\PC\Downloads\WordReplica-Remote-Harness`.
2. Put your real-world `.docx` corpus in `realworld_input` once, or use `IMPORT_EXISTING_INPUT.cmd` to copy it from an older harness folder.
3. Run `START_HERE.cmd` for every corpus run.

The persistent folders `realworld_input`, `results`, `.venv_harness`, and top-level `harness_config.json` survive code updates.

## Future updates

1. Download the small update ZIP supplied for the next harness version.
2. Put exactly that ZIP in `updates`.
3. Run `UPDATE_HARNESS.cmd`.
4. The updater validates archive paths and SHA-256 hashes, stages the code, backs up the old `current`, swaps only after preflight validation, and rolls back automatically if post-install validation fails.
5. A successfully applied update ZIP is moved to `updates\applied` so the next update can use the same folder.

Updates never modify `realworld_input` or old `results` directories.

## Results

Each run creates a new versioned directory under `results`, such as:

`results\2026-08-11_190500_v2.0.0-persistent.1`

The upload ZIP for ChatGPT is stored inside that run directory and includes the harness version in its filename.

## Rollback

Run `ROLLBACK_HARNESS.cmd` to restore the most recent valid code backup. Rollback changes only `current` and `version.json`; input documents and result history remain untouched.
