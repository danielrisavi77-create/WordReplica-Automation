# WordReplica Codex Windows setup

The one-time installer creates `C:\WordReplica-Automation`, clones the private `WordReplica-Automation` repository into `repo\`, copies the Golden source into `golden\`, creates `.venv`, installs WordReplica in editable mode with test dependencies, checks out `automation-dev`, and runs the local non-Word regression suite.

After installation, open `C:\WordReplica-Automation\repo` in the Codex desktop app. The repository-level `AGENTS.md` is the persistent operating contract. Start the task with:

> Continue WordReplica Golden #1 autonomously according to AGENTS.md. Run the existing Golden pipeline, analyze the first divergence, use strict RED/GREEN TDD for each production fix, rerun the full regression suite and real Word Golden after every green fix, and continue until promotion_ready=true or a defined fail-safe requires stopping. Do not ask me to move ZIPs or manually install updates.

Codex should invoke `RUN_GOLDEN_CODEX.ps1` for the real Microsoft Word test. Heavy diagnostics remain in `C:\WordReplica-Automation\diagnostics` and are not uploaded to GitHub during normal development.
