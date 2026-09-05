from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parents[2]


def test_local_golden_and_heavy_diagnostics_are_not_inside_repo():
    forbidden_roots = {"golden", "work", "diagnostics", "archive", "state"}
    assert not any((ROOT / name).exists() for name in forbidden_roots)


def test_gitignore_allows_versioned_docx_test_fixtures_only():
    text = (ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "*.docx" in text
    assert "!tests/fixtures/**/*.docx" in text


def test_gitignore_excludes_local_build_and_generated_golden_artifacts():
    local_only_paths = (
        ".venv_appbuild/Scripts/python.exe",
        "_g0_dump.txt",
        "_golden2_run.log",
        "_p46_events.pkl",
        "_toc_repro_rebuild/example/logs/audit.jsonl",
        "tests/fixtures/corpus/07_headers_footers_numbers_rebuild/example/output/rebuilt.docx",
        "tests/unit/test_interactive_word_bookmark_names.py.orig",
    )

    for relative_path in local_only_paths:
        result = subprocess.run(
            ["git", "check-ignore", "--no-index", "-q", relative_path],
            cwd=ROOT,
            check=False,
        )
        assert result.returncode == 0, f"local artifact is not ignored: {relative_path}"
