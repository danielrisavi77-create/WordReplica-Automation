from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def test_local_golden_and_heavy_diagnostics_are_not_inside_repo():
    forbidden_roots = {"golden", "work", "diagnostics", "archive", "state"}
    assert not any((ROOT / name).exists() for name in forbidden_roots)


def test_gitignore_allows_versioned_docx_test_fixtures_only():
    text = (ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "*.docx" in text
    assert "!tests/fixtures/**/*.docx" in text
