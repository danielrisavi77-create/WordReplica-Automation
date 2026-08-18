import json
from pathlib import Path

from scripts.codex_automation.audit_cli import run_audit_to_file


def test_run_audit_to_file_writes_exact_golden_report_name(tmp_path):
    report_path = tmp_path / "golden_report.json"
    payload = {"full_pass": False, "gates": {"G0": False}}

    written = run_audit_to_file(
        source=tmp_path / "source.docx",
        output=tmp_path / "output.docx",
        qa_dir=tmp_path / "qa",
        report_path=report_path,
        run_id="r1",
        source_sha256="abc",
        commit_sha="def",
        reconstruction_status="WARN",
        audit_fn=lambda *args, **kwargs: payload,
    )

    assert written == report_path
    assert json.loads(report_path.read_text(encoding="utf-8")) == payload
