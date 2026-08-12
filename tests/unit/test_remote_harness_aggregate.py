import json
from zipfile import ZipFile

from scripts.remote_harness.aggregate import aggregate_results
from scripts.remote_harness.package_results import package_results


def test_aggregate_groups_same_rpc_failure_once(tmp_path):
    root=tmp_path/"results"; docs=root/"documents"
    for i in (1,2):
        stage=docs/f"00{i}_doc{i}"/"interactive_maximum"; stage.mkdir(parents=True)
        (stage/"result.json").write_text(json.dumps({
            "document":f"doc{i}.docx","stage":"interactive_maximum","status":"COM_FAIL",
            "run_status":"FAIL","reasons":["Call was rejected by callee."]
        }),encoding="utf-8")
    summary=aggregate_results(root)
    assert summary["failure_groups"]["RPC_E_CALL_REJECTED"]["runs"] == 2
    assert summary["failure_groups"]["RPC_E_CALL_REJECTED"]["documents"] == 2
    assert (root/"summary.json").exists()
    assert (root/"summary.csv").exists()


def test_logs_only_zip_excludes_docx_and_pdf(tmp_path):
    root=tmp_path/"results"; root.mkdir(); (root/"summary.json").write_text("{}")
    (root/"output.docx").write_bytes(b"x"); (root/"page.pdf").write_bytes(b"x"); (root/"event_trace.jsonl").write_text("{}\n")
    out=package_results(root,tmp_path/"out.zip",logs_only=True)
    with ZipFile(out) as zf:
        names=zf.namelist()
        assert not any(name.lower().endswith((".docx",".pdf")) for name in names)
        assert any(name.endswith("summary.json") for name in names)
