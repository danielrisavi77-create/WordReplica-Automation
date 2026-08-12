import json
from pathlib import Path
import sys

from scripts.remote_harness.config import HarnessConfig
from scripts.remote_harness.runner import RemoteHarnessRunner


def test_runner_continues_and_runs_standard_only_after_expected_block(tmp_path):
    input_dir=tmp_path/"input"; input_dir.mkdir(); (input_dir/"demo.docx").write_bytes(b"fake")
    child=tmp_path/"fake_child.py"
    child.write_text(r'''
import argparse,json
from pathlib import Path
p=argparse.ArgumentParser(); p.add_argument('--stage'); p.add_argument('--source'); p.add_argument('--run-dir'); p.add_argument('--logs-only',action='store_true'); a=p.parse_args()
r=Path(a.run_dir); r.mkdir(parents=True,exist_ok=True)
status='EXPECTED_BLOCK' if a.stage=='interactive_maximum' else 'PASS'
(r/'result.json').write_text(json.dumps({'document':Path(a.source).name,'stage':a.stage,'status':status,'run_status':'FAIL' if status=='EXPECTED_BLOCK' else 'PASS','reasons':[]}))
''', encoding='utf-8')
    result_root=tmp_path/"results"/"run1"
    runner=RemoteHarnessRunner(source_root=tmp_path,input_dir=input_dir,result_root=result_root,config=HarnessConfig(),python_executable=sys.executable,child_runner_path=child)
    runner.run()
    doc_dir=next((result_root/"documents").iterdir())
    assert (doc_dir/"static"/"result.json").exists()
    assert (doc_dir/"instant"/"result.json").exists()
    assert (doc_dir/"interactive_maximum"/"result.json").exists()
    assert (doc_dir/"interactive_standard"/"result.json").exists()
    summary=json.loads((doc_dir/"document_summary.json").read_text())
    assert summary["source_sha256_before"] == summary["source_sha256_after"]
