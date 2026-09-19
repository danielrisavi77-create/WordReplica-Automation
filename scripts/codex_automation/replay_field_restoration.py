"""Test current field restoration on a temporary copy of retained Golden output.

Does not launch Word, edit source/retained output, write Golden state, or authorize
resuming a stopped loop. Prints only counts and gate booleans, not document text.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile

from scripts.codex_automation.config import load_config
from word_replica.parser.parser import DocxParser
from word_replica.qa.golden_audit import build_model_gates
from word_replica.services.interactive_rebuild import InteractiveRebuildService


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def replay_pair(source: Path, output: Path) -> dict:
    original_hashes = (digest(source), digest(output))
    parser = DocxParser()
    expected = parser.parse(source)
    before = parser.parse(output)
    with tempfile.TemporaryDirectory(prefix='wordreplica-field-replay-') as directory:
        candidate = Path(directory) / 'candidate.docx'
        shutil.copy2(output, candidate)
        restored = InteractiveRebuildService._restore_source_cross_paragraph_field_shells(candidate, source)
        after = parser.parse(candidate)
    if original_hashes != (digest(source), digest(output)):
        raise RuntimeError('Retained document integrity changed during replay')
    before_gates = {k: v.passed for k, v in build_model_gates(expected, before).items()}
    after_gates = {k: v.passed for k, v in build_model_gates(expected, after).items()}
    return {
        'restored_shells': restored,
        'fields': {'source': len(expected.fields), 'before': len(before.fields), 'after': len(after.fields)},
        'before': before_gates,
        'after': after_gates,
        'regressed_gates': [k for k, passed in before_gates.items() if passed and not after_gates[k]],
        'retained_documents_unchanged': True,
        'full_golden_verified': False,
        'note': 'Offline G0-G7 replay only. Fresh Word reconstruction and G8-G9 validation are still required.',
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--local-root', default=r'C:\WordReplica-Automation')
    parser.add_argument('--golden', default='golden_2')
    args = parser.parse_args(argv)
    root = Path(args.local_root)
    repo = Path(__file__).resolve().parents[2]
    config = load_config(repo / 'codex_automation.json', local_root_override=root)
    config.golden_filename_for(args.golden)  # reject unknown ids before constructing paths
    folder = root / 'diagnostics'
    if args.golden != 'golden_1':
        folder /= args.golden
    reports = list(folder.glob('*/golden_report.json'))
    if not reports:
        raise RuntimeError(f'No retained report for {args.golden}')
    report_path = max(reports, key=lambda p: p.stat().st_mtime_ns)
    report = json.loads(report_path.read_text(encoding='utf-8-sig'))
    source = report_path.parent / 'source.docx'
    output = report_path.parent / 'interactive/output.docx'
    if digest(source) != report.get('source_sha256'):
        raise RuntimeError('Archived source does not match report SHA-256')
    result = replay_pair(source, output)
    result.update({
        'golden_id': args.golden,
        'baseline_commit': report.get('commit_sha'),
        'tested_commit': subprocess.check_output(['git', '-C', str(repo), 'rev-parse', 'HEAD'], text=True).strip(),
        'run_id': report.get('run_id'),
        'stored_stop_required': (report.get('automation_decision') or {}).get('stop_required'),
    })
    print(json.dumps(result, indent=2))
    return 2 if result['regressed_gates'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
