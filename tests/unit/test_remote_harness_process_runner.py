import sys

from scripts.remote_harness.process_runner import run_child


def test_timeout_is_normalized_and_next_stage_can_run(tmp_path):
    first = run_child([sys.executable, "-c", "import time; time.sleep(5)"], 1, tmp_path/"a.out", tmp_path/"a.err")
    second = run_child([sys.executable, "-c", "print('ok')"], 5, tmp_path/"b.out", tmp_path/"b.err")
    assert first.timed_out is True
    assert second.exit_code == 0
    assert "ok" in (tmp_path/"b.out").read_text()
