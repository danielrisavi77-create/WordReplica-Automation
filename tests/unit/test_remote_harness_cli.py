from scripts.remote_harness.main import build_parser


def test_cli_accepts_logs_only_and_paths():
    args=build_parser().parse_args(["--input-dir","in","--config","cfg.json","--logs-only","--result-parent","out"])
    assert args.input_dir == "in"
    assert args.logs_only is True
    assert args.result_parent == "out"


def test_cli_accepts_explicit_persistent_paths_and_version():
    args=build_parser().parse_args(["--input-dir","input","--config","cfg.json","--result-parent","results","--harness-version","2.0.0"])
    assert args.harness_version == "2.0.0"
