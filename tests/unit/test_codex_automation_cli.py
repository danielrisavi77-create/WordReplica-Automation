from scripts.codex_automation.cli import build_parser


def test_cli_defaults_to_repo_config_and_optional_local_root():
    parser = build_parser()
    args = parser.parse_args([])
    assert args.config == "codex_automation.json"
    assert args.local_root is None
    assert args.visible_word is False
    assert args.golden == "golden_1"


def test_cli_accepts_visible_word_opt_in():
    args = build_parser().parse_args(["--visible-word"])

    assert args.visible_word is True


def test_cli_accepts_golden_id_override():
    args = build_parser().parse_args(["--golden", "golden_2"])

    assert args.golden == "golden_2"
