from word_replica.cli import build_parser


def test_rebuild_command_maps_modes():
    args = build_parser().parse_args([
        'rebuild', 'paper.docx', '--renderer', 'auto', '--visibility', 'visible',
        '--fidelity', 'full', '--metadata', 'fresh'
    ])
    assert args.command == 'rebuild'
    assert args.visibility == 'visible'
    assert args.fidelity == 'full'
