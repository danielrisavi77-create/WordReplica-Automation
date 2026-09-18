from word_replica.qa.report import write_qa_report


def test_report_explains_status_and_l4_limitations(tmp_path):
    path = write_qa_report(
        tmp_path / 'qa_report.html',
        status='WARN',
        levels={},
        warnings=[{'code': 'L4_UNAVAILABLE', 'message': 'Word not installed'}],
        render_result=None,
    )
    html = path.read_text(encoding='utf-8')
    assert 'WARN' in html
    assert 'L4_UNAVAILABLE' in html
    assert 'Word not installed' in html
    assert 'automated process' in html
    assert 'not proof of manual authorship' in html
    assert 'does not fabricate timestamps' in html


def test_report_renders_environment_when_provided(tmp_path):
    path = write_qa_report(
        tmp_path / 'qa_report.html',
        status='PASS',
        levels={},
        warnings=[],
        render_result=None,
        environment={'os': {'system': 'Windows'}, 'word': {'available': True, 'build': '16.0.1'}},
    )
    html = path.read_text(encoding='utf-8')
    assert 'Environment' in html
    assert 'Windows' in html
    assert '16.0.1' in html


def test_report_omits_environment_section_when_absent(tmp_path):
    path = write_qa_report(
        tmp_path / 'qa_report.html',
        status='PASS',
        levels={},
        warnings=[],
        render_result=None,
    )
    html = path.read_text(encoding='utf-8')
    assert '<h2>Environment</h2>' not in html
