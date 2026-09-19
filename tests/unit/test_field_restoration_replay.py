import hashlib
from zipfile import ZipFile, ZIP_DEFLATED

from tests.fixtures.build_fixtures import build_plain_text
from scripts.codex_automation.replay_field_restoration import replay_pair


def test_replay_recovers_graph_list_field_without_modifying_retained_documents(tmp_path):
    source = tmp_path / 'source.docx'
    output = tmp_path / 'output.docx'
    build_plain_text(source)
    with ZipFile(source) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    start = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\h \\z \\c "Grafikon" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    nested = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    for path, shell in ((source, True), (output, False)):
        xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        xml += '<w:p>' + (start if shell else '') + '<w:r><w:rPr><w:b/></w:rPr><w:t>Grafikon 1</w:t></w:r>' + nested + '</w:p>'
        xml += '<w:p><w:r><w:t>Grafikon 2</w:t></w:r>' + (end if shell else '') + '</w:p></w:body></w:document>'
        with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, xml.encode() if name == 'word/document.xml' else data)
    hashes = [hashlib.sha256(p.read_bytes()).hexdigest() for p in (source, output)]
    result = replay_pair(source, output)
    assert result['restored_shells'] == 1
    assert result['fields'] == {'source': 2, 'before': 1, 'after': 2}
    assert result['before']['G1'] is False
    assert result['before']['G7'] is False
    assert result['after']['G1'] is True
    assert result['after']['G7'] is True
    assert result['regressed_gates'] == []
    assert result['full_golden_verified'] is False
    assert hashes == [hashlib.sha256(p.read_bytes()).hexdigest() for p in (source, output)]
