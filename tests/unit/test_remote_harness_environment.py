from scripts.remote_harness.prerequisites import check_prerequisites


def test_prerequisite_report_requires_docx_and_word(tmp_path):
    input_dir=tmp_path/"input"; input_dir.mkdir()
    report = check_prerequisites(input_dir, tmp_path, os_name="nt", word_probe=lambda: (False, {"error":"no word"}), manifest_probe=lambda root:(True, []), free_space_probe=lambda root: 5000)
    assert report.ok is False
    assert any("Microsoft Word" in item for item in report.errors)
    assert any(".docx" in item for item in report.errors)


def test_prerequisite_report_can_pass_with_injected_probes(tmp_path):
    input_dir=tmp_path/"input"; input_dir.mkdir(); (input_dir/"a.docx").write_bytes(b"x")
    report = check_prerequisites(input_dir, tmp_path, os_name="nt", word_probe=lambda: (True, {"version":"16"}), manifest_probe=lambda root:(True, []), free_space_probe=lambda root: 5000)
    assert report.ok is True
    assert report.word["version"] == "16"
