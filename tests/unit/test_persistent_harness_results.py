from scripts.remote_harness.main import build_run_name, build_result_zip_name


def test_run_names_include_version():
    assert build_run_name("2.0.0", "20260811-190500") == "2026-08-11_190500_v2.0.0"
    assert build_result_zip_name("2.0.0", "20260811-190500") == "WordReplica-Remote-Results-20260811-190500-v2.0.0.zip"
