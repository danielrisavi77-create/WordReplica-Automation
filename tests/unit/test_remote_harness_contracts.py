from pathlib import Path

from scripts.remote_harness.config import HarnessConfig
from scripts.remote_harness.contracts import HarnessStatus, normalize_failure


def test_default_harness_timeouts_match_spec(tmp_path):
    cfg = HarnessConfig.load(tmp_path / "missing.json")
    assert cfg.static_timeout_seconds == 120
    assert cfg.instant_timeout_seconds == 600
    assert cfg.interactive_timeout_seconds == 1800
    assert cfg.l4_timeout_seconds == 600
    assert cfg.logs_only is False


def test_config_file_can_override_defaults(tmp_path):
    path = tmp_path / "harness_config.json"
    path.write_text('{"interactive_timeout_seconds": 42, "logs_only": true}', encoding="utf-8")
    cfg = HarnessConfig.load(path)
    assert cfg.interactive_timeout_seconds == 42
    assert cfg.logs_only is True
    assert cfg.instant_timeout_seconds == 600


def test_failure_normalizer_groups_known_root_causes():
    assert normalize_failure("(-2147418111, 'Call was rejected by callee.', None, None)") == "RPC_E_CALL_REJECTED"
    assert normalize_failure("RPC server unavailable 0x800706ba") == "RPC_SERVER_UNAVAILABLE"
    assert normalize_failure("'str' object has no attribute 'InsertAfter'") == "INVALID_ACTIVE_RANGE_TYPE"
    assert HarnessStatus.COM_FAIL.value == "COM_FAIL"


def test_normalize_failure_groups_invalid_word_header_root_errors():
    text = "Word encountered an error processing the XML file X; Location: Part: /word/header1.xml, Line: 2, Column: 89"
    assert normalize_failure(text) == "INVALID_WORD_HEADER_XML"


def test_normalize_failure_groups_undefined_image_extension():
    assert normalize_failure("Unsupported image format: .undefined") == "IMAGE_FORMAT_UNDEFINED"
