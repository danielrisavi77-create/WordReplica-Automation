import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

import pytest

from scripts.persistent_harness.archive import UpdateValidationError, validate_update_zip


def _make_update_zip(tmp_path: Path, *, member_name="payload/app.txt", content=b"hello", bad_hash=False, extra_member=None):
    digest = hashlib.sha256(content).hexdigest()
    manifest = {
        "schema_version": 1,
        "package_id": "word-replica-remote-update-2.1.0",
        "target_version": "2.1.0",
        "minimum_installed_version": "2.0.0",
        "files": {member_name.removeprefix("payload/"): ("0" * 64 if bad_hash else digest)},
    }
    path = tmp_path / "update.zip"
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("update_manifest.json", json.dumps(manifest))
        zf.writestr(member_name, content)
        if extra_member:
            zf.writestr(extra_member, b"evil")
    return path


def test_update_accepts_valid_payload(tmp_path):
    validated = validate_update_zip(_make_update_zip(tmp_path), "2.0.0")
    assert validated.target_version == "2.1.0"
    assert validated.payload_members == ("payload/app.txt",)
    assert len(validated.package_sha256) == 64


def test_update_rejects_traversal_member(tmp_path):
    archive = _make_update_zip(tmp_path, extra_member="payload/../../realworld_input/x.docx")
    with pytest.raises(UpdateValidationError, match="unsafe archive path"):
        validate_update_zip(archive, "2.0.0")


def test_update_rejects_payload_hash_mismatch(tmp_path):
    archive = _make_update_zip(tmp_path, bad_hash=True)
    with pytest.raises(UpdateValidationError, match="SHA-256"):
        validate_update_zip(archive, "2.0.0")


def test_update_rejects_unlisted_payload_file(tmp_path):
    archive = _make_update_zip(tmp_path, extra_member="payload/unlisted.txt")
    with pytest.raises(UpdateValidationError, match="unlisted"):
        validate_update_zip(archive, "2.0.0")


def test_update_rejects_incompatible_installed_version(tmp_path):
    archive = _make_update_zip(tmp_path)
    with pytest.raises(UpdateValidationError, match="minimum installed version"):
        validate_update_zip(archive, "1.9.9")
