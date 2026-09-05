from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile

from scripts.persistent_harness.build_bootstrap import BOOTSTRAP_VERSION, build_bootstrap
from scripts.persistent_harness.build_update import build_update_zip
from scripts.persistent_harness.versioning import VersionMetadata
from scripts.persistent_harness.archive import validate_update_zip

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_bootstrap_places_code_only_under_current(tmp_path):
    root = build_bootstrap(PROJECT_ROOT, tmp_path / "persistent")
    assert (root / "current" / "scripts" / "remote_harness" / "main.py").exists()
    assert (root / "current" / "scripts" / "persistent_harness" / "updater.py").exists()
    assert (root / "realworld_input").is_dir()
    assert (root / "results").is_dir()
    assert (root / "backup").is_dir()
    assert (root / "updates").is_dir()
    assert not (root / "current" / "realworld_input").exists()
    assert not (root / "current" / "results").exists()
    assert (root / "START_HERE.cmd").exists()
    assert (root / "UPDATE_HARNESS.cmd").exists()
    assert (root / "ROLLBACK_HARNESS.cmd").exists()
    assert VersionMetadata.load(root / "version.json").harness_version == BOOTSTRAP_VERSION


def test_update_builder_produces_validator_compatible_code_only_zip(tmp_path):
    source = tmp_path / "source"
    (source / "src").mkdir(parents=True)
    (source / "src" / "x.py").write_text("x=1\n", encoding="utf-8")
    (source / "pyproject.toml").write_text("[project]\nname='x'\nversion='1'\ndependencies=[]\n", encoding="utf-8")
    destination = tmp_path / "update.zip"
    build_update_zip(source, destination, target_version="2.1.0", minimum_installed_version="2.0.0")
    validated = validate_update_zip(destination, "2.0.0")
    assert validated.target_version == "2.1.0"
    with ZipFile(destination) as zf:
        names = set(zf.namelist())
    assert "update_manifest.json" in names
    assert all(name == "update_manifest.json" or name.startswith("payload/") for name in names)


def test_bootstrap_generates_valid_remote_source_manifest(tmp_path):
    from scripts.remote_harness.prerequisites import _default_manifest_probe

    root = build_bootstrap(PROJECT_ROOT, tmp_path / "persistent")
    current = root / "current"
    manifest = current / "REMOTE_HARNESS_SOURCE_SHA256.txt"
    assert manifest.is_file()
    ok, errors = _default_manifest_probe(current)
    assert ok, errors


def test_update_builder_embeds_valid_remote_source_manifest(tmp_path):
    from scripts.remote_harness.prerequisites import _default_manifest_probe

    source = tmp_path / "source"
    (source / "src" / "word_replica").mkdir(parents=True)
    (source / "src" / "word_replica" / "__init__.py").write_text("x=1\n", encoding="utf-8")
    (source / "pyproject.toml").write_text("[project]\nname='x'\nversion='1'\ndependencies=[]\n", encoding="utf-8")
    destination = tmp_path / "update.zip"
    build_update_zip(source, destination, target_version="2.1.0", minimum_installed_version="2.0.0")

    extracted = tmp_path / "extracted"
    with ZipFile(destination) as zf:
        for name in zf.namelist():
            if name.startswith("payload/"):
                zf.extract(name, extracted)
    payload = extracted / "payload"
    assert (payload / "REMOTE_HARNESS_SOURCE_SHA256.txt").is_file()
    ok, errors = _default_manifest_probe(payload)
    assert ok, errors


def test_payload_file_iterator_excludes_app_build_virtual_environment(tmp_path):
    from scripts.persistent_harness.build_update import iter_payload_files

    source = tmp_path / "source"
    (source / "src").mkdir(parents=True)
    (source / "src" / "app.py").write_text("x = 1\n", encoding="utf-8")
    app_build = source / ".venv_appbuild" / "Lib" / "site-packages"
    app_build.mkdir(parents=True)
    (app_build / "native.pyd").write_bytes(b"binary")

    relative_paths = {rel.as_posix() for _path, rel in iter_payload_files(source)}

    assert relative_paths == {"src/app.py"}


def test_manifest_probe_preserves_leading_dot_in_directory_name(tmp_path):
    from scripts.remote_harness.prerequisites import _default_manifest_probe

    root = tmp_path / "root"
    hidden_file = root / ".config" / "settings.json"
    hidden_file.parent.mkdir(parents=True)
    payload = b"{}"
    hidden_file.write_bytes(payload)
    (root / "REMOTE_HARNESS_SOURCE_SHA256.txt").write_text(
        f"{sha256(payload).hexdigest()}  .config/settings.json\n",
        encoding="utf-8",
    )

    ok, errors = _default_manifest_probe(root)

    assert ok, errors
