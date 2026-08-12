from pathlib import Path

from word_replica.domain.enums import FidelityMode, MetadataMode, RendererChoice, VisibilityMode


def test_desktop_defaults_are_safe_and_match_release_defaults():
    from word_replica.desktop import DesktopState

    state = DesktopState()
    assert state.source_path is None
    assert state.renderer == "auto"
    assert state.visibility == "background"
    assert state.fidelity == "clean"
    assert state.metadata == "fresh"
    assert state.allow_source_overwrite is False
    assert state.preserve_author_fields is False


def test_desktop_state_builds_engine_options_without_mutation(tmp_path):
    from word_replica.desktop import DesktopState, options_from_desktop_state

    source = tmp_path / "paper.docx"
    state = DesktopState(
        source_path=source,
        renderer="word",
        visibility="visible",
        fidelity="full",
        metadata="preserve",
        preserve_author_fields=True,
        custom_metadata_allowlist=("Department", "Course"),
    )
    options = options_from_desktop_state(state)

    assert options.renderer is RendererChoice.WORD
    assert options.visibility is VisibilityMode.VISIBLE
    assert options.fidelity is FidelityMode.FULL
    assert options.metadata is MetadataMode.PRESERVE
    assert options.preserve_author_fields is True
    assert options.custom_metadata_allowlist == ("Department", "Course")
    assert state.source_path == source


def test_desktop_rejects_non_docx_before_start(tmp_path):
    from word_replica.desktop import DesktopState, validate_source

    state = DesktopState(source_path=tmp_path / "notes.pdf")
    ok, message = validate_source(state)
    assert ok is False
    assert ".docx" in message


def test_desktop_accepts_existing_docx(tmp_path):
    from word_replica.desktop import DesktopState, validate_source

    source = tmp_path / "paper.docx"
    source.write_bytes(b"placeholder")
    ok, message = validate_source(DesktopState(source_path=source))
    assert ok is True
    assert message == ""


def test_build_script_targets_single_exe_without_pyside_dependency():
    script = Path("BUILD_WINDOWS_APP.ps1").read_text(encoding="utf-8")
    assert "PyInstaller" in script
    assert "--onefile" in script
    assert "--windowed" in script
    assert "word_replica.desktop" in script
    assert "PySide6" not in script
    assert "WORD_REPLICA_WORD_TESTS" in script


def test_one_click_installer_uses_per_user_location_and_desktop_shortcut():
    installer = Path("INSTALL_WORD_REPLICA.ps1").read_text(encoding="utf-8")
    wrapper = Path("BUILD_AND_INSTALL_WORD_REPLICA.cmd").read_text(encoding="utf-8")
    assert "LOCALAPPDATA" in installer
    assert "Desktop" in installer
    assert "Word Replica.lnk" in installer
    assert "BUILD_WINDOWS_APP.ps1" in wrapper
    assert "INSTALL_WORD_REPLICA.ps1" in wrapper
