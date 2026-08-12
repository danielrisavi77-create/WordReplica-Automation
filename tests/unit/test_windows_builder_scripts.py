from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def test_build_and_install_cmd_fails_fast_when_companion_scripts_are_missing():
    text = (ROOT / 'BUILD_AND_INSTALL_WORD_REPLICA.cmd').read_text(encoding='utf-8', errors='replace')
    assert 'if not exist "%~dp0BUILD_WINDOWS_APP.ps1" goto :not_extracted' in text
    assert 'if not exist "%~dp0INSTALL_WORD_REPLICA.ps1" goto :not_extracted' in text
    assert 'if not exist "%~dp0pyproject.toml" goto :not_extracted' in text


def test_build_and_install_cmd_verifies_exe_before_reporting_pass():
    text = (ROOT / 'BUILD_AND_INSTALL_WORD_REPLICA.cmd').read_text(encoding='utf-8', errors='replace')
    assert 'if not exist "%~dp0dist\\WordReplica.exe" goto :fail' in text
    assert 'if not exist "%LOCALAPPDATA%\\WordReplica\\WordReplica.exe" goto :fail' in text


def test_start_here_detects_zip_preview_and_never_claims_success_itself():
    text = (ROOT / 'START_HERE.cmd').read_text(encoding='utf-8', errors='replace')
    assert 'if not exist "%~dp0pyproject.toml" goto :not_extracted' in text
    assert 'Extract All' in text
    assert 'BUILD + INSTALL: PASS' not in text


def test_installer_is_windows_powershell_51_safe_ascii_and_verifies_hash():
    raw = (ROOT / 'INSTALL_WORD_REPLICA.ps1').read_bytes()
    assert all(byte < 128 for byte in raw), 'installer must be ASCII-safe for Windows PowerShell 5.1'
    text = raw.decode('ascii')
    assert '$sourceHash = (Get-FileHash $sourceExe -Algorithm SHA256).Hash' in text
    assert '$installedHash = (Get-FileHash $installExe -Algorithm SHA256).Hash' in text
    assert 'if ($sourceHash -ne $installedHash)' in text


def test_all_shipping_powershell_scripts_are_ascii_safe_for_windows_powershell_51():
    for name in [
        'INSTALL_WORD_REPLICA.ps1',
        'INSTALL_EXISTING_EXE_NOW.ps1',
    ]:
        raw = (ROOT / name).read_bytes()
        assert all(byte < 128 for byte in raw), f'{name} must be ASCII-safe for Windows PowerShell 5.1'


def test_existing_exe_installer_searches_downloads_and_propagates_failure():
    text = (ROOT / 'INSTALL_EXISTING_EXE_NOW.ps1').read_text(encoding='ascii')
    assert 'Downloads\\dist\\WordReplica.exe' in text
    assert 'INSTALL_WORD_REPLICA.ps1' in text
    assert 'if ($LASTEXITCODE -ne 0)' in text


def test_release_gate_has_independent_instant_and_interactive_sections_and_fail_fast_checks():
    text=(ROOT/'RUN_WINDOWS_RELEASE_GATE.ps1').read_text(encoding='utf-8',errors='replace')
    assert '=== INSTANT GATE ===' in text
    assert '=== INTERACTIVE GATE ===' in text
    assert 'INSTANT RESULT: 0 failed' in text
    assert 'INTERACTIVE RESULT: 0 failed' in text
    assert '$ErrorActionPreference = "Stop"' in text
    assert text.count('$LASTEXITCODE -ne 0') >= 2
    assert 'WORD REPLICA WINDOWS RELEASE GATE: PASS' in text


def test_build_script_invokes_dual_release_gate_before_pyinstaller():
    text=(ROOT/'BUILD_WINDOWS_APP.ps1').read_text(encoding='utf-8',errors='replace')
    gate_index=text.index('RUN_WINDOWS_RELEASE_GATE.ps1')
    build_index=text.index('PyInstaller')
    assert gate_index < build_index
