$ErrorActionPreference = "Stop"
python -m pytest tests/unit tests/integration/test_pure_docx_e2e.py tests/acceptance -m "not word" -v
python -m PyInstaller --noconfirm --clean --name WordReplica --windowed --collect-all PySide6 src/word_replica/gui/app.py
Write-Host "Built: dist\WordReplica\WordReplica.exe"
