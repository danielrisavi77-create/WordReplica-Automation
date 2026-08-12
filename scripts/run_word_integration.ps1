$env:WORD_REPLICA_WORD_TESTS = "1"
python -m pytest tests/integration/word -m word -v
exit $LASTEXITCODE
