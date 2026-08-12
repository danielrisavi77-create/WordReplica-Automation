from word_replica.domain.reconstruction import ReconstructionBlueprint, ReconstructionEvent


def test_blueprint_fingerprint_is_deterministic_and_payload_sensitive():
    a = ReconstructionBlueprint.build(
        source_sha256="a" * 64,
        source_model_fingerprint="model",
        events=(ReconstructionEvent("InsertCharacter", "run_1", {"character": "A"}),),
    )
    b = ReconstructionBlueprint.build(
        source_sha256="a" * 64,
        source_model_fingerprint="model",
        events=(ReconstructionEvent("InsertCharacter", "run_1", {"character": "A"}),),
    )
    c = ReconstructionBlueprint.build(
        source_sha256="a" * 64,
        source_model_fingerprint="model",
        events=(ReconstructionEvent("InsertCharacter", "run_1", {"character": "B"}),),
    )
    assert a.fingerprint == b.fingerprint
    assert a.fingerprint != c.fingerprint
    assert a.total_visible_characters == 1
    assert a.total_events == 1
