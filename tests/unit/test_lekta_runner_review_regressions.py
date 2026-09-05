from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from urllib.request import Request

import pytest

from word_replica.repair_contract.signature import RepairContractSignatureError, decode_spki
from word_replica.runner.http_transport import TransportError, _HttpsOnlyRedirectHandler
from word_replica.runner.lekta_claim import LaunchTicket
from word_replica.runner.one_shot import OneShotRunner, OneShotRunnerConfig
from word_replica.runner.secure_retry import SecureRetryStore


FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"


def test_redirect_handler_rejects_https_to_http_before_following() -> None:
    handler = _HttpsOnlyRedirectHandler()

    with pytest.raises(TransportError, match="HTTPS"):
        handler.redirect_request(
            Request("https://storage.example/source"),
            None,
            302,
            "Found",
            {},
            "http://storage.example/source",
        )


def test_one_shot_surfaces_signature_mismatch_before_any_download(tmp_path: Path) -> None:
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    contract["sourceSize"] += 1
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))

    class ClaimOnlyTransport:
        def post_json(self, _url: str, payload: dict) -> dict:
            return {
                "ok": True,
                "jobId": payload["jobId"],
                "contract": contract,
                "sourceDownloadUrl": "https://storage.example/source",
                "targetDownloadUrl": "https://storage.example/target",
                "expiresInSeconds": 300,
            }

        def get_bytes(self, _url: str, *, maximum_bytes: int) -> bytes:
            raise AssertionError(f"download must not start: {maximum_bytes}")

    runner = OneShotRunner(
        preflight=lambda: None,
        transport=ClaimOnlyTransport(),
        package_service=SimpleNamespace(run=lambda *_a, **_k: (_ for _ in ()).throw(AssertionError("Word must not run"))),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=SecureRetryStore(tmp_path / "retry"),
    )

    with pytest.raises(RepairContractSignatureError, match="signature-mismatch"):
        runner.run(
            LaunchTicket.parse({"version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43}),
                OneShotRunnerConfig(
                    "https://project.supabase.co/claim",
                    "https://project.supabase.co/status",
                    tmp_path / "out",
                ),
        )
