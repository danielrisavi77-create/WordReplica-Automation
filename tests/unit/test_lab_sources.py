"""Corpus source adapters.

Every test here runs offline against an injected transport. Network calls in a
unit suite are a reliability tax nobody chose to pay, and the interesting
behaviour -- retry, backoff, the circuit breaker, listing shapes, licence
policy -- is exactly the part that does not need a real server.
"""
import base64
import json

import pytest

from word_replica.lab.sources import (
    CircuitOpen,
    DocRef,
    PoliteFetcher,
    RetentionPolicy,
    apache_poi_source,
    libreoffice_source,
    local_tree_source,
)

GITILES_LISTING = ")]}'\n" + json.dumps({
    "id": "abc",
    "entries": [
        {"mode": 33188, "type": "blob", "id": "1" * 40, "name": "1-table-1-page.docx"},
        {"mode": 33188, "type": "blob", "id": "2" * 40, "name": "tdf99631.docx"},
        {"mode": 33188, "type": "blob", "id": "3" * 40, "name": "notes.odt"},
        {"mode": 16384, "type": "tree", "id": "4" * 40, "name": "subdir"},
    ],
})

POI_TREE = json.dumps({
    "truncated": False,
    "tree": [
        {"path": "test-data/document/51921-Word-Crash067.docx", "type": "blob", "size": 20006},
        {"path": "test-data/document/Bug53453Test.doc", "type": "blob", "size": 100},
        {"path": "test-data/spreadsheet/x.xlsx", "type": "blob", "size": 100},
    ],
})


class FakeTransport:
    """Records calls and replays canned responses, so retries are observable."""

    def __init__(self, responses):
        self.responses = responses
        self.calls: list[str] = []

    def __call__(self, url: str, *, timeout: float = 0) -> bytes:
        self.calls.append(url)
        result = self.responses.get(url)
        if result is None:
            for pattern, value in self.responses.items():
                if pattern in url:
                    result = value
                    break
        if result is None:
            raise FileNotFoundError(url)
        if isinstance(result, list):
            result = result.pop(0)
        if isinstance(result, Exception):
            raise result
        return result


# --- polite fetching ---------------------------------------------------------

def test_a_transient_failure_is_retried_and_then_succeeds():
    transport = FakeTransport({"x": [TimeoutError("reset"), b"payload"]})
    fetcher = PoliteFetcher(transport=transport, sleep=lambda _s: None, max_attempts=3)

    assert fetcher.get("https://example.invalid/x") == b"payload"
    assert len(transport.calls) == 2


def test_backoff_grows_between_attempts():
    slept: list[float] = []
    transport = FakeTransport({"x": [TimeoutError("a"), TimeoutError("b"), b"ok"]})
    fetcher = PoliteFetcher(transport=transport, sleep=slept.append, max_attempts=4, base_delay=1.0, jitter=0.0)

    fetcher.get("https://example.invalid/x")

    assert slept == [1.0, 2.0]


def test_giving_up_reports_the_last_error():
    transport = FakeTransport({"x": TimeoutError("always")})
    fetcher = PoliteFetcher(transport=transport, sleep=lambda _s: None, max_attempts=2)

    with pytest.raises(TimeoutError):
        fetcher.get("https://example.invalid/x")


def test_the_circuit_opens_after_repeated_failures_instead_of_hammering_the_host():
    # A 20,000-document crawl against a host that has started refusing must stop
    # and say so, not spend hours recording the same error.
    transport = FakeTransport({"x": TimeoutError("down")})
    fetcher = PoliteFetcher(
        transport=transport, sleep=lambda _s: None, max_attempts=1, circuit_threshold=3
    )

    for _ in range(3):
        with pytest.raises(TimeoutError):
            fetcher.get("https://example.invalid/x")

    with pytest.raises(CircuitOpen):
        fetcher.get("https://example.invalid/x")
    assert len(transport.calls) == 3  # the fourth call never reached the transport


def test_a_success_closes_the_circuit_again():
    transport = FakeTransport({"x": [TimeoutError("a"), TimeoutError("b"), b"ok", b"ok"]})
    fetcher = PoliteFetcher(
        transport=transport, sleep=lambda _s: None, max_attempts=1, circuit_threshold=3
    )

    for _ in range(2):
        with pytest.raises(TimeoutError):
            fetcher.get("https://example.invalid/x")
    assert fetcher.get("https://example.invalid/x") == b"ok"
    assert fetcher.get("https://example.invalid/x") == b"ok"


# --- LibreOffice via gitiles -------------------------------------------------

def test_libreoffice_listing_keeps_only_word_documents():
    transport = FakeTransport({"format=JSON": GITILES_LISTING.encode()})
    source = libreoffice_source(directories=("ooxmlexport",), fetcher=PoliteFetcher(transport=transport))

    refs = list(source.list_candidates())

    assert [ref.name for ref in refs] == ["1-table-1-page.docx", "tdf99631.docx"]
    assert all(isinstance(ref, DocRef) for ref in refs)


def test_libreoffice_records_the_blob_id_for_change_detection():
    # The blob sha1 is what makes a monthly refresh cheap: only changed files
    # need downloading, and it costs nothing extra to record.
    transport = FakeTransport({"format=JSON": GITILES_LISTING.encode()})
    source = libreoffice_source(directories=("ooxmlexport",), fetcher=PoliteFetcher(transport=transport))

    first = next(iter(source.list_candidates()))

    assert first.revision == "1" * 40
    assert first.source_ref == "ooxmlexport/data/1-table-1-page.docx"


def test_libreoffice_strips_the_gitiles_xssi_prefix():
    # Gitiles prefixes JSON with )]}' to defeat cross-site script inclusion.
    # Forgetting to strip it is the classic first bug against this API.
    transport = FakeTransport({"format=JSON": GITILES_LISTING.encode()})
    source = libreoffice_source(directories=("ooxmlexport",), fetcher=PoliteFetcher(transport=transport))

    assert list(source.list_candidates())


def test_libreoffice_fetch_decodes_the_base64_blob():
    payload = b"PK\x03\x04 pretend docx"
    transport = FakeTransport({
        "format=JSON": GITILES_LISTING.encode(),
        "format=TEXT": base64.b64encode(payload),
    })
    source = libreoffice_source(directories=("ooxmlexport",), fetcher=PoliteFetcher(transport=transport))
    ref = next(iter(source.list_candidates()))

    assert source.fetch(ref) == payload


def test_libreoffice_test_data_may_be_kept_verbatim():
    source = libreoffice_source(directories=("ooxmlexport",), fetcher=PoliteFetcher(transport=FakeTransport({})))

    assert source.retention is RetentionPolicy.KEEP_VERBATIM
    assert "MPL" in source.licence


# --- Apache POI --------------------------------------------------------------

def test_poi_listing_selects_only_word_documents_under_test_data():
    transport = FakeTransport({"git/trees": POI_TREE.encode()})
    source = apache_poi_source(fetcher=PoliteFetcher(transport=transport))

    refs = list(source.list_candidates())

    assert [ref.name for ref in refs] == ["51921-Word-Crash067.docx"]
    assert refs[0].source_ref == "test-data/document/51921-Word-Crash067.docx"


def test_poi_reports_a_truncated_tree_rather_than_silently_listing_part_of_it():
    truncated = json.dumps({"truncated": True, "tree": []})
    transport = FakeTransport({"git/trees": truncated.encode()})
    source = apache_poi_source(fetcher=PoliteFetcher(transport=transport))

    with pytest.raises(RuntimeError, match="truncated"):
        list(source.list_candidates())


def test_poi_test_data_may_be_kept_verbatim():
    source = apache_poi_source(fetcher=PoliteFetcher(transport=FakeTransport({})))

    assert source.retention is RetentionPolicy.KEEP_VERBATIM
    assert "Apache" in source.licence


# --- local tree --------------------------------------------------------------

def test_local_tree_lists_documents_relative_to_its_root(tmp_path):
    (tmp_path / "nested").mkdir()
    (tmp_path / "a.docx").write_bytes(b"PK a")
    (tmp_path / "nested" / "b.docx").write_bytes(b"PK b")
    (tmp_path / "ignore.txt").write_text("no")

    source = local_tree_source(tmp_path)
    refs = sorted(source.list_candidates(), key=lambda ref: ref.source_ref)

    assert [ref.source_ref for ref in refs] == ["a.docx", "nested/b.docx"]
    assert source.fetch(refs[0]) == b"PK a"


def test_local_tree_needs_no_network_at_all(tmp_path):
    (tmp_path / "a.docx").write_bytes(b"PK a")

    def _explode(*_args, **_kwargs):
        raise AssertionError("local_tree_source must not use the network")

    source = local_tree_source(tmp_path, fetcher=PoliteFetcher(transport=_explode))

    assert source.fetch(next(iter(source.list_candidates()))) == b"PK a"
