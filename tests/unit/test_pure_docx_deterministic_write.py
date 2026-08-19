"""The fidelity lab's generated corpus is addressed by content hash: a seed must
produce a byte-identical .docx forever, or GOLDEN-CORE silently rebaselines and
every accumulated regression history is orphaned.

MutableDocxPackage.write_atomic cannot provide that -- ZipFile.writestr stamps
each member with the current local time and iterates dict insertion order. These
tests pin the deterministic variant.
"""
import os
from pathlib import Path
from zipfile import ZipFile

from word_replica.renderers.pure_docx import MutableDocxPackage

PARTS = {
    "[Content_Types].xml": b"<Types/>",
    "word/document.xml": b"<document/>",
    "_rels/.rels": b"<Relationships/>",
    "word/media/image1.png": b"\x89PNG\r\n\x1a\n" + b"x" * 64,
}


def test_write_deterministic_is_byte_identical_across_writes(tmp_path):
    a, b = tmp_path / "a.docx", tmp_path / "b.docx"
    MutableDocxPackage(dict(PARTS)).write_deterministic(a)
    MutableDocxPackage(dict(PARTS)).write_deterministic(b)

    assert a.read_bytes() == b.read_bytes()


def test_write_deterministic_ignores_part_insertion_order(tmp_path):
    forward, reverse = tmp_path / "f.docx", tmp_path / "r.docx"
    MutableDocxPackage(dict(PARTS)).write_deterministic(forward)
    MutableDocxPackage({k: PARTS[k] for k in reversed(list(PARTS))}).write_deterministic(reverse)

    assert forward.read_bytes() == reverse.read_bytes()


def test_write_deterministic_pins_member_timestamps(tmp_path):
    out = tmp_path / "out.docx"
    MutableDocxPackage(dict(PARTS)).write_deterministic(out)

    with ZipFile(out) as archive:
        assert {info.date_time for info in archive.infolist()} == {(1980, 1, 1, 0, 0, 0)}


def test_write_deterministic_orders_opc_entries_canonically(tmp_path):
    # Readers tolerate any order, but a stable order is what makes the bytes
    # stable. [Content_Types].xml first, then _rels/.rels, then sorted.
    out = tmp_path / "out.docx"
    MutableDocxPackage(dict(PARTS)).write_deterministic(out)

    with ZipFile(out) as archive:
        assert archive.namelist() == [
            "[Content_Types].xml",
            "_rels/.rels",
            "word/document.xml",
            "word/media/image1.png",
        ]


def test_write_deterministic_round_trips_every_part(tmp_path):
    out = tmp_path / "out.docx"
    MutableDocxPackage(dict(PARTS)).write_deterministic(out)

    with ZipFile(out) as archive:
        assert {name: archive.read(name) for name in archive.namelist()} == PARTS


def test_write_deterministic_replaces_an_existing_file_atomically(tmp_path):
    out = tmp_path / "out.docx"
    out.write_bytes(b"stale")
    MutableDocxPackage(dict(PARTS)).write_deterministic(out)

    with ZipFile(out) as archive:
        assert archive.read("word/document.xml") == b"<document/>"
    assert not list(tmp_path.glob("*.tmp"))


def test_write_atomic_remains_non_deterministic_and_untouched(tmp_path):
    # write_atomic is what the live renderer uses; this change must not alter it.
    out = tmp_path / "legacy.docx"
    MutableDocxPackage(dict(PARTS)).write_atomic(out)

    with ZipFile(out) as archive:
        assert archive.read("word/document.xml") == b"<document/>"


def test_write_deterministic_is_stable_across_processes(tmp_path):
    # Same-process equality cannot catch hash-seed-dependent iteration order.
    # The determinism contract is a cross-process one, so assert it that way.
    import subprocess
    import sys

    script = (
        "import sys, hashlib\n"
        "from pathlib import Path\n"
        "from word_replica.renderers.pure_docx import MutableDocxPackage\n"
        f"parts = {PARTS!r}\n"
        "out = Path(sys.argv[1])\n"
        "MutableDocxPackage(dict(parts)).write_deterministic(out)\n"
        "print(hashlib.sha256(out.read_bytes()).hexdigest())\n"
    )

    digests = set()
    for seed in ("0", "1", "2"):
        result = subprocess.run(
            [sys.executable, "-c", script, str(tmp_path / f"out{seed}.docx")],
            capture_output=True, text=True, check=True,
            env={**os.environ, "PYTHONHASHSEED": seed},
        )
        digests.add(result.stdout.strip())

    assert len(digests) == 1, digests
