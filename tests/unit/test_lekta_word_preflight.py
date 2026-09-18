from __future__ import annotations

from pathlib import Path

import pytest

from word_replica.runner.word_preflight import WordPreflightError, run_word_preflight


class FakeContent:
    Text = ""


class FakeDocument:
    def __init__(self, order: list[str], *, fail_save: bool = False) -> None:
        self.order = order
        self.fail_save = fail_save
        self.Content = FakeContent()
        self.saved_path: Path | None = None

    def SaveAs2(self, path: str, *, FileFormat: int) -> None:
        self.order.append("save")
        assert FileFormat == 16
        if self.fail_save:
            raise RuntimeError("editing is disabled")
        self.saved_path = Path(path)
        self.saved_path.write_bytes(b"PK-scratch-docx")

    def Close(self, *, SaveChanges: int) -> None:
        self.order.append("close")
        assert SaveChanges == 0


class FakeDocuments:
    def __init__(self, document: FakeDocument, order: list[str]) -> None:
        self.document = document
        self.order = order

    def Add(self) -> FakeDocument:
        self.order.append("add")
        return self.document


class FakeWord:
    def __init__(self, document: FakeDocument, order: list[str]) -> None:
        self.Documents = FakeDocuments(document, order)
        self.order = order
        self.Visible: bool | None = None
        self.DisplayAlerts: int | None = None
        self.AutomationSecurity: int | None = None

    def Quit(self) -> None:
        self.order.append("quit")


def test_preflight_proves_word_can_create_save_and_delete_its_scratch_docx(tmp_path: Path) -> None:
    order: list[str] = []
    document = FakeDocument(order)
    word = FakeWord(document, order)

    def dispatch(name: str) -> FakeWord:
        order.append(f"dispatch:{name}")
        return word

    result = run_word_preflight(
        temp_parent=tmp_path,
        platform_name="win32",
        dispatch_factory=dispatch,
        co_initialize=lambda: order.append("coinit"),
        co_uninitialize=lambda: order.append("couninit"),
    )

    assert result.can_create_and_save is True
    assert order == ["coinit", "dispatch:Word.Application", "add", "save", "close", "quit", "couninit"]
    assert document.Content.Text == "Lekta Word preflight"
    assert word.Visible is False
    assert word.DisplayAlerts == 0
    assert word.AutomationSecurity == 3
    assert document.saved_path is not None
    assert not document.saved_path.exists()
    assert list(tmp_path.iterdir()) == []


def test_preflight_failure_closes_only_its_document_and_word_instance(tmp_path: Path) -> None:
    order: list[str] = []
    document = FakeDocument(order, fail_save=True)
    word = FakeWord(document, order)

    with pytest.raises(WordPreflightError, match="create and save"):
        run_word_preflight(
            temp_parent=tmp_path,
            platform_name="win32",
            dispatch_factory=lambda _name: word,
            co_initialize=lambda: order.append("coinit"),
            co_uninitialize=lambda: order.append("couninit"),
        )

    assert order == ["coinit", "add", "save", "close", "quit", "couninit"]
    assert list(tmp_path.iterdir()) == []


def test_preflight_rejects_non_windows_before_dispatch(tmp_path: Path) -> None:
    called = False

    def dispatch(_name: str) -> FakeWord:
        nonlocal called
        called = True
        raise AssertionError("must not dispatch")

    with pytest.raises(WordPreflightError, match="Windows"):
        run_word_preflight(temp_parent=tmp_path, platform_name="linux", dispatch_factory=dispatch)

    assert called is False
