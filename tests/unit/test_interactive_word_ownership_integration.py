import sys
from pathlib import Path
from types import SimpleNamespace

from word_replica.renderers import interactive_word


class FakeRange:
    pass


class FakeDocument:
    def Range(self, start, end):
        return FakeRange()
    def Close(self, *_args, **_kwargs):
        pass


class FakeDocuments:
    def Add(self):
        return FakeDocument()


class FakeApplication:
    Hwnd = 555
    def __init__(self):
        self.Documents = FakeDocuments()
        self.Visible = False
    def Quit(self):
        pass


class RejectedCall(Exception):
    hresult = interactive_word.RPC_E_CALL_REJECTED


def test_interactive_controller_records_and_clears_owned_word(monkeypatch):
    fake_pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    fake_client = SimpleNamespace(DispatchEx=lambda _name: FakeApplication())
    fake_win32com = SimpleNamespace(client=fake_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    calls = []
    monkeypatch.setattr(interactive_word, "record_owned_word", lambda app, role, **kwargs: calls.append(("record", role)) or 777, raising=False)
    monkeypatch.setattr(interactive_word, "clear_owned_word", lambda pid: calls.append(("clear", pid)), raising=False)

    controller = interactive_word.InteractiveWordController()
    controller.open_blank()
    controller.close()

    assert calls == [("record", "interactive"), ("clear", 777)]


def test_open_existing_retries_range_and_defers_visibility_until_checkpoint_restore(monkeypatch, tmp_path):
    monkeypatch.setattr(interactive_word.time, "sleep", lambda _seconds: None)
    calls = []

    class RejectOnceDocument:
        def __init__(self):
            self.Content = SimpleNamespace(Start=0, End=11)
            self.range_attempts = 0

        def Range(self, start, end):
            self.range_attempts += 1
            if self.range_attempts == 1:
                raise RejectedCall("range busy")
            calls.append("document.Range")
            return (start, end)

        def Activate(self):
            calls.append("document.Activate")

        def Close(self, *_args, **_kwargs):
            pass

    document = RejectOnceDocument()

    class ExistingDocuments:
        def Open(self, *_args, **_kwargs):
            calls.append("Documents.Open")
            return document

    class ExistingApplication:
        Hwnd = 555

        def __init__(self):
            object.__setattr__(self, "Documents", ExistingDocuments())
            object.__setattr__(self, "Visible", False)

        def __setattr__(self, name, value):
            if name in {"DisplayAlerts", "Visible"}:
                calls.append((name, value))
            object.__setattr__(self, name, value)

        def Quit(self):
            pass

        def Activate(self):
            calls.append("application.Activate")

    application = ExistingApplication()
    fake_pythoncom = SimpleNamespace(
        CoInitialize=lambda: None,
        CoUninitialize=lambda: None,
        PumpWaitingMessages=lambda: None,
    )
    fake_client = SimpleNamespace(DispatchEx=lambda _name: application)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    monkeypatch.setattr(interactive_word, "word_process_pids", lambda: set())
    monkeypatch.setattr(interactive_word, "record_owned_word", lambda *args, **kwargs: 777)
    monkeypatch.setattr(interactive_word, "clear_owned_word", lambda _pid: None)

    controller = interactive_word.InteractiveWordController()
    controller.open_existing(tmp_path / "checkpoint.docx")
    try:
        assert calls.index(("DisplayAlerts", 0)) < calls.index("Documents.Open")
        assert calls.index("Documents.Open") < calls.index("document.Range")
        assert ("Visible", True) not in calls
        assert "document.Activate" not in calls
        assert "application.Activate" not in calls
        checkpoint = SimpleNamespace(
            story="body",
            range_start=10,
            range_end=10,
            paragraph_started=True,
            table_element_id=None,
            cell_element_id=None,
            resume_state={"section_index": 0, "section_started": False},
        )
        controller.restore_checkpoint_state(checkpoint)
        assert calls.index("document.Range") < calls.index(("Visible", True))
        assert calls[-2:] == ["document.Activate", "application.Activate"]
        assert document.range_attempts == 3
        assert controller.active_range == (10, 10)
        assert controller._last_saved_path == tmp_path / "checkpoint.docx"
    finally:
        controller.close()


def test_close_retries_rejected_document_and_application_calls_before_clearing_ownership(monkeypatch):
    monkeypatch.setattr(interactive_word.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(interactive_word, "word_process_pids", lambda: set())
    calls = []

    class RejectOnceDocument:
        attempts = 0

        def Close(self, _save_changes):
            self.attempts += 1
            if self.attempts == 1:
                raise RejectedCall("document busy")

    class RejectOnceApplication:
        attempts = 0

        def Quit(self):
            self.attempts += 1
            if self.attempts == 1:
                raise RejectedCall("application busy")

    controller = interactive_word.InteractiveWordController()
    controller.document = RejectOnceDocument()
    controller.application = RejectOnceApplication()
    controller._owned_word_pid = 777
    monkeypatch.setattr(interactive_word, "clear_owned_word", lambda pid: calls.append(pid))

    document = controller.document
    application = controller.application
    controller.close()

    assert document.attempts == 2
    assert application.attempts == 2
    assert calls == [777]


def test_close_preserves_ownership_record_when_quit_returns_but_word_pid_is_still_live(monkeypatch):
    calls = []

    controller = interactive_word.InteractiveWordController()
    controller.application = FakeApplication()
    controller._owned_word_pid = 777
    monkeypatch.setattr(interactive_word, "word_process_pids", lambda: {777})
    monkeypatch.setattr(interactive_word, "clear_owned_word", lambda pid: calls.append(pid))

    controller.close()

    assert calls == []


def test_close_preserves_ownership_record_when_application_quit_never_succeeds(monkeypatch):
    monkeypatch.setattr(interactive_word.time, "sleep", lambda _seconds: None)
    calls = []

    class RejectingApplication:
        def Quit(self):
            raise RejectedCall("application busy")

    controller = interactive_word.InteractiveWordController()
    controller.application = RejectingApplication()
    controller._owned_word_pid = 777
    monkeypatch.setattr(interactive_word, "clear_owned_word", lambda pid: calls.append(pid))

    controller.close()

    assert calls == []


def test_close_retries_rotating_checkpoint_cleanup_after_application_quit(tmp_path, monkeypatch):
    rotating = tmp_path / "wr-1234567890abcdef.docx"
    delivery = tmp_path / "checkpoint.docx"
    rotating.write_bytes(b"durable-checkpoint")
    calls = []

    class ClosingDocument:
        def Close(self, _save_changes):
            calls.append("document.Close")

    class UnlockingApplication:
        quit_called = False

        def Quit(self):
            assert controller.active_range is None
            assert controller._active_section is None
            assert controller._table_stack == []
            assert controller._story_stack == []
            assert controller._paragraph_style_cache == {}
            calls.append("application.Quit")
            self.quit_called = True

    application = UnlockingApplication()
    original_unlink = Path.unlink

    def locked_until_quit(path, *args, **kwargs):
        if path == rotating and not application.quit_called:
            raise PermissionError("Word still owns rotating checkpoint")
        return original_unlink(path, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", locked_until_quit)
    monkeypatch.setattr(interactive_word, "word_process_pids", lambda: set())
    monkeypatch.setattr(interactive_word, "clear_owned_word", lambda _pid: None)
    controller = interactive_word.InteractiveWordController()
    controller.document = ClosingDocument()
    controller.application = application
    controller.active_range = object()
    controller._active_section = object()
    controller._table_stack = [{"table": object()}]
    controller._story_stack = [(object(), False)]
    controller._paragraph_style_cache = {("p", "s"): object()}
    controller._last_saved_path = rotating
    controller._checkpoint_delivery_path = delivery

    controller.close()

    assert delivery.read_bytes() == b"durable-checkpoint"
    assert not rotating.exists()
    assert calls == ["document.Close", "application.Quit"]


def test_close_does_not_fail_a_durable_checkpoint_when_redundant_rotating_file_stays_locked(tmp_path, monkeypatch):
    rotating = tmp_path / "wr-1234567890abcdef.docx"
    delivery = tmp_path / "checkpoint.docx"
    rotating.write_bytes(b"durable")
    controller = interactive_word.InteractiveWordController()
    controller.document = FakeDocument()
    controller.application = FakeApplication()
    controller._last_saved_path = rotating
    controller._checkpoint_delivery_path = delivery
    monkeypatch.setattr(interactive_word, "word_process_pids", lambda: {777})
    monkeypatch.setattr(
        interactive_word, "_unlink_with_retry",
        lambda _path: (_ for _ in ()).throw(PermissionError("still locked")),
    )

    controller.close()

    assert delivery.read_bytes() == b"durable"
    assert rotating.exists()


def test_interactive_controller_honors_background_visibility(monkeypatch):
    fake_pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    application = FakeApplication()
    application.Visible = True
    fake_client = SimpleNamespace(DispatchEx=lambda _name: application)
    fake_win32com = SimpleNamespace(client=fake_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    monkeypatch.setattr(interactive_word, "record_owned_word", lambda *args, **kwargs: 777)
    monkeypatch.setattr(interactive_word, "clear_owned_word", lambda _pid: None)

    controller = interactive_word.InteractiveWordController(visible=False)
    controller.open_blank()
    try:
        assert application.Visible is False
    finally:
        controller.close()
