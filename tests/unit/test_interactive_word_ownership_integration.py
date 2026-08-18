import sys
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
