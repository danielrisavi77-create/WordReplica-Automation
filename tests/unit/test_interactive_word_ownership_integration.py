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
