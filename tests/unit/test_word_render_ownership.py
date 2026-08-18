import sys
from types import SimpleNamespace

from word_replica.qa import word_render


class FakeDoc:
    def ExportAsFixedFormat(self, **kwargs):
        pass
    def Close(self, **kwargs):
        pass


class FakeDocuments:
    def Open(self, **kwargs):
        return FakeDoc()


class FakeApplication:
    Hwnd = 333
    def __init__(self):
        self.Documents = FakeDocuments()
        self.Visible = False
        self.DisplayAlerts = 1
    def Quit(self):
        pass


def test_pdf_export_records_and_clears_owned_word(tmp_path, monkeypatch):
    fake_pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    fake_client = SimpleNamespace(DispatchEx=lambda _name: FakeApplication())
    fake_win32com = SimpleNamespace(client=fake_client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    calls = []
    monkeypatch.setattr(word_render, "record_owned_word", lambda app, role, **kwargs: calls.append(("record", role)) or 888, raising=False)
    monkeypatch.setattr(word_render, "clear_owned_word", lambda pid: calls.append(("clear", pid)), raising=False)

    word_render.export_docx_to_pdf_with_word(tmp_path / "a.docx", tmp_path / "a.pdf")

    assert calls == [("record", "pdf-export"), ("clear", 888)]
