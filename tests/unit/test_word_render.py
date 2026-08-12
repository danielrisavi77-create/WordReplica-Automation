import pytest

from word_replica.qa.word_render import _export_with_application


class FakeDocument:
    def __init__(self, log, *, export_error=None, close_error=None):
        self.log=log; self.export_error=export_error; self.close_error=close_error
    def ExportAsFixedFormat(self, **kwargs):
        self.log.append(("export", kwargs))
        if self.export_error: raise self.export_error
    def Close(self, SaveChanges=False):
        self.log.append(("close", SaveChanges))
        if self.close_error: raise self.close_error


class FakeDocuments:
    def __init__(self, log, doc): self.log=log; self.doc=doc
    def Open(self, **kwargs): self.log.append(("open", kwargs)); return self.doc


class FakeApplication:
    def __init__(self, log, doc): self.Documents=FakeDocuments(log,doc)


def test_export_contract_opens_read_only_exports_pdf_and_closes_without_save(tmp_path):
    log=[]; doc=FakeDocument(log); app=FakeApplication(log,doc)
    source=tmp_path/"source.docx"; source.write_bytes(b"x")
    pdf=tmp_path/"out.pdf"
    _export_with_application(app, source, pdf)
    assert log[0][0] == "open"
    assert log[0][1]["ReadOnly"] is True
    assert log[0][1]["AddToRecentFiles"] is False
    assert log[1][0] == "export"
    assert log[1][1]["ExportFormat"] == 17
    assert log[-1] == ("close", False)


def test_export_failure_is_not_overwritten_by_cleanup_failure(tmp_path):
    log=[]
    export_error=RuntimeError("real export failure")
    doc=FakeDocument(log, export_error=export_error, close_error=RuntimeError("cleanup failure"))
    app=FakeApplication(log,doc)
    source=tmp_path/"source.docx"; source.write_bytes(b"x")
    with pytest.raises(RuntimeError, match="real export failure"):
        _export_with_application(app, source, tmp_path/"out.pdf")
