from __future__ import annotations

from contextlib import suppress
from pathlib import Path

WD_EXPORT_FORMAT_PDF = 17


def _export_with_application(application, docx_path: Path, pdf_path: Path) -> Path:
    doc = None
    try:
        doc = application.Documents.Open(
            FileName=str(Path(docx_path).resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        destination = Path(pdf_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        doc.ExportAsFixedFormat(
            OutputFileName=str(destination.resolve()),
            ExportFormat=WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
        )
        return destination
    finally:
        if doc is not None:
            with suppress(Exception):
                doc.Close(SaveChanges=False)


def export_docx_to_pdf_with_word(docx_path: Path, pdf_path: Path, visible: bool = False) -> Path:
    import pythoncom
    import win32com.client

    application = None
    pythoncom.CoInitialize()
    try:
        application = win32com.client.DispatchEx("Word.Application")
        application.Visible = bool(visible)
        with suppress(Exception):
            application.DisplayAlerts = 0
        return _export_with_application(application, Path(docx_path), Path(pdf_path))
    finally:
        if application is not None:
            with suppress(Exception):
                application.Quit()
        with suppress(Exception):
            pythoncom.CoUninitialize()
