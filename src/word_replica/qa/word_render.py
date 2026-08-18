from __future__ import annotations

from contextlib import suppress
from pathlib import Path

from word_replica.renderers.word_ownership import clear_owned_word, record_owned_word, word_process_pids

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


def _export_pair_with_application(
    application,
    source_docx_path: Path,
    source_pdf_path: Path,
    output_docx_path: Path,
    output_pdf_path: Path,
) -> tuple[Path, Path]:
    source_pdf = _export_with_application(application, source_docx_path, source_pdf_path)
    output_pdf = _export_with_application(application, output_docx_path, output_pdf_path)
    return source_pdf, output_pdf


def export_docx_to_pdf_with_word(docx_path: Path, pdf_path: Path, visible: bool = False) -> Path:
    import pythoncom
    import win32com.client

    application = None
    owned_word_pid = None
    existing_word_pids = word_process_pids()
    pythoncom.CoInitialize()
    try:
        application = win32com.client.DispatchEx("Word.Application")
        owned_word_pid = record_owned_word(
            application, role="pdf-export", existing_word_pids=existing_word_pids
        )
        application.Visible = bool(visible)
        with suppress(Exception):
            application.DisplayAlerts = 0
        return _export_with_application(application, Path(docx_path), Path(pdf_path))
    finally:
        if application is not None:
            with suppress(Exception):
                application.Quit()
        clear_owned_word(owned_word_pid)
        with suppress(Exception):
            pythoncom.CoUninitialize()


def export_docx_pair_to_pdf_with_word(
    source_docx_path: Path,
    source_pdf_path: Path,
    output_docx_path: Path,
    output_pdf_path: Path,
    visible: bool = False,
) -> tuple[Path, Path]:
    import pythoncom
    import win32com.client

    application = None
    owned_word_pid = None
    existing_word_pids = word_process_pids()
    pythoncom.CoInitialize()
    try:
        application = win32com.client.DispatchEx("Word.Application")
        owned_word_pid = record_owned_word(
            application, role="pdf-pair-export", existing_word_pids=existing_word_pids
        )
        application.Visible = bool(visible)
        with suppress(Exception):
            application.DisplayAlerts = 0
        return _export_pair_with_application(
            application,
            Path(source_docx_path),
            Path(source_pdf_path),
            Path(output_docx_path),
            Path(output_pdf_path),
        )
    finally:
        if application is not None:
            with suppress(Exception):
                application.Quit()
        clear_owned_word(owned_word_pid)
        with suppress(Exception):
            pythoncom.CoUninitialize()
