from __future__ import annotations

from contextlib import suppress
from pathlib import Path

from word_replica.renderers.word_ownership import clear_owned_word, record_owned_word, word_process_pids

WD_EXPORT_FORMAT_PDF = 17


def _clear_owned_word_after_confirmed_exit(pid: int | None, *, quit_succeeded: bool) -> None:
    if pid is None or not quit_succeeded:
        return
    try:
        if pid in word_process_pids():
            return
    except Exception:
        return
    clear_owned_word(pid)


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
        quit_succeeded = application is None
        if application is not None:
            try:
                application.Quit()
                quit_succeeded = True
            except Exception:
                quit_succeeded = False
        _clear_owned_word_after_confirmed_exit(owned_word_pid, quit_succeeded=quit_succeeded)
        with suppress(Exception):
            pythoncom.CoUninitialize()


def _document_compatibility_mode(application, docx_path: Path) -> int | None:
    doc = None
    try:
        doc = application.Documents.Open(
            FileName=str(Path(docx_path).resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        return int(doc.CompatibilityMode)
    except Exception:
        return None
    finally:
        if doc is not None:
            with suppress(Exception):
                doc.Close(SaveChanges=False)


def read_docx_pair_compatibility_mode(
    source_docx_path: Path,
    output_docx_path: Path,
) -> dict[str, int | None]:
    import pythoncom
    import win32com.client

    application = None
    owned_word_pid = None
    existing_word_pids = word_process_pids()
    pythoncom.CoInitialize()
    try:
        application = win32com.client.DispatchEx("Word.Application")
        owned_word_pid = record_owned_word(
            application, role="compatibility-mode-probe", existing_word_pids=existing_word_pids
        )
        application.Visible = False
        with suppress(Exception):
            application.DisplayAlerts = 0
        return {
            "source": _document_compatibility_mode(application, Path(source_docx_path)),
            "output": _document_compatibility_mode(application, Path(output_docx_path)),
        }
    finally:
        quit_succeeded = application is None
        if application is not None:
            try:
                application.Quit()
                quit_succeeded = True
            except Exception:
                quit_succeeded = False
        _clear_owned_word_after_confirmed_exit(owned_word_pid, quit_succeeded=quit_succeeded)
        with suppress(Exception):
            pythoncom.CoUninitialize()


def _open_and_repair_signature(application, docx_path: Path, *, open_and_repair: bool) -> str:
    doc = None
    try:
        doc = application.Documents.Open(
            FileName=str(Path(docx_path).resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=open_and_repair,
        )
        return str(doc.Content.Text)
    finally:
        if doc is not None:
            with suppress(Exception):
                doc.Close(SaveChanges=False)


def _open_and_repair_changed_content(application, docx_path: Path) -> bool:
    """True if forcing OpenAndRepair yields different content than a normal
    open — evidence Word silently repaired latent corruption in the file."""
    normal = _open_and_repair_signature(application, docx_path, open_and_repair=False)
    repaired = _open_and_repair_signature(application, docx_path, open_and_repair=True)
    return normal != repaired


def detect_open_and_repair(docx_path: Path) -> bool | None:
    """Best-effort real-Word check for src/word_replica/repair_contract's
    requireOpenAndRepairFalse policy. Returns None (not True or False) if
    the probe itself could not complete, e.g. no Word available — callers
    must treat None as "not verified", never silently as a pass.
    """
    import pythoncom
    import win32com.client

    application = None
    owned_word_pid = None
    existing_word_pids = word_process_pids()
    pythoncom.CoInitialize()
    try:
        application = win32com.client.DispatchEx("Word.Application")
        owned_word_pid = record_owned_word(
            application, role="open-and-repair-probe", existing_word_pids=existing_word_pids
        )
        application.Visible = False
        with suppress(Exception):
            application.DisplayAlerts = 0
        try:
            return _open_and_repair_changed_content(application, Path(docx_path))
        except Exception:
            return None
    except Exception:
        return None
    finally:
        quit_succeeded = application is None
        if application is not None:
            try:
                application.Quit()
                quit_succeeded = True
            except Exception:
                quit_succeeded = False
        _clear_owned_word_after_confirmed_exit(owned_word_pid, quit_succeeded=quit_succeeded)
        with suppress(Exception):
            pythoncom.CoUninitialize()


def _fields_update_leaves_content_unchanged(application, docx_path: Path) -> bool:
    doc = None
    try:
        doc = application.Documents.Open(
            FileName=str(Path(docx_path).resolve()),
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False,
        )
        before = str(doc.Content.Text)
        doc.Fields.Update()
        after = str(doc.Content.Text)
        return before == after
    finally:
        if doc is not None:
            with suppress(Exception):
                doc.Close(SaveChanges=False)


def check_fields_update_equality(docx_path: Path) -> bool | None:
    """Best-effort real-Word check for requireFieldsUpdateEquality: do the
    document's cached field results already match what Fields.Update()
    would recompute? Returns None (not verified) if the probe could not
    complete; callers must not treat None as a pass.
    """
    import pythoncom
    import win32com.client

    application = None
    owned_word_pid = None
    existing_word_pids = word_process_pids()
    pythoncom.CoInitialize()
    try:
        application = win32com.client.DispatchEx("Word.Application")
        owned_word_pid = record_owned_word(
            application, role="fields-update-probe", existing_word_pids=existing_word_pids
        )
        application.Visible = False
        with suppress(Exception):
            application.DisplayAlerts = 0
        try:
            return _fields_update_leaves_content_unchanged(application, Path(docx_path))
        except Exception:
            return None
    except Exception:
        return None
    finally:
        quit_succeeded = application is None
        if application is not None:
            try:
                application.Quit()
                quit_succeeded = True
            except Exception:
                quit_succeeded = False
        _clear_owned_word_after_confirmed_exit(owned_word_pid, quit_succeeded=quit_succeeded)
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
        quit_succeeded = application is None
        if application is not None:
            try:
                application.Quit()
                quit_succeeded = True
            except Exception:
                quit_succeeded = False
        _clear_owned_word_after_confirmed_exit(owned_word_pid, quit_succeeded=quit_succeeded)
        with suppress(Exception):
            pythoncom.CoUninitialize()
