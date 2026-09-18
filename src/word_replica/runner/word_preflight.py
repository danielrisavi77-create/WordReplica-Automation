"""Functional Microsoft Word preflight for the one-shot Lekta runner."""
from __future__ import annotations

import sys
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any


class WordPreflightError(RuntimeError):
    """Microsoft Word cannot safely create and save the runner's document."""


@dataclass(frozen=True, slots=True)
class WordPreflightResult:
    can_create_and_save: bool


def run_word_preflight(
    *,
    temp_parent: Path | None = None,
    platform_name: str | None = None,
    dispatch_factory: Callable[[str], Any] | None = None,
    co_initialize: Callable[[], None] | None = None,
    co_uninitialize: Callable[[], None] | None = None,
) -> WordPreflightResult:
    if (platform_name or sys.platform) != "win32":
        raise WordPreflightError("Microsoft Word desktop requires Windows")

    if dispatch_factory is None or co_initialize is None or co_uninitialize is None:
        import pythoncom
        import win32com.client

        dispatch_factory = dispatch_factory or win32com.client.DispatchEx
        co_initialize = co_initialize or pythoncom.CoInitialize
        co_uninitialize = co_uninitialize or pythoncom.CoUninitialize

    parent = Path(temp_parent).resolve() if temp_parent is not None else None
    application = None
    document = None
    operation_error: Exception | None = None
    cleanup_error: Exception | None = None

    with TemporaryDirectory(prefix="lekta-word-preflight-", dir=parent) as temporary:
        scratch_path = Path(temporary) / "preflight.docx"
        co_initialize()
        try:
            application = dispatch_factory("Word.Application")
            application.Visible = False
            application.DisplayAlerts = 0
            application.AutomationSecurity = 3
            document = application.Documents.Add()
            document.Content.Text = "Lekta Word preflight"
            document.SaveAs2(str(scratch_path), FileFormat=16)
            if not scratch_path.is_file() or scratch_path.stat().st_size == 0:
                raise RuntimeError("Word did not create the scratch DOCX")
        except Exception as exc:
            operation_error = exc
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=0)
                except Exception as exc:
                    cleanup_error = cleanup_error or exc
            if application is not None:
                try:
                    application.Quit()
                except Exception as exc:
                    cleanup_error = cleanup_error or exc
            try:
                co_uninitialize()
            except Exception as exc:
                cleanup_error = cleanup_error or exc

        if operation_error is not None:
            raise WordPreflightError("Microsoft Word cannot create and save a DOCX") from operation_error
        if cleanup_error is not None:
            raise WordPreflightError("Microsoft Word preflight cleanup failed") from cleanup_error

    return WordPreflightResult(can_create_and_save=True)
