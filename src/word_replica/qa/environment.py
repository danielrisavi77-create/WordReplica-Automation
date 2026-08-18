from __future__ import annotations

from datetime import datetime, timezone
import locale
import platform
import sys
from typing import Any


def _word_application_info() -> dict[str, Any]:
    if sys.platform != "win32":
        return {"available": False, "error": "not Windows"}
    pythoncom = None
    app = None
    try:
        import pythoncom as _pythoncom
        import win32com.client

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Word.Application")
        return {
            "available": True,
            "version": str(getattr(app, "Version", "unknown")),
            "build": str(getattr(app, "Build", "unknown")),
            "name": str(getattr(app, "Name", "Microsoft Word")),
        }
    except Exception as exc:
        return {"available": False, "error": str(exc), "exception_type": type(exc).__name__}
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _enumerate_windows_fonts() -> dict[str, Any]:
    if sys.platform != "win32":
        return {"available": False, "count": 0, "names": []}
    try:
        import winreg

        names: set[str] = set()
        for hive, key_path in (
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        ):
            try:
                with winreg.OpenKey(hive, key_path) as key:
                    index = 0
                    while True:
                        try:
                            name, _value, _type = winreg.EnumValue(key, index)
                        except OSError:
                            break
                        names.add(name.rsplit(" (", 1)[0].strip())
                        index += 1
            except OSError:
                continue
        return {"available": True, "count": len(names), "names": sorted(names)}
    except Exception as exc:
        return {"available": False, "count": 0, "names": [], "error": str(exc)}


def capture_environment_fingerprint() -> dict[str, Any]:
    """Best-effort snapshot of the machine a QA/Golden run was verified on.

    Every sub-probe swallows its own failures so this never raises — a
    fingerprint with partial/error fields is still useful, and QA report
    generation must not be blocked by it.
    """
    loc = locale.getlocale()
    return {
        "captured_at_utc": datetime.now(timezone.utc).isoformat(),
        "os": {
            "platform": platform.platform(),
            "system": platform.system(),
            "release": platform.release(),
            "version": platform.version(),
            "machine": platform.machine(),
        },
        "python": {
            "version": platform.python_version(),
            "implementation": platform.python_implementation(),
            "architecture": platform.architecture()[0],
        },
        "locale": {
            "default": loc[0],
            "encoding": loc[1] or sys.getdefaultencoding(),
        },
        "word": _word_application_info(),
        "fonts": _enumerate_windows_fonts(),
    }
