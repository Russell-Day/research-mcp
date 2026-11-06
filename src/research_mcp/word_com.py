"""Low-level Windows COM initialization and Word process helpers."""

from __future__ import annotations

import os
import signal
import pythoncom
import win32com.client
import win32process

from .constants import WD_ALERTS_NONE, WD_DO_NOT_SAVE_CHANGES
from .processes import process_creation_time


def ensure_default_desktop() -> bool:
    """Ensure the calling thread is attached to the interactive Default desktop.

    On Windows, when running inside services or sandboxes (e.g. CI, harness subshells),
    the process thread may be attached to a non-default desktop. Zotero runs on the
    interactive 'Default' desktop. If Word is created on a separate desktop, Zotero's
    VBA macro FindWindow("MozillaWindowClass", ...) cannot locate Zotero, causing
    an invisible modal dialog that hangs COM automation.
    """
    try:
        import ctypes

        user32 = ctypes.windll.user32
        hwinsta = user32.OpenWindowStationW("WinSta0", False, 0x37F)
        if hwinsta:
            user32.SetProcessWindowStation(hwinsta)
            hdesk = user32.OpenDesktopW("Default", 0, False, 0x1FF)
            if hdesk:
                user32.SetThreadDesktop(hdesk)
                return True
    except Exception:
        pass
    return False


def co_initialize() -> None:
    ensure_default_desktop()
    pythoncom.CoInitialize()


def co_uninitialize() -> None:
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


def create_word_application(*, visible: bool = False):
    ensure_default_desktop()
    word = win32com.client.DispatchEx("Word.Application")
    try:
        word.Visible = visible
        word.DisplayAlerts = WD_ALERTS_NONE
        word.ScreenUpdating = False
        # Prevent document-embedded macros from opening dialogs during unattended
        # file opens. Installed global-template commands remain separate.
        word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    except Exception:
        # The caller cannot clean up an instance that was never returned.
        try:
            word.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            pass
        raise
    return word


def word_pid(word) -> int | None:
    try:
        hwnd = int(word.Hwnd)
        return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception:
        pass
    try:
        control = word.Documents.Add(Visible=False)
        try:
            return int(win32process.GetWindowThreadProcessId(int(control.ActiveWindow.Hwnd))[1])
        finally:
            control.Close(WD_DO_NOT_SAVE_CHANGES)
    except Exception:
        return None


def process_exists(pid: int) -> bool:
    return process_creation_time(pid) is not None


def kill_process(pid: int) -> None:
    try:
        os.kill(pid, signal.SIGTERM)
    except OSError:
        pass


def com_error_code(exc: Exception) -> int | None:
    args = getattr(exc, "args", None)
    if args and isinstance(args[0], int):
        return args[0]
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        return hresult
    return None
