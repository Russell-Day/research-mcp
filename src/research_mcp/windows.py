"""Read-only top-level window inventory for modal detection."""

from __future__ import annotations

from dataclasses import dataclass
import ctypes
from ctypes import wintypes


@dataclass(frozen=True)
class WindowInfo:
    hwnd: int
    pid: int
    title: str
    class_name: str
    visible: bool


def enumerate_windows() -> tuple[WindowInfo, ...]:
    user32 = ctypes.windll.user32
    windows: list[WindowInfo] = []
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _):
        length = user32.GetWindowTextLengthW(hwnd)
        title_buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title_buffer, length + 1)
        class_buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buffer, 256)
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        windows.append(
            WindowInfo(
                hwnd=int(hwnd),
                pid=int(pid.value),
                title=title_buffer.value,
                class_name=class_buffer.value,
                visible=bool(user32.IsWindowVisible(hwnd)),
            )
        )
        return True

    user32.EnumWindows(callback_type(callback), 0)
    return tuple(windows)


def visible_dialogs_for_pids(pids: set[int]) -> tuple[WindowInfo, ...]:
    return tuple(
        window
        for window in enumerate_windows()
        if window.visible and window.pid in pids and window.title.strip()
    )


def is_suspicious_zotero_window(window: WindowInfo) -> bool:
    """Distinguish integration dialogs from Zotero's main/launcher windows."""

    title = window.title.strip()
    lowered = title.lower()
    if window.class_name == "#32770":
        return True
    if window.class_name == "MozillaWindowClass":
        return False
    if lowered.endswith("\\zotero.exe") or lowered.endswith("/zotero.exe"):
        return False
    return any(token in lowered for token in ("warning", "integration error", "zotero error"))
