"""Zotero integration recovery.

Zotero binds a Word integration session to a Word *application instance*. If
that instance dies uncleanly while a session is open, Zotero keeps the dead
reference and raises "Zotero experienced an error updating your document" on
every subsequent command. No amount of lock clearing, fresh output paths, or
document repair helps; only restarting Zotero clears it.

Two debugging cycles were lost to this because the damage was self-inflicted
during cleanup and looked exactly like a code defect. So this module makes the
condition explicit rather than mysterious:

- `mark_zotero_suspect` records that we terminated a Word that may have held a
  session, so the *next* run knows why it is failing;
- `recover_zotero` performs the documented restart, gracefully; and
- `zotero_is_suspect` lets a preflight refuse to start rather than produce a
  confusing modal.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import time

from .zotero import ZoteroClient

ZOTERO_EXE = Path(r"C:\Program Files\Zotero\zotero.exe")


def state_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "research_mcp"
    base.mkdir(parents=True, exist_ok=True)
    return base / "zotero-suspect.json"


def mark_zotero_suspect(reason: str) -> None:
    """Call whenever a Word instance is terminated rather than quit cleanly."""
    state_path().write_text(
        json.dumps({"reason": reason, "at": time.time()}), encoding="utf8"
    )


def clear_zotero_suspect() -> None:
    try:
        state_path().unlink()
    except FileNotFoundError:
        pass


def zotero_is_suspect() -> dict | None:
    try:
        return json.loads(state_path().read_text(encoding="utf8"))
    except (FileNotFoundError, ValueError):
        return None


def _zotero_processes() -> list[int]:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq zotero.exe", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
    ).stdout
    pids = []
    for line in out.strip().splitlines():
        parts = [p.strip('"') for p in line.split(",")]
        if len(parts) > 1 and parts[1].isdigit():
            pids.append(int(parts[1]))
    return pids


def dismiss_zotero_dialogs() -> list[str]:
    """Close Zotero modal dialogs. Mozilla dialogs expose no child buttons, so
    WM_CLOSE is the only reliable route; it is the cancel path, never 'Yes'."""
    import ctypes

    from .windows import is_suspicious_zotero_window, visible_dialogs_for_pids
    from .integration import zotero_pids

    closed = []
    for window in visible_dialogs_for_pids(zotero_pids()):
        if window.class_name == "MozillaDialogClass" or is_suspicious_zotero_window(window):
            ctypes.windll.user32.SendMessageW(window.hwnd, 0x0010, 0, 0)
            closed.append(window.title)
    if closed:
        time.sleep(1.5)
    return closed


def wait_for_zotero_api(timeout: float = 90.0) -> bool:
    client = ZoteroClient()
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            if client.ping():
                return True
        except Exception:
            pass
        time.sleep(1.0)
    return False


def recover_zotero(*, timeout: float = 90.0) -> dict:
    """Dismiss dialogs, stop Zotero gracefully, restart it, wait for its API.

    Graceful first: post WM_CLOSE to the one verified Zotero main window. /F is
    used only if Zotero does not exit within the deadline.
    """
    report: dict = {"dismissed": dismiss_zotero_dialogs()}

    before = _zotero_processes()
    report["pids_before"] = before

    from .windows import enumerate_windows
    main_windows = [
        window for window in enumerate_windows()
        if window.pid in set(before)
        and window.visible
        and window.class_name == "MozillaWindowClass"
        and window.title.strip().lower().endswith(" - zotero")
    ]
    report["graceful_window_count"] = len(main_windows)
    if len(main_windows) == 1:
        import ctypes
        ctypes.windll.user32.PostMessageW(main_windows[0].hwnd, 0x0010, 0, 0)

    deadline = time.monotonic() + 25.0
    while time.monotonic() < deadline and _zotero_processes():
        time.sleep(0.5)

    forced = []
    for pid in _zotero_processes():
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, text=True)
        forced.append(pid)
    report["forced"] = forced

    deadline = time.monotonic() + 15.0
    while time.monotonic() < deadline and _zotero_processes():
        time.sleep(0.5)

    if not ZOTERO_EXE.exists():
        report["relaunched"] = False
        report["error"] = f"{ZOTERO_EXE} not found"
        return report

    subprocess.Popen([str(ZOTERO_EXE)], close_fds=True)
    report["relaunched"] = True
    report["api_ready"] = wait_for_zotero_api(timeout)
    report["pids_after"] = _zotero_processes()
    if report["api_ready"]:
        clear_zotero_suspect()
    return report
