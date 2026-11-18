"""Narrow Windows process identity helpers; a PID alone is not an identity."""

from __future__ import annotations

import ctypes
from ctypes import wintypes
from functools import cache

ERROR_INVALID_PARAMETER = 87  # OpenProcess: no process has this PID
IDENTITY_MATCH = "match"
IDENTITY_GONE = "gone"
IDENTITY_UNKNOWN = "unknown"
PROCESS_TERMINATE = 0x0001
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
SYNCHRONIZE = 0x00100000
STILL_ACTIVE = 259
WAIT_OBJECT_0 = 0


def _handle_creation_time(kernel32, handle) -> int | None:
    creation = wintypes.FILETIME()
    exit_time = wintypes.FILETIME()
    kernel = wintypes.FILETIME()
    user = wintypes.FILETIME()
    exit_code = wintypes.DWORD()
    if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
        return None
    if exit_code.value != STILL_ACTIVE:
        return None
    if not kernel32.GetProcessTimes(
        handle, ctypes.byref(creation), ctypes.byref(exit_time), ctypes.byref(kernel), ctypes.byref(user)
    ):
        return None
    return (int(creation.dwHighDateTime) << 32) | int(creation.dwLowDateTime)


def process_creation_time(pid: int) -> int | None:
    if pid <= 0:
        return None
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    try:
        return _handle_creation_time(kernel32, handle)
    finally:
        kernel32.CloseHandle(handle)


def process_identity_matches(pid: int, created: int | None) -> bool:
    return created is not None and process_creation_time(pid) == created


@cache
def _kernel32():
    return ctypes.WinDLL("kernel32", use_last_error=True)


def _handle_identity(kernel32, handle, created: int) -> str:
    exit_code = wintypes.DWORD()
    if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
        return IDENTITY_UNKNOWN
    if exit_code.value != STILL_ACTIVE:
        return IDENTITY_GONE
    current = _handle_creation_time(kernel32, handle)
    if current is None:
        return IDENTITY_UNKNOWN
    return IDENTITY_MATCH if current == created else IDENTITY_GONE


def process_identity_status(pid: int, created: int) -> str:
    """``match``, ``gone`` (no such process, exited, or a reused PID) or ``unknown``.

    ``unknown`` means the process could not be inspected (access denied, a
    failed query): it may well be alive, so callers must not treat it as gone.
    """

    if pid <= 0:
        return IDENTITY_GONE
    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return IDENTITY_GONE if ctypes.get_last_error() == ERROR_INVALID_PARAMETER else IDENTITY_UNKNOWN
    try:
        return _handle_identity(kernel32, handle, created)
    finally:
        kernel32.CloseHandle(handle)


def terminate_process_identity(pid: int, created: int | None, *, timeout: float = 10.0) -> bool:
    """Terminate ``pid`` only if it is still the process created at ``created``.

    The identity is checked on the same handle that is terminated, so a PID
    reused by another process between the check and the kill is never hit.
    Returns True only when that process is confirmed gone (it had already
    exited or been replaced, or it was terminated and its exit observed within
    ``timeout``); False when it is still running or could not be inspected.
    """

    if pid <= 0 or created is None:
        return False
    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(
        PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, False, pid
    )
    if not handle:
        return process_identity_status(pid, created) == IDENTITY_GONE
    try:
        status = _handle_identity(kernel32, handle, created)
        if status != IDENTITY_MATCH:
            return status == IDENTITY_GONE
        if not kernel32.TerminateProcess(handle, 1):
            return False
        return kernel32.WaitForSingleObject(handle, int(timeout * 1000)) == WAIT_OBJECT_0
    finally:
        kernel32.CloseHandle(handle)
