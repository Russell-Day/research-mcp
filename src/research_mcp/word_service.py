"""Lifecycle and authenticated IPC for the project-owned persistent Word worker."""

from __future__ import annotations

import json
import os
from multiprocessing.connection import Client
from pathlib import Path
import secrets
import subprocess
import sys
import time

from .errors import IntegrationBusy
from .processes import process_identity_matches


# A ping must outlast ordinary COM latency but must not wait behind a
# refresh: the worker serves one request at a time, so a ping that blocks
# means the worker is busy or wedged, and a clean refusal beats a hang.
PING_TIMEOUT = 30.0


def service_root() -> Path:
    root = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "research_mcp"
    root.mkdir(parents=True, exist_ok=True)
    return root


def worker_state_path() -> Path:
    return service_root() / "word-worker.json"


def worker_pipe() -> str:
    return rf"\\.\pipe\research-mcp-word-{os.getlogin()}"


def refresh_state_path() -> Path:
    """Where the worker records live refresh progress.

    `Application.Run("ZoteroRefresh")` is a synchronous COM call, so nothing
    inside the worker can report progress while it is executing and nothing
    the worker writes to stderr is readable while it is still running. A
    sidecar file is the only channel that survives a supervisor timeout.
    """

    return service_root() / "refresh-progress.json"


def read_worker_state() -> dict | None:
    try:
        state = json.loads(worker_state_path().read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError, OSError):
        return None
    return state if isinstance(state, dict) else None


def worker_state_is_live(state: dict | None) -> bool:
    if not state:
        return False
    try:
        return process_identity_matches(int(state["pid"]), int(state["process_created"]))
    except (KeyError, TypeError, ValueError):
        return False


def _request_from_state(state: dict, payload: dict, *, timeout: float | None = None) -> dict:
    connection = Client(state["pipe"], family="AF_PIPE", authkey=bytes.fromhex(state["authkey"]))
    try:
        connection.send(payload)
        # An unbounded recv() was the reason a wedged worker never surfaced:
        # the relay blocked here forever, so the supervisor's "helper exited
        # non-zero" reconciliation path could never run and the worker's own
        # deadline error was never delivered.
        if timeout is not None and not connection.poll(timeout):
            raise IntegrationBusy(
                f"Word worker did not answer {payload.get('command')!r} within {timeout}s; "
                "the worker was left running for explicit recovery"
            )
        response = connection.recv()
    finally:
        connection.close()
    if not isinstance(response, dict):
        raise RuntimeError("Word worker returned a malformed response")
    if not response.get("ok"):
        raise RuntimeError(str(response.get("error", "Word worker failed")))
    return response


def ensure_word_worker(*, timeout: float = 25.0) -> dict:
    state = read_worker_state()
    if worker_state_is_live(state):
        try:
            _request_from_state(state, {"command": "ping"}, timeout=PING_TIMEOUT)
            return state
        except Exception as exc:
            raise IntegrationBusy(
                "the recorded Word worker is alive but not responding; explicit recovery is required"
            ) from exc
    worker_state_path().unlink(missing_ok=True)
    authkey = secrets.token_hex(32)
    startupinfo = None
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.lpDesktop = r"WinSta0\Default"
    subprocess.Popen(
        [sys.executable, "-m", "research_mcp.word_worker", worker_pipe(), authkey],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
        startupinfo=startupinfo,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        state = read_worker_state()
        if worker_state_is_live(state):
            try:
                _request_from_state(state, {"command": "ping"}, timeout=PING_TIMEOUT)
                return state
            except Exception:
                pass
        time.sleep(0.2)
    raise RuntimeError("project-owned Word worker did not become ready")


def request_word_worker(payload: dict, *, timeout: float | None = None) -> dict:
    return _request_from_state(ensure_word_worker(), payload, timeout=timeout)


def shutdown_word_worker() -> dict:
    state = read_worker_state()
    if not state:
        return {"ok": True, "running": False}
    if not worker_state_is_live(state):
        worker_state_path().unlink(missing_ok=True)
        return {"ok": True, "running": False, "stale_state_removed": True}
    response = _request_from_state(state, {"command": "shutdown"}, timeout=PING_TIMEOUT)
    deadline = time.monotonic() + 25.0
    while time.monotonic() < deadline and worker_state_is_live(state):
        time.sleep(0.25)
    if worker_state_is_live(state):
        raise IntegrationBusy("Word worker did not shut down cleanly; it was not terminated")
    worker_state_path().unlink(missing_ok=True)
    return {**response, "running": False}
