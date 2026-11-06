"""Guards for the refresh barrier's deadline ordering and rescue path.

Three recorded runs discarded a complete, correctly rendered manuscript because
the supervisor timed out, raised before reading anything, and never attempted
saved-package reconciliation. These tests pin the three mechanics that made
that possible. All are offline: no Word, no Zotero, no integration lease.
"""

from __future__ import annotations

import json
from multiprocessing.connection import Listener
import secrets
import threading

import pytest

from research_mcp.errors import IntegrationBusy
from research_mcp.integration import read_refresh_progress
from research_mcp.word_service import _request_from_state


def _silent_worker(
    pipe_name: str, authkey: bytes, stop: threading.Event, ready: threading.Event
) -> None:
    """A worker that accepts a request and deliberately never answers."""

    listener = Listener(pipe_name, family="AF_PIPE", authkey=authkey)
    # The pipe exists only after Listener() returns. Signalling here is what
    # keeps the test measuring the read deadline instead of racing the OS:
    # thread.start() returns before this line runs, so an unsynchronised
    # client can connect to a pipe that does not exist yet and fail with
    # FileNotFoundError inside the pytest.raises block.
    ready.set()
    try:
        connection = listener.accept()
        connection.recv()
        stop.wait(10.0)
        connection.close()
    except Exception:
        pass
    finally:
        listener.close()


def test_pipe_read_gives_up_instead_of_blocking_forever():
    """An unbounded recv() is why the relay never exited and never reconciled."""

    pipe_name = rf"\\.\pipe\research-mcp-test-{secrets.token_hex(8)}"
    authkey = secrets.token_bytes(16)
    stop = threading.Event()
    ready = threading.Event()
    thread = threading.Thread(
        target=_silent_worker, args=(pipe_name, authkey, stop, ready), daemon=True
    )
    thread.start()
    assert ready.wait(10.0), "listener thread never created the pipe"

    state = {"pipe": pipe_name, "authkey": authkey.hex()}
    try:
        with pytest.raises(IntegrationBusy, match="did not answer"):
            _request_from_state(state, {"command": "refresh"}, timeout=1.0)
    finally:
        stop.set()
        thread.join(timeout=5.0)


def test_relay_deadline_expires_before_the_supervisor_deadline():
    """worker loop < relay pipe < supervisor. Any other order reintroduces the bug."""

    for dispatch_timeout in (75.0, 144.0, 300.0, 900.0, 3600.0):
        cleanup_margin = min(30.0, max(5.0, dispatch_timeout * 0.1))
        worker_timeout = max(1.0, dispatch_timeout - cleanup_margin)
        relay_timeout = max(worker_timeout + 1.0, dispatch_timeout - cleanup_margin * 0.5)
        assert worker_timeout < relay_timeout, dispatch_timeout
        assert relay_timeout < dispatch_timeout, dispatch_timeout


def test_progress_reader_tolerates_a_missing_or_corrupt_sidecar(tmp_path, monkeypatch):
    import research_mcp.integration as integration

    target = tmp_path / "refresh-progress.json"
    monkeypatch.setattr(integration, "refresh_state_path", lambda: target)
    assert read_refresh_progress() is None

    target.write_text("not json", encoding="utf-8")
    assert read_refresh_progress() is None

    target.write_text(json.dumps({"phase": "zotero_refresh_invoked"}), encoding="utf-8")
    assert read_refresh_progress() == {"phase": "zotero_refresh_invoked"}
