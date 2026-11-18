"""A worker whose Word process died is recycled instead of refusing (chk_64874a43)."""

from __future__ import annotations

import ctypes
import json
import os
import subprocess
import sys

import pytest

from research_mcp import word_service
from research_mcp.errors import IntegrationBusy
from research_mcp.processes import (
    process_creation_time,
    process_identity_status,
    terminate_process_identity,
)

OLD = {
    "pid": 100, "process_created": 1, "word_pid": 200, "word_process_created": 2,
    "pipe": r"\\.\pipe\old", "authkey": "00", "started": 0.0,
}
NEW = {**OLD, "pid": 101, "process_created": 3, "word_pid": 201, "word_process_created": 4, "pipe": r"\\.\pipe\new"}


@pytest.fixture
def fake_service(tmp_path, monkeypatch):
    """Patch process identity, termination, ping and spawn around a temp state file."""

    state_file = tmp_path / "word-worker.json"
    monkeypatch.setattr(word_service, "worker_state_path", lambda: state_file)
    alive: set[tuple[int, int]] = set()
    calls: dict = {"terminated": [], "pinged": [], "spawned": []}
    ping_error: dict[str, Exception | None] = {"old": None}

    def identity(pid, created):
        return (pid, created) in alive

    def terminate(pid, created, *, timeout=10.0):
        calls["terminated"].append(pid)
        alive.discard((pid, created))
        return True

    def ping(state, payload, *, timeout=None):
        calls["pinged"].append(state["pid"])
        if state["pid"] == OLD["pid"] and ping_error["old"] is not None:
            raise ping_error["old"]
        return {"ok": True}

    def spawn(*args, **kwargs):
        calls["spawned"].append(args)
        alive.update({(NEW["pid"], NEW["process_created"]), (NEW["word_pid"], NEW["word_process_created"])})
        state_file.write_text(json.dumps(NEW), encoding="utf-8")

    unknown: set[int] = set()

    def status(pid, created):
        if pid in unknown:
            return "unknown"
        return "match" if (pid, created) in alive else "gone"

    monkeypatch.setattr(word_service, "process_identity_matches", identity)
    monkeypatch.setattr(word_service, "process_identity_status", status)
    monkeypatch.setattr(word_service, "terminate_process_identity", terminate)
    monkeypatch.setattr(word_service, "_request_from_state", ping)
    monkeypatch.setattr(word_service.subprocess, "Popen", spawn)
    if sys.platform == "win32":
        monkeypatch.setattr(word_service.subprocess, "STARTUPINFO", lambda: type("S", (), {})())
    calls["uninspectable"] = unknown
    return state_file, alive, calls, ping_error


def test_dead_word_recycles_worker_and_starts_fresh_one(fake_service):
    state_file, alive, calls, ping_error = fake_service
    state_file.write_text(json.dumps(OLD), encoding="utf-8")
    alive.add((OLD["pid"], OLD["process_created"]))  # worker alive, its Word gone
    ping_error["old"] = RuntimeError("(-2147023174, 'The RPC server is unavailable.')")

    state = word_service.ensure_word_worker(timeout=5)

    assert state == NEW
    assert calls["terminated"] == [OLD["pid"]]
    assert OLD["pid"] not in calls["pinged"]
    assert len(calls["spawned"]) == 1


def test_word_dying_between_check_and_ping_still_recycles(fake_service, monkeypatch):
    state_file, alive, calls, _ = fake_service
    state_file.write_text(json.dumps(OLD), encoding="utf-8")
    alive.update({(OLD["pid"], OLD["process_created"]), (OLD["word_pid"], OLD["word_process_created"])})

    def rpc_dies(*args, **kwargs):
        alive.discard((OLD["word_pid"], OLD["word_process_created"]))
        raise RuntimeError("The RPC server is unavailable.")

    real_ping = word_service._request_from_state

    def ping(state, payload, *, timeout=None):
        if state["pid"] == OLD["pid"]:
            rpc_dies()
        return real_ping(state, payload, timeout=timeout)

    monkeypatch.setattr(word_service, "_request_from_state", ping)

    assert word_service.ensure_word_worker(timeout=5) == NEW
    assert calls["terminated"] == [OLD["pid"]]


def test_live_word_with_unresponsive_worker_still_refuses(fake_service):
    state_file, alive, calls, ping_error = fake_service
    state_file.write_text(json.dumps(OLD), encoding="utf-8")
    alive.update({(OLD["pid"], OLD["process_created"]), (OLD["word_pid"], OLD["word_process_created"])})
    ping_error["old"] = IntegrationBusy("busy refreshing")

    with pytest.raises(IntegrationBusy, match="explicit recovery is required"):
        word_service.ensure_word_worker(timeout=5)
    assert calls["terminated"] == []
    assert calls["spawned"] == []
    assert json.loads(state_file.read_text(encoding="utf-8")) == OLD


def test_state_without_word_identity_is_never_recycled(fake_service):
    state_file, alive, calls, ping_error = fake_service
    legacy = {key: value for key, value in OLD.items() if key != "word_process_created"}
    state_file.write_text(json.dumps(legacy), encoding="utf-8")
    alive.add((OLD["pid"], OLD["process_created"]))
    ping_error["old"] = RuntimeError("no answer")

    assert not word_service.worker_word_is_dead(legacy)
    with pytest.raises(IntegrationBusy, match="explicit recovery is required"):
        word_service.ensure_word_worker(timeout=5)
    assert calls["terminated"] == []


def test_worker_that_cannot_be_stopped_refuses_and_keeps_state(fake_service, monkeypatch):
    state_file, alive, calls, _ = fake_service
    state_file.write_text(json.dumps(OLD), encoding="utf-8")
    alive.add((OLD["pid"], OLD["process_created"]))
    monkeypatch.setattr(word_service, "terminate_process_identity", lambda pid, created, **kw: False)

    with pytest.raises(IntegrationBusy, match="lost its Word process 200 and could not be stopped"):
        word_service.ensure_word_worker(timeout=5)
    assert calls["spawned"] == []
    assert state_file.exists()


def test_word_that_cannot_be_inspected_is_never_treated_as_dead(fake_service):
    state_file, alive, calls, ping_error = fake_service
    state_file.write_text(json.dumps(OLD), encoding="utf-8")
    alive.add((OLD["pid"], OLD["process_created"]))
    calls["uninspectable"].add(OLD["word_pid"])  # e.g. access denied
    ping_error["old"] = IntegrationBusy("busy refreshing")

    assert not word_service.worker_word_is_dead(OLD)
    with pytest.raises(IntegrationBusy, match="alive but not responding"):
        word_service.ensure_word_worker(timeout=5)
    assert calls["terminated"] == []
    assert calls["spawned"] == []


def test_shutdown_stops_a_worker_whose_word_died(fake_service):
    state_file, alive, calls, _ = fake_service
    state_file.write_text(json.dumps(OLD), encoding="utf-8")
    alive.add((OLD["pid"], OLD["process_created"]))

    result = word_service.shutdown_word_worker()

    assert result == {"ok": True, "running": False, "orphaned_worker_stopped": True}
    assert calls["terminated"] == [OLD["pid"]]
    assert calls["pinged"] == []
    assert not state_file.exists()


def test_recycle_leaves_a_replaced_state_file_alone(fake_service):
    state_file, alive, calls, _ = fake_service
    state_file.write_text(json.dumps(NEW), encoding="utf-8")  # another caller already replaced it
    alive.add((OLD["pid"], OLD["process_created"]))

    word_service._recycle_orphaned_worker(OLD)

    assert calls["terminated"] == [OLD["pid"]]
    assert json.loads(state_file.read_text(encoding="utf-8")) == NEW


@pytest.mark.skipif(sys.platform != "win32", reason="Windows process identity")
def test_process_identity_status_separates_gone_from_unknown():
    created = process_creation_time(os.getpid())
    assert process_identity_status(os.getpid(), created) == "match"
    assert process_identity_status(os.getpid(), created + 1) == "gone"
    child = subprocess.Popen([sys.executable, "-c", "pass"])
    child.wait(timeout=30)
    assert process_identity_status(child.pid, 1) == "gone"
    # A protected system process cannot be opened by a normal user: unknown, not gone.
    if not ctypes.windll.shell32.IsUserAnAdmin():
        assert process_identity_status(_protected_pid(), 1) == "unknown"


def _protected_pid() -> int:
    output = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq lsass.exe", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, check=True,
    ).stdout
    return int(output.split(",")[1].strip('"'))


@pytest.mark.skipif(sys.platform != "win32", reason="Windows process identity")
def test_terminate_process_identity_checks_identity_on_the_handle():
    child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
    try:
        created = process_creation_time(child.pid)
        assert created is not None
        # A wrong creation time (a reused PID) is never terminated.
        assert terminate_process_identity(child.pid, created + 1) is True
        assert child.poll() is None
        assert terminate_process_identity(child.pid, created) is True
        assert child.wait(timeout=10) == 1
        # Already gone: reported as stopped, nothing else touched.
        assert terminate_process_identity(child.pid, created) is True
    finally:
        if child.poll() is None:
            child.kill()
