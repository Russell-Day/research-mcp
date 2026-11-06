from __future__ import annotations

import json
from pathlib import Path
import subprocess

import pytest

from research_mcp.errors import IntegrationRefused, RefusedUnsafe, VerificationFailed
from research_mcp import operation_process


def test_run_operation_passes_hard_timeout_and_hidden_worker_flag(monkeypatch):
    calls = {}

    def fake_run(command, **kwargs):
        calls["command"] = command
        calls.update(kwargs)
        response = Path(command[4])
        response.write_text(json.dumps({"ok": True, "result": {"done": 1}}), encoding="utf-8")
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(operation_process.subprocess, "run", fake_run)
    result = operation_process.run_operation("edit", {"source": "a.docx"}, timeout_seconds=7)

    assert result == {"done": 1}
    assert calls["timeout"] == 7
    assert calls["check"] is False
    assert calls["stdout"] == subprocess.DEVNULL
    assert calls["stderr"] == subprocess.DEVNULL
    assert calls["text"] is True
    assert calls["env"]["RESEARCH_MCP_OWNED_WORD_FILE"].endswith("word.json")
    expected_flag = subprocess.CREATE_NO_WINDOW if operation_process.os.name == "nt" else 0
    assert calls["creationflags"] == expected_flag


def test_timeout_cleans_up_only_registered_owned_identity(monkeypatch):
    killed = []

    def fake_run(command, **kwargs):
        registry = Path(kwargs["env"]["RESEARCH_MCP_OWNED_WORD_FILE"])
        registry.write_text(json.dumps({"pid": 1234, "created": 55}), encoding="utf-8")
        raise subprocess.TimeoutExpired(command, kwargs["timeout"])

    monkeypatch.setattr(operation_process.subprocess, "run", fake_run)
    monkeypatch.setattr(operation_process, "process_identity_matches", lambda pid, created: (pid, created) == (1234, 55))
    monkeypatch.setattr("research_mcp.word_com.kill_process", lambda pid: killed.append(pid))

    with pytest.raises(IntegrationRefused, match="exceeded 3 seconds"):
        operation_process.run_operation("style", {}, timeout_seconds=3)
    assert killed == [1234]


def test_timeout_refuses_to_kill_reused_pid(monkeypatch):
    killed = []

    def fake_run(command, **kwargs):
        registry = Path(kwargs["env"]["RESEARCH_MCP_OWNED_WORD_FILE"])
        registry.write_text(json.dumps({"pid": 1234, "created": 55}), encoding="utf-8")
        raise subprocess.TimeoutExpired(command, kwargs["timeout"])

    monkeypatch.setattr(operation_process.subprocess, "run", fake_run)
    monkeypatch.setattr(operation_process, "process_identity_matches", lambda pid, created: False)
    monkeypatch.setattr("research_mcp.word_com.kill_process", lambda pid: killed.append(pid))

    with pytest.raises(IntegrationRefused):
        operation_process.run_operation("render", {}, timeout_seconds=1)
    assert killed == []


def test_worker_error_maps_known_exception_types(monkeypatch):
    def fake_run(command, **kwargs):
        Path(command[4]).write_text(
            json.dumps({"ok": False, "error": {"type": "RefusedUnsafe", "message": "bad package"}}),
            encoding="utf-8",
        )
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(operation_process.subprocess, "run", fake_run)
    with pytest.raises(RefusedUnsafe, match="bad package"):
        operation_process.run_operation("edit", {})

    def fake_verify(command, **kwargs):
        Path(command[4]).write_text(
            json.dumps({"ok": False, "error": {"type": "VerificationFailed", "message": "invariant"}}),
            encoding="utf-8",
        )
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(operation_process.subprocess, "run", fake_verify)
    with pytest.raises(VerificationFailed, match="invariant"):
        operation_process.run_operation("style", {})


def test_unknown_kind_and_invalid_timeout_are_rejected():
    with pytest.raises(ValueError, match="Unknown isolated operation"):
        operation_process.run_operation("other", {})
    with pytest.raises(ValueError, match="positive"):
        operation_process.run_operation("edit", {}, timeout_seconds=0)


def test_real_stalled_subprocess_is_bounded(monkeypatch):
    import sys
    import time
    real_run = subprocess.run
    def stalled(command, **kwargs):
        return real_run([getattr(sys, "_base_executable", sys.executable), "-c", "import time; time.sleep(30)"], **kwargs)
    monkeypatch.setattr(operation_process.subprocess, "run", stalled)
    started = time.monotonic()
    with pytest.raises(IntegrationRefused, match="exceeded"):
        operation_process.run_operation("edit", {}, timeout_seconds=0.2)
    assert time.monotonic() - started < 5


def test_real_worker_dispatch_preserves_refusal_without_starting_word(tmp_path):
    with pytest.raises(RefusedUnsafe, match="unknown style profile"):
        operation_process.run_operation("style", {"source": str(tmp_path / "none.docx"), "profile_name": "invalid"})
