from __future__ import annotations

import json

import pytest

from research_mcp.errors import IntegrationBusy
from research_mcp.integration import clear_integration_lock, integration_status


def test_checked_recovery_requires_confirmation_and_dead_owner(tmp_path, monkeypatch):
    lock = tmp_path / "integration.lock"
    lock.write_text(json.dumps({"pid": 123, "reason": "failed"}))
    monkeypatch.setattr("research_mcp.integration.process_exists", lambda pid: False)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    with pytest.raises(IntegrationBusy, match="confirmation"):
        clear_integration_lock(lock)
    result = clear_integration_lock(lock, confirmed=True)
    assert result["cleared"]
    assert not lock.exists()


def test_checked_recovery_refuses_running_owner(tmp_path, monkeypatch):
    lock = tmp_path / "integration.lock"
    lock.write_text(json.dumps({"pid": 123}))
    monkeypatch.setattr("research_mcp.integration.process_exists", lambda pid: True)
    with pytest.raises(IntegrationBusy, match="legacy lock PID"):
        clear_integration_lock(lock, confirmed=True)
    assert integration_status(lock)["locked"]


def test_checked_recovery_refuses_matching_process_identity(tmp_path, monkeypatch):
    lock = tmp_path / "integration.lock"
    lock.write_text(json.dumps({"pid": 123, "process_created": 456}))
    monkeypatch.setattr("research_mcp.integration.process_exists", lambda pid: True)
    monkeypatch.setattr("research_mcp.integration.process_identity_matches", lambda pid, created: True)
    with pytest.raises(IntegrationBusy, match="live process"):
        clear_integration_lock(lock, confirmed=True)


def test_checked_recovery_refuses_invalid_partial_lock(tmp_path, monkeypatch):
    lock = tmp_path / "integration.lock"
    lock.write_text("")
    with pytest.raises(IntegrationBusy, match="invalid or partially written"):
        clear_integration_lock(lock, confirmed=True)
    assert lock.exists()


def test_lease_refuses_to_delete_changed_ownership(tmp_path):
    from research_mcp.integration import IntegrationLease

    lock = tmp_path / "integration.lock"
    lease = IntegrationLease.acquire(lock)
    lock.write_text(json.dumps({"pid": 999, "lease_id": "different"}))
    with pytest.raises(IntegrationBusy, match="ownership changed"):
        lease.release_success()
    assert lock.exists()
