"""Where the tool-call log goes: a source checkout keeps <repo>/logs, an installed copy uses
%LOCALAPPDATA%\\research_mcp\\logs, RESEARCH_MCP_LOG_DIR (environment or config.env) overrides both."""

from __future__ import annotations

from pathlib import Path

import pytest

from research_mcp import config, logger


@pytest.fixture
def installed(monkeypatch, tmp_path):
    """Pretend the package is installed (no git checkout around it), with LOCALAPPDATA under tmp_path."""
    monkeypatch.delenv("RESEARCH_MCP_LOG_DIR", raising=False)
    monkeypatch.delenv("RESEARCH_MCP_HOME", raising=False)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    monkeypatch.setattr(logger, "_checkout_root", lambda: None)
    return tmp_path / "local" / "research_mcp" / "logs"


def test_source_checkout_keeps_repo_logs(monkeypatch):
    monkeypatch.delenv("RESEARCH_MCP_LOG_DIR", raising=False)
    root = logger._checkout_root()
    if root is None:
        pytest.skip("not running from a git checkout (e.g. an unpacked sdist)")
    assert (root / "src" / "research_mcp" / "logger.py").exists()
    assert logger._default_log_dir() == root / "logs"


def test_installed_copy_logs_under_localappdata(installed):
    assert logger.get_log_dir() == installed.resolve()
    logger.log_tool_call("inspect_document", {"path": "x.docx"}, {"ok": True})
    assert (installed / "mcp_traffic.jsonl").read_text(encoding="utf-8").count("\n") == 1


def test_log_dir_from_config_file(installed, tmp_path, isolated_config):
    target = tmp_path / "custom-logs"
    isolated_config.write_text(f"RESEARCH_MCP_LOG_DIR={target}\n", encoding="utf-8")
    config.reload()
    assert logger.get_log_dir() == target.resolve()


def test_default_log_rolls_over_once(installed, monkeypatch):
    monkeypatch.setattr(logger, "ROLLOVER_BYTES", 200)
    for _ in range(10):
        logger.log_tool_call("inspect_document", {"path": "x" * 50}, {"ok": True})
    files = sorted(p.name for p in installed.iterdir())
    assert files == ["mcp_traffic.1.jsonl", "mcp_traffic.jsonl"]
    assert (installed / "mcp_traffic.jsonl").stat().st_size < 400


def test_configured_log_never_rolls_over(monkeypatch, tmp_path):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path / "mine"))
    monkeypatch.setattr(logger, "ROLLOVER_BYTES", 200)
    for _ in range(10):
        logger.log_tool_call("inspect_document", {"path": "x" * 50}, {"ok": True})
    assert [p.name for p in (tmp_path / "mine").iterdir()] == ["mcp_traffic.jsonl"]


def test_report_issue_never_raises_when_the_log_folder_cannot_be_made(monkeypatch):
    def broken() -> Path:
        raise PermissionError("read-only install folder")

    monkeypatch.setattr(logger, "get_log_dir", broken)
    result = logger.record_agent_issue("title", "details")
    assert result["ok"] is False and "read-only install folder" in result["message"]
