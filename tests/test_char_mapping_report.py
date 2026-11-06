"""report_issue: every report is logged locally; the result is a prefilled GitHub issue link by default, or,
with RESEARCH_MCP_ISSUES=agent, the improvement + file_with steps (plan 551 S4)."""

import asyncio
import json
from urllib.parse import parse_qs, urlsplit

import pytest

from research_mcp import server, textclean


def _report(**kw):
    out = server.report_issue(**kw)
    return out if isinstance(out, dict) else out.model_dump()


@pytest.fixture
def agent_mode(monkeypatch):
    monkeypatch.setenv("RESEARCH_MCP_ISSUES", "agent")


def _query(url: str) -> dict[str, str]:
    parts = urlsplit(url)
    assert f"{parts.scheme}://{parts.netloc}{parts.path}" == server.GITHUB_ISSUES
    return {key: values[0] for key, values in parse_qs(parts.query).items()}


def test_char_mapping_report_builds_improvement(tmp_path, monkeypatch, agent_mode):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    table_before = textclean.MAP_PATH.read_bytes()
    out = _report(title="NBSP in units", details="10 mm must not wrap", category="char_mapping",
                  char="U+00A0", should_be="keep", snippet="width 10\u00a0mm")
    assert out["ok"] is True
    imp = out["improvement"]
    assert imp == {"text": "Char mapping: U+00A0 NO-BREAK SPACE -> U+0020 (should be kept)",
                   "notes": imp["notes"], "project": "research"}
    assert "research MCP" in imp["notes"] and "10 mm must not wrap" in imp["notes"]
    assert out["char_mapping"]["title_prefix"] == "Char mapping: U+00A0 "
    assert out["file_with"] == textclean.FILE_WITH and out["report_count"] == 1
    assert _report(title="again", details="d", category="char_mapping", char="\u00a0",
                   should_be="keep")["report_count"] == 2
    lines = (tmp_path / "char_reports.jsonl").read_text(encoding="utf-8").splitlines()
    assert [json.loads(x)["codepoint"] for x in lines] == ["U+00A0", "U+00A0"]
    issues = (tmp_path / "ISSUES.md").read_text(encoding="utf-8")
    assert "Char mapping: U+00A0" in issues and "`char_mapping`" in issues
    assert textclean.MAP_PATH.read_bytes() == table_before  # reports never edit the table


def test_char_mapping_removal_report_through_mcp(tmp_path, monkeypatch, agent_mode):
    """The documented call, through the MCP tool layer; should_be="" means "remove it"."""
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    table_before = textclean.MAP_PATH.read_bytes()
    result = asyncio.run(server.server.call_tool("report_issue", {
        "title": "ALM left in", "details": "bidi mark should be removed", "category": "char_mapping",
        "char": "U+061C", "should_be": "", "snippet": "abc"}))
    out = result.structured_content
    assert out["improvement"]["text"].endswith("(should be (removed))")
    assert out["report_count"] == 1 and "ALM left in" in (tmp_path / "ISSUES.md").read_text(encoding="utf-8")
    assert textclean.MAP_PATH.read_bytes() == table_before


def test_char_mapping_needs_char_and_should_be(tmp_path, monkeypatch):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    with pytest.raises(ValueError):
        _report(title="x", details="y", category="char_mapping")
    with pytest.raises(ValueError):
        _report(title="x", details="y", category="char_mapping", char="U+00A0")


def test_other_categories_in_agent_mode_only_log(tmp_path, monkeypatch, agent_mode):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    out = _report(title="x", details="y", category="friction")
    assert out["ok"] is True and out.get("improvement") is None and out.get("issue_url") is None
    assert not (tmp_path / "char_reports.jsonl").exists()


# -- default: a prefilled GitHub issue -----------------------------------------------------------------


def test_default_char_mapping_report_returns_a_github_issue_link(tmp_path, monkeypatch):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    table_before = textclean.MAP_PATH.read_bytes()
    out = _report(title="NBSP in units", details="10 mm must not wrap", category="char_mapping",
                  char="U+00A0", should_be="keep", snippet="width 10\u00a0mm")
    assert out["ok"] is True and out["report_count"] == 1
    assert out.get("improvement") is None and out.get("file_with") is None
    title = "Char mapping: U+00A0 NO-BREAK SPACE -> U+0020 (should be kept)"
    assert out["issue_title"] == title
    assert "10 mm must not wrap" in out["issue_body"] and "Fix: " not in out["issue_body"]
    assert _query(out["issue_url"]) == {"title": title, "body": out["issue_body"]}
    for text in (out["issue_url"], out["issue_body"], out["message"]):
        assert "file_with" not in text and "title_prefix" not in text
    # still logged locally, table untouched
    assert json.loads((tmp_path / "char_reports.jsonl").read_text(encoding="utf-8"))["codepoint"] == "U+00A0"
    assert "Char mapping: U+00A0" in (tmp_path / "ISSUES.md").read_text(encoding="utf-8")
    assert textclean.MAP_PATH.read_bytes() == table_before


def test_default_other_category_returns_a_github_issue_link(tmp_path, monkeypatch):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    out = _report(title="Table  style\nlost", details="Borders vanish & reappear?", category="formatting_defect")
    assert out["ok"] is True and out.get("improvement") is None
    assert out["issue_title"] == "Table style lost"
    assert out["issue_body"] == "Category: formatting_defect\n\nBorders vanish & reappear?"
    assert _query(out["issue_url"]) == {"title": "Table style lost", "body": out["issue_body"]}
    assert "&" not in urlsplit(out["issue_url"]).query.split("&body=")[1]  # the body's own & is encoded
    assert "Borders vanish & reappear?" in (tmp_path / "ISSUES.md").read_text(encoding="utf-8")


@pytest.mark.parametrize("unit", ["x", "a traceback line: File \"x.py\", line 3\n", "\u201cquote\u201d \u2014 ", "\U0001F600"])
def test_default_issue_url_is_capped_after_encoding(tmp_path, monkeypatch, unit):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    out = _report(title="long", details=unit * 20_000, category="other")
    assert server.ISSUE_URL_MAX - 20 < len(out["issue_url"]) <= server.ISSUE_URL_MAX
    query = _query(out["issue_url"])
    assert query["body"] == out["issue_body"] and out["issue_body"].startswith("Category: other")
    # Short reports are not trimmed.
    short = _report(title="short", details="fits easily", category="other")
    assert short["issue_body"] == "Category: other\n\nfits easily"


def test_default_report_through_mcp_carries_the_link(tmp_path, monkeypatch):
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    result = asyncio.run(server.server.call_tool("report_issue", {
        "title": "Lock stuck", "details": "edit refused for an hour", "category": "lock"}))
    out = result.structured_content
    assert out["issue_url"].startswith(server.GITHUB_ISSUES + "?")
    assert out["issue_title"] == "Lock stuck"


def test_issues_mode_reads_the_config_file(tmp_path, monkeypatch):
    from research_mcp import config

    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path))
    target = tmp_path / "config.env"
    target.write_text("RESEARCH_MCP_ISSUES=Agent\nRESEARCH_MCP_ISSUES_FILE_WITH=\"Use tracker X, project 'docs'.\"\n", encoding="utf-8")
    monkeypatch.setenv("RESEARCH_MCP_CONFIG", str(target))
    config.reload()
    out = _report(title="x", details="y", category="char_mapping", char="U+00A0", should_be="keep")
    assert "improvement" in out and out.get("issue_url") is None
    assert out["file_with"] == "Use tracker X, project 'docs'."
