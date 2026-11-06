"""config.py: environment > config file > default, and the comment-author fallback chain.

conftest's `isolated_config` points RESEARCH_MCP_CONFIG at a scratch file (yielded here as the fixture
value) and pins RESEARCH_MCP_AUTHOR; tests that exercise the chain clear it.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from research_mcp import config
from research_mcp.errors import ZoteroUnavailable


def _write(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")
    config.reload()


# -- precedence ------------------------------------------------------------------------------------------


def test_env_beats_file_beats_default(isolated_config, monkeypatch):
    _write(isolated_config, "RESEARCH_MCP_X=from-file\n")
    monkeypatch.setenv("RESEARCH_MCP_X", "from-env")
    assert config.get("RESEARCH_MCP_X", "default") == "from-env"
    monkeypatch.delenv("RESEARCH_MCP_X")
    assert config.get("RESEARCH_MCP_X", "default") == "from-file"
    assert config.get("RESEARCH_MCP_UNSET", "default") == "default"
    assert config.get("RESEARCH_MCP_UNSET") is None


def test_empty_env_value_counts_as_unset(isolated_config, monkeypatch):
    _write(isolated_config, "RESEARCH_MCP_X=from-file\n")
    monkeypatch.setenv("RESEARCH_MCP_X", "  ")
    assert config.get("RESEARCH_MCP_X") == "from-file"


def test_missing_file_is_fine(tmp_path, monkeypatch):
    monkeypatch.setenv("RESEARCH_MCP_CONFIG", str(tmp_path / "nowhere" / "config.env"))
    config.reload()
    assert config.get("RESEARCH_MCP_X", "default") == "default"


def test_file_syntax_comments_quotes_export_and_bom(isolated_config):
    isolated_config.write_bytes(
        "﻿# a comment\n\nexport A=1\nB = \"two words\"\nC='3'\nnot a setting\nD=\nE=x=y\n".encode("utf-8")
    )
    config.reload()
    assert [config.get(name) for name in "ABCDE"] == ["1", "two words", "3", None, "x=y"]
    assert config.parse("# only=comment\n") == {}


def test_inline_comments_quoted_values_and_export_tab(isolated_config):
    _write(isolated_config, 'A=user # a note\nB="q" # c\nC=\'it # stays\'\nexport\tD=1\nE=x#y\nF="unterminated\n')
    assert [config.get(name) for name in "ABCDEF"] == ["user", "q", "it # stays", "1", "x#y", '"unterminated']


def test_utf16_file_as_powershell_writes_it(isolated_config):
    # Windows PowerShell 5.1 `"..." > config.env` and Out-File write UTF-16LE with a BOM.
    isolated_config.write_bytes(b"\xff\xfe" + "RESEARCH_MCP_NAME=Jos\u00e9 Doe\r\nRESEARCH_MCP_X=1\r\n".encode("utf-16-le"))
    config.reload()
    assert config.get("RESEARCH_MCP_NAME") == "Jos\u00e9 Doe" and config.get("RESEARCH_MCP_X") == "1"
    isolated_config.write_bytes(b"\xfe\xff" + "RESEARCH_MCP_X=big\n".encode("utf-16-be"))
    config.reload()
    assert config.get("RESEARCH_MCP_X") == "big"


def test_ansi_byte_does_not_break_the_other_settings(isolated_config):
    isolated_config.write_bytes("RESEARCH_MCP_NAME=Jos\u00e9\nRESEARCH_MCP_X=1\n".encode("cp1252"))
    config.reload()
    assert config.get("RESEARCH_MCP_X") == "1"
    assert config.get("RESEARCH_MCP_NAME") == "Jos\ufffd"


def test_a_directory_counts_as_a_missing_file(tmp_path, monkeypatch):
    monkeypatch.setenv("RESEARCH_MCP_CONFIG", str(tmp_path))
    config.reload()
    assert config.get("RESEARCH_MCP_X", "default") == "default"
    from research_mcp import checkpoints

    checkpoints.limits()  # in-place edits read their limits through config


def test_file_is_cached_until_reload(isolated_config):
    _write(isolated_config, "RESEARCH_MCP_X=old\n")
    assert config.get("RESEARCH_MCP_X") == "old"
    isolated_config.write_text("RESEARCH_MCP_X=new\n", encoding="utf-8")
    assert config.get("RESEARCH_MCP_X") == "old"
    config.reload()
    assert config.get("RESEARCH_MCP_X") == "new"


def test_default_path_is_under_appdata(monkeypatch, tmp_path):
    monkeypatch.delenv("RESEARCH_MCP_CONFIG")
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert config.config_path() == tmp_path / "research-mcp" / "config.env"


def test_zotero_write_client_reads_the_config_file(isolated_config, monkeypatch):
    for name in ("ZOTERO_API_KEY", "ZOTERO_LIBRARY_ID", "ZOTERO_LIBRARY_TYPE"):
        monkeypatch.delenv(name, raising=False)
    from research_mcp.zotero import _write_client_from_env

    _write(isolated_config, "")
    with pytest.raises(ZoteroUnavailable, match="ZOTERO_LIBRARY_ID, ZOTERO_API_KEY") as excinfo:
        _write_client_from_env()
    assert str(isolated_config) in str(excinfo.value)

    _write(isolated_config, "ZOTERO_API_KEY=k" + "0" * 23 + "\nZOTERO_LIBRARY_ID=42\n")
    client = _write_client_from_env()
    assert (client.library_id, client.library_type) == ("42", "users")


# -- comment author ----------------------------------------------------------------------------------------


@pytest.fixture
def no_configured_author(monkeypatch):
    for name in ("RESEARCH_MCP_AUTHOR", "RESEARCH_MCP_AUTHOR_INITIALS"):
        monkeypatch.delenv(name, raising=False)
    config.reload()


def _word(monkeypatch, worker=None, registry=None):
    calls = []

    def source(value, label):
        def read():
            calls.append(label)
            if isinstance(value, Exception):
                raise value
            return value
        return read

    monkeypatch.setattr(config, "_word_user_from_worker", source(worker, "worker"))
    monkeypatch.setattr(config, "_word_user_from_registry", source(registry, "registry"))
    config.reload()
    return calls


def test_configured_author_wins_and_derives_initials(isolated_config, no_configured_author, monkeypatch):
    calls = _word(monkeypatch, worker=("Word Name", "WN"))
    _write(isolated_config, "RESEARCH_MCP_AUTHOR=Ada Lovelace\n")
    assert config.comment_author() == ("Ada Lovelace", "AL")
    _write(isolated_config, "RESEARCH_MCP_AUTHOR=Ada Lovelace\nRESEARCH_MCP_AUTHOR_INITIALS=AdL\n")
    assert config.comment_author() == ("Ada Lovelace", "AdL")
    assert calls == []  # Word is never asked when the author is configured


def test_word_user_name_is_the_fallback(no_configured_author, monkeypatch):
    calls = _word(monkeypatch, worker=("Grace Hopper", "GMH"))
    assert config.comment_author() == ("Grace Hopper", "GMH")
    assert config.comment_author() == ("Grace Hopper", "GMH")
    assert calls == ["worker"]  # cached once per process


def test_word_initials_are_derived_when_word_has_none(no_configured_author, monkeypatch):
    _word(monkeypatch, worker=None, registry=("grace brewster hopper", ""))
    assert config.comment_author() == ("grace brewster hopper", "GBH")


def test_a_failing_worker_falls_through_to_the_registry(no_configured_author, monkeypatch):
    calls = _word(monkeypatch, worker=OSError("pipe gone"), registry=("Grace Hopper", "GH"))
    assert config.comment_author() == ("Grace Hopper", "GH")
    assert calls == ["worker", "registry"]


def test_author_is_the_last_resort(no_configured_author, monkeypatch):
    calls = _word(monkeypatch, worker=None, registry=None)
    assert config.comment_author() == ("Author", "A")
    assert config.comment_author() == ("Author", "A")
    assert calls == ["worker", "registry", "worker", "registry"]  # a miss is not cached: Word may start later


def test_a_busy_worker_never_blocks_the_author_lookup(no_configured_author, monkeypatch):
    """The Word worker serves one request at a time; during a long refresh its handshake hangs."""

    import threading
    import time

    release = threading.Event()

    def busy_worker():
        release.wait(30)
        return ("Busy Worker", "BW")

    monkeypatch.setattr(config, "WORD_NAME_TIMEOUT", 0.3)
    monkeypatch.setattr(config, "_word_user_from_worker", busy_worker)
    monkeypatch.setattr(config, "_word_user_from_registry", lambda: ("Grace Hopper", "GH"))
    config.reload()
    try:
        started = time.monotonic()
        assert config.comment_author() == ("Grace Hopper", "GH")
        assert time.monotonic() - started < 2.0
        monkeypatch.setattr(config, "_word_user_from_registry", lambda: None)
        config.reload()
        started = time.monotonic()
        assert config.comment_author() == ("Author", "A")
        assert time.monotonic() - started < 2.0
    finally:
        release.set()


def test_a_given_author_never_asks_for_the_default(no_configured_author, monkeypatch):
    from research_mcp.comment_ops import _clean_author

    calls = _word(monkeypatch, worker=("Grace Hopper", "GH"), registry=("Grace Hopper", "GH"))
    monkeypatch.setenv("RESEARCH_MCP_AUTHOR", "Ada Lovelace")
    monkeypatch.setenv("RESEARCH_MCP_AUTHOR_INITIALS", "AdL")
    monkeypatch.setattr(config, "comment_author", lambda: pytest.fail("default author resolved for a given author"))
    assert _clean_author("Ada Lovelace", None, "add_comment") == ("Ada Lovelace", "AL")
    assert _clean_author(" Alan Turing ", "T", "reply_comment") == ("Alan Turing", "T")
    assert calls == []


def test_comment_ops_use_the_chain(no_configured_author, monkeypatch):
    from research_mcp.comment_ops import _clean_author

    _word(monkeypatch, worker=("Grace Hopper", "GH"))
    assert _clean_author(None, None, "add_comment") == ("Grace Hopper", "GH")
    assert _clean_author("Alan Turing", None, "add_comment") == ("Alan Turing", "AT")
    assert _clean_author("Alan Turing", "T", "add_comment") == ("Alan Turing", "T")
    _word(monkeypatch, worker=None, registry=None)
    assert _clean_author(None, None, "add_comment") == ("Author", "A")


def test_tracked_edits_still_default_to_claude():
    from research_mcp.edit_models import TrackedReplaceTextOperation

    op = TrackedReplaceTextOperation(op_id="r", type="tracked_replace_text", paragraph_index=1, find="a", replace="b")
    assert op.author == "Claude"
