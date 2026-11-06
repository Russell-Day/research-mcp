from __future__ import annotations

from pathlib import Path

import pytest

from research_mcp.errors import IntegrationRefused
from research_mcp.word_worker import _configure_word_application, _refresh


def test_persistent_worker_configures_hidden_no_macro_word():
    class Word:
        pass

    word = Word()
    _configure_word_application(word)
    assert word.Visible is False
    assert word.ScreenUpdating is False
    assert word.AutomationSecurity == 3


def test_worker_refresh_preflight_rejects_before_documents_open(monkeypatch, minimal_docx: Path):
    opened = []

    class Documents:
        Count = 0

        def Open(self, **kwargs):
            opened.append(kwargs)
            raise AssertionError("Word must not open a preflight-invalid package")

    documents = Documents()

    class Word:
        Documents = documents

    monkeypatch.setattr(
        "research_mcp.word_worker.document_structure_errors",
        lambda path: ["missing required package part"],
    )
    with pytest.raises(RuntimeError, match="preflight failed"):
        _refresh(Word(), minimal_docx, (), 1.0)
    assert opened == []


def test_worker_refresh_uses_hidden_safe_open_flags(monkeypatch, minimal_docx: Path):
    opened = {}

    class Document:
        ReadOnly = False

        def Activate(self):
            pass

        class Application:
            @staticmethod
            def Run(command):
                assert command == "ZoteroRefresh"

        def Save(self):
            pass

        def Close(self, **kwargs):
            pass

    class Documents:
        Count = 0

        def Open(self, **kwargs):
            opened.update(kwargs)
            return Document()

    documents = Documents()

    class Word:
        Documents = documents

    monkeypatch.setattr("research_mcp.word_worker.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word_worker._mark_refresh_pending", lambda document: "marker")
    monkeypatch.setattr("research_mcp.word_worker._refresh_completed", lambda document, marker: True)
    monkeypatch.setattr("research_mcp.word_worker._item_fields", lambda document: [])
    monkeypatch.setattr("research_mcp.word_worker._bibliography_fields", lambda document: [])
    monkeypatch.setattr("research_mcp.word_worker._record_progress", lambda **fields: None)
    monkeypatch.setattr("research_mcp.word_worker.time.sleep", lambda seconds: None)

    _refresh(Word(), minimal_docx, (), 1.0)
    assert opened == {
        "FileName": str(minimal_docx.resolve()),
        "ReadOnly": False,
        "AddToRecentFiles": False,
        "Visible": False,
        "ConfirmConversions": False,
        "OpenAndRepair": False,
        "NoEncodingDialog": True,
    }


def _fake_word():
    class Document:
        ReadOnly = False

        def Activate(self):
            pass

        class Application:
            @staticmethod
            def Run(command):
                pass

        def Save(self):
            pass

        def Close(self, *args, **kwargs):
            pass

    class Documents:
        Count = 0

        def Open(self, **kwargs):
            return Document()

    class Word:
        pass

    Word.Documents = Documents()
    return Word()


def _field(code: str, result: str):
    class Text:
        def __init__(self, text):
            self.Text = text

    class Field:
        Code = Text(code)
        Result = Text(result)

    return Field()


def _patch_loop(monkeypatch, polls):
    """Each poll yields (item fields, refresh command complete)."""
    state = {"i": 0}

    def item_fields(document):
        state["i"] = min(state["i"] + 1, len(polls))
        return polls[state["i"] - 1][0]

    monkeypatch.setattr("research_mcp.word_worker.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word_worker._mark_refresh_pending", lambda document: "marker")
    monkeypatch.setattr("research_mcp.word_worker._refresh_completed", lambda document, marker: polls[state["i"] - 1][1])
    monkeypatch.setattr("research_mcp.word_worker._item_fields", item_fields)
    monkeypatch.setattr("research_mcp.word_worker._bibliography_fields", lambda document: [])
    monkeypatch.setattr("research_mcp.word_worker._record_progress", lambda **fields: None)
    monkeypatch.setattr("research_mcp.word_worker.time.sleep", lambda seconds: None)


def test_worker_refresh_tolerates_field_mid_rewrite(monkeypatch, minimal_docx: Path):
    done = _field('ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"a1"}', "[1]")
    polls = [
        ([_field("ADDIN ZOTERO_ITEM CSL_CITATION {\"citat", "")], False),  # Zotero rewriting the code
        ([_field("ADDIN ZOTERO_ITEM", "")], False),
        ([done], True),
    ]
    _patch_loop(monkeypatch, polls)
    assert _refresh(_fake_word(), minimal_docx, ("a1",), 5.0)["rendered"] == {"a1": "[1]"}


def test_worker_refresh_fails_when_field_gone_after_refresh(monkeypatch, minimal_docx: Path):
    _patch_loop(monkeypatch, [([], True)])
    with pytest.raises(RuntimeError, match="citation fields disappeared"):
        _refresh(_fake_word(), minimal_docx, ("a1",), 5.0)
