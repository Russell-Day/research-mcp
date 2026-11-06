from __future__ import annotations

from pathlib import Path
import json
import threading
import time
import zipfile

import pytest

from research_mcp.edit_models import EditDocumentRequest
from research_mcp.errors import IntegrationRefused, RefusedUnsafe
from research_mcp.word import IsolatedWordDocument


def test_malformed_preflight_does_not_initialize_com(monkeypatch, tmp_path: Path):
    path = tmp_path / "broken.docx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
    calls = []
    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda _: ["missing document.xml"])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: calls.append("init"))

    with pytest.raises(RefusedUnsafe, match="preflight"):
        with IsolatedWordDocument(path):
            pass
    assert calls == []


def test_creation_failure_uninitializes_com(monkeypatch, minimal_docx: Path):
    calls = []
    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: calls.append("init"))
    monkeypatch.setattr("research_mcp.word.co_uninitialize", lambda: calls.append("uninit"))
    monkeypatch.setattr(
        "research_mcp.word.create_word_application",
        lambda **kwargs: (_ for _ in ()).throw(RuntimeError("creation failed")),
    )

    with pytest.raises(IntegrationRefused, match="could not open"):
        with IsolatedWordDocument(minimal_docx):
            pass
    assert calls == ["init", "uninit"]


def test_open_uses_hidden_safe_flags(monkeypatch, minimal_docx: Path):
    opened = {}

    class Documents:
        def Open(self, **kwargs):
            opened.update(kwargs)
            return object()

    documents = Documents()

    class Word:
        Documents = documents

        def Quit(self, **kwargs):
            pass

    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.co_uninitialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.create_word_application", lambda **kwargs: Word())
    monkeypatch.setattr("research_mcp.word.word_pid", lambda word: 4242)
    monkeypatch.setattr("research_mcp.word.process_creation_time", lambda pid: 100)
    monkeypatch.setattr("research_mcp.word.process_identity_matches", lambda pid, created: False)

    with IsolatedWordDocument(minimal_docx, visible=False, read_only=True):
        pass

    assert opened == {
        "FileName": str(minimal_docx.resolve()),
        "ReadOnly": True,
        "AddToRecentFiles": False,
        "Visible": False,
        "ConfirmConversions": False,
        "OpenAndRepair": False,
        "NoEncodingDialog": True,
    }


def test_open_deadline_terminates_owned_word_and_skips_com_cleanup(monkeypatch, minimal_docx: Path):
    killed = []
    close_calls = []
    quit_calls = []
    release_open = threading.Event()

    class Document:
        def Close(self, **kwargs):
            close_calls.append(kwargs)

    class Documents:
        def Open(self, **kwargs):
            release_open.wait(1)
            raise RuntimeError("server terminated")

    documents = Documents()

    class Word:
        Documents = documents

        def Quit(self, **kwargs):
            quit_calls.append(kwargs)

    def kill(pid):
        killed.append(pid)
        release_open.set()

    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.co_uninitialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.create_word_application", lambda **kwargs: Word())
    monkeypatch.setattr("research_mcp.word.word_pid", lambda word: 4242)
    monkeypatch.setattr("research_mcp.word.process_creation_time", lambda pid: 100)
    monkeypatch.setattr("research_mcp.word.process_identity_matches", lambda pid, created: True)
    monkeypatch.setattr("research_mcp.word.kill_process", kill)

    with pytest.raises(IntegrationRefused, match="exceeded timeout_seconds"):
        with IsolatedWordDocument(minimal_docx, timeout_seconds=0.01):
            pass
    assert killed == [4242]
    assert close_calls == []
    assert quit_calls == []


def test_body_deadline_fails_context_before_cleanup_com_calls(monkeypatch, minimal_docx: Path):
    killed = []
    close_calls = []
    quit_calls = []

    class Document:
        def Close(self, **kwargs):
            close_calls.append(kwargs)

    class Documents:
        def Open(self, **kwargs):
            return Document()

    documents = Documents()

    class Word:
        Documents = documents

        def Quit(self, **kwargs):
            quit_calls.append(kwargs)

    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.co_uninitialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.create_word_application", lambda **kwargs: Word())
    monkeypatch.setattr("research_mcp.word.word_pid", lambda word: 4242)
    monkeypatch.setattr("research_mcp.word.process_creation_time", lambda pid: 100)
    monkeypatch.setattr("research_mcp.word.process_identity_matches", lambda pid, created: True)
    monkeypatch.setattr("research_mcp.word.kill_process", lambda pid: killed.append(pid))

    with pytest.raises(IntegrationRefused, match="exceeded timeout_seconds"):
        with IsolatedWordDocument(minimal_docx, timeout_seconds=0.01):
            time.sleep(0.05)
    assert killed == [4242]
    assert close_calls == []
    assert quit_calls == []


def test_missing_owned_word_identity_is_refused(monkeypatch, minimal_docx: Path):
    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.co_uninitialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.create_word_application", lambda **kwargs: object())
    monkeypatch.setattr("research_mcp.word.word_pid", lambda word: None)

    with pytest.raises(IntegrationRefused, match="owned Word PID"):
        with IsolatedWordDocument(minimal_docx):
            pass


def test_writable_session_rejects_document_opened_read_only(monkeypatch, minimal_docx: Path):
    class Document:
        ReadOnly = True

    class Documents:
        def Open(self, **kwargs):
            return Document()

    documents = Documents()

    class Word:
        Documents = documents

        def Quit(self, **kwargs):
            pass

    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.co_uninitialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.create_word_application", lambda **kwargs: Word())
    monkeypatch.setattr("research_mcp.word.word_pid", lambda word: 4242)
    monkeypatch.setattr("research_mcp.word.process_creation_time", lambda pid: 100)
    monkeypatch.setattr("research_mcp.word.process_identity_matches", lambda pid, created: False)

    with pytest.raises(IntegrationRefused, match="opened the document read-only"):
        with IsolatedWordDocument(minimal_docx, read_only=False):
            pass


def test_owned_word_registration_is_atomic_and_identity_checked(monkeypatch, minimal_docx: Path, tmp_path: Path):
    registration = tmp_path / "owned-word.json"

    class Documents:
        def Open(self, **kwargs):
            return object()

    documents = Documents()

    class Word:
        Documents = documents

        def Quit(self, **kwargs):
            pass

    monkeypatch.setenv("RESEARCH_MCP_OWNED_WORD_FILE", str(registration))
    monkeypatch.setattr("research_mcp.word.document_structure_errors", lambda path: [])
    monkeypatch.setattr("research_mcp.word.co_initialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.co_uninitialize", lambda: None)
    monkeypatch.setattr("research_mcp.word.create_word_application", lambda **kwargs: Word())
    monkeypatch.setattr("research_mcp.word.word_pid", lambda word: 4242)
    monkeypatch.setattr("research_mcp.word.process_creation_time", lambda pid: 100)
    monkeypatch.setattr("research_mcp.word.process_identity_matches", lambda pid, created: False)

    with IsolatedWordDocument(minimal_docx):
        assert json.loads(registration.read_text(encoding="utf-8")) == {"pid": 4242, "created": 100}
    assert not registration.exists()


def test_cleanup_does_not_kill_reused_pid(monkeypatch, minimal_docx: Path):
    document = IsolatedWordDocument(minimal_docx)
    document.pid = 4242
    document.process_created = 100
    killed = []
    monkeypatch.setattr("research_mcp.word.process_identity_matches", lambda pid, created: False)
    monkeypatch.setattr("research_mcp.word.kill_process", lambda pid: killed.append(pid))
    document._force_cleanup()
    assert killed == []


def test_property_initialization_failure_closes_new_word_instance(monkeypatch):
    from research_mcp.word_com import create_word_application
    closed = []

    class Word:
        def __setattr__(self, name, value):
            if name == "AutomationSecurity":
                raise RuntimeError("security property failed")
            object.__setattr__(self, name, value)

        def Quit(self, **kwargs):
            closed.append(kwargs)

    monkeypatch.setattr("research_mcp.word_com.ensure_default_desktop", lambda: None)
    monkeypatch.setattr("research_mcp.word_com.win32com.client.DispatchEx", lambda _: Word())
    with pytest.raises(RuntimeError, match="security property failed"):
        create_word_application()
    assert closed == [{"SaveChanges": 0}]


@pytest.mark.parametrize("operation", [
    {"op_id": "r", "type": "replace_text", "paragraph_index": 1, "find": "x" * 256, "replace": "y"},
    {"op_id": "i", "type": "insert_text", "paragraph_index": 1, "position": "after", "anchor": "x" * 256, "text": "y"},
])
def test_find_and_anchor_are_limited_to_word_search_length(operation):
    with pytest.raises(ValueError, match="at most 255"):
        EditDocumentRequest(source="x", source_sha256="0" * 64, operations=[operation])
