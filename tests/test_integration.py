from __future__ import annotations

import json
import zipfile

import pytest

from research_mcp.errors import IntegrationBusy, ZoteroModal
from research_mcp.integration import IntegrationLease, dispatch_refresh
from research_mcp.ooxml import zotero_package_fields
from research_mcp.windows import WindowInfo


def _write_refresh_docx(
    path, *, citation_result: str, bibliography_result: str, extra_result: str | None = None
):
    payload = json.dumps({"citationID": "CITE1234"}, separators=(",", ":"))
    extra = ""
    if extra_result is not None:
        extra_payload = json.dumps({"citationID": "EXTRA999"}, separators=(",", ":"))
        extra = f'''
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> ADDIN ZOTERO_ITEM CSL_CITATION {extra_payload} </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>{extra_result}</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>'''
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> ADDIN ZOTERO_ITEM CSL_CITATION {payload} </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>{citation_result}</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    {extra}
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> ADDIN ZOTERO_BIBL {{}} CSL_BIBLIOGRAPHY </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>{bibliography_result}</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p></w:body>
</w:document>'''
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", xml)


class _FakeStream:
    """Minimal stdout/stderr stand-in.

    dispatch_refresh drains both streams on reader threads instead of calling
    communicate(): the OS pipe buffer is smaller than a 36-citation response,
    so an undrained helper deadlocks mid-write and never exits.
    """

    def __init__(self, text: str = ""):
        self._lines = [text] if text else []

    def __iter__(self):
        return iter(self._lines)

    def close(self):
        pass


class _FailedHelper:
    pid = 4242
    returncode = 1
    stdout_text = ""
    stderr_text = "worker response exceeded its deadline"

    def __init__(self):
        self.stdout = _FakeStream(self.stdout_text)
        self.stderr = _FakeStream(self.stderr_text)

    def poll(self):
        return self.returncode


class _IncompleteSuccessHelper(_FailedHelper):
    returncode = 0
    stdout_text = "{}"
    stderr_text = ""


class _AliveHelper(_FailedHelper):
    returncode = None

    def poll(self):
        return None


def test_lease_serializes_and_releases_on_success(tmp_path):
    lock = tmp_path / "integration.lock"
    lease = IntegrationLease.acquire(lock)
    assert json.loads(lock.read_text())["pid"] > 0
    with pytest.raises(IntegrationBusy):
        IntegrationLease.acquire(lock)
    lease.release_success()
    assert not lock.exists()


def test_lease_preserves_failure_for_explicit_recovery(tmp_path):
    lock = tmp_path / "integration.lock"
    lease = IntegrationLease.acquire(lock)
    lease.preserve_failure("modal")
    data = json.loads(lock.read_text())
    assert data["reason"] == "modal"
    with pytest.raises(IntegrationBusy, match="requires recovery"):
        IntegrationLease.acquire(lock)


def test_dispatch_refuses_preexisting_modal_before_spawning_helper(
    tmp_path, monkeypatch
):
    target = tmp_path / "target.docx"
    target.write_bytes(b"fixture")
    prompt = WindowInfo(
        7, 11, r"C:\Program Files\Zotero\zotero.exe", "#32770", True
    )
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: {11})
    monkeypatch.setattr(
        "research_mcp.integration.visible_dialogs_for_pids", lambda pids: (prompt,)
    )
    spawned = False

    def forbidden_spawn(*args, **kwargs):
        nonlocal spawned
        spawned = True
        raise AssertionError("helper must not start")

    monkeypatch.setattr("research_mcp.integration.subprocess.Popen", forbidden_spawn)
    with pytest.raises(ZoteroModal, match="already visible"):
        dispatch_refresh(target)
    assert not spawned


def test_dispatch_refuses_suspect_state_before_spawning_helper(tmp_path, monkeypatch):
    target = tmp_path / "target.docx"
    target.write_bytes(b"fixture")
    monkeypatch.setattr(
        "research_mcp.recovery.zotero_is_suspect",
        lambda: {"reason": "forced Word termination"},
    )
    monkeypatch.setattr(
        "research_mcp.integration.subprocess.Popen",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("must not spawn")),
    )
    with pytest.raises(IntegrationBusy, match="marked suspect"):
        dispatch_refresh(target)


def test_dispatch_reconciles_rendered_saved_package_after_helper_failure(
    tmp_path, monkeypatch
):
    target = tmp_path / "target.docx"
    _write_refresh_docx(
        target, citation_result="{Citation}", bibliography_result="{Bibliography}"
    )
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    monkeypatch.setattr(
        "research_mcp.integration.visible_dialogs_for_pids", lambda pids: ()
    )

    def complete_then_fail(*args, **kwargs):
        _write_refresh_docx(
            target, citation_result="(1)", bibliography_result="1. Example entry"
        )
        return _FailedHelper()

    monkeypatch.setattr("research_mcp.integration.subprocess.Popen", complete_then_fail)
    result = dispatch_refresh(target, citation_ids=("CITE1234",))
    assert result["rendered"] == {"CITE1234": "(1)"}
    assert result["bibliography_results"] == ["1. Example entry"]
    assert result["reconciled_from_saved_package"] is True
    assert result["helper_returncode"] == 1


def test_dispatch_does_not_mask_failed_idempotent_refresh(tmp_path, monkeypatch):
    target = tmp_path / "target.docx"
    _write_refresh_docx(
        target, citation_result="(1)", bibliography_result="1. Example entry"
    )
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    monkeypatch.setattr(
        "research_mcp.integration.visible_dialogs_for_pids", lambda pids: ()
    )
    monkeypatch.setattr(
        "research_mcp.integration.subprocess.Popen", lambda *args, **kwargs: _FailedHelper()
    )
    with pytest.raises(RuntimeError, match="saved package did not change"):
        dispatch_refresh(target, citation_ids=("CITE1234",))


def test_dispatch_rejects_valid_but_incomplete_helper_json(tmp_path, monkeypatch):
    target = tmp_path / "target.docx"
    _write_refresh_docx(
        target, citation_result="{Citation}", bibliography_result="{Bibliography}"
    )
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    monkeypatch.setattr("research_mcp.integration.visible_dialogs_for_pids", lambda pids: ())
    monkeypatch.setattr(
        "research_mcp.integration.subprocess.Popen",
        lambda *args, **kwargs: _IncompleteSuccessHelper(),
    )
    with pytest.raises(RuntimeError, match="wrong citation IDs"):
        dispatch_refresh(target, citation_ids=("CITE1234",))


def test_reconciliation_supports_requested_subset_and_requires_target_transition(
    tmp_path, monkeypatch
):
    target = tmp_path / "target.docx"
    _write_refresh_docx(
        target,
        citation_result="{Citation}",
        bibliography_result="1. Existing entry",
        extra_result="(9)",
    )
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    monkeypatch.setattr("research_mcp.integration.visible_dialogs_for_pids", lambda pids: ())

    def complete_subset(*args, **kwargs):
        _write_refresh_docx(
            target,
            citation_result="(1)",
            bibliography_result="1. Existing entry",
            extra_result="(9)",
        )
        return _FailedHelper()

    monkeypatch.setattr("research_mcp.integration.subprocess.Popen", complete_subset)
    result = dispatch_refresh(target, citation_ids=("CITE1234",))
    assert result["rendered"] == {"CITE1234": "(1)"}


def test_bibliography_only_change_cannot_mask_failed_targeted_refresh(
    tmp_path, monkeypatch
):
    target = tmp_path / "target.docx"
    _write_refresh_docx(
        target, citation_result="(1)", bibliography_result="{Bibliography}"
    )
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    monkeypatch.setattr("research_mcp.integration.visible_dialogs_for_pids", lambda pids: ())

    def bibliography_only(*args, **kwargs):
        _write_refresh_docx(
            target, citation_result="(1)", bibliography_result="1. Entry"
        )
        return _FailedHelper()

    monkeypatch.setattr("research_mcp.integration.subprocess.Popen", bibliography_only)
    with pytest.raises(RuntimeError, match="targeted pending-to-rendered transition"):
        dispatch_refresh(target, citation_ids=("CITE1234",))


def test_dispatch_retains_helper_that_outlives_deadline_and_grace(tmp_path, monkeypatch):
    target = tmp_path / "target.docx"
    _write_refresh_docx(
        target, citation_result="{Citation}", bibliography_result="{Bibliography}"
    )
    clock = {"now": 0.0}
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    monkeypatch.setattr("research_mcp.integration.visible_dialogs_for_pids", lambda pids: ())
    monkeypatch.setattr("research_mcp.integration.subprocess.Popen", lambda *a, **k: _AliveHelper())
    monkeypatch.setattr("research_mcp.integration.time.monotonic", lambda: clock["now"])
    monkeypatch.setattr(
        "research_mcp.integration.time.sleep",
        lambda seconds: clock.__setitem__("now", clock["now"] + seconds),
    )
    with pytest.raises(ZoteroModal, match="helper PID 4242 retained"):
        dispatch_refresh(
            target, citation_ids=("CITE1234",), dispatch_timeout=1.0, poll_interval=0.2
        )


def test_package_reader_preserves_paragraphs_and_nested_field_results(tmp_path):
    target = tmp_path / "nested.docx"
    xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>First</w:t></w:r></w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Link</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t> second</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
</w:body></w:document>'''
    with zipfile.ZipFile(target, "w") as archive:
        archive.writestr("word/document.xml", xml)
    assert zotero_package_fields(target).bibliography_results == ("First\nLink second",)


def test_dispatch_drains_a_helper_response_larger_than_the_pipe_buffer(
    tmp_path, monkeypatch
):
    """The defect that made every 36-citation refresh hang at any timeout.

    The OS pipe buffers roughly 4-8 KB. A 36-citation response is 7382 bytes,
    so a supervisor that polls poll() without draining stdout deadlocks the
    helper mid-write: the document is already refreshed and saved, but the
    helper can never exit and the refresh is reported as a timeout. Widening
    the timeout cannot fix it -- 144s, 300s and 900s all failed identically.
    """

    import subprocess as real_subprocess
    import sys as real_sys

    # These tests patch research_mcp.integration.subprocess.Popen, which is the
    # real subprocess module attribute -- capture the original before patching
    # or the fake calls itself.
    real_popen = real_subprocess.Popen

    target = tmp_path / "target.docx"
    _write_refresh_docx(
        target, citation_result="{Citation}", bibliography_result="{Bibliography}"
    )
    monkeypatch.setattr("research_mcp.recovery.zotero_is_suspect", lambda: None)
    monkeypatch.setattr("research_mcp.integration.zotero_pids", lambda: set())
    monkeypatch.setattr("research_mcp.integration.visible_dialogs_for_pids", lambda pids: ())

    # Build the padding inside the child: a 40 KB literal on the command line
    # exceeds the Windows 32 KB limit.
    script = (
        "import json,sys;"
        "payload={'rendered':{'CITE1234':'(1)'},"
        "'bibliography_results':['1. Example entry'],'padding':'x'*40000};"
        "sys.stdout.write(json.dumps(payload));sys.stdout.flush()"
    )

    def spawn_large_writer(*args, **kwargs):
        _write_refresh_docx(
            target, citation_result="(1)", bibliography_result="1. Example entry"
        )
        return real_popen(
            [real_sys.executable, "-c", script],
            stdout=real_subprocess.PIPE,
            stderr=real_subprocess.PIPE,
            text=True,
        )

    monkeypatch.setattr("research_mcp.integration.subprocess.Popen", spawn_large_writer)
    result = dispatch_refresh(target, citation_ids=("CITE1234",), dispatch_timeout=30.0)
    assert result["rendered"] == {"CITE1234": "(1)"}
    assert not result.get("helper_still_running")
    assert not result.get("reconciled_from_saved_package")
