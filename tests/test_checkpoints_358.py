"""Plan 358 S1: checkpoint store and one-call undo for in-place Word edits (synthetic .docx in tmp only)."""
from __future__ import annotations

import asyncio
import json
import os
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

from research_mcp import checkpoints as ck
from research_mcp.audit import audit_docx, file_sha256
from research_mcp.edit_batch import edit_document
from research_mcp.edit_models import EditDocumentRequest
from research_mcp.errors import RefusedUnsafe
from research_mcp.publication import (
    checkpoint_document,
    document_history,
    publish_inplace,
    undo_document,
)

pytestmark = pytest.mark.skipif(os.name != "nt", reason="undo store writes need Windows Python")


def fake_apply(source, output, operation):
    """Stand-in for Word: replace operation.find with operation.replace in word/document.xml."""
    if operation.find == "missing":
        raise RefusedUnsafe("anchor not found")
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(output, "w") as target:
        for item in archive.infolist():
            data = archive.read(item.filename)
            if item.filename == "word/document.xml":
                data = data.replace(operation.find.encode(), operation.replace.encode(), 1)
            target.writestr(item, data)
    return {"paragraph_count": audit_docx(output).paragraph_count}


@pytest.fixture
def word(monkeypatch):
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)


def batch(path, *pairs, output=None):
    ops = [dict(op_id=f"op{i}", type="replace_text", paragraph_index=1, find=f, replace=r)
           for i, (f, r) in enumerate(pairs, 1)]
    kwargs = {"output": str(output)} if output else {}
    return edit_document(EditDocumentRequest(source=str(path), source_sha256=file_sha256(path),
                                             operations=ops, **kwargs))


def stage(tmp_path, data: bytes) -> Path:
    s = tmp_path / f"stage-{os.urandom(4).hex()}.docx"
    s.write_bytes(data)
    return s


def write(path: Path, data: bytes, tmp_path) -> dict:
    return publish_inplace(stage(tmp_path, data), path, expected_sha256=file_sha256(path), tool="test", op="write")


def test_t1_one_call_undo_of_three_op_batch(word, minimal_docx):
    before_bytes, before_sha = minimal_docx.read_bytes(), file_sha256(minimal_docx)
    result = batch(minimal_docx, ("Fixture", "Alpha"), ("paragraph", "line"), ("Bookmarked", "Marked"))
    assert result.status == "complete", result
    assert result.undo_token and result.sha256_before == before_sha
    after_bytes = minimal_docx.read_bytes()
    assert result.sha256_after == file_sha256(minimal_docx) != before_sha
    assert not minimal_docx.with_name(minimal_docx.name + ".bak").exists()

    hist = document_history(minimal_docx)
    assert [e["kind"] for e in hist["entries"]] == ["edit"]
    assert hist["entries"][0]["id"] == result.undo_token
    assert hist["entries"][0]["tool"] == "edit_document"
    assert all(op in hist["entries"][0]["op"] for op in ("op1", "op2", "op3"))

    undone = undo_document(str(minimal_docx))
    assert undone["restored"] and minimal_docx.read_bytes() == before_bytes
    assert undone["sha256_after"] == before_sha == file_sha256(minimal_docx)
    redone = undo_document(str(minimal_docx))
    assert redone["restored"] and minimal_docx.read_bytes() == after_bytes
    assert [e["kind"] for e in document_history(minimal_docx)["entries"]] == ["undo", "undo", "edit"]


def test_t2_undo_to_label_after_three_edits(word, minimal_docx):
    start = minimal_docx.read_bytes()
    cp = checkpoint_document(str(minimal_docx), "start")
    assert cp["sha256"] == file_sha256(minimal_docx)
    for find, rep in (("Fixture", "One"), ("Early", "Two"), ("Late", "Three")):
        assert batch(minimal_docx, (find, rep)).status == "complete"
    assert minimal_docx.read_bytes() != start
    result = undo_document(str(minimal_docx), to="start")
    assert result["restored"] and minimal_docx.read_bytes() == start
    with pytest.raises(RefusedUnsafe, match="no checkpoint labelled"):
        undo_document(str(minimal_docx), to="nope")


def test_t3_stale_guard_force_and_output(word, minimal_docx, tmp_path):
    pre = minimal_docx.read_bytes()
    edit = batch(minimal_docx, ("Fixture", "Edited"))
    external = minimal_docx.read_bytes() + b"hand edit"
    minimal_docx.write_bytes(external)
    ext_sha = file_sha256(minimal_docx)

    with pytest.raises(ck.Refusal) as refused:
        undo_document(str(minimal_docx))
    assert refused.value.code == "stale"
    assert ext_sha in str(refused.value) and edit.sha256_after in str(refused.value)
    assert minimal_docx.read_bytes() == external

    out = tmp_path / "earlier.docx"
    written = undo_document(str(minimal_docx), output=str(out))
    assert out.read_bytes() == pre and minimal_docx.read_bytes() == external
    assert written["output"] == str(out.resolve())

    forced = undo_document(str(minimal_docx), force=True)
    assert forced["forced"] and minimal_docx.read_bytes() == pre
    assert ck.blob_path(ext_sha).read_bytes() == external  # overwritten state checkpointed first
    kinds = [e["kind"] for e in document_history(minimal_docx)["entries"]]
    assert kinds[:2] == ["undo", "external"]
    assert undo_document(str(minimal_docx))["restored"]
    assert minimal_docx.read_bytes() == external


def test_t4_failed_batch_creates_no_checkpoint(word, minimal_docx):
    original = minimal_docx.read_bytes()
    result = batch(minimal_docx, ("Fixture", "Alpha"), ("missing", "x"))
    assert result.status == "no_changes"
    assert result.undo_token is None
    assert minimal_docx.read_bytes() == original
    assert document_history(minimal_docx)["entries"] == []
    assert not list(ck.blobs_dir().glob("*.docx"))


def test_t5_retention_entries(minimal_docx):
    ids = [checkpoint_document(str(minimal_docx), f"c{i}")["checkpoint"] for i in range(31)]
    hist = document_history(minimal_docx)
    assert len(hist["entries"]) == 30
    assert ids[0] not in [e["id"] for e in hist["entries"]]
    assert hist["dropped"][-1]["id"] == ids[0] and "more than 30 entries" in hist["dropped"][-1]["reason"]
    assert len(list(ck.blobs_dir().glob("*.docx"))) == 1  # identical states stored once


def test_t5_retention_days(minimal_docx):
    old = checkpoint_document(str(minimal_docx), "old")["checkpoint"]
    hp = ck.doc_dir(minimal_docx) / "history.json"
    hist = json.loads(hp.read_text(encoding="utf-8"))
    hist["entries"][0]["time"] = (datetime.now(timezone.utc) - timedelta(days=15)).isoformat(timespec="seconds")
    hp.write_text(json.dumps(hist), encoding="utf-8")
    checkpoint_document(str(minimal_docx), "new")
    hist = document_history(minimal_docx)
    assert [e["label"] for e in hist["entries"]] == ["new"]
    assert hist["dropped"][-1]["id"] == old and "older than 14 days" in hist["dropped"][-1]["reason"]


def test_t5_retention_bytes_and_store_full(monkeypatch, tmp_path):
    doc = tmp_path / "bytes.docx"
    doc.write_bytes(b"A" * 1000)
    monkeypatch.setenv("RESEARCH_MCP_MAX_BYTES", "2500")
    first = write(doc, b"B" * 1000, tmp_path)
    write(doc, b"C" * 1000, tmp_path)
    write(doc, b"D" * 1000, tmp_path)  # stores C; A + B + C would exceed 2500
    hist = document_history(doc)
    assert first["undo_token"] not in [e["id"] for e in hist["entries"]]
    assert "above 2500 bytes" in hist["dropped"][-1]["reason"]
    assert not ck.blob_path(file_sha256_bytes(b"A" * 1000)).exists()
    assert doc.read_bytes() == b"D" * 1000

    monkeypatch.setenv("RESEARCH_MCP_MAX_BYTES", "500")
    with pytest.raises(ck.Refusal) as refused:
        write(doc, b"E" * 1000, tmp_path)
    assert refused.value.code == "store_full"
    assert doc.read_bytes() == b"D" * 1000


def file_sha256_bytes(data: bytes) -> str:
    import hashlib
    return hashlib.sha256(data).hexdigest()


def test_t6_refuses_locked_file(word, minimal_docx, tmp_path):
    batch(minimal_docx, ("Fixture", "Alpha"))
    current = minimal_docx.read_bytes()
    owner = minimal_docx.parent / f"~${minimal_docx.name}"
    owner.write_bytes(b"owner")
    with pytest.raises(ck.Refusal) as refused:
        undo_document(str(minimal_docx))
    assert refused.value.code == "doc_open"
    with pytest.raises(RefusedUnsafe, match="Save and close"):
        write(minimal_docx, b"new", tmp_path)
    assert minimal_docx.read_bytes() == current


def test_t7_legacy_bak_fallback_and_no_new_bak(word, monkeypatch, minimal_docx, tmp_path):
    bak = minimal_docx.with_name(minimal_docx.name + ".bak")
    legacy = minimal_docx.read_bytes()
    bak.write_bytes(legacy)
    minimal_docx.write_bytes(legacy + b"later legacy edit")
    current = minimal_docx.read_bytes()
    monkeypatch.setenv("RESEARCH_MCP_LOG_DIR", str(tmp_path / "logs"))

    with pytest.raises(ck.Refusal) as refused:
        undo_document(str(minimal_docx))
    assert refused.value.code == "stale" and minimal_docx.read_bytes() == current

    # The record comes through the real logger path: an MCP CallToolResult envelope.
    from mcp.types import CallToolResult, TextContent
    from research_mcp.logger import log_tool_call
    payload = {"source": str(minimal_docx), "output": str(minimal_docx), "status": "complete",
               "source_sha256": file_sha256(bak), "output_sha256": file_sha256(minimal_docx)}
    log_tool_call("edit_document", {"source": str(minimal_docx)}, result=CallToolResult(
        content=[TextContent(type="text", text=json.dumps(payload))], structured_content=payload))
    logged = json.loads((tmp_path / "logs" / "mcp_traffic.jsonl").read_text(encoding="utf-8").splitlines()[-1])
    assert "structured_content" in logged["result"] and "source" not in logged["result"]
    result = undo_document(str(minimal_docx))
    assert result["source"] == "bak" and result["backup_path"] == str(bak)
    assert minimal_docx.read_bytes() == legacy and bak.read_bytes() == legacy  # .bak left in place
    assert undo_document(str(minimal_docx))["restored"] and minimal_docx.read_bytes() == current

    fresh = tmp_path / "fresh"
    fresh.mkdir()
    doc = fresh / "paper.docx"
    doc.write_bytes(legacy)
    assert batch(doc, ("Fixture", "Alpha")).status == "complete"
    write(doc, legacy, tmp_path)
    checkpoint_document(str(doc), "x")
    undo_document(str(doc))
    assert sorted(p.name for p in fresh.iterdir()) == ["paper.docx"]
    bare = tmp_path / "no_history.docx"
    bare.write_bytes(legacy)
    with pytest.raises(ck.Refusal, match="OneDrive version history"):
        undo_document(str(bare))


def test_t8_versioned_write_leaves_source_identical(word, minimal_docx, tmp_path):
    original, sha = minimal_docx.read_bytes(), file_sha256(minimal_docx)
    out = tmp_path / "v2.docx"
    result = batch(minimal_docx, ("Fixture", "Alpha"), output=out)
    assert result.status == "complete"
    assert minimal_docx.read_bytes() == original
    assert result.undo_token is None
    assert result.sha256_before == sha and result.sha256_after == file_sha256(out)
    assert document_history(minimal_docx)["entries"] == []


def test_t9_same_name_different_folders_separate_histories(word, minimal_docx, tmp_path):
    docs = []
    for name in ("a", "b"):
        (tmp_path / name).mkdir()
        doc = tmp_path / name / "paper.docx"
        doc.write_bytes(minimal_docx.read_bytes())
        docs.append(doc)
    batch(docs[0], ("Fixture", "Alpha"))
    batch(docs[1], ("Fixture", "Beta"))
    batch(docs[1], ("Early", "Gamma"))
    h0, h1 = document_history(docs[0]), document_history(docs[1])
    assert h0["store"] != h1["store"]
    assert len(h0["entries"]) == 1 and len(h1["entries"]) == 2
    undo_document(str(docs[0]))
    assert docs[0].read_bytes() == minimal_docx.read_bytes()
    assert b"Gamma" in zipfile.ZipFile(docs[1]).read("word/document.xml")


def test_t10_change_between_checkpoint_and_publish_refused(monkeypatch, minimal_docx, tmp_path):
    original = minimal_docx.read_bytes()
    external = original + b"saved by Word meanwhile"
    real_store = ck.store_blob

    def store_then_change(src, sha=None):
        got = real_store(src, sha)
        minimal_docx.write_bytes(external)
        return got

    monkeypatch.setattr(ck, "store_blob", store_then_change)
    with pytest.raises(RefusedUnsafe, match="Source changed"):
        write(minimal_docx, original + b"agent edit", tmp_path)
    assert minimal_docx.read_bytes() == external
    monkeypatch.setattr(ck, "store_blob", real_store)
    assert document_history(minimal_docx)["entries"] == []
    assert not list(ck.journal_dir().glob("*.json"))


def test_server_tools_return_undo_fields(word, minimal_docx):
    from research_mcp.server import server

    async def call(name, arguments):
        return (await server.call_tool(name, arguments)).structured_content

    before = file_sha256(minimal_docx)
    ops = [{"op_id": "a", "type": "replace_text", "paragraph_index": 1, "find": "Fixture", "replace": "Alpha"}]
    edited = asyncio.run(call("edit_document", {"source": str(minimal_docx), "source_sha256": before,
                                                "output": str(minimal_docx), "operations": ops}))
    assert edited["undo_token"] and edited["sha256_before"] == before
    hist = asyncio.run(call("document_history", {"path": str(minimal_docx)}))
    assert hist["entries"][0]["id"] == edited["undo_token"]
    reverted = asyncio.run(call("revert_document", {"path": str(minimal_docx)}))
    assert reverted["restored"] and reverted["sha256_after"] == before == file_sha256(minimal_docx)


def _inject_save_before_replace(monkeypatch, target: Path, data: bytes):
    """Write ``data`` to ``target`` right before the ReplaceFileW call: after the last hash check."""
    from research_mcp import publication
    real = publication._replace_file
    calls = []

    def save_then_replace(replacement, replaced, displaced):
        if not calls:
            Path(replaced).write_bytes(data)
        calls.append(replaced)
        return real(replacement, replaced, displaced)

    monkeypatch.setattr(publication, "_replace_file", save_then_replace)
    return calls


def test_p1_save_just_before_edit_swap_is_kept(word, monkeypatch, minimal_docx):
    pre = minimal_docx.read_bytes()
    external = pre + b"saved by Word at the boundary"
    ext_sha = file_sha256_bytes(external).upper()
    calls = _inject_save_before_replace(monkeypatch, minimal_docx, external)
    result = batch(minimal_docx, ("Fixture", "Alpha"))
    assert calls and result.status == "complete"
    edited = minimal_docx.read_bytes()
    assert result.publish_warning and ext_sha in result.publish_warning
    assert result.sha256_before == ext_sha
    assert ck.blob_path(ext_sha).read_bytes() == external
    hist = document_history(minimal_docx)
    assert [e["kind"] for e in hist["entries"]] == ["edit", "external"]
    assert hist["entries"][1]["sha_after"] == ext_sha
    assert not list(ck.journal_dir().glob("*.json"))
    assert not list(minimal_docx.parent.glob("*displaced*"))

    undone = undo_document(str(minimal_docx))
    assert minimal_docx.read_bytes() == external, undone
    assert undo_document(str(minimal_docx))["restored"] and minimal_docx.read_bytes() == edited


def test_p1_save_just_before_undo_swap_is_kept(word, monkeypatch, minimal_docx):
    pre = minimal_docx.read_bytes()
    batch(minimal_docx, ("Fixture", "Alpha"))
    external = minimal_docx.read_bytes() + b"saved by Word at the boundary"
    _inject_save_before_replace(monkeypatch, minimal_docx, external)
    undone = undo_document(str(minimal_docx))
    assert minimal_docx.read_bytes() == pre
    assert undone["publish_warning"] and "WARNING" in undone["message"]
    assert undone["sha256_before"] == file_sha256_bytes(external).upper()
    assert ck.blob_path(file_sha256_bytes(external)).read_bytes() == external
    assert undo_document(str(minimal_docx))["restored"] and minimal_docx.read_bytes() == external


def _edit_lifecycle(doc: Path):
    from research_mcp.word import staged_document
    with staged_document(doc, doc) as stage:  # the real edit lifecycle: TemporaryDirectory stage beside the doc
        stage.write_bytes(stage.read_bytes() + b"agent edit")


def _batch_lifecycle(doc: Path):
    batch(doc, ("Fixture", "Alpha"))


def _undo_lifecycle(doc: Path):
    undo_document(str(doc))


LIFECYCLES = {"edit": (None, _edit_lifecycle, ["edit", "external"]),
              "batch": (None, _batch_lifecycle, ["edit", "external"]),
              "undo": (_batch_lifecycle, _undo_lifecycle, ["undo", "external", "edit"])}


def _displaced_left(doc: Path) -> list[Path]:
    return sorted(doc.parent.glob(".~research-mcp-displaced-*.docx"))


def _assert_recovered(doc: Path, external: bytes, kinds: list[str]):
    """The next call recovers the journal: X stored, history records it, the file is gone, undo brings X back."""
    hist = document_history(doc)
    assert [e["kind"] for e in hist["entries"]] == kinds
    assert hist["entries"][0]["recovered"] and hist["entries"][0]["sha_before"] == file_sha256_bytes(external).upper()
    assert ck.blob_path(file_sha256_bytes(external)).read_bytes() == external
    assert not _displaced_left(doc) and not list(ck.journal_dir().glob("*.json"))
    undo_document(str(doc))  # the interrupted write's own entry: restores the state it overwrote, X
    assert doc.read_bytes() == external


@pytest.mark.parametrize("life", sorted(LIFECYCLES))
def test_p1_interrupt_after_swap_keeps_displaced_save(word, monkeypatch, minimal_docx, life):
    setup, run, kinds = LIFECYCLES[life]
    if setup:
        setup(minimal_docx)
    external = minimal_docx.read_bytes() + b"saved by Word at the boundary"
    _inject_save_before_replace(monkeypatch, minimal_docx, external)
    real = ck.settle_displaced
    calls = []

    def interrupt(*args, **kwargs):
        if not calls:
            calls.append(args)
            raise KeyboardInterrupt("injected interrupt after swap before settle")
        return real(*args, **kwargs)

    monkeypatch.setattr(ck, "settle_displaced", interrupt)
    with pytest.raises(KeyboardInterrupt):
        run(minimal_docx)
    assert not [d for d in minimal_docx.parent.iterdir() if d.is_dir() and d.name.startswith("research-mcp-")]
    (left,) = _displaced_left(minimal_docx)  # survives the stage folder's cleanup
    assert left.read_bytes() == external
    (journal,) = ck.journal_dir().glob("*.json")
    assert json.loads(journal.read_text(encoding="utf-8"))["displaced"] == str(left)
    _assert_recovered(minimal_docx, external, kinds)


@pytest.mark.parametrize("life", sorted(LIFECYCLES))
def test_p1_failed_store_keeps_displaced_save_at_named_path(word, monkeypatch, minimal_docx, life):
    setup, run, kinds = LIFECYCLES[life]
    if setup:
        setup(minimal_docx)
    external = minimal_docx.read_bytes() + b"saved by Word at the boundary"
    _inject_save_before_replace(monkeypatch, minimal_docx, external)
    real_store, real_rename = ck.store_blob, os.replace

    def store_full(src, sha=None):
        if Path(src).name.startswith(".~research-mcp-displaced-"):
            raise OSError(28, "injected store full")
        return real_store(src, sha)

    def no_rename(src, dst, *args, **kwargs):  # any move of the displaced file (a "rescue") is refused too
        if Path(src).name.startswith(".~research-mcp-displaced-"):
            raise PermissionError(13, "injected rescue access denied")
        return real_rename(src, dst, *args, **kwargs)

    monkeypatch.setattr(ck, "store_blob", store_full)
    monkeypatch.setattr(os, "replace", no_rename)
    with pytest.raises(ck.Refusal) as info:
        run(minimal_docx)
    monkeypatch.setattr(ck, "store_blob", real_store)
    monkeypatch.setattr(os, "replace", real_rename)
    named = Path(info.value.details["displaced"])
    assert named.is_file() and named.read_bytes() == external and str(named) in str(info.value)
    assert _displaced_left(minimal_docx) == [named]
    assert not [d for d in minimal_docx.parent.iterdir() if d.is_dir() and d.name.startswith("research-mcp-")]
    (journal,) = ck.journal_dir().glob("*.json")
    assert json.loads(journal.read_text(encoding="utf-8"))["displaced"] == str(named)
    _assert_recovered(minimal_docx, external, kinds)


def _publish_with_stuck_recovery(monkeypatch, doc: Path):
    """A, then an outside save X lands just before an MCP batch publishes B, and storing X keeps failing.
    Returns (X, B, the sidecar path, a function that makes the store work again)."""
    external = doc.read_bytes() + b"saved by Word at the boundary"
    _inject_save_before_replace(monkeypatch, doc, external)
    real_store = ck.store_blob

    def store_full(src, sha=None):  # full for new states; one already stored needs no room
        if Path(src).name.startswith(".~research-mcp-displaced-") and not ck.has_blob(sha or file_sha256(src)):
            raise OSError(28, "injected store full")
        return real_store(src, sha)

    monkeypatch.setattr(ck, "store_blob", store_full)
    with pytest.raises(ck.Refusal) as info:
        _batch_lifecycle(doc)
    return external, doc.read_bytes(), Path(info.value.details["displaced"]), \
        lambda: monkeypatch.setattr(ck, "store_blob", real_store)


def _assert_pending_refuses(doc: Path, sidecar: Path, tmp_path: Path):
    """While the earlier write is unresolved: every write on this document refuses; history and others work."""
    b = doc.read_bytes()
    refused = batch(doc, ("Alpha", "Beta"))
    assert refused.status == "refused" and str(sidecar) in refused.message and doc.read_bytes() == b
    for call in (lambda: undo_document(str(doc)), lambda: checkpoint_document(doc, "x"),
                 lambda: _edit_lifecycle(doc)):
        with pytest.raises(ck.Refusal) as info:
            call()
        assert info.value.code == "recovery_pending" and str(sidecar) in str(info.value)
        assert info.value.details["pending"][0]["displaced"] == str(sidecar)
    assert doc.read_bytes() == b and sidecar.is_file()
    hist = document_history(doc)
    assert [p["displaced"] for p in hist["recovery_pending"]] == [str(sidecar)]
    other = tmp_path / "other.docx"
    other.write_bytes(b)
    assert batch(other, ("Alpha", "Beta")).status == "complete"  # other documents are unaffected


def test_p1_pending_recovery_refuses_then_edit_and_undo_keep_order(word, monkeypatch, minimal_docx, tmp_path):
    external, b, sidecar, store_works = _publish_with_stuck_recovery(monkeypatch, minimal_docx)
    _assert_pending_refuses(minimal_docx, sidecar, tmp_path)
    store_works()
    assert batch(minimal_docx, ("Alpha", "Beta")).status == "complete"  # recovers first, then B -> C
    c = minimal_docx.read_bytes()
    assert not sidecar.exists() and ck.blob_path(file_sha256_bytes(external)).read_bytes() == external
    first = undo_document(str(minimal_docx))
    assert not first["forced"] and minimal_docx.read_bytes() == b
    second = undo_document(str(minimal_docx))  # reverses the undo: C, not the outside save X
    assert not second["forced"] and minimal_docx.read_bytes() == c
    kinds = [e["kind"] for e in document_history(minimal_docx)["entries"]]
    assert kinds == ["undo", "undo", "edit", "edit", "external"]
    assert document_history(minimal_docx)["recovery_pending"] == []


def test_p1_pending_recovery_refuses_then_later_edit_undoes_cleanly(word, monkeypatch, minimal_docx, tmp_path):
    external, b, sidecar, store_works = _publish_with_stuck_recovery(monkeypatch, minimal_docx)
    _assert_pending_refuses(minimal_docx, sidecar, tmp_path)
    store_works()
    assert batch(minimal_docx, ("Alpha", "Beta")).status == "complete"
    undone = undo_document(str(minimal_docx))  # not stale: the recovered write was recorded before this edit
    assert not undone["forced"] and minimal_docx.read_bytes() == b
    assert undo_document(str(minimal_docx), undo_token=document_history(minimal_docx)["entries"][2]["id"])[
        "restored"] and minimal_docx.read_bytes() == external  # the recovered write's own token restores X


def test_displaced_name_never_reuses_an_existing_file(tmp_path):
    doc = tmp_path / "d.docx"
    taken = ck.displaced_path(doc, "u-1")
    taken.write_bytes(b"earlier leftover")
    assert ck.displaced_path(doc, "u-1") != taken and taken.read_bytes() == b"earlier leftover"


def test_p2_logged_payload_unwraps_envelope_and_requires_success():
    from research_mcp.publication import _logged_payload
    flat = {"source": "a.docx", "output": "a.docx", "sha256_before": "0" * 64, "sha256_after": "1" * 64}
    assert _logged_payload(flat) == flat
    env = {"meta": None, "content": [], "structured_content": flat, "is_error": False, "result_type": "complete"}
    assert _logged_payload(env) == flat
    assert _logged_payload(dict(env, is_error=True)) is None
    assert _logged_payload({"content": [{"type": "text", "text": json.dumps(flat)}], "is_error": False}) == flat
    assert _logged_payload(dict(env, structured_content=dict(flat, status="refused"))) is None
    assert _logged_payload(None) is None


def test_norm_path_converts_wsl_spelling_before_resolving():
    assert ck.norm_path("/mnt/c/Users/x/Paper.docx") == "c:\\users\\x\\paper.docx"
    assert ck.doc_key(Path("/mnt/c/Users/x/Paper.docx")) == ck.doc_key(Path("C:/Users/x/Paper.docx"))
