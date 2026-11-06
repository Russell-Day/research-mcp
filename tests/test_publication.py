from pathlib import Path

import pytest

from research_mcp.audit import file_sha256
from research_mcp.errors import RefusedUnsafe
from research_mcp.publication import publish_inplace
from research_mcp.word import staged_document


def files(tmp_path):
    target = tmp_path / "paper.docx"
    stage = tmp_path / "candidate.docx"
    target.write_bytes(b"original")
    stage.write_bytes(b"verified edit")
    return target, stage, file_sha256(target)


def test_success_stores_exact_pre_edit_state(tmp_path):
    from research_mcp import checkpoints as ck
    target, stage, sha = files(tmp_path)
    result = publish_inplace(stage, target, expected_sha256=sha)
    assert target.read_bytes() == b"verified edit"
    assert not target.with_name("paper.docx.bak").exists()
    assert result["undo_token"]
    assert result["sha256_before"] == sha
    assert result["sha256_after"] == file_sha256(target)
    assert ck.blob_path(sha).read_bytes() == b"original"


def test_failed_store_preserves_source_and_existing_backup(monkeypatch, tmp_path):
    from research_mcp import checkpoints as ck
    target, stage, sha = files(tmp_path)
    backup = target.with_name("paper.docx.bak")
    backup.write_bytes(b"older rollback")

    def fail_copy(source, destination):
        Path(destination).write_bytes(b"partial")
        raise OSError("disk full")

    monkeypatch.setattr("research_mcp.checkpoints.shutil.copyfile", fail_copy)
    with pytest.raises(RefusedUnsafe, match="disk full"):
        publish_inplace(stage, target, expected_sha256=sha)
    assert target.read_bytes() == b"original"
    assert backup.read_bytes() == b"older rollback"
    assert stage.read_bytes() == b"verified edit"
    assert ck.load(target)["entries"] == []
    assert not list(ck.blobs_dir().glob("*.tmp"))


def test_staging_preserves_external_change(tmp_path):
    target, _, _ = files(tmp_path)
    with pytest.raises(RefusedUnsafe, match="Source changed during editing"):
        with staged_document(target, target) as stage:
            stage.write_bytes(b"agent edit")
            target.write_bytes(b"user edit")
    assert target.read_bytes() == b"user edit"
    assert not target.with_name("paper.docx.bak").exists()


def test_change_during_store_is_detected(monkeypatch, tmp_path):
    target, stage, sha = files(tmp_path)

    def copy_then_change(source, destination):
        Path(destination).write_bytes(Path(source).read_bytes())
        target.write_bytes(b"new external content")

    monkeypatch.setattr("research_mcp.checkpoints.shutil.copyfile", copy_then_change)
    with pytest.raises(RefusedUnsafe, match="Source changed"):
        publish_inplace(stage, target, expected_sha256=sha)
    assert target.read_bytes() == b"new external content"


def test_inexact_store_copy_refuses_publish(monkeypatch, tmp_path):
    target, stage, sha = files(tmp_path)
    monkeypatch.setattr("research_mcp.checkpoints.shutil.copyfile", lambda src, dst: Path(dst).write_bytes(b"bad snapshot"))
    with pytest.raises(RefusedUnsafe, match="stored copy"):
        publish_inplace(stage, target, expected_sha256=sha)
    assert target.read_bytes() == b"original"


def test_locked_document_refused_before_staging(monkeypatch, tmp_path):
    target, _, _ = files(tmp_path)
    monkeypatch.setattr("research_mcp.word.is_file_locked", lambda p: (True, "owner lock present"))
    with pytest.raises(RefusedUnsafe, match="Save and close"):
        with staged_document(target, target):
            pytest.fail("must not yield a staging document")


def test_batch_preserves_external_change_and_requires_reinspection(monkeypatch, minimal_docx):
    from research_mcp.edit_batch import edit_document
    from research_mcp.edit_models import EditDocumentRequest
    original = minimal_docx.read_bytes()

    def apply(source, output, operation):
        output.write_bytes(original)
        minimal_docx.write_bytes(original + b"user saved concurrently")
        return {"paragraph_count": 8}

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", apply)
    request = EditDocumentRequest(
        source=str(minimal_docx), source_sha256=file_sha256(minimal_docx), output=str(minimal_docx),
        operations=[{"op_id": "edit", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"}],
    )
    result = edit_document(request)
    assert result.status == "refused"
    assert result.output is None
    assert result.retry.requires_reinspection is True
    assert minimal_docx.read_bytes() == original + b"user saved concurrently"
