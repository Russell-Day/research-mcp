from __future__ import annotations

from pathlib import Path
import os
import shutil
import sys

import pytest

from research_mcp.audit import audit_docx, file_sha256
from research_mcp.edit_batch import edit_document
from research_mcp.edit_models import EditDocumentRequest
from research_mcp.errors import RefusedUnsafe


DATA = Path(__file__).resolve().parent / "data"
CITATION = DATA / "zotero_citation_verified.docx"
WORD_LIVE = pytest.mark.skipif(
    os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1",
    reason="set RESEARCH_MCP_WORD_LIVE=1 to run disposable live Word tests",
)


def _request(source: Path, output: Path, operations: list[dict], *, sha: str | None = None):
    return EditDocumentRequest(
        source=str(source),
        source_sha256=sha or file_sha256(source),
        output=str(output),
        operations=operations,
        allow_partial=True,
    )


def test_verified_partial_publishes_success_failure_success(monkeypatch, minimal_docx, tmp_path):
    def fake_apply(source, output, operation):
        shutil.copy2(source, output)
        if operation.op_id == "bad":
            raise RefusedUnsafe("deliberate refusal")
        return {
            "output_sha256": file_sha256(output),
            "paragraph_count": audit_docx(output).paragraph_count,
            "indices_after_shifted": False,
        }

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)
    output = tmp_path / "published.docx"
    request = _request(
        minimal_docx,
        output,
        [
            {"op_id": "first", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"},
            {"op_id": "bad", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "c"},
            {"op_id": "last", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "d"},
        ],
    )
    result = edit_document(request)
    assert result.status == "partial"
    assert [item.status for item in result.operations] == ["applied", "failed", "applied"]
    assert result.counts == {"applied": 2, "failed": 1, "skipped_dependency": 0}
    assert output.exists()
    assert result.output_sha256 == file_sha256(output)
    assert result.retry.source == str(output.resolve())
    assert result.retry.source_sha256 == result.output_sha256
    assert not output.with_name(output.name + ".research-mcp.lock").exists()


def test_failed_dependency_is_skipped(monkeypatch, minimal_docx, tmp_path):
    def fake_apply(source, output, operation):
        if operation.op_id == "bad":
            raise RefusedUnsafe("no")
        shutil.copy2(source, output)
        return {
            "output_sha256": file_sha256(output),
            "paragraph_count": audit_docx(output).paragraph_count,
            "indices_after_shifted": False,
        }

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)
    result = edit_document(
        _request(
            minimal_docx,
            tmp_path / "unused.docx",
            [
                {"op_id": "bad", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"},
                {
                    "op_id": "reply",
                    "type": "reply_comment",
                    "comment_id": "ABC",
                    "disposition": "addressed",
                    "reply_text": "Handled.",
                    "depends_on": ["bad"],
                },
            ],
        )
    )
    assert result.status == "no_changes"
    assert [item.status for item in result.operations] == ["failed", "skipped_dependency"]
    assert result.output is None


def test_structural_success_rebases_later_original_index(monkeypatch, minimal_docx, tmp_path):
    seen = []

    def fake_apply(source, output, operation):
        seen.append(operation)
        shutil.copy2(source, output)
        count = audit_docx(output).paragraph_count
        return {
            "output_sha256": file_sha256(output),
            "paragraph_count": count + (1 if operation.type == "insert_paragraph" else 0),
            "indices_after_shifted": operation.type == "insert_paragraph",
        }

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)
    result = edit_document(
        _request(
            minimal_docx,
            tmp_path / "rebased.docx",
            [
                {"op_id": "paragraph", "type": "insert_paragraph", "after_paragraph_index": 1, "text": "New"},
                {"op_id": "edit", "type": "replace_text", "paragraph_index": 2, "find": "a", "replace": "b"},
            ],
        )
    )
    assert result.status == "complete"
    assert seen[1].paragraph_index == 3
    assert result.operations[1].details["requested_index"] == 2
    assert result.operations[1].details["resolved_index"] == 3


def test_sha_mismatch_refuses_without_reserving_output(minimal_docx, tmp_path):
    output = tmp_path / "never.docx"
    result = edit_document(
        _request(
            minimal_docx,
            output,
            [{"op_id": "edit", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"}],
            sha="0" * 64,
        )
    )
    assert result.status == "refused"
    assert result.retry.requires_reinspection is True
    assert not output.exists()
    assert not output.with_name(output.name + ".research-mcp.lock").exists()


def test_existing_output_is_never_overwritten(minimal_docx, tmp_path):
    output = tmp_path / "existing.docx"
    output.write_bytes(b"owned")
    result = edit_document(
        _request(
            minimal_docx,
            output,
            [{"op_id": "edit", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"}],
        )
    )
    assert result.status == "refused"
    assert output.read_bytes() == b"owned"


def test_inplace_edit_updates_source_directly(monkeypatch, minimal_docx, tmp_path):
    target = tmp_path / "working.docx"
    shutil.copy2(minimal_docx, target)
    before_sha = file_sha256(target)

    def fake_apply(source, output, operation):
        shutil.copy2(source, output)
        with output.open("ab") as f_out:
            f_out.write(b"modified")
        return {
            "output_sha256": file_sha256(output),
            "paragraph_count": audit_docx(output).paragraph_count,
            "indices_after_shifted": False,
        }

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)

    request = _request(
        target,
        target,
        [{"op_id": "op1", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"}],
    )
    result = edit_document(request)
    assert result.status == "complete"
    assert result.output == str(target.resolve())
    assert result.output_sha256 == file_sha256(target)
    assert result.output_sha256 != before_sha
    assert not target.with_name(target.name + ".research-mcp.lock").exists()


def test_inplace_edit_fails_cleanly_when_locked(monkeypatch, minimal_docx, tmp_path):
    target = tmp_path / "working.docx"
    shutil.copy2(minimal_docx, target)
    before_sha = file_sha256(target)

    def fake_apply(source, output, operation):
        # A real change: identical bytes are a no-op that never reaches os.replace.
        import zipfile
        with zipfile.ZipFile(source) as archive, zipfile.ZipFile(output, "w") as staged:
            for item in archive.infolist():
                data = archive.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(b"Fixture", b"Edited", 1)
                staged.writestr(item, data)
        return {
            "output_sha256": file_sha256(output),
            "paragraph_count": audit_docx(output).paragraph_count,
            "indices_after_shifted": False,
        }

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)

    def fake_replace(src, dst):
        raise PermissionError("[WinError 32] File is locked by Word")

    monkeypatch.setattr("os.replace", fake_replace)

    request = _request(
        target,
        target,
        [{"op_id": "op1", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"}],
    )
    result = edit_document(request)
    assert result.status == "refused"
    assert "locked by Word" in result.message
    assert "Original document was not replaced" in result.message
    assert file_sha256(target) == before_sha


@pytest.mark.word_live
@WORD_LIVE
@pytest.mark.skipif(sys.platform != "win32" or not CITATION.exists(), reason="requires Word citation fixture")
def test_live_partial_batch_publishes_both_safe_edits_around_zotero_field(tmp_path):
    output = tmp_path / "partial-live.docx"
    request = _request(
        CITATION,
        output,
        [
            {
                "op_id": "before",
                "type": "replace_text",
                "paragraph_index": 1,
                "find": "Disposable",
                "replace": "Temporary",
            },
            {
                "op_id": "refused",
                "type": "replace_text",
                "paragraph_index": 1,
                "find": "text that is absent",
                "replace": "never published",
            },
            {
                "op_id": "after",
                "type": "insert_text",
                "paragraph_index": 1,
                "anchor": "source",
                "position": "after",
                "text": " verified",
            },
        ],
    )
    result = edit_document(request)
    assert result.status == "partial"
    assert [item.status for item in result.operations] == ["applied", "failed", "applied"]
    assert result.retry.source == str(output.resolve())
    assert result.retry.source_sha256 == file_sha256(output)
    assert audit_docx(output).zotero_item_fields == 1
    from research_mcp.paragraphs import read_paragraphs

    assert read_paragraphs(output)[0].text == (
        "Temporary Zotero integration source verified(1)."
    )
