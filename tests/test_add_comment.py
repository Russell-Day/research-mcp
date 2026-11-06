from pathlib import Path
import os

import pytest
from pydantic import ValidationError

from research_mcp.audit import audit_docx, file_sha256
from research_mcp.comments import inspect_comments
from research_mcp.edit_models import AddCommentOperation, EditDocumentRequest
from research_mcp.edit_batch import edit_document
from research_mcp.editing import add_comment_to_version
from research_mcp.errors import RefusedUnsafe
from research_mcp.paragraphs import read_snapshot
from research_mcp.server import MCPEditOperation, _strict_operation_payload


def test_comment_envelope_and_validation():
    op = MCPEditOperation(op_id="c1", type="add_comment", paragraph_index=2,
                          anchor="Evidence", text="Please verify.", author="Sam")
    strict = _strict_operation_payload(op)
    assert isinstance(strict, AddCommentOperation)
    assert strict.author == "Sam"
    with pytest.raises(ValidationError):
        AddCommentOperation(op_id="c", type="add_comment", paragraph_index=1,
                            anchor="", text="test")


def test_empty_comment_refused_before_word():
    with pytest.raises(RefusedUnsafe, match="must not be empty"):
        add_comment_to_version("unused.docx", "unused2.docx", paragraph_index=1,
                               anchor="Anchor", text=" ")


@pytest.mark.word_live
@pytest.mark.skipif(os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1", reason="requires Word")
def test_add_comments_in_table_preserves_prose_and_existing_thread(tmp_path):
    from docx import Document
    path = tmp_path / "review.docx"
    doc = Document()
    doc.add_paragraph("Introduction anchor.")
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "R2.1a. Screening criteria"
    doc.add_paragraph("")  # Word requires a final body paragraph after a table.
    doc.save(path)
    before = audit_docx(path)
    for i, anchor in enumerate(["R2.1a.", "Screening criteria"]):
        req = EditDocumentRequest(source=str(path), source_sha256=file_sha256(path),
            operations=[dict(op_id=f"c{i}", type="add_comment", paragraph_index=2,
                             anchor=anchor, text=f"Author note {i}\nSecond line\n\nLast paragraph", author="Sam", initials="S")])
        result = edit_document(req)
        assert result.status == "complete", result
    after = audit_docx(path)
    assert after.prose_text_sha256 == before.prose_text_sha256
    assert after.table_count == 1
    assert after.live_comment_count == 2
    threads = inspect_comments(read_snapshot(path))
    assert [t.text for t in threads] == [f"Author note {i}\nSecond line\n\nLast paragraph" for i in range(2)]
    assert all(t.author == "Sam" and not t.done for t in threads)
    assert not path.with_suffix(".docx.bak").exists()
    assert result.undo_token


@pytest.mark.word_live
@pytest.mark.skipif(os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1", reason="requires Word")
def test_bad_anchor_batch_rolls_back(tmp_path):
    from docx import Document
    path = tmp_path / "review.docx"
    doc = Document(); doc.add_paragraph("One anchor."); doc.save(path)
    before = file_sha256(path)
    req = EditDocumentRequest(source=str(path), source_sha256=before, operations=[
        dict(op_id="a", type="add_comment", paragraph_index=1, anchor="One", text="Good"),
        dict(op_id="b", type="add_comment", paragraph_index=1, anchor="Missing", text="Bad"),
    ])
    result = edit_document(req)
    assert result.status != "complete"
    assert file_sha256(path) == before
