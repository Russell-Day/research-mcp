from pathlib import Path
import shutil

from research_mcp.audit import file_sha256, audit_docx
from research_mcp.edit_batch import edit_document
from research_mcp.edit_models import EditDocumentRequest
from research_mcp.errors import RefusedUnsafe


def request(path, **kwargs):
    return EditDocumentRequest(source=str(path), source_sha256=file_sha256(path), operations=[
        dict(op_id="good", type="replace_text", paragraph_index=1, find="a", replace="b"),
        dict(op_id="bad", type="replace_text", paragraph_index=1, find="a", replace="c"),
    ], **kwargs)


def fake_apply(source, output, operation):
    if operation.op_id == "bad":
        raise RefusedUnsafe("anchor not found")
    shutil.copy2(source, output)
    return {"paragraph_count": audit_docx(output).paragraph_count}


def test_atomic_default_discards_verified_subset(monkeypatch, minimal_docx):
    original = minimal_docx.read_bytes()
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)
    result = edit_document(request(minimal_docx))
    assert result.status == "no_changes"
    assert [op.status for op in result.operations] == ["rolled_back", "failed"]
    assert result.counts["applied"] == 0
    assert result.counts["rolled_back"] == 1
    assert result.output is None and result.summary_diffs == []
    assert minimal_docx.read_bytes() == original
    assert not minimal_docx.with_name(minimal_docx.name + ".bak").exists()


def test_partial_requires_explicit_opt_in(monkeypatch, minimal_docx):
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)
    result = edit_document(request(minimal_docx, allow_partial=True))
    assert result.status == "partial"
    assert result.counts["applied"] == 1


def test_publication_refusal_does_not_report_staged_success(monkeypatch, minimal_docx):
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake_apply)
    def refuse(*args, **kwargs):
        raise RefusedUnsafe("external save")
    monkeypatch.setattr("research_mcp.edit_batch._publish", refuse)
    result = edit_document(request(minimal_docx, allow_partial=True))
    assert result.status == "refused"
    assert result.counts["applied"] == 0
    assert result.operations[0].status == "rolled_back"
    assert result.operations[0].details["published"] is False
    assert result.retry.requires_reinspection


def test_locked_source_never_starts_word(monkeypatch, minimal_docx):
    monkeypatch.setattr("research_mcp.edit_batch.is_file_locked", lambda _: (True, "Word owner lock"))
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", lambda *args: (_ for _ in ()).throw(AssertionError("must not start")))
    result = edit_document(request(minimal_docx))
    assert result.status == "refused"
    assert result.operations == []


def test_same_count_citation_identity_change_is_rejected(monkeypatch, tmp_path):
    import zipfile
    source = tmp_path / "citation.docx"
    shutil.copy2(Path(__file__).parent / "data" / "zotero_citation_verified.docx", source)
    original = source.read_bytes()
    def damage(source, output, operation):
        with zipfile.ZipFile(source) as archive, zipfile.ZipFile(output, "w") as target:
            for item in archive.infolist():
                data = archive.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(b'citationID', b'citationZZ')
                target.writestr(item, data)
        return {"paragraph_count": audit_docx(output).paragraph_count}
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", damage)
    result = edit_document(request(source))
    assert result.counts["applied"] == 0
    assert result.status == "no_changes"
    assert source.read_bytes() == original
    assert all(op.status == "failed" for op in result.operations)
