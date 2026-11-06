import zipfile
from pathlib import Path
import pytest
from research_mcp.audit import file_sha256
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.style_tools import style_text
from research_mcp.styling import apply_journal_style


def test_preview_has_exact_scope_and_never_starts_word(monkeypatch, minimal_docx):
    def must_not_run(*args, **kwargs):
        raise AssertionError("Preview started Word")
    monkeypatch.setattr("research_mcp.style_tools.run_operation", must_not_run)
    before = minimal_docx.read_bytes()
    result = style_text(str(minimal_docx), file_sha256(minimal_docx), 1, "Fixture", bold=True)
    assert result["status"] == "preview"
    assert result["scope"] == "exact_text"
    assert result["context_after"] == " paragraph."
    assert minimal_docx.read_bytes() == before


def test_preview_stale_hash_refuses(minimal_docx):
    with pytest.raises(RefusedUnsafe, match="Source changed"):
        style_text(str(minimal_docx), "0"*64, 1, "Fixture", bold=True)


def test_whole_paragraph_requires_explicit_scope(minimal_docx):
    with pytest.raises(RefusedUnsafe, match="whole-paragraph"):
        style_text(str(minimal_docx), file_sha256(minimal_docx), 1, "Fixture paragraph.", bold=True)
    result = style_text(str(minimal_docx), file_sha256(minimal_docx), 1, "Fixture paragraph.", bold=False, allow_whole_paragraph=True)
    assert result["scope"] == "whole_paragraph"


def test_failed_inline_verification_never_publishes(monkeypatch, minimal_docx):
    before = minimal_docx.read_bytes()
    def fail(*args, **kwargs):
        raise VerificationFailed("Whole paragraph was accidentally bolded")
    monkeypatch.setattr("research_mcp.style_tools.run_operation", fail)
    with pytest.raises(VerificationFailed):
        style_text(str(minimal_docx), file_sha256(minimal_docx), 1, "Fixture", bold=True, preview_only=False)
    assert minimal_docx.read_bytes() == before
    assert not minimal_docx.with_name(minimal_docx.name + ".bak").exists()


def test_layout_default_refuses_accidental_emphasis_change(monkeypatch, minimal_docx):
    before = minimal_docx.read_bytes()
    def wrong_style(kind, payload, **kwargs):
        path = Path(payload["source"])
        with zipfile.ZipFile(path) as archive:
            entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
        with zipfile.ZipFile(path, "w") as archive:
            for info, data in entries:
                if info.filename == "word/document.xml":
                    data = data.replace(b"<w:r>", b"<w:r><w:rPr><w:b/></w:rPr>", 1)
                archive.writestr(info, data)
        return {"paragraphs_styled": 1}
    monkeypatch.setattr("research_mcp.operation_process.run_operation", wrong_style)
    with pytest.raises(VerificationFailed, match="changed bold or italic"):
        apply_journal_style(minimal_docx)
    assert minimal_docx.read_bytes() == before


def test_source_change_between_inspection_and_staging_is_preserved(monkeypatch, minimal_docx):
    from contextlib import contextmanager
    from research_mcp.word import staged_document as real_stage
    original = minimal_docx.read_bytes()
    @contextmanager
    def external_change(source, output):
        minimal_docx.write_bytes(original + b"external save")
        with real_stage(source, output) as stage:
            yield stage
    monkeypatch.setattr("research_mcp.style_tools.staged_document", external_change)
    with pytest.raises(RefusedUnsafe, match="changed while staging"):
        style_text(str(minimal_docx), file_sha256(minimal_docx), 1, "Fixture", bold=True, preview_only=False)
    assert minimal_docx.read_bytes() == original + b"external save"
