from pathlib import Path

import pytest

from research_mcp.citation_ops import insert_citation_to_version
from research_mcp.errors import RefusedUnsafe


def test_citation_pass_requires_explicit_item_keys(tmp_path, minimal_docx):
    output = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="item_keys"):
        insert_citation_to_version(
            minimal_docx,
            output,
            item_donor=tmp_path / "missing.docx",
            paragraph_index=1,
            anchor="x",
            item_keys=(),
        )
    assert not output.exists()


@pytest.mark.parametrize(
    ("donor", "paragraph"),
    [(Path("bibliography.docx"), None), (None, 2)],
)
def test_citation_pass_requires_complete_bibliography_pair(
    tmp_path, minimal_docx, donor, paragraph
):
    with pytest.raises(RefusedUnsafe, match="supplied together"):
        insert_citation_to_version(
            minimal_docx,
            tmp_path / "out.docx",
            item_donor=tmp_path / "item.docx",
            paragraph_index=1,
            anchor="x",
            item_keys=("ITEMKEY1",),
            bibliography_donor=donor,
            bibliography_paragraph_index=paragraph,
        )


def test_citation_pass_refuses_overwrite_before_opening_word(tmp_path, minimal_docx):
    output = tmp_path / "existing.docx"
    output.write_bytes(b"do not replace")
    with pytest.raises(RefusedUnsafe, match="overwrite"):
        insert_citation_to_version(
            minimal_docx,
            output,
            item_donor=tmp_path / "item.docx",
            paragraph_index=1,
            anchor="x",
            item_keys=("ITEMKEY1",),
        )
    assert output.read_bytes() == b"do not replace"


def test_citation_pass_refuses_mismatched_source_sha256(tmp_path, minimal_docx):
    output = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="source document SHA-256 mismatch"):
        insert_citation_to_version(
            minimal_docx,
            output,
            item_donor=tmp_path / "item.docx",
            paragraph_index=1,
            anchor="x",
            item_keys=("ITEMKEY1",),
            source_sha256="0" * 64,
        )
    assert not output.exists()


def test_citation_pass_refuses_source_saved_after_sha_check(tmp_path, minimal_docx, monkeypatch):
    """Plan 602 S2 review r1 #2: the staged copy must match the caller's sha (a save between the check
    and staging used to be edited and reported under the old sha). Word is never opened."""
    import shutil

    from research_mcp import citation_ops
    from research_mcp.audit import file_sha256

    donor = tmp_path / "donor.docx"
    shutil.copyfile(minimal_docx, donor)
    original = file_sha256(minimal_docx)
    real_audit = citation_ops.audit_docx
    calls = []

    def audit(path):
        result = real_audit(path)
        if not calls:
            with open(minimal_docx, "ab") as handle:  # another save lands right after the check
                handle.write(b"\0")
        calls.append(path)
        return result

    def no_word(*args, **kwargs):
        raise AssertionError("Word must not open")

    monkeypatch.setattr(citation_ops, "audit_docx", audit)
    monkeypatch.setattr(citation_ops, "IsolatedWordDocument", no_word)
    output = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="source changed after its SHA-256 was checked"):
        insert_citation_to_version(minimal_docx, output, item_donor=donor, paragraph_index=1, anchor="x",
                                   item_keys=("ITEMKEY1",), source_sha256=original)
    assert not output.exists()
