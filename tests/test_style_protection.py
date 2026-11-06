"""Offline tests for the styling post-save protection gate."""

from dataclasses import replace

import pytest

from research_mcp.audit import Audit, TableAudit
from research_mcp.errors import VerificationFailed
from research_mcp.styling import _verify_protected_structure


def _audit() -> Audit:
    table = TableAudit(
        index=1,
        rows=2,
        columns=2,
        cell_hashes=("a" * 64, "b" * 64, "c" * 64, "d" * 64),
    )
    return Audit(
        path="fixture.docx",
        sha256="0" * 64,
        body_text_sha256="body",
        prose_text_sha256="prose",
        paragraph_count=3,
        table_count=1,
        tables=(table,),
        bookmark_names=("anchor",),
        bookmark_references={"anchor": ()},
        field_count=2,
        zotero_item_fields=1,
        zotero_bibliography_fields=1,
        zotero_pref_properties=1,
        revision_count=0,
        comment_count=1,
        live_comment_count=1,
        orphan_comment_ids=(),
        drawing_count=0,
        media_count=0,
        live_media_count=0,
        dangling_media_targets=(),
    )


def test_unchanged_audits_pass():
    audit = _audit()
    _verify_protected_structure(audit, audit)


def test_same_table_count_but_edited_cell_is_rejected():
    before = _audit()
    edited = replace(before.tables[0], cell_hashes=("x" * 64, *before.tables[0].cell_hashes[1:]))

    with pytest.raises(VerificationFailed, match="tables changed"):
        _verify_protected_structure(before, replace(before, tables=(edited,)))


def test_citation_loss_is_rejected_even_when_field_count_changes_consistently():
    before = _audit()
    after = replace(before, field_count=1, zotero_item_fields=0)

    with pytest.raises(VerificationFailed, match="field_count changed"):
        _verify_protected_structure(before, after)


@pytest.mark.parametrize("field", ["body_text_sha256", "prose_text_sha256"])
def test_body_or_prose_change_is_rejected(field):
    before = _audit()

    with pytest.raises(VerificationFailed, match=field):
        _verify_protected_structure(before, replace(before, **{field: "changed"}))


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("bookmark_names", ()),
        ("live_comment_count", 0),
        ("revision_count", 1),
    ],
)
def test_bookmarks_comments_and_revisions_are_rejected(field, value):
    before = _audit()

    with pytest.raises(VerificationFailed, match=field):
        _verify_protected_structure(before, replace(before, **{field: value}))


def test_structural_errors_are_rejected():
    before = _audit()

    with pytest.raises(VerificationFailed, match="structural errors"):
        _verify_protected_structure(before, replace(before, structural_errors=("bad field",)))
