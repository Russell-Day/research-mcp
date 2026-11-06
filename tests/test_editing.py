from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
import os
import sys

import pytest

from research_mcp.audit import audit_docx
from research_mcp.editing import (
    _locate,
    _refuse_table_damage,
    _refuse_crossing_comment_boundary,
    _refuse_unsafe_candidate,
    insert_paragraph_to_version,
    insert_table_to_version,
    insert_text_to_version,
    locate_anchor,
    reply_comment_to_version,
    replace_text_to_version,
)
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.fixtures import create_table_fixture
from research_mcp.paragraphs import read_paragraphs


DATA = Path(__file__).resolve().parent / "data"
CITATION = DATA / "zotero_citation_verified.docx"
WORD_LIVE = pytest.mark.skipif(
    os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1",
    reason="set RESEARCH_MCP_WORD_LIVE=1 to run disposable live Word tests",
)


class _Find:
    def ClearFormatting(self):
        self.cleared = True

    def Execute(self):
        return True


def test_locate_explicitly_sets_every_find_option():
    finder = _Find()
    target = SimpleNamespace(Find=finder)
    document = SimpleNamespace(Paragraphs=lambda _: SimpleNamespace(Range=target))
    assert _locate(document, 1, "literal") is target
    assert finder.Text == "literal"
    assert finder.Forward is True
    assert finder.Wrap == 0
    assert finder.MatchCase is True
    assert finder.MatchWildcards is False
    assert finder.Format is False
    assert finder.MatchWholeWord is False
    assert finder.MatchSoundsLike is False
    assert finder.MatchAllWordForms is False


def test_locate_anchor_checks_word_resolved_text_and_find_limit():
    finder = _Find()
    target = SimpleNamespace(Find=finder, Text="literal")
    document = SimpleNamespace(Paragraphs=lambda _: SimpleNamespace(Range=target))

    assert locate_anchor(document, 1, "literal") is target

    target.Text = "different"
    with pytest.raises(RefusedUnsafe, match="resolved range does not hold"):
        locate_anchor(document, 1, "literal")

    with pytest.raises(RefusedUnsafe, match="cannot exceed 255"):
        locate_anchor(document, 1, "x" * 256)


def test_crossing_a_comment_boundary_is_refused_but_containment_is_allowed():
    class Comments:
        Count = 1

        def __call__(self, _):
            return SimpleNamespace(Scope=SimpleNamespace(Start=10, End=20))

    document = SimpleNamespace(Comments=Comments())
    _refuse_crossing_comment_boundary(document, SimpleNamespace(Start=12, End=18))
    with pytest.raises(RefusedUnsafe, match="comment anchor boundary"):
        _refuse_crossing_comment_boundary(document, SimpleNamespace(Start=8, End=12))
    with pytest.raises(RefusedUnsafe, match="comment anchor boundary"):
        _refuse_crossing_comment_boundary(document, SimpleNamespace(Start=18, End=22))


@pytest.mark.parametrize(
    ("attribute", "message"),
    [
        ("Fields", "field"),
        ("ContentControls", "content control"),
        ("Hyperlinks", "hyperlink"),
        ("Revisions", "tracked revision"),
        ("InlineShapes", "inline drawing"),
    ],
)
def test_unsafe_range_collections_are_refused(attribute, message):
    empty = SimpleNamespace(Count=0)
    candidate = SimpleNamespace(
        Fields=empty,
        ContentControls=empty,
        Hyperlinks=empty,
        Revisions=empty,
        InlineShapes=empty,
    )
    setattr(candidate, attribute, SimpleNamespace(Count=1))
    document = SimpleNamespace(Comments=SimpleNamespace(Count=0))
    with pytest.raises(RefusedUnsafe, match=message):
        _refuse_unsafe_candidate(document, candidate)


def test_insert_table_rejects_ragged_grid_before_opening_word():
    with pytest.raises(RefusedUnsafe, match="same width"):
        insert_table_to_version(
            "missing.docx",
            "unused.docx",
            after_paragraph_index=1,
            headers=["A", "B"],
            rows=[["one"]],
            descriptive_columns=[1],
        )


def test_reply_comment_rejects_unknown_disposition_before_opening_word():
    with pytest.raises(RefusedUnsafe, match="unsupported comment disposition"):
        reply_comment_to_version(
            "missing.docx",
            "unused.docx",
            comment_id="ABC",
            disposition="resolved",
            reply_text="Done.",
        )


@pytest.mark.word_live
@WORD_LIVE
@pytest.mark.skipif(sys.platform != "win32" or not CITATION.exists(), reason="requires Word fixture on Windows")
def test_live_replacement_before_and_after_large_zotero_field(tmp_path):
    before_output = tmp_path / "before.docx"
    before = replace_text_to_version(
        CITATION,
        before_output,
        paragraph_index=1,
        find="Disposable",
        replace="Temporary",
    )
    assert before["text_after"] == "Temporary Zotero integration source(1)."
    assert audit_docx(before_output).zotero_item_fields == 1

    after_output = tmp_path / "after.docx"
    after = replace_text_to_version(
        CITATION,
        after_output,
        paragraph_index=1,
        find=".",
        replace="!",
    )
    assert after["text_after"] == "Disposable Zotero integration source(1)!"
    assert read_paragraphs(after_output)[0].text.endswith("(1)!")
    assert audit_docx(after_output).zotero_item_fields == 1


@pytest.mark.word_live
@WORD_LIVE
@pytest.mark.skipif(sys.platform != "win32" or not CITATION.exists(), reason="requires Word fixture on Windows")
def test_live_replacement_of_field_result_is_refused(tmp_path):
    output = tmp_path / "field-result.docx"
    with pytest.raises(RefusedUnsafe, match="overlaps a field"):
        replace_text_to_version(
            CITATION,
            output,
            paragraph_index=1,
            find="(1)",
            replace="(2)",
        )
    assert audit_docx(output).zotero_item_fields == 1


@pytest.mark.word_live
@WORD_LIVE
@pytest.mark.skipif(sys.platform != "win32" or not CITATION.exists(), reason="requires Word fixture on Windows")
@pytest.mark.parametrize(
    ("position", "text", "expected"),
    [
        ("before", "verified ", "Disposable Zotero integration verified source(1)."),
        ("after", " [checked]", "Disposable Zotero integration source [checked](1)."),
    ],
)
def test_live_insert_text_around_anchor_preserves_zotero_field(tmp_path, position, text, expected):
    output = tmp_path / f"insert-{position}.docx"
    anchor = "source"
    result = insert_text_to_version(
        CITATION,
        output,
        paragraph_index=1,
        anchor=anchor,
        position=position,
        text=text,
    )
    assert result["text_after"] == expected
    assert audit_docx(output).zotero_item_fields == 1


def test_every_mcp_reply_disposition_has_a_label():
    """The MCP schema offers accepted/rejected/addressed; each must reach a label (chk_203f50a3774347b5)."""
    from research_mcp.editing import _DISPOSITION_LABELS
    from research_mcp.edit_models import ReplyCommentOperation

    offered = ReplyCommentOperation.model_fields["disposition"].annotation.__args__
    assert set(offered) <= set(_DISPOSITION_LABELS)

# A document whose tables never follow a top-level paragraph directly cannot reach this boundary.
# The synthetic fixture can: paragraph 2 is
# body prose and paragraph 3 is the first cell of a 7x3 table.
@pytest.mark.word_live
@WORD_LIVE
@pytest.mark.skipif(sys.platform != "win32", reason="requires Word")
def test_live_insert_paragraph_before_a_table_stays_out_of_the_table(tmp_path):
    fixture = tmp_path / "fixture.docx"
    create_table_fixture(fixture)
    before = audit_docx(fixture)
    assert read_paragraphs(fixture)[2].in_table, "fixture must abut a table at 3"

    result = insert_paragraph_to_version(
        fixture,
        tmp_path / "paragraph-before-table.docx",
        after_paragraph_index=2,
        text="Body prose, not a table cell.",
    )
    after = audit_docx(result["output"])
    saved = read_paragraphs(result["output"])

    # Before the fix this reported success with the text inside cell (1,1).
    assert saved[result["new_paragraph_index"] - 1].in_table is False
    assert saved[result["new_paragraph_index"] - 1].text == "Body prose, not a table cell."
    assert [table.cell_hashes for table in after.tables] == [
        table.cell_hashes for table in before.tables
    ]
    assert after.table_count == before.table_count


@pytest.mark.word_live
@WORD_LIVE
@pytest.mark.skipif(sys.platform != "win32", reason="requires Word")
def test_live_insert_table_before_an_existing_table_preserves_it(tmp_path):
    fixture = tmp_path / "fixture.docx"
    create_table_fixture(fixture)
    before = audit_docx(fixture)

    result = insert_table_to_version(
        fixture,
        tmp_path / "table-before-table.docx",
        after_paragraph_index=2,
        headers=["Measure", "Value"],
        rows=[["Accuracy", "0.95"], ["F1", "0.93"]],
        descriptive_columns=[1],
    )
    after = audit_docx(result["output"])

    # Before the fix Word nested the new table inside cell (1,1) of the
    # existing one and verification refused the whole operation.
    assert after.table_count == before.table_count + 1
    remaining = list(after.tables)
    remaining.pop(result["table_index"] - 1)
    assert [table.cell_hashes for table in remaining] == [
        table.cell_hashes for table in before.tables
    ]
    assert (result["rows"], result["columns"]) == (3, 2)


def _audit_with_tables(*cell_hashes: tuple[str, ...]) -> SimpleNamespace:
    return SimpleNamespace(
        tables=tuple(
            SimpleNamespace(
                rows=1,
                columns=len(hashes),
                cell_hashes=hashes,
                vertical_merges=(),
                horizontal_spans=(),
            )
            for hashes in cell_hashes
        )
    )


# These drive the guard that the opt-in live tests cover, without Word. The bug
# it guards against reported success, so it must not be reachable only by a
# suite that is skipped by default.
def test_paragraph_insertion_into_a_table_cell_is_refused():
    saved = [
        SimpleNamespace(in_table=False, text="Prose."),
        SimpleNamespace(in_table=True, text="Text that went into a cell."),
    ]
    audit = _audit_with_tables(("a", "b"))
    with pytest.raises(VerificationFailed, match="landed inside a table"):
        _refuse_table_damage(saved, 2, audit, audit)


def test_paragraph_insertion_that_changed_a_table_is_refused():
    saved = [SimpleNamespace(in_table=False, text="Prose.")]
    before = _audit_with_tables(("a", "b"))
    after = _audit_with_tables(("a", "CHANGED"))
    with pytest.raises(VerificationFailed, match="changed an existing table"):
        _refuse_table_damage(saved, 1, before, after)


def test_body_paragraph_insertion_leaving_tables_alone_is_accepted():
    saved = [SimpleNamespace(in_table=False, text="Prose.")]
    audit = _audit_with_tables(("a", "b"))
    _refuse_table_damage(saved, 1, audit, audit)
