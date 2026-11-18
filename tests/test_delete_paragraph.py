"""Plan 864 S3c: the `delete_paragraph` edit_document op (tracked or plain, body and table cells).

Direct XML always (Word never opens). Tracked: each paragraph's live content in a new `w:del` by the author and
its mark deleted; a paragraph whose text is already a tracked deletion (any author) keeps that `w:del`
byte-identical. Untracked: the paragraphs are removed and later paragraph-indexed ops in the batch are shifted.
Reject and Accept are checked with the string engine (`resolve_revisions`).
"""
from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.errors import RefusedUnsafe
from research_mcp.paragraph_delete import delete_paragraph
from research_mcp.sections import canonical
from test_edit_feedback import feedback_dir  # noqa: F401 - a fixture
from test_revisions import document_xml, other_parts_identical, raw_revisions, run_batch, windows_only
from test_tracked_sections import (
    AUTHOR, LINE_PPR, NOW, SECT, STAMP, body, deleted_run, first_id, make, new_ids, para, resolved, rev, run, texts,
)

OLD = 'w:date="2026-10-10T17:37:00Z"'


def old_del(rev_id: int, text: str, author: str = "Claude") -> str:
    return f'<w:del w:id="{rev_id}" w:author="{author}" {OLD}>{deleted_run(text)}</w:del>'


EMPTY = para(LINE_PPR)


def delete(source: Path, out: Path, first: int, last: int | None = None, *, tracked: bool = True) -> dict:
    return delete_paragraph(source, out, paragraph_index=first, through_paragraph_index=last, tracked=tracked,
                            author=AUTHOR, now=NOW)


def assert_reject_is_input(source: Path, out: Path, tmp_path: Path, result: dict) -> None:
    back = resolved(out, tmp_path, "reject", new_ids(result))
    assert canonical(body(back)) == canonical(body(source))


def accept_all(out: Path, tmp_path: Path) -> list[str]:
    target = tmp_path / f"all-{out.stem}.docx"
    revisions.resolve_revisions(out, target, mode="accept", ids=all_ids(out))
    return texts(target)


def all_ids(path: Path) -> list[str]:
    return re.findall(r'<w:(?:ins|del|pPrChange|rPrChange)\b[^>]*?w:id="(\d+)"', document_xml(path))


# -- 1-4: body paragraphs -------------------------------------------------------------------------------------

BODY = para("", run("Before.")) + EMPTY + para("", run("After."))


def test_body_empty_paragraph_tracked(tmp_path):
    source = make(tmp_path, "empty", BODY)
    out = tmp_path / "empty-out.docx"
    n = first_id(source)
    result = delete(source, out, 2)
    want = f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr></w:p>"
    assert document_xml(out) == document_xml(source).replace(EMPTY, want, 1)
    assert (result["engine"], result["tracked"], result["author"]) == ("direct_xml", True, AUTHOR)
    assert (result["index_delta"], result["accepted_index_delta"]) == (0, -1)
    assert result["deleted"] == [{"paragraph_index": 2, "text": ""}]
    assert new_ids(result) == [str(n)]
    assert other_parts_identical(source, out)
    assert_reject_is_input(source, out, tmp_path, result)
    assert texts(resolved(out, tmp_path, "accept", new_ids(result))) == ["Before.", "After."]


def test_body_empty_paragraph_untracked(tmp_path):
    source = make(tmp_path, "empty-plain", BODY)
    out = tmp_path / "empty-plain-out.docx"
    result = delete(source, out, 2, tracked=False)
    assert document_xml(out) == document_xml(source).replace(EMPTY, "", 1)
    assert (result["index_delta"], result["accepted_index_delta"], result["revisions_added"]) == (-1, -1, [])
    assert result["index_shifts"] == [[3, -1]] and result["deleted_range"] == [2, 2]
    assert texts(out) == ["Before.", "After."]


@pytest.mark.parametrize("owner", ["Claude", "Reviewer B"])
def test_paragraph_wholly_an_existing_deletion(tmp_path, owner):
    """Case (a): the text is already a tracked deletion, the mark is not: the old w:del stays byte-identical."""

    old = old_del(950, "Old text.", owner)
    source = make(tmp_path, "olddel", para("", run("Before.")) + para(LINE_PPR, old) + para("", run("After.")))
    out = tmp_path / "olddel-out.docx"
    n = first_id(source)
    result = delete(source, out, 2)
    want = f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr>{old}</w:p>"
    assert document_xml(out) == document_xml(source).replace(para(LINE_PPR, old), want, 1)
    assert raw_revisions(source)[0] in raw_revisions(out)
    assert new_ids(result) == [str(n)]
    assert_reject_is_input(source, out, tmp_path, result)
    assert texts(resolved(out, tmp_path, "accept", new_ids(result))) == ["Before.", "After."]
    assert accept_all(out, tmp_path) == ["Before.", "After."]


def test_existing_deletion_in_the_middle_two_new_deletions(tmp_path):
    old = old_del(950, "gone")
    line = para("", run("Live a "), old, run(" live b"))
    source = make(tmp_path, "mid", para("", run("Before.")) + line + para("", run("After.")))
    out = tmp_path / "mid-out.docx"
    n = first_id(source)
    result = delete(source, out, 2)
    want = (f"<w:p><w:pPr><w:rPr>{rev('del', n)}</w:rPr></w:pPr>{rev('del', n + 1, deleted_run('Live a '))}"
            f"{old}{rev('del', n + 2, deleted_run(' live b'))}</w:p>")
    assert document_xml(out) == document_xml(source).replace(line, want, 1)
    assert new_ids(result) == [str(n), str(n + 1), str(n + 2)]
    assert_reject_is_input(source, out, tmp_path, result)
    assert texts(resolved(out, tmp_path, "accept", new_ids(result))) == ["Before.", "After."]


def test_range_of_three_body_paragraphs(tmp_path):
    old = old_del(950, "Already deleted.")
    three = EMPTY + para("", run("Plain text.")) + para("", old)
    source = make(tmp_path, "three", para("", run("Before.")) + three + para("", run("After.")))
    out = tmp_path / "three-out.docx"
    n = first_id(source)
    result = delete(source, out, 2, 4)
    want = (f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr></w:p>"
            f"<w:p><w:pPr><w:rPr>{rev('del', n + 1)}</w:rPr></w:pPr>{rev('del', n + 2, deleted_run('Plain text.'))}</w:p>"
            f"<w:p><w:pPr><w:rPr>{rev('del', n + 3)}</w:rPr></w:pPr>{old}</w:p>")
    assert document_xml(out) == document_xml(source).replace(three, want, 1)
    assert result["deleted"] == [{"paragraph_index": 2, "text": ""}, {"paragraph_index": 3, "text": "Plain text."},
                                 {"paragraph_index": 4, "text": ""}]
    assert (result["index_delta"], result["accepted_index_delta"]) == (0, -3)
    assert_reject_is_input(source, out, tmp_path, result)
    assert texts(resolved(out, tmp_path, "accept", new_ids(result))) == ["Before.", "After."]
    # untracked: a paragraph holding a tracked deletion would drop it
    with pytest.raises(RefusedUnsafe, match="paragraph 4 holds tracked deletion .*an untracked delete would drop"):
        delete(source, tmp_path / "three-plain.docx", 2, 4, tracked=False)
    plain = delete(source, tmp_path / "two-plain.docx", 2, 3, tracked=False)
    assert plain["index_delta"] == -2
    assert texts(tmp_path / "two-plain.docx") == ["Before.", "", "After."]


# -- 5: table cells ---------------------------------------------------------------------------------------------

def cell(width: int, *paragraphs: str) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{"".join(paragraphs)}</w:tc>'


def table(*rows: str) -> str:
    return ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/></w:tblGrid>' + "".join(rows) + "</w:tbl>")


CELL_A = [para(LINE_PPR, run("Q1")), EMPTY, EMPTY, EMPTY, para(LINE_PPR, run("A1")), EMPTY, EMPTY]
TABLE_BODY = (para("", run("Intro."))
              + table("<w:tr>" + cell(4000, *CELL_A) + cell(2000, para("", run("Only"))) + "</w:tr>",
                      "<w:tr>" + cell(4000, para("", run("R2")), EMPTY) + cell(2000, para("", run("x"))) + "</w:tr>")
              + para("", run("End.")))
# layout: 1 Intro. | 2-8 cell A | 9 Only | 10 row end | 11 R2, 12 "" | 13 x | 14 row end | 15 End.
TABLE_TEXTS = ["Intro.", "Q1", "", "", "", "A1", "", "", "Only", "", "R2", "", "x", "", "End."]


def test_table_middle_empty_paragraphs(tmp_path):
    source = make(tmp_path, "cellmid", TABLE_BODY)
    assert texts(source) == TABLE_TEXTS
    out = tmp_path / "cellmid-out.docx"
    n = first_id(source)
    result = delete(source, out, 3, 5)
    deleted = "".join(f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n + k)}</w:rPr></w:pPr></w:p>" for k in range(3))
    assert document_xml(out) == document_xml(source).replace(EMPTY * 3, deleted, 1)
    assert result["trailing"] is False and result["container"] == "table 1 row 1 cell 1"
    assert_reject_is_input(source, out, tmp_path, result)
    accepted = texts(resolved(out, tmp_path, "accept", new_ids(result)))
    assert accepted == TABLE_TEXTS[:2] + TABLE_TEXTS[5:]


def test_table_last_two_of_a_cell_words_way(tmp_path):
    source = make(tmp_path, "celllast", TABLE_BODY)
    out = tmp_path / "celllast-out.docx"
    n = first_id(source)
    result = delete(source, out, 7, 8)
    assert result["trailing"] is True and result["previous_paragraph_index"] == 6
    xml = document_xml(out)
    # the mark of "A1" (the line before) and of the first empty line are deleted; the cell's last mark stays
    want = (f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr>{run('A1')}</w:p>"
            f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n + 1)}</w:rPr></w:pPr></w:p>"
            f"{EMPTY}</w:tc>")
    assert want in xml
    assert_reject_is_input(source, out, tmp_path, result)
    accepted = texts(resolved(out, tmp_path, "accept", new_ids(result)))
    assert accepted == TABLE_TEXTS[:6] + TABLE_TEXTS[8:]
    plain = tmp_path / "celllast-plain.docx"
    assert delete(source, plain, 7, 8, tracked=False)["index_delta"] == -2
    assert texts(plain) == TABLE_TEXTS[:6] + TABLE_TEXTS[8:]


def test_table_last_lines_one_op_at_a_time_refused_with_hint(tmp_path):
    source = make(tmp_path, "celltail", TABLE_BODY)
    middle = tmp_path / "celltail-1.docx"
    delete(source, middle, 7)
    with pytest.raises(RefusedUnsafe, match=r"give them in one op .*ending at 8"):
        delete(middle, tmp_path / "celltail-2.docx", 8)


def test_table_last_line_already_a_deletion(tmp_path):
    """The cell's last line is a tracked deletion with a live mark: deleted Word's way, the old w:del kept."""

    old = old_del(950, "Old line.", "Reviewer B")
    last = para(LINE_PPR, old)
    source = make(tmp_path, "celldel", TABLE_BODY.replace(EMPTY * 2 + "</w:tc>", EMPTY + last + "</w:tc>", 1))
    out = tmp_path / "celldel-out.docx"
    n = first_id(source)
    result = delete(source, out, 8)
    assert result["trailing"] is True
    xml = document_xml(out)
    assert f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr></w:p>{last}</w:tc>" in xml
    assert new_ids(result) == [str(n)]
    assert_reject_is_input(source, out, tmp_path, result)
    assert texts(resolved(out, tmp_path, "accept", new_ids(result))) == TABLE_TEXTS[:7] + TABLE_TEXTS[8:]
    assert accept_all(out, tmp_path) == TABLE_TEXTS[:7] + TABLE_TEXTS[8:]


@pytest.mark.parametrize("first, last", [(9, 9), (11, 12), (2, 8)])
def test_table_whole_cell_refused(tmp_path, first, last):
    source = make(tmp_path, "cellonly", TABLE_BODY)
    out = tmp_path / "cellonly-out.docx"
    with pytest.raises(RefusedUnsafe, match="a cell keeps at least one paragraph"):
        delete(source, out, first, last)
    assert not out.exists()


def test_table_range_across_cells_refused_named(tmp_path):
    source = make(tmp_path, "cross", TABLE_BODY)
    with pytest.raises(RefusedUnsafe) as caught:
        delete(source, tmp_path / "cross-out.docx", 8, 9)
    message = str(caught.value)
    assert "paragraph 8 is in table 1 row 1 cell 1 and paragraph 9 in table 1 row 1 cell 2" in message
    assert "give one op per cell" in message
    with pytest.raises(RefusedUnsafe, match="paragraph 1 is in the body and paragraph 2 in table 1 row 1 cell 1"):
        delete(source, tmp_path / "cross-out.docx", 1, 3)
    with pytest.raises(RefusedUnsafe, match="paragraph 10 is a table row-end marker"):
        delete(source, tmp_path / "cross-out.docx", 10)


# -- 6: refusals, named ------------------------------------------------------------------------------------------

OTHER = 'w:author="Other Author" w:date="2026-10-10T17:37:00Z"'
NESTED = table("<w:tr>" + cell(4000, para("", run("Outer")),
                                table("<w:tr>" + cell(2000, para("", run("Inner"))) + "</w:tr>"),
                                para("", run("Tail"))) + cell(2000, para("", run("x"))) + "</w:tr>")


@pytest.mark.parametrize("line, named", [
    (para("", run("Text "), f'<w:ins w:id="901" {OTHER}><w:r><w:t>new</w:t></w:r></w:ins>'),
     r"holds a tracked revision \(w:ins id=901 by 'Other Author'.*accept or reject it first"),
    (para("", '<w:commentRangeStart w:id="0"/>', run("Text"), '<w:commentRangeEnd w:id="0"/>'), "holds a comment"),
    (para("", '<w:bookmarkStart w:id="5" w:name="mark"/>', run("Text"), '<w:bookmarkEnd w:id="5"/>'),
     "holds a bookmark"),
    (para("", '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>'
              '<w:r><w:fldChar w:fldCharType="end"/></w:r>'), "holds a field"),
    (para("", '<w:r><w:drawing/></w:r>'), "holds a drawing"),
    (para(SECT, run("Text")), "holds a section break"),
    (f'<w:p><w:pPr><w:rPr><w:del w:id="903" {OTHER}/></w:rPr></w:pPr>{run("Text")}</w:p>',
     r"holds a tracked revision \(w:del id=903"),
    (f'<w:p><w:pPr><w:rPr><w:ins w:id="904" {OTHER}/></w:rPr></w:pPr>{run("Text")}</w:p>',
     r"holds a tracked revision \(w:ins id=904"),
    (para("", f'<w:del w:id="905" {OTHER}><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:del>'),
     r"holds a field inside tracked deletion \(w:del id=905"),
    (para("", f'<w:del w:id="906" {OTHER}><w:r><w:rPr><w:b/><w:rPrChange w:id="907" {OTHER}><w:rPr/>'
              f'</w:rPrChange></w:rPr><w:delText>x</w:delText></w:r></w:del>'),
     "holds a tracked formatting change"),
])
def test_refusals_named(tmp_path, line, named):
    source = make(tmp_path, "refuse", para("", run("Before.")) + line + para("", run("After.")))
    out = tmp_path / "refuse-out.docx"
    with pytest.raises(RefusedUnsafe, match=named):
        delete(source, out, 2)
    assert not out.exists()


def test_nested_table_paragraph_refused(tmp_path):
    source = make(tmp_path, "nested", para("", run("Before.")) + NESTED + para("", run("After.")))
    index = texts(source).index("Inner") + 1
    with pytest.raises(RefusedUnsafe, match=f"paragraph {index} is inside a nested table"):
        delete(source, tmp_path / "nested-out.docx", index)


# -- 7: batches --------------------------------------------------------------------------------------------------

def rows_body(count: int) -> str:
    rows = "".join("<w:tr>" + cell(4000, para(LINE_PPR, run(f"Label {k}")), EMPTY, EMPTY, EMPTY, EMPTY,
                                  para(LINE_PPR, run(f"Tail {k}")))
                   + cell(2000, para("", run(f"Date {k}"))) + "</w:tr>" for k in range(count))
    return para("", run("Intro.")) + table(rows) + para("", run("Closing line."))


def batch_failure(result) -> str:
    return f"{result.message} {[(o.op_id, o.status, o.error) for o in result.operations]}"


@windows_only
def test_batch_twenty_tracked_deletes_then_a_replace(tmp_path):
    source = make(tmp_path, "twenty", rows_body(5))
    before = texts(source)
    # each row: Label (r), 4 empties, Tail, Date, row end -> 8 layout entries from 2 + 8k
    empties = [2 + 8 * k + j for k in range(5) for j in range(1, 5)]
    assert all(before[i - 1] == "" for i in empties) and len(empties) == 20
    # 20 ops per call (the batch cap; a leading set_track_changes does not count): 19 deletes + the replace,
    # then the 20th delete alone
    ops = [{"op_id": f"d{i}", "type": "delete_paragraph", "paragraph_index": i} for i in empties[:19]]
    closing = before.index("Closing line.") + 1
    ops.append({"op_id": "r", "type": "replace_paragraph", "paragraph_index": closing, "text": "Closing words."})
    result = run_batch(source, *ops)
    assert result.status == "complete", batch_failure(result)
    after = texts(source)
    assert after[closing - 1] == "Closing words." and len(after) == len(before)
    assert all(o.details["index_delta"] == 0 and o.details["tracked"] is True for o in result.operations[:19])
    last = run_batch(source, {"op_id": "d", "type": "delete_paragraph", "paragraph_index": empties[19]})
    assert last.status == "complete", batch_failure(last)
    accepted = tmp_path / "twenty-accepted.docx"
    revisions.resolve_revisions(source, accepted, mode="accept", ids=all_ids(source))
    assert texts(accepted) == [t if t != "Closing line." else "Closing words."
                               for i, t in enumerate(before, 1) if i not in empties]


@pytest.mark.skipif(os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1", reason="requires Word (RESEARCH_MCP_WORD_LIVE=1)")
def test_batch_untracked_later_ops_shifted(tmp_path):
    source = make(tmp_path, "shift", para("", run("One.")) + EMPTY + EMPTY + para("", run("Four.")))
    result = run_batch(source, {"op_id": "d", "type": "delete_paragraph", "paragraph_index": 2,
                                "through_paragraph_index": 3, "tracked": False},
                       {"op_id": "r", "type": "replace_paragraph", "paragraph_index": 4, "text": "Four, edited."})
    assert result.status == "complete", batch_failure(result)
    assert result.operations[0].details["index_delta"] == -2
    assert texts(source) == ["One.", "Four, edited."]


@windows_only
def test_batch_untracked_op_on_a_deleted_paragraph_refused(tmp_path):
    source = make(tmp_path, "stale", para("", run("One.")) + EMPTY + para("", run("Three.")))
    result = run_batch(source, {"op_id": "d", "type": "delete_paragraph", "paragraph_index": 2, "tracked": False},
                       {"op_id": "r", "type": "replace_paragraph", "paragraph_index": 2, "text": "Two."})
    assert result.status != "complete"
    assert texts(source) == ["One.", "", "Three."]


@windows_only
@pytest.mark.parametrize("later", [
    {"type": "insert_paragraph", "after_paragraph_index": 2, "text": "New."},
    {"type": "delete_paragraph", "paragraph_index": 1, "through_paragraph_index": 2, "tracked": False},
], ids=["insert-after-deleted", "range-ending-in-deleted"])
def test_batch_untracked_selector_naming_a_deleted_paragraph_refused(tmp_path, later):
    # review S3c: every source selector is checked, not only paragraph_index
    source = make(tmp_path, "stale-sel", para("", run("One.")) + EMPTY + para("", run("Three.")))
    result = run_batch(source, {"op_id": "d", "type": "delete_paragraph", "paragraph_index": 2, "tracked": False},
                       {"op_id": "x", **later})
    assert result.status != "complete"
    field = "after_paragraph_index" if "after_paragraph_index" in later else "through_paragraph_index"
    assert f"{field} 2 lies inside" in result.operations[1].error.message
    assert texts(source) == ["One.", "", "Three."]


@windows_only
def test_batch_untracked_range_spanning_a_deleted_paragraph_maps(tmp_path):
    source = make(tmp_path, "span", para("", run("One.")) + EMPTY + para("", run("Three.")) + para("", run("Four.")))
    result = run_batch(source, {"op_id": "d", "type": "delete_paragraph", "paragraph_index": 2, "tracked": False},
                       {"op_id": "d2", "type": "delete_paragraph", "paragraph_index": 1, "through_paragraph_index": 3,
                        "tracked": False})
    assert result.status == "complete", batch_failure(result)
    assert texts(source) == ["Four."]


@windows_only
def test_batch_allow_partial_failed_row_op_rolls_back_that_rows_delete(tmp_path):
    source = make(tmp_path, "partial", rows_body(2))
    before = texts(source)
    result = run_batch(source,
                       {"op_id": "d1", "type": "delete_paragraph", "paragraph_index": 3},
                       {"op_id": "bad", "type": "replace_text", "paragraph_index": 2, "find": "missing",
                        "replace": "x"},
                       {"op_id": "d2", "type": "delete_paragraph", "paragraph_index": 11},
                       allow_partial=True)
    status = {o.op_id: o.status for o in result.operations}
    assert status["d2"] == "applied" and status["d1"] != "applied", batch_failure(result)
    xml = document_xml(source)
    assert len(re.findall(r"<w:del\b", xml)) == 1  # row 2's mark only
    assert texts(source) == before


# -- 8: the Track Changes switch ---------------------------------------------------------------------------------

@windows_only
@pytest.mark.parametrize("switch, op_fields, tracked, author", [
    ({"on": True, "author": "Reviewer B"}, {}, True, "Reviewer B"),
    ({"on": True, "author": "Reviewer B"}, {"tracked": False}, False, None),
    ({"on": False, "author": "Reviewer B"}, {}, False, None),
    (None, {}, True, "Claude"),
    (None, {"author": "Editor"}, True, "Editor"),
])
def test_switch_decides_tracked_and_author(tmp_path, switch, op_fields, tracked, author):
    source = make(tmp_path, "switch", BODY)
    if switch is not None:
        on = run_batch(source, {"op_id": "s", "type": "set_track_changes", **switch})
        assert on.status == "complete", batch_failure(on)
    result = run_batch(source, {"op_id": "d", "type": "delete_paragraph", "paragraph_index": 2, **op_fields})
    assert result.status == "complete", batch_failure(result)
    details = result.operations[0].details
    assert details["tracked"] is tracked and details["author"] == author
    xml = document_xml(source)
    if tracked:
        assert re.search(rf'<w:del w:id="\d+" w:author="{author}"', xml)
        assert texts(source) == ["Before.", "", "After."]
    else:
        assert "<w:del " not in xml and texts(source) == ["Before.", "After."]


# -- 9: a revision-free CV-style document, Word never opened ------------------------------------------------------

@windows_only
def test_cv_fixture_without_word(tmp_path, monkeypatch):
    import win32com.client

    import research_mcp.edit_batch as edit_batch
    from research_mcp.fixtures import create_mixed_run_fixture

    cv = create_mixed_run_fixture(tmp_path / "cv.docx")
    from research_mcp.sections import read_layout

    layout = read_layout(cv)
    target = next(item for item in layout.paragraphs
                  if item.table is not None and not item.row_end and item.text
                  and len(layout.row_of(item.index).cells[item.cell]) > 1
                  and item.index != layout.row_of(item.index).cells[item.cell][-1])

    def no_word(*args, **kwargs):
        raise AssertionError("Word was opened")

    monkeypatch.setattr(edit_batch, "_apply_operation", edit_batch._apply_operation_local)
    monkeypatch.setattr(win32com.client, "DispatchEx", no_word)
    monkeypatch.setattr(win32com.client, "Dispatch", no_word)
    before = texts(cv)
    result = run_batch(cv, {"op_id": "d", "type": "delete_paragraph", "paragraph_index": target.index})
    assert result.status == "complete", batch_failure(result)
    assert result.operations[0].details["engine"] == "direct_xml"
    accepted = tmp_path / "cv-accepted.docx"
    revisions.resolve_revisions(cv, accepted, mode="accept", ids=all_ids(cv))
    assert texts(accepted) == before[:target.index - 1] + before[target.index:]


# -- 10: edit feedback by place ----------------------------------------------------------------------------------

@pytest.mark.parametrize("resolution, outcome", [("accept", "accepted"), ("reject", "rejected")])
def test_feedback_tracked_delete_resolved(tmp_path, feedback_dir, resolution, outcome):
    from test_edit_feedback import ASSISTANT, events, feedback, model_edit
    from test_edit_feedback_insert import user_resolve_all

    path = make(tmp_path, "fb")
    model_edit(path, {"type": "delete_paragraph", "paragraph_index": 9, "author": ASSISTANT})
    [prop] = [e for e in events(path) if e["event"] == "proposal" and e.get("deleted")]
    assert (prop["before"], prop["nb_prev"], prop["nb_next"]) == ("Second body.", "First body.", "Third body.")
    assert feedback(path)["edit_feedback"] is None  # pending while its w:del is there
    user_resolve_all(path, resolution)
    assert feedback(path)["edit_feedback"] == f"Since your last edit: 1 {outcome}. Call edit_feedback for details."



@pytest.mark.parametrize("tracked", [True, False])
def test_feedback_blank_paragraph_not_recorded(tmp_path, feedback_dir, tracked):
    # Lead decision (review S3c): a paragraph with no visible text (blank, or text already a tracked deletion)
    # is not a proposal: there is no text to judge it by, and an empty paragraph between its neighbours
    # decides nothing. Deleting it records nothing and warns nothing.
    from test_edit_feedback import ASSISTANT, events, feedback, model_edit
    from test_edit_feedback_insert import user_resolve_all

    path = make(tmp_path, "fb-blank", BODY)
    out = model_edit(path, {"type": "delete_paragraph", "paragraph_index": 2, "author": ASSISTANT,
                            "tracked": tracked})
    assert not out.get("feedback_warning")
    assert not [e for e in events(path) if e["event"] == "proposal" and e.get("deleted")]
    if tracked:
        user_resolve_all(path, "accept")
    assert feedback(path)["edit_feedback"] is None
