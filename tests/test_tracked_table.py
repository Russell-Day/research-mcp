"""Plan 916 S5: insert_table as direct XML, tracked (switch on / tracked: true) or plain on a revised document.

Word's way, measured in a private Word with Track Changes on (lead decision 8): the anchor's paragraph mark
becomes the inserted one, the caption and the table follow, and an empty paragraph after the table keeps the
anchor's original pPr; every row's w:trPr, every cell paragraph mark and every run carry a w:ins; tblPr and
tblGrid are not tracked; ids ascend in document order. Merges are plain vMerge plus one comment per block
(decision 3A). Reject is checked with the string engine (`resolve_revisions`) and the ElementTree engine
(`et_resolve`); the insert verifies itself with its own check, never these assertions. Synthetic fixtures only.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import editing, revisions, tracked_table
from research_mcp.audit import audit_docx, file_sha256
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.sections import canonical, read_layout
from research_mcp.tables import MergeGroup
from test_revisions import W, document_xml, raw_revisions, run_batch, windows_only
from test_revisions_insert_paragraph import STYLES, TABLE, make, para, run
from test_revisions_marks_moves import AUTHOR as GUEST
from test_tracked_rows import build, plain_row, table, tracked_row

CLAUDE = "Claude"
CAPTION_STYLES = STYLES.replace(
    "</w:styles>", '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/></w:style></w:styles>')
BODY = (para('<w:spacing w:after="120"/>', run("Alpha one.")) + para('<w:jc w:val="both"/>', run("Beta two."))
        + para("", run("Gamma three.")))
HEADERS = ("Group", "N", "Mean", "SD")
ROWS = (("A", "10", "1.5", "0.2"), ("A", "12", "1.7", "0.3"), ("B", "9", "1.1", "0.1"))
GRID = [list(HEADERS), *[list(row) for row in ROWS]]
CAPTION = "Table 1. Synthetic values."


REVIEWER = 'w:author="Reviewer B" w:date="2026-10-10T17:37:00Z"'
REVISED = (para("", run("Intro paragraph."))
           + para("", run("Prior work "), f'<w:ins w:id="5" {REVIEWER}>' + run("clearly ") + "</w:ins>",
                  f'<w:del w:id="6" {REVIEWER}><w:r><w:delText xml:space="preserve">never </w:delText></w:r></w:del>',
                  run("showed this."))
           + para("", run("Middle paragraph."))
           + TABLE
           + para("", run("After the table.")))


def body_of(path: Path) -> ET.Element:
    return ET.fromstring(document_xml(path)).find(W + "body")


def insert(source: Path, out: Path, *, tracked: bool | None = True, merges=(), caption: str | None = CAPTION,
           after: int = 2, author: str | None = CLAUDE) -> dict:
    return editing.insert_table_to_version(source, out, after_paragraph_index=after, headers=HEADERS, rows=ROWS,
                                           descriptive_columns=(1,), merges=merges, caption=caption,
                                           tracked=tracked, author=author)


def clean(tmp_path: Path, name: str = "clean") -> Path:
    path = make(tmp_path, name, BODY, CAPTION_STYLES)
    assert revisions.revision_count(path) == 0
    return path


def revised(tmp_path: Path, name: str = "revised") -> Path:
    """A Word-valid revised body: Reviewer B's w:ins and w:del in paragraph 2, a one-cell table after paragraph 3."""
    path = make(tmp_path, name, REVISED, CAPTION_STYLES)
    assert revisions.revision_count(path) == 2
    return path


def cell_texts(tbl: ET.Element) -> list[list[str]]:
    return [["\n".join("".join(t.text or "" for t in p.iter(W + "t")) for p in tc.findall(W + "p"))
             for tc in tr.findall(W + "tc")] for tr in tbl.findall(W + "tr")]


def ins_ids(xml: str, author: str) -> list[int]:
    return [int(m.group(1)) for m in re.finditer(rf'<w:ins w:id="(\d+)" w:author="{author}"', xml)]


def assert_tracked_table(out: Path, grid: list[list[str]], author: str = CLAUDE) -> ET.Element:
    """Every row mark, cell paragraph mark and run of the one new table is inserted by `author`."""
    body = body_of(out)
    parents = {child: parent for parent in body.iter() for child in parent}
    tables = [tbl for tbl in body.findall(W + "tbl") if cell_texts(tbl)[:1] == grid[:1]]
    assert len(tables) == 1
    tbl = tables[0]
    for tr in tbl.findall(W + "tr"):
        mark = tr.find(f"{W}trPr/{W}ins")
        assert mark is not None and mark.get(W + "author") == author
    for p in tbl.iter(W + "p"):
        mark = p.find(f"{W}pPr/{W}rPr/{W}ins")
        assert mark is not None and mark.get(W + "author") == author
    for r in tbl.iter(W + "r"):
        if r.find(W + "commentReference") is not None:
            continue  # a merge comment's reference run is not tracked (comments never are)
        assert parents[r].tag == W + "ins" and parents[r].get(W + "author") == author
    assert not [el for el in tbl.find(W + "tblPr").iter() if el.tag in (W + "ins", W + "del")]
    assert tbl.find(W + "tblPr").find(W + "tblLook").get(W + "val") == "0000"
    assert len(tbl.find(W + "tblGrid").findall(W + "gridCol")) == len(grid[0])
    return tbl


def assert_journal_neutral(tbl: ET.Element, grid: list[list[str]], descriptive=(1,)) -> None:
    """Header row top+bottom, last row bottom, no verticals; header and descriptive columns left, others right."""
    rows = tbl.findall(W + "tr")
    for r, tr in enumerate(rows, start=1):
        for c, tc in enumerate(tr.findall(W + "tc"), start=1):
            borders = tc.find(f"{W}tcPr/{W}tcBorders")
            sides = {el.tag[len(W):] for el in borders} if borders is not None else set()
            want = ({"top", "bottom"} if r == 1 else set()) | ({"bottom"} if r == len(rows) else set())
            assert sides == want, (r, c, sides)
            jc = {p.find(f"{W}pPr/{W}jc").get(W + "val") for p in tc.findall(W + "p")}
            assert jc == {"left" if r == 1 or c in descriptive else "right"}, (r, c)


def reject_mcp(out: Path, tmp_path: Path, name: str = "rejected.docx") -> tuple[Path, dict]:
    target = tmp_path / name
    return target, revisions.resolve_revisions(out, target, mode="reject", author=CLAUDE)


# -- 1. switch on, revision-free document ----------------------------------------------------------------

def test_tracked_table_on_clean_doc(tmp_path):
    source = clean(tmp_path)
    before = file_sha256(source)
    out = tmp_path / "out.docx"
    result = insert(source, out)
    assert file_sha256(source) == before
    assert (result["engine"], result["tracked"], result["author"], result["anchor_mark_moved"]) == (
        "direct_xml", True, CLAUDE, True)
    assert (result["rows"], result["columns"], result["caption_style"], result["untracked"]) == (4, 4, "Caption", [])
    # caption + 4 rows x (4 cells + row end) + the empty paragraph after the table
    assert result["index_delta"] == 1 + 4 * 5 + 1 and result["new_paragraph_index"] == 3
    assert len(result["revisions_added"]) == len(ins_ids(document_xml(out), CLAUDE))

    tbl = assert_tracked_table(out, GRID)
    assert cell_texts(tbl) == GRID
    assert_journal_neutral(tbl, GRID)
    ids = ins_ids(document_xml(out), CLAUDE)
    assert ids == sorted(ids) and len(set(ids)) == len(ids)

    out_body = body_of(out)
    paragraphs = out_body.findall(W + "p")
    anchor, caption, trailing = paragraphs[1], paragraphs[2], paragraphs[3]
    assert anchor.find(f"{W}pPr/{W}rPr/{W}ins") is not None  # Word's way: the anchor's mark is the inserted one
    assert caption.find(f"{W}pPr/{W}pStyle").get(W + "val") == "Caption"
    assert "".join(t.text for t in caption.iter(W + "t")) == CAPTION
    assert caption.find(f"{W}pPr/{W}rPr/{W}ins") is not None and caption.find(W + "ins") is not None
    assert [el.tag for el in out_body][1:5] == [W + "p", W + "p", W + "tbl", W + "p"] and trailing.find(W + "r") is None
    assert canonical(trailing.find(W + "pPr")) == canonical(body_of(source).findall(W + "p")[1].find(W + "pPr"))

    # Reject (string engine): the whole table goes, as in Word; the input body comes back
    rejected, resolved = reject_mcp(out, tmp_path)
    assert canonical(body_of(rejected)) == canonical(body_of(source))
    assert resolved["tables_removed"] == [{"table": "table 1", "rows": 4, "paragraphs": [4, 23]}]
    # Reject (ElementTree engine) agrees
    et_body = body_of(out)
    revisions.et_resolve(et_body, lambda el: el.get(W + "author") == CLAUDE, "reject")
    assert canonical(et_body) == canonical(body_of(source))
    # Accept: the table and caption stay, nothing tracked is left
    accepted = tmp_path / "accepted.docx"
    revisions.resolve_revisions(out, accepted, mode="accept", author=CLAUDE)
    assert revisions.revision_count(accepted) == 0
    assert cell_texts(body_of(accepted).find(W + "tbl")) == GRID
    texts = [item.text for item in read_layout(accepted).paragraphs]
    assert texts[:3] == ["Alpha one.", "Beta two.", CAPTION] and texts[-2:] == ["", "Gamma three."]


def test_plain_insert_on_clean_doc_by_direct_xml_reads_like_the_word_path(tmp_path):
    """The plain direct-XML table (same builder) is anchor, caption, table, empty paragraph, next paragraph."""
    source = clean(tmp_path)
    out = tmp_path / "out.docx"
    result = tracked_table.insert_table(source, out, after_paragraph_index=2, headers=HEADERS, rows=ROWS,
                                        descriptive_columns=(1,), caption=CAPTION, caption_style_id="Caption",
                                        caption_style_name="Caption", tracked=False)
    assert result["revisions_added"] == [] and result["anchor_mark_moved"] is False
    assert revisions.revision_count(out) == 0
    texts = [item.text for item in read_layout(out).paragraphs]
    assert texts[:3] == ["Alpha one.", "Beta two.", CAPTION] and texts[-2:] == ["", "Gamma three."]
    assert_journal_neutral(body_of(out).find(W + "tbl"), GRID)


# -- 2. sibling: another author's revisions -------------------------------------------------------------

@pytest.mark.parametrize("fixture", ["tracked_fixture", "reviewer_b"])
def test_tracked_table_on_revised_doc_keeps_other_revisions(tmp_path, fixture):
    if fixture == "tracked_fixture":
        source, after, caption = create_tracked_fixture(tmp_path / "revised.docx"), 1, None
    else:
        source, after, caption = revised(tmp_path), 2, CAPTION
    old = raw_revisions(source)
    assert old
    out = tmp_path / "out.docx"
    result = insert(source, out, after=after, caption=caption)
    new = raw_revisions(out)
    assert [raw for raw in new if raw in old] == old  # every existing revision byte-identical, in order
    added = [raw for raw in new if raw not in old]
    assert len(added) == len(result["revisions_added"]) and all(f'w:author="{CLAUDE}"' in raw for raw in added)
    existing = {int(m) for m in re.findall(r'w:id="(\d+)"', document_xml(source))}
    ids = ins_ids(document_xml(out), CLAUDE)
    assert ids == sorted(ids) and min(ids) > max(existing) and len(set(ids)) == len(ids)
    assert_tracked_table(out, GRID)
    rejected, _ = reject_mcp(out, tmp_path)
    assert canonical(body_of(rejected)) == canonical(body_of(source))


# -- 3. merges: plain vMerge + one comment ---------------------------------------------------------------

def test_tracked_table_merge_adds_one_comment(tmp_path):
    source = clean(tmp_path)
    out = tmp_path / "out.docx"
    result = insert(source, out, merges=(MergeGroup(column=1, start_row=2, end_row=3, label="A"),))
    text = "Cells in rows 2-3 of column 1 were merged (not tracked)"
    [item] = result["untracked"]
    assert (item["kind"], item["column"], item["start_row"], item["end_row"], item["text"]) == (
        "merge", 1, 2, 3, text)
    def comments(path: Path) -> str:
        with zipfile.ZipFile(path) as archive:
            return archive.read("word/comments.xml").decode("utf-8")

    assert comments(out).count("<w:comment ") == comments(source).count("<w:comment ") + 1
    assert comments(out).count(text) == 1 and f'w:author="{CLAUDE}"' in comments(out)
    assert audit_docx(out).live_comment_count == audit_docx(source).live_comment_count + 1
    tbl = assert_tracked_table(out, GRID)
    first, second = (tr.findall(W + "tc")[0] for tr in tbl.findall(W + "tr")[1:3])
    assert first.find(f"{W}tcPr/{W}vMerge").get(W + "val") == "restart"
    later = second.find(f"{W}tcPr/{W}vMerge")
    assert later is not None and later.get(W + "val") is None
    assert cell_texts(tbl)[2][0] == ""  # the continuation cell is empty
    assert not [el for el in tbl.iter() if el.tag in (W + "tcPrChange", W + "cellMerge")]  # the merge is plain
    # The comment sits on the merged cell's text; layout paragraph of that cell is the one reported.
    assert read_layout(out).paragraphs[item["paragraph_index"] - 1].text == "A"
    # MCP Reject of the whole table removes it and the merge comment with it, as Word does
    rejected, rejection = reject_mcp(out, tmp_path)
    assert canonical(body_of(rejected)) == canonical(body_of(source))
    [gone] = rejection["comments_removed"]
    assert gone["comment_id"] == item["comment_id"] and gone["author"] == CLAUDE
    assert comments(rejected).count("<w:comment ") == comments(source).count("<w:comment ")
    assert audit_docx(rejected).live_comment_count == audit_docx(source).live_comment_count
    assert text not in comments(rejected)
    with zipfile.ZipFile(rejected) as archive:
        for name in archive.namelist():
            if name.startswith("word/comments") and name != "word/comments.xml":
                assert item["comment_id"].upper() not in archive.read(name).decode("utf-8").upper()


def test_untracked_merge_adds_no_comment(tmp_path):
    source = create_tracked_fixture(tmp_path / "revised.docx")
    out = tmp_path / "out.docx"
    result = insert(source, out, tracked=False, after=1, caption=None,
                    merges=(MergeGroup(column=1, start_row=2, end_row=3, label="A"),))
    assert result["untracked"] == [] and result["revisions_added"] == []
    assert audit_docx(out).live_comment_count == audit_docx(source).live_comment_count


# -- 4. tracking off, revised document --------------------------------------------------------------------

def test_plain_table_on_revised_doc(tmp_path):
    source = create_tracked_fixture(tmp_path / "revised.docx")
    out = tmp_path / "out.docx"
    result = insert(source, out, tracked=False, after=1, caption=None)
    assert (result["engine"], result["tracked"], result["revisions_added"]) == ("direct_xml", False, [])
    assert raw_revisions(out) == raw_revisions(source)
    tbl = body_of(out).findall(W + "tbl")
    assert any(cell_texts(t) == GRID for t in tbl)
    assert audit_docx(out).table_count == audit_docx(source).table_count + 1


# -- 5. tracking off, revision-free document: the Word path ----------------------------------------------

class _WordPath(Exception):
    pass


@pytest.mark.parametrize("tracked", [None, False])
def test_untracked_clean_doc_takes_word_path(tmp_path, monkeypatch, tracked):
    source = clean(tmp_path)

    def word(*args, **kwargs):
        raise _WordPath

    def direct(*args, **kwargs):
        raise AssertionError("the direct-XML path ran on a revision-free document with tracking off")

    monkeypatch.setattr(editing, "IsolatedWordDocument", word)
    monkeypatch.setattr(tracked_table, "insert_table", direct)
    with pytest.raises(_WordPath):
        insert(source, tmp_path / "out.docx", tracked=tracked)


def test_tracked_caption_style_added_when_undefined(tmp_path, monkeypatch):
    """Plan 973 S3 (was: refused). A document without a Caption style gets the Caption definition Word adds
    (Word faked here; tests/test_builtin_styles.py checks the fake against Word), and only that; a caption style
    that is no built-in is still refused."""

    from test_builtin_styles import WORD_ADDS, install_fake_word

    asked = install_fake_word(monkeypatch)
    source = make(tmp_path, "nocaption", BODY, STYLES)
    out = tmp_path / "out.docx"
    result = insert(source, out)
    assert result["styles_added"] == ["Caption"] and asked == [-35]
    block = WORD_ADDS[-35][0].encode("utf-8")
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(out) as after:
        assert after.read("word/styles.xml") == before.read("word/styles.xml").replace(
            b"</w:styles>", block + b"</w:styles>")
    other = tmp_path / "other.docx"
    with pytest.raises(RefusedUnsafe, match="'Table Caption' is not a paragraph style of this document"):
        editing.insert_table_to_version(source, other, after_paragraph_index=2, headers=HEADERS, rows=ROWS,
                                        descriptive_columns=(1,), caption=CAPTION, caption_style="Table Caption",
                                        tracked=True, author=CLAUDE)
    assert not other.exists()


# -- 6. refusals ---------------------------------------------------------------------------------------------

CELL_TABLE = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/>'
              "</w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"4320\" w:type=\"dxa\"/></w:tcPr>{cell}</w:tc></w:tr>"
              "</w:tbl>")


REFUSALS = [
    ("cell", para("", run("Before.")) + CELL_TABLE.format(cell=para("", run("In a cell."))) + para("", run("After.")),
     2, "paragraph 2 is inside table 1"),
    ("nested", para("", run("Before.")) + CELL_TABLE.format(
        cell=CELL_TABLE.format(cell=para("", run("Nested."))) + para("", run("Outer."))) + para("", run("After.")),
     2, "paragraph 2 is inside a nested table"),
    ("deleted", para('<w:rPr><w:del w:id="7" w:author="Reviewer B" w:date="2026-10-10T17:37:00Z"/></w:rPr>',
                     run("Mark deleted.")) + para("", run("Next.")),
     1, "paragraph 1's mark is a tracked deletion"),
]


@pytest.mark.parametrize("name, body, after, message", REFUSALS, ids=[case[0] for case in REFUSALS])
@pytest.mark.parametrize("tracked", [True, False])
def test_refusals(tmp_path, name, body, after, message, tracked):
    source = make(tmp_path, f"{name}{tracked}", body, CAPTION_STYLES)
    before = file_sha256(source)
    out = tmp_path / "out.docx"
    if not tracked and not revisions.revision_count(source):
        pytest.skip("a revision-free document with tracking off takes the Word path")
    with pytest.raises(RefusedUnsafe, match=re.escape(message)):
        insert(source, out, tracked=tracked, after=after, caption=None)
    assert not out.exists() and file_sha256(source) == before


def test_bad_requests_refused_before_anything_is_written(tmp_path):
    source = clean(tmp_path)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="caption must be one line"):
        insert(source, out, caption="Table 1.\nTwo lines.")
    with pytest.raises(RefusedUnsafe, match="do not uniformly contain label"):
        insert(source, out, merges=(MergeGroup(column=1, start_row=3, end_row=4, label="A"),))
    assert not out.exists()


# -- 7. the switch routes insert_table in edit_document -------------------------------------------------

@windows_only
def test_batch_routes_insert_table_under_switch(tmp_path):
    doc = clean(tmp_path)
    source = tmp_path / "source-copy.docx"
    source.write_bytes(doc.read_bytes())
    result = run_batch(doc, {"op_id": "s", "type": "set_track_changes", "on": True, "author": CLAUDE},
                       {"op_id": "t", "type": "insert_table", "after_paragraph_index": 2, "headers": list(HEADERS),
                        "rows": [list(row) for row in ROWS], "descriptive_columns": [1]},
                       {"op_id": "r", "type": "replace_text", "paragraph_index": 3, "find": "Gamma",
                        "replace": "Delta"})
    assert result.status == "complete", result
    table_op = result.operations[1]
    assert table_op.details["routed_as"] == "insert_table" and table_op.details["engine"] == "direct_xml"
    assert table_op.details["tracked"] is True and table_op.details["author"] == CLAUDE
    assert table_op.details["index_delta"] == 4 * 5 + 1
    # the later op named the input's paragraph 3; it was shifted past the table and still found "Gamma"
    texts = [item.text for item in read_layout(doc).paragraphs]
    assert texts[-1] == "Delta three."
    assert_tracked_table(doc, GRID)
    accepted = tmp_path / "accepted.docx"
    revisions.resolve_revisions(doc, accepted, mode="reject", author=CLAUDE)
    assert [item.text for item in read_layout(accepted).paragraphs] == ["Alpha one.", "Beta two.", "Gamma three."]


@windows_only
def test_batch_caption_fields_reach_the_table_and_are_cleaned(tmp_path):
    doc = revised(tmp_path)
    result = run_batch(doc, {"op_id": "t", "type": "insert_table", "after_paragraph_index": 2, "tracked": True,
                             "author": CLAUDE, "headers": list(HEADERS), "rows": [list(row) for row in ROWS],
                             "descriptive_columns": [1], "caption": "Table 1.\u00a0Synthetic values.",
                             "caption_style": "caption"})
    assert result.status == "complete", result
    details = result.operations[0].details
    assert details["caption"] == CAPTION and details["caption_style"] == "caption"
    assert [item["codepoint"] for item in details["cleaned"]] == ["U+00A0"]
    caption = list(body_of(doc))[2]
    assert caption.find(f"{W}pPr/{W}pStyle").get(W + "val") == "Caption"
    assert "".join(t.text or "" for t in caption.iter(W + "t")) == CAPTION


# -- revisions.py: rejecting a whole tracked table removes it (both engines); the other refusals stay -----

def _tracked_table_doc(tmp_path: Path, name: str, *rows: str) -> Path:
    return build(tmp_path, para("", run("x")) + table(*rows) + para("", run("y")), name)


def test_reject_whole_inserted_table_removes_it_in_both_engines(tmp_path):
    doc = _tracked_table_doc(tmp_path, "whole", tracked_row("ins", 501, ["A1"], ["B1"]),
                             tracked_row("ins", 511, ["A2"], ["B2"]))
    out, result = tmp_path / "o.docx", None
    result = revisions.resolve_revisions(doc, out, mode="reject", author=GUEST)
    assert body_of(out).find(W + "tbl") is None
    assert [item.text for item in read_layout(out).paragraphs] == ["x", "y"]
    assert result["tables_removed"] == [{"table": "table 1", "rows": 2, "paragraphs": [2, 7]}]
    et_body = body_of(doc)
    revisions.et_resolve(et_body, lambda el: True, "reject")
    assert canonical(et_body) == canonical(body_of(out))


def _commented_tracked_table(tmp_path: Path) -> tuple[Path, Path, dict]:
    """A fully inserted table with a comment on cell B2's text (the comment lies wholly inside the table)."""
    from research_mcp.comment_ops import add_comment

    doc = _tracked_table_doc(tmp_path, "cmt-src", tracked_row("ins", 501, ["A1"], ["B1"]),
                             tracked_row("ins", 511, ["A2"], ["B2"]))
    out = tmp_path / "cmt.docx"
    paragraph = [item.text for item in read_layout(doc).paragraphs].index("B2") + 1
    added = add_comment(doc, out, paragraph_index=paragraph, anchor="B2", text="Check this cell.", author="Reviewer B")
    return doc, out, added


def _comment_entries(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as archive:
        return {name: len(re.findall(r"<w(?:15|16cid|16cex)?:comment(?:Ex|Extensible)?\b",
                                     archive.read(name).decode("utf-8")))
                for name in archive.namelist() if name.startswith("word/comments")}


def test_reject_whole_table_removes_a_comment_wholly_inside_it_in_both_engines(tmp_path):
    doc, commented, added = _commented_tracked_table(tmp_path)
    assert audit_docx(commented).live_comment_count == 1
    out = tmp_path / "o.docx"
    result = revisions.resolve_revisions(commented, out, mode="reject", author=GUEST)
    assert [(item["comment_id"], item["author"]) for item in result["comments_removed"]] == [
        (added["comment_id"], "Reviewer B")]
    assert result["tables_removed"][0]["rows"] == 2
    assert [item.text for item in read_layout(out).paragraphs] == ["x", "y"]
    assert body_of(out).find(W + "tbl") is None
    assert not [el for el in body_of(out).iter() if "comment" in el.tag]
    assert audit_docx(out).live_comment_count == 0
    assert _comment_entries(out) == {**_comment_entries(doc), **{name: 0 for name in _comment_entries(out)
                                                               if name not in _comment_entries(doc)}}
    assert "Check this cell." not in zipfile.ZipFile(out).read("word/comments.xml").decode("utf-8")
    # the ElementTree engine removes the table with its comment markers too; without the table the body is the
    # input's (the comment run and markers went with it)
    et_body = body_of(commented)
    revisions.et_resolve(et_body, lambda el: True, "reject")
    assert canonical(et_body) == canonical(body_of(out))
    expected = body_of(doc)
    expected.remove(expected.find(W + "tbl"))
    assert canonical(body_of(out)) == canonical(expected)


def test_reject_whole_table_refuses_a_comment_reaching_outside_it(tmp_path):
    _, commented, added = _commented_tracked_table(tmp_path)
    reaching = tmp_path / "reaching.docx"

    def move_start(xml: str) -> str:
        start = re.search(r'<w:commentRangeStart w:id="(\d+)"/>', xml).group(0)
        xml = xml.replace(start, "", 1)
        return xml.replace("<w:r><w:t", start + "<w:r><w:t", 1)  # into paragraph "x", before the table

    from test_revisions_guards import variant

    variant(commented, reaching, move_start)
    xml = document_xml(reaching)
    assert xml.index("commentRangeStart") < xml.index("<w:tbl>") < xml.index("commentRangeEnd")
    with pytest.raises(RefusedUnsafe, match="comment"):
        revisions.resolve_revisions(reaching, tmp_path / "o.docx", mode="reject", author=GUEST)
    assert not (tmp_path / "o.docx").exists()
    et_body = body_of(reaching)
    with pytest.raises(RefusedUnsafe):
        revisions.et_resolve(et_body, lambda el: True, "reject")


@windows_only
def test_tracked_merge_then_reject_in_a_batch_gives_the_input_back(tmp_path):
    doc = clean(tmp_path)
    source = tmp_path / "source-copy.docx"
    source.write_bytes(doc.read_bytes())
    result = run_batch(doc, {"op_id": "s", "type": "set_track_changes", "on": True, "author": CLAUDE},
                       {"op_id": "t", "type": "insert_table", "after_paragraph_index": 2, "headers": list(HEADERS),
                        "rows": [list(row) for row in ROWS], "descriptive_columns": [1],
                        "merges": [{"column": 1, "start_row": 2, "end_row": 3, "label": "A"}]})
    assert result.status == "complete", result
    [untracked] = result.operations[1].details["untracked"]
    assert audit_docx(doc).live_comment_count == 1
    result = run_batch(doc, {"op_id": "r", "type": "reject_revisions", "author": CLAUDE})
    assert result.status == "complete", result
    assert [item["comment_id"] for item in result.operations[0].details["comments_removed"]] == [
        untracked["comment_id"]]
    assert canonical(body_of(doc)) == canonical(body_of(source))
    assert audit_docx(doc).live_comment_count == 0


def test_whole_table_resolution_keeps_its_refusals(tmp_path):
    # accept of a fully deleted table: still refused (Word's result not matched yet)
    deleted = _tracked_table_doc(tmp_path, "del", tracked_row("del", 501, ["A1"], ["B1"]))
    with pytest.raises(RefusedUnsafe, match="leave the table with no rows; resolve it in Word"):
        revisions.resolve_revisions(deleted, tmp_path / "a.docx", mode="accept", author=GUEST)
    # an unselected revision in the table keeps it
    row2 = tracked_row("ins", 511, ["A2"], ["B2"])
    cut = row2.rindex(f'w:author="{GUEST}"')
    row2 = row2[:cut] + 'w:author="Reviewer B"' + row2[cut + len(f'w:author="{GUEST}"'):]  # B2's run: another author
    mixed = _tracked_table_doc(tmp_path, "mixed", tracked_row("ins", 501, ["A1"], ["B1"]), row2)
    with pytest.raises(RefusedUnsafe, match="table 1 row 2"):
        revisions.resolve_revisions(mixed, tmp_path / "b.docx", mode="reject", author=GUEST)
    # a plain row stays, so the table stays: only the inserted row goes
    partial = _tracked_table_doc(tmp_path, "partial", plain_row(["A1"], ["B1"]),
                                 tracked_row("ins", 511, ["A2"], ["B2"]))
    out = tmp_path / "c.docx"
    revisions.resolve_revisions(partial, out, mode="reject", author=GUEST)
    assert cell_texts(body_of(out).find(W + "tbl")) == [["A1", "B1"]]
    # a bookmark in an inserted row is still refused (Word would delete it with the row)
    marked = tracked_row("ins", 501, ["A1"], ["B1"]).replace(
        "</w:rPr></w:pPr>", '</w:rPr></w:pPr><w:bookmarkStart w:id="9" w:name="Mark"/><w:bookmarkEnd w:id="9"/>', 1)
    bm = _tracked_table_doc(tmp_path, "bm", marked)
    with pytest.raises(RefusedUnsafe, match="w:bookmarkStart"):
        revisions.resolve_revisions(bm, tmp_path / "d.docx", mode="reject", author=GUEST)


# -- live: Word opens it; Word Reject All == input; Accept All == the untracked Word insert ----------------

live = pytest.mark.skipif(os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1", reason="needs RESEARCH_MCP_WORD_LIVE=1")


def _layout(path: Path) -> tuple[list[str], list[tuple[int, ...]]]:
    texts = [item.text for item in read_layout(path).paragraphs]
    shapes = [tuple(len(tr.findall(W + "tc")) for tr in tbl.findall(W + "tr"))
              for tbl in body_of(path).iter(W + "tbl")]
    return texts, shapes


@live
@pytest.mark.parametrize("merged", [False, True])
@pytest.mark.parametrize("revised_source", [False, True])
def test_live_tracked_table_word(tmp_path, merged, revised_source):
    from live_word import word_resolve

    source = revised(tmp_path) if revised_source else clean(tmp_path)
    after = 2
    merges = (MergeGroup(column=1, start_row=2, end_row=3, label="A"),) if merged else ()
    caption = CAPTION
    out = tmp_path / "tracked.docx"
    insert(source, out, after=after, merges=merges, caption=caption)
    word, word_src = word_resolve([out, source])
    assert word["opens_clean"] and not word.get("error"), word
    assert word["AcceptAll_left"] == 0 and word["RejectAll_left"] == 0, word
    # the merge comment stays with the accepted table and goes with the rejected one, as Word handles it
    assert word["word_comments"] == word_src["word_comments"] + len(merges), (word, word_src)
    assert word["AcceptAll_comments"] == word_src["word_comments"] + len(merges)
    assert word["RejectAll_comments"] == word_src["word_comments"]
    # Reject All == input's Reject All: text, paragraph count, table count
    assert _layout(word["reject"]) == _layout(word_src["reject"])
    # Accept All == the untracked Word-path insert on the input's Accept All (revision-free)
    base = word_src["accept"]
    assert revisions.revision_count(base) == 0
    expected = tmp_path / "word-path.docx"
    editing.insert_table_to_version(base, expected, after_paragraph_index=after, headers=HEADERS, rows=ROWS,
                                    descriptive_columns=(1,), merges=merges, caption=caption, tracked=False)
    assert _layout(word["accept"]) == _layout(expected)
    accepted_tbl = [t for t in body_of(word["accept"]).iter(W + "tbl") if cell_texts(t)[0] == list(HEADERS)]
    expected_tbl = [t for t in body_of(expected).iter(W + "tbl") if cell_texts(t)[0] == list(HEADERS)]
    assert cell_texts(accepted_tbl[0]) == cell_texts(expected_tbl[0])
    vmerge = lambda tbl: [[(tc.find(f"{W}tcPr/{W}vMerge") is not None) for tc in tr.findall(W + "tc")]
                          for tr in tbl.findall(W + "tr")]
    assert vmerge(accepted_tbl[0]) == vmerge(expected_tbl[0])
