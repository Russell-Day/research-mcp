"""Plan 864 S3b: tracked table rows.

Group 1: the resolver learns rows. Word records a tracked row deletion or insertion as a mark in the row's
w:trPr plus a mark on every paragraph of every cell and a wrapper around every run (shape probed in hidden
Word, Track Changes on, Rows(n).Delete() / Rows.Add()). Rejecting the insertion or accepting the deletion
removes the whole w:tr, as Word does, when everything in the row goes with the same selection; anything else
is refused, naming the row and the revision that keeps it.

Group 2: the delete_rows / insert_rows primitives write that shape. Every check reads the output with this
file's own ElementTree helpers.
"""
from __future__ import annotations

import datetime as dt
import itertools
import re
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.errors import RefusedUnsafe
from research_mcp.sections import canonical
from test_revisions import W, document_xml
from test_revisions_guards import variant
from test_revisions_marks_moves import AUTHOR, FIELD, build, para, rev, resolve, run, text

OTHER = "Reviewer B"
GRID = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid>'
TCPR = '<w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>'


def plain_row(*cells: list[str], trpr: str = "") -> str:
    out = "<w:tr>" + (f"<w:trPr>{trpr}</w:trPr>" if trpr else "")
    for cell in cells:
        out += "<w:tc>" + TCPR + "".join(para("", run(line)) if line else para("") for line in cell) + "</w:tc>"
    return out + "</w:tr>"


def tracked_row(kind: str, first_id: int, *cells: list[str], author: str = AUTHOR) -> str:
    """A row as Word writes it with Track Changes on: row mark, then per paragraph its mark and its run."""
    ids = itertools.count(first_id)
    tag = "delText" if kind == "del" else "t"
    out = f"<w:tr><w:trPr>{rev(kind, next(ids), author=author)}</w:trPr>"
    for cell in cells:
        out += "<w:tc>" + TCPR
        for line in cell:
            out += f"<w:p><w:pPr><w:rPr>{rev(kind, next(ids), author=author)}</w:rPr></w:pPr>"
            if line:
                out += rev(kind, next(ids), f'<w:r><w:{tag} xml:space="preserve">{line}</w:{tag}></w:r>', author)
            out += "</w:p>"
        out += "</w:tc>"
    return out + "</w:tr>"


def table(*rows: str) -> str:
    return "<w:tbl>" + GRID + "".join(rows) + "</w:tbl>"


R2 = tracked_row("del", 501, ["Gone one", "Gone two"], ["Gone three"])          # ids 501-507
R4 = tracked_row("ins", 601, ["New row"], [""])                                  # ids 601-604
ROWS_BODY = (para("", run("Before."))
             + table(plain_row(["A1"], ["B1"]), R2, plain_row(["A3"], ["B3"]), R4)
             + para("", run("After.")))
R2_IDS = list(range(501, 508))
R4_IDS = list(range(601, 605))


@pytest.fixture
def rows_doc(tmp_path: Path) -> Path:
    return build(tmp_path, ROWS_BODY, "rows")


def table_rows(path: Path, number: int = 0) -> list[list[str]]:
    """Each row's cells as their visible text (w:t only), paragraphs joined with '|'."""
    tbl = list(ET.fromstring(document_xml(path)).iter(W + "tbl"))[number]
    return [["|".join(text(p) for p in tc.findall(W + "p")) for tc in tr.findall(W + "tc")]
            for tr in tbl.findall(W + "tr")]


def revision_ids(path: Path) -> list[str]:
    root = ET.fromstring(document_xml(path))
    return [el.get(W + "id") for el in root.iter() if el.tag in (W + "ins", W + "del")]


def row_xml(path: Path, number: int) -> ET.Element:
    return list(next(ET.fromstring(document_xml(path)).iter(W + "tbl")).iter(W + "tr"))[number]


# -- group 1: resolve_revisions ---------------------------------------------------------------------------

def test_accept_deleted_row_removes_it(rows_doc):
    out, result = resolve(rows_doc, "accept", ids=R2_IDS)
    assert table_rows(out) == [["A1", "B1"], ["A3", "B3"], ["New row", ""]]
    assert sorted(revision_ids(out)) == [str(i) for i in R4_IDS]  # the inserted row is untouched
    # Before = 1; R1 = 2, 3, end 4; R2 = 5, 6, 7, end 8
    assert result["index_delta"] == -4
    assert result["index_shifts"] == [[6, -1], [7, -1], [8, -1], [9, -1]]
    assert result["rows_removed"] == [{"row": "table 1 row 2 (paragraphs 5-8)", "paragraphs": [5, 6, 7],
                                       "row_end_index": 8, "text": ["Gone one", "Gone two", "Gone three"]}]


def test_reject_deleted_row_keeps_it(rows_doc):
    out, result = resolve(rows_doc, "reject", ids=R2_IDS)
    assert table_rows(out) == [["A1", "B1"], ["Gone one|Gone two", "Gone three"], ["A3", "B3"], ["New row", ""]]
    row = row_xml(out, 1)
    assert row.find(W + "trPr") is None  # Word drops the emptied trPr
    assert not [el for el in row.iter() if el.tag in (W + "del", W + "delText", W + "ins")]
    assert result["index_delta"] == 0 and result["rows_removed"] == []


def test_reject_inserted_row_removes_it(rows_doc):
    out, result = resolve(rows_doc, "reject", ids=R4_IDS)
    assert table_rows(out) == [["A1", "B1"], ["|", ""], ["A3", "B3"]]  # R2 still a pending deletion
    assert sorted(revision_ids(out)) == [str(i) for i in R2_IDS]
    assert result["index_delta"] == -3  # R4 = 12, 13, end 14
    assert result["index_shifts"] == [[13, -1], [14, -1], [15, -1]]


def test_accept_inserted_row_keeps_it(rows_doc):
    out, result = resolve(rows_doc, "accept", ids=R4_IDS)
    assert table_rows(out)[3] == ["New row", ""]
    row = row_xml(out, 3)
    assert row.find(W + "trPr") is None
    assert not [el for el in row.iter() if el.tag in (W + "ins", W + "del")]
    assert row.find(f"{W}tc/{W}p/{W}pPr") is None  # emptied mark rPr / pPr dropped, as Word does
    assert result["index_delta"] == 0


@pytest.mark.parametrize("mode, expected", [
    ("accept", [["A1", "B1"], ["A3", "B3"], ["New row", ""]]),
    ("reject", [["A1", "B1"], ["Gone one|Gone two", "Gone three"], ["A3", "B3"]]),
])
def test_all_by_author_both_modes(rows_doc, mode, expected):
    out, result = resolve(rows_doc, mode, author=AUTHOR)
    assert table_rows(out) == expected
    assert revision_ids(out) == []
    assert result["index_delta"] == (-4 if mode == "accept" else -3)


@pytest.mark.parametrize("mode, ids, row, keeper", [
    ("reject", [601], "table 1 row 4", "w:ins id=602"),
    ("accept", [501], "table 1 row 2", "w:del id=502"),
    ("accept", [501, 502, 503, 504, 505, 506], "table 1 row 2", "w:del id=507"),
])
def test_partial_row_selection_refused(rows_doc, mode, ids, row, keeper):
    out = rows_doc.parent / "partial.docx"
    with pytest.raises(RefusedUnsafe, match=rf"would remove {row} .*{keeper}.*select it too"):
        revisions.resolve_revisions(rows_doc, out, mode=mode, ids=ids)
    assert not out.exists()


def test_row_with_other_author_revision_refused(tmp_path):
    other = tracked_row("del", 701, ["Theirs"], [""], author=OTHER).replace(
        f'w:author="{OTHER}"', f'w:author="{AUTHOR}"', 1)  # row mark by AUTHOR, its content by OTHER
    doc = build(tmp_path, para("", run("x")) + table(plain_row(["A"], ["B"]), other) + para("", run("y")),
                "other")
    with pytest.raises(RefusedUnsafe, match=rf"table 1 row 2 .*w:del id=702 .* by '{OTHER}'.*select it too"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="accept", author=AUTHOR)
    out, _ = resolve(doc, "accept", ids=[701, 702, 703, 704])  # selecting them all removes it
    assert table_rows(out) == [["A", "B"]]


def test_row_with_plain_text_refused(tmp_path):
    row = tracked_row("del", 501, ["Gone"], ["Gone too"]).replace(
        "</w:p></w:tc></w:tr>", run("Kept") + "</w:p></w:tc></w:tr>")
    doc = build(tmp_path, para("", run("x")) + table(plain_row(["A"], ["B"]), row) + para("", run("y")), "plain")
    with pytest.raises(RefusedUnsafe, match=r"table 1 row 2 .*holds text 'Kept', which is not part of the "
                                            r"selection; resolve it in Word"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="accept", author=AUTHOR)


def test_last_row_refused(tmp_path):
    doc = build(tmp_path, para("", run("x")) + table(tracked_row("del", 501, ["Only"], ["Row"]))
                + para("", run("y")), "last")
    with pytest.raises(RefusedUnsafe, match="leave the table with no rows; resolve it in Word"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="accept", author=AUTHOR)
    out, _ = resolve(doc, "reject", author=AUTHOR)  # rejecting the deletion keeps it: fine
    assert table_rows(out) == [["Only", "Row"]]


def test_row_with_bookmark_refused(tmp_path):
    row = tracked_row("del", 501, ["Gone"], ["Gone too"]).replace(
        "</w:rPr></w:pPr>", '</w:rPr></w:pPr><w:bookmarkStart w:id="9" w:name="Mark"/><w:bookmarkEnd w:id="9"/>', 1)
    doc = build(tmp_path, para("", run("x")) + table(plain_row(["A"], ["B"]), row) + para("", run("y")), "bm")
    with pytest.raises(RefusedUnsafe, match=r"table 1 row 2 .*w:bookmarkStart, which Word would delete with the "
                                            r"row; remove it first"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="accept", author=AUTHOR)


def test_index_shifts_combine_rows_and_merges(tmp_path):
    body = (para(f"<w:rPr>{rev('del', 400)}</w:rPr>", run("Merges.")) + para("", run("Into this."))
            + table(plain_row(["A1"], ["B1"]), R2) + para("", run("After.")))
    doc = build(tmp_path, body, "both")
    out, result = resolve(doc, "accept", author=AUTHOR)
    # paragraph 1 merges into 2; R1 = 3, 4, end 5; R2 = 6, 7, 8, end 9
    assert result["index_shifts"] == [[2, -1], [7, -1], [8, -1], [9, -1], [10, -1]]
    assert table_rows(out) == [["A1", "B1"]]
    assert result["merged"][0]["text"] == "Merges.Into this."


# -- group 1: et_resolve on its own -----------------------------------------------------------------------

def _body(path: Path) -> ET.Element:
    return ET.fromstring(document_xml(path)).find(W + "body")


def _et_rows(body: ET.Element) -> list[list[str]]:
    tbl = body.find(W + "tbl")
    return [["|".join(text(p) for p in tc.findall(W + "p")) for tc in tr.findall(W + "tc")]
            for tr in tbl.findall(W + "tr")]


@pytest.mark.parametrize("mode, ids, expected", [
    ("accept", R2_IDS, [["A1", "B1"], ["A3", "B3"], ["New row", ""]]),
    ("reject", R2_IDS, [["A1", "B1"], ["Gone one|Gone two", "Gone three"], ["A3", "B3"], ["New row", ""]]),
    ("reject", R4_IDS, [["A1", "B1"], ["|", ""], ["A3", "B3"]]),
    ("accept", R4_IDS, [["A1", "B1"], ["|", ""], ["A3", "B3"], ["New row", ""]]),
])
def test_et_resolve_rows(rows_doc, mode, ids, expected):
    body = _body(rows_doc)
    chosen = {str(i) for i in ids}
    revisions.et_resolve(body, lambda el: el.get(W + "id") in chosen, mode)
    assert _et_rows(body) == expected
    # the resolved row, when kept, loses its emptied trPr; the other tracked row keeps its own
    marks = [tr.find(W + "trPr") is not None for tr in body.find(W + "tbl").findall(W + "tr")]
    assert sum(marks) == 1


@pytest.mark.parametrize("mode, ids", [("reject", [601]), ("accept", [501, 502, 503])])
def test_et_resolve_partial_row_refused(rows_doc, mode, ids):
    body = _body(rows_doc)
    chosen = {str(i) for i in ids}
    with pytest.raises(RefusedUnsafe, match="row"):
        revisions.et_resolve(body, lambda el: el.get(W + "id") in chosen, mode)


def test_et_resolve_last_row_refused(tmp_path):
    doc = build(tmp_path, para("", run("x")) + table(tracked_row("del", 501, ["Only"], ["Row"]))
                + para("", run("y")), "etlast")
    body = _body(doc)
    with pytest.raises(RefusedUnsafe, match="no rows"):
        revisions.et_resolve(body, lambda el: True, "accept")


# -- group 2: delete_rows / insert_rows primitives ----------------------------------------------------------

NOW = dt.datetime(2026, 11, 17, 12, 0, tzinfo=dt.timezone.utc)
CLAUDE = revisions.DEFAULT_AUTHOR
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W14_NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
# Row 1: a bold, centred first cell and an empty second cell (I6), with Word's paraId/textId on row and paragraphs.
ROW1 = ('<w:tr w14:paraId="1A2B3C01" w14:textId="1A2B3C02"><w:tc>' + TCPR
        + '<w:p w14:paraId="1A2B3C03" w14:textId="1A2B3C04"><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr>'
        + run("A1", "<w:b/>") + '</w:p></w:tc><w:tc>' + TCPR + '<w:p w14:paraId="1A2B3C05" w14:textId="1A2B3C06"/>'
        + "</w:tc></w:tr>")
ROW2 = plain_row(["B1 one", "B1 two"], ["B2"])
ROW3 = plain_row(["C1"], ["C2"])
ELSEWHERE = rev("ins", 900, run("added "), author=OTHER)  # a revision outside the table, by another author
# Layout: Before = 1; row 1 = 2, 3, end 4; row 2 = 5, 6, 7, end 8; row 3 = 9, 10, end 11; After = 12.
ORIGINAL = [["A1", ""], ["B1 one|B1 two", "B2"], ["C1", "C2"]]


def build_rows(tmp_path: Path, name: str, *rows: str, revised: bool = True) -> Path:
    body = (para("", run("Before "), ELSEWHERE if revised else "", run("text."))
            + table(*rows) + para("", run("After.")))
    plain = build(tmp_path, body, f"{name}-raw")
    return variant(plain, tmp_path / f"{name}.docx", lambda xml: xml.replace("<w:document ", "<w:document " + W14_NS, 1))


@pytest.fixture
def row_doc(tmp_path: Path) -> Path:
    return build_rows(tmp_path, "rowdoc", ROW1, ROW2, ROW3)


def table_span(xml: str) -> tuple[int, int]:
    return xml.index("<w:tbl>"), xml.index("</w:tbl>") + len("</w:tbl>")


def assert_outside_table_identical(before: str, after: str) -> None:
    lo, hi = table_span(before)
    assert after[:lo] == before[:lo]
    assert after.endswith(before[hi:])


def new_ids(path: Path, author: str = CLAUDE) -> list[int]:
    root = ET.fromstring(document_xml(path))
    return [int(el.get(W + "id")) for el in root.iter() if el.get(W + "author") == author]


def assert_tracked_row(tr: ET.Element, kind: str, author: str = CLAUDE) -> None:
    """Word's shape: row mark in trPr, every paragraph mark marked first in its rPr, the runs in one wrapper."""
    marks = [el for el in tr.find(W + "trPr") if el.tag in (W + "ins", W + "del")]
    assert [el.tag for el in marks] == [W + kind]
    for p in tr.iter(W + "p"):
        assert p.find(f"{W}pPr/{W}rPr")[0].tag == W + kind
        content = [c for c in p if c.tag != W + "pPr"]
        if content:
            assert [c.tag for c in content] == [W + kind]
            assert content[0].find(f"{W}r/{W}{'delText' if kind == 'del' else 't'}") is not None
            if kind == "del":
                assert content[0].find(f".//{W}t") is None
    stamped = [el for el in tr.iter() if el.tag in (W + "ins", W + "del")]
    ids = [int(el.get(W + "id")) for el in stamped]
    assert ids == sorted(ids) and len(set(ids)) == len(ids)
    assert {el.get(W + "author") for el in stamped} == {author}


def test_delete_row_tracked_word_shape(row_doc, tmp_path):
    out = tmp_path / "d1.docx"
    result = revisions.delete_rows(row_doc, out, first_indices=[6], tracked=True, now=NOW)
    before, after = document_xml(row_doc), document_xml(out)
    assert_outside_table_identical(before, after)
    assert ELSEWHERE in after
    tr = row_xml(out, 1)
    assert list(tr)[0].tag == W + "trPr"  # created first in the row (no tblPrEx)
    assert_tracked_row(tr, "del")
    assert len(new_ids(out)) == 1 + 2 * 3  # row mark, then per paragraph its mark and its runs
    assert result["index_delta"] == 0 and result["indices_after_shifted"] is False
    assert result["rows_deleted"] == [{"row": "table 1 row 2 (paragraphs 5-8)", "paragraphs": [5, 6, 7],
                                       "row_end_index": 8, "text": ["B1 one", "B1 two", "B2"]}]
    assert [item["type"] for item in result["revisions_added"]] == ["del"] * 7
    accepted, _ = resolve(out, "accept", author=CLAUDE)
    assert table_rows(accepted) == [ORIGINAL[0], ORIGINAL[2]]
    rejected, _ = resolve(out, "reject", author=CLAUDE)
    assert table_rows(rejected) == ORIGINAL
    assert canonical(row_xml(rejected, 1)) == canonical(row_xml(row_doc, 1))


def test_delete_row_untracked(row_doc, tmp_path):
    out = tmp_path / "d2.docx"
    result = revisions.delete_rows(row_doc, out, first_indices=[8], tracked=False)  # its row-end marker
    assert document_xml(out) == document_xml(row_doc).replace(ROW2, "", 1)
    assert result["index_delta"] == -4 and result["indices_after_shifted"] is True
    assert result["revisions_added"] == [] and result["author"] is None


def test_delete_two_rows_tracked(row_doc, tmp_path):
    out = tmp_path / "d3.docx"
    result = revisions.delete_rows(row_doc, out, first_indices=[9, 5, 6], tracked=True, now=NOW)
    for number in (1, 2):
        assert_tracked_row(row_xml(out, number), "del")
    ids = new_ids(out)
    assert ids == sorted(ids) and len(ids) == 7 + 5
    assert [item["row"] for item in result["rows_deleted"]] == ["table 1 row 2 (paragraphs 5-8)",
                                                               "table 1 row 3 (paragraphs 9-11)"]
    assert row_xml(out, 0).find(W + "trPr") is None  # row 1 untouched
    accepted, _ = resolve(out, "accept", author=CLAUDE)
    assert table_rows(accepted) == [ORIGINAL[0]]


def _cell(*paragraphs: str, tcpr: str = TCPR) -> str:
    return "<w:tc>" + tcpr + "".join(paragraphs) + "</w:tc>"


def _merged(prop: str) -> str:
    return '<w:tcPr><w:tcW w:w="4320" w:type="dxa"/>' + prop + "</w:tcPr>"


ROW_CASES = {
    "other_author": ("<w:tr>" + _cell(para("", run("B1 "), rev("ins", 910, run("theirs"), author=OTHER)))
                     + _cell(para("", run("B2"))) + "</w:tr>",
                     r"holds a tracked revision \(w:ins id=910 by 'Reviewer B'.*accept or reject it first"),
    "already_tracked": (tracked_row("ins", 920, ["B1"], ["B2"]),
                        r"holds a tracked revision \(w:ins id=920 by 'Guest User'"),
    "nested_table": ("<w:tr>" + _cell(para("", run("B1")), table(plain_row(["n1"], ["n2"])), para(""))
                     + _cell(para("", run("B2"))) + "</w:tr>", r"holds a nested table; {way}"),
    "vMerge": ("<w:tr>" + _cell(para("", run("B1")), tcpr=_merged('<w:vMerge w:val="restart"/>'))
               + _cell(para("", run("B2"))) + "</w:tr>", r"has merged cells \(<w:vMerge>\).*; {way}"),
    "gridSpan": ("<w:tr>" + _cell(para("", run("B1")), tcpr=_merged('<w:gridSpan w:val="2"/>')) + "</w:tr>",
                 r"has merged cells \(<w:gridSpan>\)"),
    "hMerge": ("<w:tr>" + _cell(para("", run("B1")), tcpr=_merged('<w:hMerge w:val="restart"/>'))
               + _cell(para("", run("B2"))) + "</w:tr>", r"has merged cells \(<w:hMerge>\)"),
    "field": ("<w:tr>" + _cell(para("", run("B1 "), FIELD)) + _cell(para("", run("B2"))) + "</w:tr>",
              r"holds a field, which {verb} the row would {effect}; {way}"),
    "hyperlink": ("<w:tr>" + _cell(para("", '<w:hyperlink w:anchor="x">' + run("B1") + "</w:hyperlink>"))
                  + _cell(para("", run("B2"))) + "</w:tr>", r"holds a hyperlink"),
    "drawing": ("<w:tr>" + _cell(para("", "<w:r><w:drawing/></w:r>")) + _cell(para("", run("B2"))) + "</w:tr>",
                r"holds a drawing"),
}
DELETE_ONLY = {
    "bookmark": ("<w:tr>" + _cell(para("", '<w:bookmarkStart w:id="70" w:name="Mark"/>', run("B1"),
                                        '<w:bookmarkEnd w:id="70"/>')) + _cell(para("", run("B2"))) + "</w:tr>",
                 r"holds a bookmark, which deleting the row would remove; delete the row in Word, or keep it"),
    "comment": ("<w:tr>" + _cell(para("", '<w:commentRangeStart w:id="0"/>', run("B1"), '<w:commentRangeEnd w:id="0"/>'))
                + _cell(para("", run("B2"))) + "</w:tr>", r"holds a comment"),
}
DELETE_WAY = "delete the row in Word, or keep it"
COPY_WAY = "add the row in Word, or use a row without it as the template"


@pytest.mark.parametrize("case", sorted(ROW_CASES) + sorted(DELETE_ONLY))
def test_delete_rows_refusals(tmp_path, case):
    row, pattern = {**ROW_CASES, **DELETE_ONLY}[case]
    doc = build_rows(tmp_path, case, ROW1, row, ROW3)
    out = tmp_path / "refused.docx"
    pattern = pattern.format(way=re.escape(DELETE_WAY), verb="deleting", effect="remove")
    with pytest.raises(RefusedUnsafe, match=r"delete_rows: table 1 row 2 \(paragraphs 5-\d+\) " + pattern):
        revisions.delete_rows(doc, out, first_indices=[5], tracked=True, now=NOW)
    assert not out.exists()


@pytest.mark.parametrize("rows, indices, message", [
    ((ROW1,), [2], "deleting its only row would leave table 1 with no rows; a table keeps at least one row: "
                   "delete the table in Word"),
    ((ROW1, ROW2, ROW3), [2, 5, 10], "deleting all 3 of its rows would leave table 1 with no rows"),
])
def test_delete_rows_refuses_emptying_a_table(tmp_path, rows, indices, message):
    doc = build_rows(tmp_path, f"empty{len(rows)}", *rows)
    out = tmp_path / "refused.docx"
    for tracked in (True, False):
        with pytest.raises(RefusedUnsafe, match=re.escape(message)):
            revisions.delete_rows(doc, out, first_indices=indices, tracked=tracked, now=NOW)
    assert not out.exists()


def test_delete_rows_refuses_a_paragraph_outside_tables(row_doc, tmp_path):
    with pytest.raises(RefusedUnsafe, match="row paragraph 1 is not in a table"):
        revisions.delete_rows(row_doc, tmp_path / "o.docx", first_indices=[1], tracked=True)


@pytest.mark.parametrize("row, expected", [
    (plain_row(["B1"], ["B2"], trpr='<w:trHeight w:val="400"/>'), [W + "trHeight", W + "del"]),
    ('<w:tr><w:tblPrEx><w:tblBorders/></w:tblPrEx>' + plain_row(["B1"], ["B2"])[len("<w:tr>"):], [W + "del"]),
    (plain_row(["B1"], ["B2"]).replace("<w:tr>", "<w:tr><w:trPr/>"), [W + "del"]),
])
def test_delete_row_existing_trpr(tmp_path, row, expected):
    doc = build_rows(tmp_path, "trpr", ROW1, row, ROW3)
    out = tmp_path / "d5.docx"
    revisions.delete_rows(doc, out, first_indices=[5], tracked=True, now=NOW)
    tr = row_xml(out, 1)
    assert [el.tag for el in tr.find(W + "trPr")] == expected
    assert [el.tag for el in tr][:2] == ([W + "tblPrEx", W + "trPr"] if "tblPrEx" in row else [W + "trPr", W + "tc"])
    rejected, _ = resolve(out, "reject", author=CLAUDE)
    # trHeight kept, a created trPr gone; an input `<w:trPr/>` comes back absent (Word drops an empty trPr)
    original = row_xml(doc, 1)
    for trpr in original.findall(W + "trPr"):  # independent of the prefix ElementTree has registered for w
        if not len(trpr) and not trpr.attrib:
            original.remove(trpr)
    assert canonical(row_xml(rejected, 1)) == canonical(original)


def _new_row_checks(source: Path, out: Path, number: int) -> ET.Element:
    tr = row_xml(out, number)
    assert not [key for el in tr.iter() for key in el.attrib if key.startswith(W14)]  # no paraId / textId
    assert_tracked_row(tr, "ins")
    first, empty = list(tr.iter(W + "p"))
    assert first.find(f"{W}pPr/{W}jc").get(W + "val") == "center"  # the template's formatting
    assert [el.tag for el in first.find(f"{W}pPr/{W}rPr")] == [W + "ins", W + "b"]
    assert first.find(f"{W}ins/{W}r/{W}rPr/{W}b") is not None
    assert [c.tag for c in empty] == [W + "pPr"]  # I6: an empty paragraph has only its marked mark
    assert len(new_ids(out)) == 4  # row mark, A1's mark and run, the empty paragraph's mark
    assert_outside_table_identical(document_xml(source), document_xml(out))
    assert ELSEWHERE in document_xml(out)
    template = row_xml(out, 0)  # the template keeps its own ids
    assert template.get(W14 + "paraId") == "1A2B3C01"
    accepted, _ = resolve(out, "accept", author=CLAUDE)
    assert canonical(row_xml(accepted, number)) == canonical(row_xml(source, 0))
    rejected, _ = resolve(out, "reject", author=CLAUDE)
    assert table_rows(rejected) == ORIGINAL
    return tr


def test_insert_row_tracked_before(row_doc, tmp_path):
    out = tmp_path / "i1.docx"
    result = revisions.insert_rows(row_doc, out, template=3, before=10, tracked=True, now=NOW)
    assert table_rows(out) == [ORIGINAL[0], ORIGINAL[1], ["A1", ""], ORIGINAL[2]]
    _new_row_checks(row_doc, out, 2)
    assert result["new_row"] == {"table": 1, "row": 3, "first_index": 9, "end_index": 11, "text": ["A1", ""]}
    assert result["index_delta"] == 3 and result["indices_after_shifted"] is True
    assert result["template"] == "table 1 row 1 (paragraphs 2-4)" and result["stripped"] == []


def test_insert_row_tracked_append(row_doc, tmp_path):
    out = tmp_path / "i2.docx"
    result = revisions.insert_rows(row_doc, out, template=2, append_after=11, tracked=True, now=NOW)
    assert table_rows(out) == ORIGINAL + [["A1", ""]]
    _new_row_checks(row_doc, out, 3)
    assert result["new_row"]["first_index"] == 12 and result["new_row"]["row"] == 4


def test_insert_row_append_after_must_be_last_row(row_doc, tmp_path):
    with pytest.raises(RefusedUnsafe, match=r"append_after names table 1 row 2 .*not the table's last row"):
        revisions.insert_rows(row_doc, tmp_path / "o.docx", template=2, append_after=5, tracked=True)
    with pytest.raises(RefusedUnsafe, match="exactly one of before or append_after"):
        revisions.insert_rows(row_doc, tmp_path / "o.docx", template=2, tracked=True)


def test_insert_row_untracked(row_doc, tmp_path):
    out = tmp_path / "i3.docx"
    result = revisions.insert_rows(row_doc, out, template=2, before=9, tracked=False)
    before = document_xml(row_doc)
    copy = re.sub(r' w14:(?:paraId|textId)="[^"]*"', "", ROW1)
    at = before.index(ROW3)
    assert document_xml(out) == before[:at] + copy + before[at:]
    assert result["revisions_added"] == [] and result["index_delta"] == 3


def test_insert_row_then_text_edit_stays_inserted(row_doc, tmp_path):
    out, edited = tmp_path / "i4.docx", tmp_path / "i4-edit.docx"
    revisions.insert_rows(row_doc, out, template=2, before=9, tracked=True, now=NOW)
    result = revisions.text_edit(out, edited, op="tracked_replace_paragraph", paragraph_index=9, tracked=True,
                                 author=CLAUDE, text="Changed", now=NOW)
    tr = row_xml(edited, 2)
    assert not [el for el in tr.iter() if el.tag in (W + "del", W + "delText")]
    first = next(tr.iter(W + "p"))
    assert [c.tag for c in first if c.tag != W + "pPr"] == [W + "ins"]
    assert "".join(t.text for t in first.iter(W + "t")) == "Changed"
    assert result["revised_in_place"]
    accepted, _ = resolve(edited, "accept", author=CLAUDE)
    assert table_rows(accepted)[2] == ["Changed", ""]
    rejected, _ = resolve(edited, "reject", author=CLAUDE)
    assert table_rows(rejected) == ORIGINAL


@pytest.mark.parametrize("case", sorted(ROW_CASES))
def test_insert_rows_refusals(tmp_path, case):
    row, pattern = ROW_CASES[case]
    doc = build_rows(tmp_path, case, row, ROW2, ROW3)
    out = tmp_path / "refused.docx"
    pattern = pattern.format(way=re.escape(COPY_WAY), verb="copying", effect="duplicate")
    with pytest.raises(RefusedUnsafe, match=r"insert_rows: template table 1 row 1 \(paragraphs 2-\d+\) " + pattern):
        revisions.insert_rows(doc, out, template=2, append_after=revisions.load(doc)[1].layout.tables[0][2].end_index,
                              tracked=True, now=NOW)
    assert not out.exists()


def test_insert_row_strips_bookmarks(tmp_path):
    template = ("<w:tr>" + _cell(para("", '<w:bookmarkStart w:id="70" w:name="Mark"/>', '<w:commentRangeStart w:id="0"/>',
                                       run("A1"), '<w:bookmarkEnd w:id="70"/>', '<w:commentRangeEnd w:id="0"/>',
                                       '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                                       '<w:commentReference w:id="0"/></w:r>'))
                + '<w:bookmarkStart w:id="71" w:name="Cell"/>' + _cell(para("", run("A2")))
                + '<w:bookmarkEnd w:id="71"/></w:tr>')
    doc = build_rows(tmp_path, "anchors", template, ROW2, ROW3)
    for tracked in (True, False):
        out = tmp_path / f"i5-{tracked}.docx"
        result = revisions.insert_rows(doc, out, template=2, before=5, tracked=tracked, now=NOW)
        copy = row_xml(out, 1)
        names = {el.tag for el in copy.iter()}
        assert not names & {W + "bookmarkStart", W + "bookmarkEnd", W + "commentRangeStart", W + "commentRangeEnd",
                            W + "commentReference"}
        assert table_rows(out)[:2] == [["A1", "A2"], ["A1", "A2"]]
        assert canonical(row_xml(out, 0)) == canonical(row_xml(doc, 0))  # the template keeps every anchor
        assert len(result["stripped"]) == 7
        assert "<w:bookmarkStart w:id=70 w:name=Mark>" in result["stripped"]
        assert "<w:commentReference w:id=0>" in result["stripped"]


def test_reviewer_b_on_a_revision_free_document(tmp_path):
    doc = build_rows(tmp_path, "clean", ROW1, ROW2, ROW3, revised=False)
    assert not revisions.has_revisions(doc)
    deleted, added = tmp_path / "rb-del.docx", tmp_path / "rb-ins.docx"
    result = revisions.delete_rows(doc, deleted, first_indices=[5], tracked=True, author=OTHER, now=NOW)
    assert {item["author"] for item in result["revisions_added"]} == {OTHER}
    assert_tracked_row(row_xml(deleted, 1), "del", author=OTHER)
    result = revisions.insert_rows(deleted, added, template=2, before=9, tracked=True, author=OTHER, now=NOW)
    assert {item["author"] for item in result["revisions_added"]} == {OTHER}
    assert_tracked_row(row_xml(added, 2), "ins", author=OTHER)
    ids = new_ids(added, OTHER)
    assert len(set(ids)) == len(ids) == 7 + 4
    accepted, _ = resolve(added, "accept", author=OTHER)
    assert table_rows(accepted) == [ORIGINAL[0], ["A1", ""], ORIGINAL[2]]
    rejected, _ = resolve(added, "reject", author=OTHER)
    assert table_rows(rejected) == ORIGINAL
    with pytest.raises(RefusedUnsafe, match="author must not be empty"):
        revisions.delete_rows(doc, tmp_path / "x.docx", first_indices=[5], tracked=True, author=" ")


def test_insert_row_stays_in_its_table(tmp_path):
    body = table(ROW2, ROW3) + para("", run("Between.")) + table(plain_row(["D1"], ["D2"]))
    doc = build(tmp_path, body, "twotables")
    # table 1 = 1..4 (end 4), 5, 6 (end 7); Between = 8; table 2 = 9, 10 (end 11)
    with pytest.raises(RefusedUnsafe, match=r"the template is in table 1 but the copy would go in table 2; a row "
                                            r"is copied within its own table"):
        revisions.insert_rows(doc, tmp_path / "o.docx", template=1, before=9, tracked=True)
    assert not (tmp_path / "o.docx").exists()
