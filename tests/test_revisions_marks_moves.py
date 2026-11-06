"""Plan 562 S2: accept/reject paragraph-mark revisions (merges) and moves on the direct-XML engine.

The fixtures mirror what Word itself wrote (Track Changes on, Word COM) in C:\\tmp\\p562\\s2\\fx, and the
expectations are what Word's own Accept All / Reject All produced there: a merged paragraph keeps the
*following* paragraph's pPr and mark rPr (the last one in a chain); a paragraph-level move removes the
source paragraphs on accept and the destination paragraphs on reject. Every check reads the output with
this file's own ElementTree helpers, not with revisions.py.
"""
from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.audit import file_sha256
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_tracked_fixture
from test_revisions import W, document_xml, paragraphs, run_batch, view, windows_only
from test_revisions_guards import variant
from research_mcp.setup_cmd import _word_registered

# These edits go through a (hidden, separate) Word instance; a PC or CI runner without Word skips them.
needs_word = pytest.mark.skipif(not _word_registered(), reason="needs Microsoft Word (Word.Application)")

AUTHOR = "Guest User"
DATE = "2026-11-02T10:00:00Z"
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'


def rev(tag: str, rev_id: int, inner: str = "", author: str = AUTHOR) -> str:
    attrs = f'w:id="{rev_id}" w:author="{author}" w:date="{DATE}"'
    return f"<w:{tag} {attrs}/>" if not inner else f"<w:{tag} {attrs}>{inner}</w:{tag}>"


def run(text: str, rpr: str = "") -> str:
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(ppr: str, *content: str) -> str:
    return "<w:p>" + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + "".join(content) + "</w:p>"


def mark(*revs: str, rpr: str = "") -> str:
    return f"<w:rPr>{rpr}{''.join(revs)}</w:rPr>"


FIELD = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
         '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"abc"} </w:instrText></w:r>'
         '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(1)</w:t></w:r>'
         '<w:r><w:fldChar w:fldCharType="end"/></w:r>')

MARKS_BODY = (
    para(f'<w:spacing w:after="0"/><w:jc w:val="center"/>{mark(rev("del", 101), rpr="<w:b/>")}', run("Alpha one."))
    + para(f'<w:ind w:left="720"/><w:jc w:val="right"/>{mark(rpr="<w:i/>")}', run("Bravo two."))
    + para(mark(rev("ins", 102)), run("Charlie."))
    + para('<w:jc w:val="both"/>', run("Delta."))
    + para(mark(rev("del", 103)), run("Hotel."))
    + para(f'<w:ind w:left="360"/>{mark(rev("del", 104))}', run("India."))
    + para('<w:ind w:left="1440"/>', run("Juliet."))
    + para(f'<w:jc w:val="center"/>{mark(rev("del", 110))}')
    + para("", run("Lima."))
    + para(f'{mark(rev("del", 111))}{rev("pPrChange", 112, "<w:pPr/>")}', run("Mike."))
    + para("", run("November."))
    + para(f'<w:jc w:val="center"/>{mark(rev("del", 105))}', run("Si"), '<w:bookmarkStart w:id="300" w:name="AcrossJoin"/>',
           run("erra "), '<w:commentRangeStart w:id="0"/>', run("commented start."))
    + '<w:bookmarkStart w:id="302" w:name="Between"/>'
    + para('<w:jc w:val="right"/>', run("Tan"), '<w:bookmarkEnd w:id="300"/>', run("go "), FIELD,
           '<w:bookmarkEnd w:id="302"/>', '<w:commentRangeEnd w:id="0"/>',
           '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>',
           run(" commented end."))
    + para(mark(rev("del", 106)), run("Uniform before table."))
    + "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4320\"/>"
      "<w:gridCol w:w=\"4320\"/></w:tblGrid><w:tr>"
    + '<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>' + para(mark(rev("del", 107)), run("Papa"))
    + para("", run("Quebec")) + "</w:tc>"
    + '<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>' + para(mark(rev("del", 108)), run("Victor")) + "</w:tc>"
    + "</w:tr></w:tbl>"
    + para(f'{mark(rev("del", 113))}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>', run("Whiskey ends a section."))
    + para(mark(rev("ins", 109)), run("Xray last."))
)

MOVES_BODY = (
    para("", run("Stay "), f'<w:moveFromRangeStart w:id="201" w:author="{AUTHOR}" w:date="{DATE}" w:name="move1"/>',
         rev("moveFrom", 202, run("Moved words. ")), '<w:moveFromRangeEnd w:id="201"/>', run("rest."))
    + para("", run("Arrival "), f'<w:moveToRangeStart w:id="203" w:author="{AUTHOR}" w:date="{DATE}" w:name="move1"/>',
           rev("moveTo", 204, run("Moved words. ")), '<w:moveToRangeEnd w:id="203"/>', run("end."))
    # paragraph-level move, as Word's Range.Relocate wrote it
    + para(f'<w:jc w:val="center"/>{mark(rev("moveFrom", 205))}',
           f'<w:moveFromRangeStart w:id="206" w:author="{AUTHOR}" w:date="{DATE}" w:name="move2"/>',
           rev("moveFrom", 207, run("Mike paragraph that moves.")))
    + para(mark(rev("moveFrom", 208)), f'<w:moveFromRangeStart w:id="209" w:author="{AUTHOR}" w:date="{DATE}" w:name="move3"/>',
           '<w:moveFromRangeEnd w:id="206"/>', rev("moveFrom", 210, run("November.")))
    + '<w:moveFromRangeEnd w:id="209"/>'
    + para("", run("Oscar."))
    + para(f'<w:jc w:val="center"/>{mark(rev("moveTo", 211))}',
           f'<w:moveToRangeStart w:id="212" w:author="{AUTHOR}" w:date="{DATE}" w:name="move2"/>',
           rev("moveTo", 213, run("Mike paragraph that moves.")))
    + para(mark(rev("del", 214), rev("moveTo", 215)),
           f'<w:moveToRangeStart w:id="216" w:author="{AUTHOR}" w:date="{DATE}" w:name="move3"/>',
           '<w:moveToRangeEnd w:id="212"/>', rev("moveTo", 217, run("November.")))
    + '<w:moveToRangeEnd w:id="216"/>'
    + para('<w:ind w:left="720"/>', run("Papa."))
)


def build(tmp_path: Path, body: str, name: str) -> Path:
    base = tmp_path / f"base-{name}.docx"
    if not base.exists():
        create_tracked_fixture(base)

    def swap(xml: str) -> str:
        start = xml.index("<w:body>") + len("<w:body>")
        return xml[:start] + body + SECT + "</w:body></w:document>"

    return variant(base, tmp_path / f"{name}.docx", swap)


@pytest.fixture
def marks(tmp_path: Path) -> Path:
    return build(tmp_path, MARKS_BODY, "marks")


@pytest.fixture
def moves(tmp_path: Path) -> Path:
    return build(tmp_path, MOVES_BODY, "moves")


def resolve(doc: Path, mode: str, **selectors):
    out = doc.parent / f"{mode}-{os.urandom(3).hex()}.docx"
    return out, revisions.resolve_revisions(doc, out, mode=mode, **selectors)


def text(p) -> str:
    """Visible text of an already-resolved paragraph: w:t only (no delText, no field codes)."""
    return "".join(t.text or "" for t in p.iter(W + "t"))


def texts(path: Path) -> list[str]:
    return [text(p) for p in paragraphs(path)]


def find(path: Path, starts: str):
    return next(p for p in paragraphs(path) if text(p).startswith(starts))


def ppr(p) -> dict:
    """pPr as {local name: w:val or w:left}, mark rPr children included under 'mark'."""
    out: dict = {}
    node = p.find(W + "pPr")
    if node is None:
        return out
    for child in node:
        name = child.tag.replace(W, "")
        if name == "rPr":
            out["mark"] = [item.tag.replace(W, "") for item in child]
        else:
            out[name] = child.get(W + "val") or child.get(W + "left") or child.get(W + "after")
    return out


MARKUP = re.compile(r"<w:(moveFrom|moveTo|moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd)\b")


def anchors(xml: str) -> list[str]:
    return re.findall(r"<w:(?:bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|commentReference)\b[^>]*>", xml)


# -- paragraph marks --------------------------------------------------------------------------------------

def test_accept_deleted_mark_merges_with_following_ppr(marks):
    before = len(paragraphs(marks))
    out, result = resolve(marks, "accept", ids=[101])
    assert len(paragraphs(out)) == before - 1
    merged = paragraphs(out)[0]
    assert text(merged) == "Alpha one.Bravo two."
    assert ppr(merged) == {"ind": "720", "jc": "right", "mark": ["i"]}
    assert result["merged"] == [{"paragraphs": [1, 2], "kept_properties_of": 2, "text": "Alpha one.Bravo two."}]
    assert result["index_delta"] == -1 and result["index_shifts"] == [[2, -1]]
    assert result["indices_after_shifted"] is True


def test_reject_inserted_mark_merges_with_following_ppr(marks):
    out, _ = resolve(marks, "reject", ids=[102])
    p = find(out, "Charlie")
    assert text(p) == "Charlie.Delta." and ppr(p) == {"jc": "both"}
    assert "Delta." not in texts(out)


def test_accept_inserted_or_reject_deleted_mark_only_drops_marker(marks):
    accepted, result = resolve(marks, "accept", ids=[102])
    rejected, _ = resolve(marks, "reject", ids=[101])
    assert len(paragraphs(accepted)) == len(paragraphs(rejected)) == len(paragraphs(marks))
    assert ppr(find(accepted, "Charlie")) == {} and ppr(find(accepted, "Delta")) == {"jc": "both"}
    assert ppr(find(rejected, "Alpha")) == {"spacing": "0", "jc": "center", "mark": ["b"]}
    assert result["merged"] == [] and result["index_delta"] == 0


def test_chain_of_deleted_marks_keeps_last_ppr(marks):
    out, result = resolve(marks, "accept", ids=[103, 104])
    p = find(out, "Hotel")
    assert text(p) == "Hotel.India.Juliet." and ppr(p) == {"ind": "1440"}
    assert len(paragraphs(out)) == len(paragraphs(marks)) - 2
    assert result["merged"] == [{"paragraphs": [5, 6, 7], "kept_properties_of": 7, "text": "Hotel.India.Juliet."}]
    assert result["index_shifts"] == [[6, -1], [7, -1]]


def test_partial_chain(marks):
    out, _ = resolve(marks, "accept", ids=[104])
    assert "Hotel." in texts(out) and "India.Juliet." in texts(out)
    assert ppr(find(out, "Hotel")) == {"mark": ["del"]}


def test_empty_paragraph_with_deleted_mark_disappears(marks):
    out, _ = resolve(marks, "accept", ids=[110])
    assert "" not in texts(out)
    assert ppr(find(out, "Lima")) == {}
    assert len(paragraphs(out)) == len(paragraphs(marks)) - 1


def test_merge_keeps_comment_bookmark_and_field(marks):
    out, _ = resolve(marks, "accept", ids=[105])
    p = find(out, "Si")
    assert text(p) == "Sierra commented start.Tango (1) commented end."
    assert ppr(p) == {"jc": "right"}
    assert anchors(document_xml(out)) == anchors(document_xml(marks))
    inner = ET_string(p)
    assert inner.count(":fldChar ") == 3 and "ZOTERO_ITEM" in inner
    order = re.findall(r'<w:(bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|commentReference) w:id="(\d+)"',
                       document_xml(out))
    assert order == [("bookmarkStart", "300"), ("commentRangeStart", "0"), ("bookmarkStart", "302"),
                     ("bookmarkEnd", "300"), ("bookmarkEnd", "302"), ("commentRangeEnd", "0"),
                     ("commentReference", "0")]
    with zipfile.ZipFile(marks) as a, zipfile.ZipFile(out) as b:
        assert a.read("word/comments.xml") == b.read("word/comments.xml")


def ET_string(p) -> str:
    import xml.etree.ElementTree as ET
    return ET.tostring(p, encoding="unicode")


def test_merge_inside_one_cell(marks):
    out, _ = resolve(marks, "accept", ids=[107])
    import xml.etree.ElementTree as ET
    cells = list(ET.fromstring(document_xml(out)).iter(W + "tc"))
    assert [[text(p) for p in cell.iter(W + "p")] for cell in cells] == [["PapaQuebec"], ["Victor"]]


@pytest.mark.parametrize("mode, ids, message", [
    ("accept", [106], "would merge paragraph 14 with a table"),
    ("accept", [108], "is the last paragraph of its table cell"),
    ("reject", [109], "is the last paragraph of the document"),
    ("accept", [113], "ends a section"),
    ("accept", [111], "would drop w:pPrChange id=112 by 'Guest User'"),
])
def test_merge_refusals(marks, mode, ids, message):
    out = marks.parent / "refused.docx"
    with pytest.raises(RefusedUnsafe, match=re.escape(message)):
        revisions.resolve_revisions(marks, out, mode=mode, ids=ids)
    assert not out.exists()


def test_mark_with_its_ppr_change_selected_merges(marks):
    out, _ = resolve(marks, "accept", ids=[111, 112])
    assert "Mike.November." in texts(out)


# -- moves ------------------------------------------------------------------------------------------------

def test_accept_run_move_pulls_partner(moves):
    out, result = resolve(moves, "accept", ids=[202])
    assert texts(out)[:2] == ["Stay rest.", "Arrival Moved words. end."]
    first_two = "".join(ET_string(p) for p in paragraphs(out)[:2])
    assert not MARKUP.search(first_two)
    assert [r["id"] for r in result["pulled_in"]] == ["204"]
    assert result["moved"][0]["name"] == "move1" and result["moved"][0]["kept"] == "to"
    assert result["moved"][0]["ids"] == ["202", "204"]
    assert result["resolved_count"] == 2


def test_reject_run_move_pulls_partner(moves):
    out, result = resolve(moves, "reject", ids=[204])
    assert texts(out)[:2] == ["Stay Moved words. rest.", "Arrival end."]
    first_two = "".join(ET_string(p) for p in paragraphs(out)[:2])
    assert not MARKUP.search(first_two) and "delText" not in first_two
    assert [r["id"] for r in result["pulled_in"]] == ["202"]
    assert result["moved"][0]["kept"] == "from"


def test_paragraph_move_accept_all_matches_word(moves):
    out, result = resolve(moves, "accept", author=AUTHOR)
    assert texts(out) == ["Stay rest.", "Arrival Moved words. end.", "Oscar.", "Mike paragraph that moves.",
                          "November.Papa."]
    assert ppr(paragraphs(out)[3]) == {"jc": "center"}
    assert ppr(paragraphs(out)[4]) == {"ind": "720"}
    assert not MARKUP.search(document_xml(out))
    assert result["index_delta"] == -3


def test_paragraph_move_reject_all_matches_word(moves):
    out, _ = resolve(moves, "reject", author=AUTHOR)
    assert texts(out) == ["Stay Moved words. rest.", "Arrival end.", "Mike paragraph that moves.", "November.",
                          "Oscar.", "Papa."]
    assert ppr(paragraphs(out)[2]) == {"jc": "center"}
    assert ppr(paragraphs(out)[3]) == {}
    assert ppr(paragraphs(out)[5]) == {"ind": "720"}
    assert not MARKUP.search(document_xml(out))


def test_one_side_selects_whole_pair(moves):
    out, result = resolve(moves, "accept", ids=[213])
    assert sorted(r["id"] for r in result["resolved"]) == ["205", "207", "211", "213"]
    xml = document_xml(out)
    for gone in ('w:name="move2"', 'w:id="206"', 'w:id="212"'):
        assert gone not in xml
    for kept in ('w:name="move3"', 'w:id="208"', 'w:id="210"', 'w:id="214"', 'w:id="215"', 'w:id="217"'):
        assert kept in xml
    assert "Mike paragraph that moves." in texts(out)
    assert texts(out).count("Mike paragraph that moves.") == 1
    rejected, result_r = resolve(moves, "reject", ids=[207])
    assert sorted(r["id"] for r in result_r["resolved"]) == ["205", "207", "211", "213"]
    assert texts(rejected).count("Mike paragraph that moves.") == 1


NESTED_MOVE = (para("", f'<w:moveFromRangeStart w:id="1" w:author="A" w:date="{DATE}" w:name="m"/>',
                    rev("moveFrom", 2, run("Moved text.")), '<w:moveFromRangeEnd w:id="1"/>', run(" Stay."))
               + para("", run("Here: "), f'<w:moveToRangeStart w:id="3" w:author="A" w:date="{DATE}" w:name="m"/>',
                      rev("moveTo", 4, run("Moved ") + rev("del", 5, '<w:r><w:delText>old </w:delText></w:r>',
                                                            author="B") + run("text.")),
                      '<w:moveToRangeEnd w:id="3"/>'))


def test_reject_move_holding_another_authors_deletion_refused(tmp_path):
    """Word keeps B's pending deletion when A's move is rejected (plan 562 brief). Until that relocation is
    implemented and Word-checked, the rejection is refused before anything is written, naming both ids."""
    doc = build(tmp_path, NESTED_MOVE, "nested")
    before = doc.read_bytes()
    with pytest.raises(RefusedUnsafe, match=r"rejecting w:moveTo id=4 .*not selected \(w:del id=5 by 'B'.*"
                                            r"Word keeps them\. Select id 5 too"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="reject", ids=[4])
    assert not (tmp_path / "o.docx").exists() and doc.read_bytes() == before
    # selecting the nested deletion too is a complete decision: B's text goes with the rejected move
    out, result = resolve(doc, "reject", ids=[4, 5])
    assert texts(out) == ["Moved text. Stay.", "Here: "]
    assert sorted(r["id"] for r in result["resolved"]) == ["2", "4", "5"] and result["removed_with"] == []
    # accepting the move keeps B's deletion where it is, untouched
    accepted, result_a = resolve(doc, "accept", ids=[4])
    assert texts(accepted) == [" Stay.", "Here: Moved text."]
    assert '<w:del w:id="5" w:author="B"' in document_xml(accepted) and result_a["removed_with"] == []


def test_accept_move_whose_source_holds_another_authors_insertion_refused(tmp_path):
    body = (para("", f'<w:moveFromRangeStart w:id="1" w:author="A" w:date="{DATE}" w:name="m"/>',
                 rev("moveFrom", 2, run("Moved ") + rev("ins", 6, run("new "), author="B") + run("text.")),
                 '<w:moveFromRangeEnd w:id="1"/>', run(" Stay."))
            + para("", run("Here: "), f'<w:moveToRangeStart w:id="3" w:author="A" w:date="{DATE}" w:name="m"/>',
                   rev("moveTo", 4, run("Moved text.")), '<w:moveToRangeEnd w:id="3"/>'))
    doc = build(tmp_path, body, "nestedfrom")
    with pytest.raises(RefusedUnsafe, match=r"accepting w:moveFrom id=2 .*w:ins id=6 "):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="accept", ids=[4])
    out, _ = resolve(doc, "reject", ids=[4])  # rejecting drops the destination side only: id 6 stays
    assert '<w:ins w:id="6" w:author="B"' in document_xml(out)


def _range(kind: str, range_id: int, name: str) -> str:
    return f'<w:{kind}RangeStart w:id="{range_id}" w:author="A" w:date="{DATE}" w:name="{name}"/>'


def _range_end(kind: str, range_id: int) -> str:
    return f'<w:{kind}RangeEnd w:id="{range_id}"/>'


# outer move 11 -> 13; the destination of a second move (21 -> 23) sits inside moveTo 13 (recheck sol61 P1)
NESTED_PAIRS = (para("", _range("moveFrom", 10, "outer"), rev("moveFrom", 11, run("Outer.")), _range_end("moveFrom", 10))
                + para("", _range("moveFrom", 20, "inner"), rev("moveFrom", 21, run("Inner.")),
                       _range_end("moveFrom", 20))
                + para("", _range("moveTo", 12, "outer"),
                       rev("moveTo", 13, run("Outer.") + _range("moveTo", 22, "inner")
                           + rev("moveTo", 23, run("Inner.")) + _range_end("moveTo", 22)),
                       _range_end("moveTo", 12)))


@pytest.mark.parametrize("mode", ["accept", "reject"])
@pytest.mark.parametrize("ids", [[13], [11], [21], [23], [11, 13, 21, 23]])
def test_move_nested_in_another_move_refused(tmp_path, mode, ids):
    """Either side of either pair would resolve one move and leave half of the other (or a later decision
    duplicates or loses 'Inner.'), and Word's handling of nested moves is not verified: refused up front."""
    doc = build(tmp_path, NESTED_PAIRS, "nestedpairs")
    before, out = doc.read_bytes(), tmp_path / "o.docx"
    with pytest.raises(RefusedUnsafe, match=rf"{mode}_revisions: {mode}ing w:moveTo id=23 .*inside w:moveTo "
                                            r"id=13 .*nested moves are not supported"):
        revisions.resolve_revisions(doc, out, mode=mode, ids=ids)
    assert not out.exists() and doc.read_bytes() == before


@windows_only
@pytest.mark.parametrize("mode", ["accept", "reject"])
def test_move_nested_in_another_move_refused_through_batch(tmp_path, mode):
    doc = build(tmp_path, NESTED_PAIRS, "nestedbatch")
    before = doc.read_bytes()
    result = run_batch(doc, {"op_id": "outer", "type": f"{mode}_revisions", "ids": [13]})
    assert result.status != "complete"
    assert "nested moves are not supported" in str(result.operations[0].error)
    assert doc.read_bytes() == before
    assert {r.id for r in revisions.load(doc)[1].revisions} == {"11", "13", "21", "23"}


@pytest.mark.parametrize("mode", ["accept", "reject"])
def test_move_wrapper_inside_move_wrapper_refused(tmp_path, mode):
    """terra61 probe: a moveTo directly inside another moveTo (no range of its own)."""
    body = NESTED_MOVE.replace(rev("del", 5, '<w:r><w:delText>old </w:delText></w:r>', author="B"),
                               rev("moveTo", 8, run("nested")))
    doc = build(tmp_path, body, "wrapperinwrapper")
    before, out = doc.read_bytes(), tmp_path / "o.docx"
    with pytest.raises(RefusedUnsafe, match=r"w:moveTo id=8 .*inside w:moveTo id=4 .*nested moves are not supported"):
        revisions.resolve_revisions(doc, out, mode=mode, ids=[4])
    assert not out.exists() and doc.read_bytes() == before


def test_inner_pair_with_own_ranges_inside_move_refused(tmp_path):
    """terra61 probe: rejecting outer moveTo 4 used to remove inner moveTo 14 and leave moveFrom 12 pending."""
    source = _range("moveFrom", 11, "inner") + rev("moveFrom", 12, run("inner")) + _range_end("moveFrom", 11)
    nested = _range("moveTo", 13, "inner") + rev("moveTo", 14, run("inner")) + _range_end("moveTo", 13)
    body = para("", source) + NESTED_MOVE.replace(
        rev("del", 5, '<w:r><w:delText>old </w:delText></w:r>', author="B"), nested)
    doc = build(tmp_path, body, "innerpair")
    before, out = doc.read_bytes(), tmp_path / "o.docx"
    with pytest.raises(RefusedUnsafe, match=r"w:moveTo id=14 .*inside w:moveTo id=4 .*nested moves are not supported"):
        revisions.resolve_revisions(doc, out, mode="reject", ids=[4])
    assert not out.exists() and doc.read_bytes() == before


def test_one_move_in_several_pieces_is_not_nested(tmp_path):
    """Like a real manuscript's move: one range holding two moveTo pieces resolves as one pair."""
    body = (para("", _range("moveFrom", 1, "m"), rev("moveFrom", 2, run("One two.")), _range_end("moveFrom", 1))
            + para("", run("Here: "), _range("moveTo", 3, "m"), rev("moveTo", 4, run("One ")), run("[kept] "),
                   rev("moveTo", 5, run("two.")), _range_end("moveTo", 3)))
    doc = build(tmp_path, body, "pieces")
    accepted, result = resolve(doc, "accept", ids=[4])
    assert sorted(r["id"] for r in result["resolved"]) == ["2", "4", "5"]
    assert texts(accepted) == ["", "Here: One [kept] two."]
    rejected, _ = resolve(doc, "reject", ids=[5])
    assert texts(rejected) == ["One two.", "Here: [kept] "]


def test_move_with_complete_field(tmp_path):
    body = (para("", f'<w:moveFromRangeStart w:id="1" w:author="A" w:date="{DATE}" w:name="m"/>',
                 rev("moveFrom", 2, run("Cite ") + FIELD), '<w:moveFromRangeEnd w:id="1"/>', run(" stays."))
            + para("", run("New place "), f'<w:moveToRangeStart w:id="3" w:author="A" w:date="{DATE}" w:name="m"/>',
                   rev("moveTo", 4, run("Cite ") + FIELD), '<w:moveToRangeEnd w:id="3"/>'))
    doc = build(tmp_path, body, "field")
    for mode, expected in (("accept", [" stays.", "New place Cite (1)"]), ("reject", ["Cite (1) stays.", "New place "])):
        out, _ = resolve(doc, mode, ids=[2])
        assert texts(out) == expected
        xml = document_xml(out)
        assert xml.count("ZOTERO_ITEM") == 1 and xml.count('w:fldCharType="begin"') == 1
        assert "delInstrText" not in xml and not MARKUP.search(xml)


def test_move_with_drawing_refused(tmp_path):
    pict = '<w:r><w:pict><v:rect style="width:10pt;height:10pt"/></w:pict></w:r>'
    body = (para("", f'<w:moveFromRangeStart w:id="1" w:author="A" w:date="{DATE}" w:name="m"/>',
                 rev("moveFrom", 2, run("x") + pict), '<w:moveFromRangeEnd w:id="1"/>')
            + para("", f'<w:moveToRangeStart w:id="3" w:author="A" w:date="{DATE}" w:name="m"/>',
                   rev("moveTo", 4, run("x") + pict), '<w:moveToRangeEnd w:id="3"/>'))
    doc = build(tmp_path, body, "drawing")
    with pytest.raises(RefusedUnsafe, match=r"w:moveFrom id=2 .*w:pict"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="accept", ids=[4])
    assert not (tmp_path / "o.docx").exists()


# -- S1 interaction: anchors inside removed wrappers are hoisted, never lost --------------------------------

def test_reject_ins_hoists_comment_start(tmp_path):
    body = para("", run("Keep "), rev("ins", 1, run("new ") + '<w:commentRangeStart w:id="0"/>' + run("words")),
                run(" tail"), '<w:commentRangeEnd w:id="0"/>',
                '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>')
    doc = build(tmp_path, body, "hoist")
    out, result = resolve(doc, "reject", ids=[1])
    assert texts(out) == ["Keep  tail"]
    assert result["detached_comments"] == []  # the range still holds " tail"
    assert anchors(document_xml(out)) == anchors(document_xml(doc))
    p = ET_string(paragraphs(out)[0])
    assert p.index("Keep ") < p.index("commentRangeStart") < p.index(" tail")


def test_accept_del_hoists_bookmark(tmp_path):
    body = para("", run("A "), '<w:bookmarkStart w:id="40" w:name="b"/>',
                rev("del", 1, '<w:r><w:delText>gone</w:delText></w:r><w:bookmarkEnd w:id="40"/>'), run(" B"))
    doc = build(tmp_path, body, "bm")
    out, _ = resolve(doc, "accept", ids=[1])
    assert texts(out) == ["A  B"]
    assert anchors(document_xml(out)) == anchors(document_xml(doc))


def test_reject_ins_holding_whole_comment_keeps_it(tmp_path):
    body = para("", run("Before "), rev("ins", 1, '<w:commentRangeStart w:id="0"/>' + run("inserted")
                                          + '<w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/>'
                                            '</w:rPr><w:commentReference w:id="0"/></w:r>'), run(" after"))
    doc = build(tmp_path, body, "whole")
    out, result = resolve(doc, "reject", ids=[1])
    assert texts(out) == ["Before  after"]
    assert result["detached_comments"] == ["0"]  # kept, now zero-length (Word would delete it)
    assert anchors(document_xml(out)) == anchors(document_xml(doc))
    with zipfile.ZipFile(doc) as a, zipfile.ZipFile(out) as b:
        assert a.read("word/comments.xml") == b.read("word/comments.xml")


def test_detached_comment_found_whatever_the_attribute_quotes(tmp_path):
    """terra61 final review: a single-quoted (and spaced) w:id is the same comment; a range left empty by the
    rejection must still be reported in detached_comments."""
    body = para("", run("Before "), rev("ins", 1, "<w:commentRangeStart w:id='0'/>" + run("inserted")
                                          + "<w:commentRangeEnd w:id = '0' />"
                                            "<w:r><w:commentReference w:id='0'/></w:r>"), run(" after"))
    doc = build(tmp_path, body, "quotes")
    out, result = resolve(doc, "reject", ids=[1])
    assert texts(out) == ["Before  after"]
    assert result["detached_comments"] == ["0"]
    assert anchors(document_xml(out)) == anchors(document_xml(doc))


def test_comment_reference_with_text_refused(tmp_path):
    body = para("", '<w:commentRangeStart w:id="0"/>', run("x"), '<w:commentRangeEnd w:id="0"/>',
                rev("ins", 1, '<w:r><w:t>y</w:t><w:commentReference w:id="0"/></w:r>'))
    doc = build(tmp_path, body, "ref")
    with pytest.raises(RefusedUnsafe, match=r"w:commentReference"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="reject", ids=[1])


# -- whole-fixture behaviour and the edit_document envelope -----------------------------------------------

def test_move_and_inserted_mark_now_resolve_on_plan_501_fixture(tmp_path):
    doc = create_tracked_fixture(tmp_path / "tracked.docx")
    out, result = resolve(doc, "accept", ids=[7])
    assert sorted(r["id"] for r in result["resolved"]) == ["5", "7"]
    assert [text(p) for p in paragraphs(out)[5:7]] == ["Stay text.", "Moved sentence. Arrival text."]
    out, _ = resolve(doc, "reject", ids=[8])
    assert "Inserted paragraph mark here.Centered by the other author." in texts(out)


@windows_only
def test_edit_document_merge_then_undo(marks):
    from research_mcp.publication import undo_document

    before = file_sha256(marks)
    result = run_batch(marks, {"op_id": "a", "type": "accept_revisions", "ids": [101]})
    assert result.status == "complete" and result.undo_token, result
    assert texts(marks)[0] == "Alpha one.Bravo two."
    assert result.operations[0].details["merged"][0]["paragraphs"] == [1, 2]
    undone = undo_document(str(marks))
    assert undone["restored"] is True and file_sha256(marks) == before


def test_affected_paragraphs_include_partner_and_joined_paragraph(marks, moves):
    from research_mcp.revisions import affected_paragraphs, load, select
    _, doc = load(marks)
    chosen, _ = select(doc, op="accept_revisions", ids=[101])
    assert affected_paragraphs(doc, chosen, "accept") == [1, 2]
    assert affected_paragraphs(doc, chosen, "reject") == [1]
    _, doc = load(moves)
    chosen, _ = select(doc, op="accept_revisions", ids=[213])
    # the source Mike paragraph (3) goes and its removed mark joins it to 4; the destination is 6
    assert affected_paragraphs(doc, chosen, "accept") == [3, 4, 6]
    assert affected_paragraphs(doc, chosen, "reject") == [3, 6, 7]


# -- fix round 1: whole fields in ins/del, the Zotero guard, batch shifts and groups, table properties ------

DEL_FIELD = (FIELD.replace("<w:instrText", "<w:delInstrText").replace("</w:instrText>", "</w:delInstrText>")
             .replace("<w:t>", "<w:delText>").replace("</w:t>", "</w:delText>"))


def field_doc(tmp_path: Path, tag: str, name: str) -> Path:
    inner = DEL_FIELD if tag == "del" else FIELD
    return build(tmp_path, para("", run("Cite "), rev(tag, 1, inner), run(" end.")) + para("", run("Last.")), name)


def test_whole_field_in_del_accept_removes_reject_restores(tmp_path):
    doc = field_doc(tmp_path, "del", "delfield")
    out, _ = resolve(doc, "accept", ids=[1])
    assert texts(out) == ["Cite  end.", "Last."]
    assert "ZOTERO" not in document_xml(out) and "fldChar" not in document_xml(out)
    out, _ = resolve(doc, "reject", ids=[1])
    xml = document_xml(out)
    assert texts(out) == ["Cite (1) end.", "Last."]
    assert xml.count("ZOTERO_ITEM") == 1 and xml.count('w:fldCharType="begin"') == 1
    assert "delInstrText" not in xml and "delText" not in xml and "<w:instrText" in xml


def test_whole_field_in_ins_reject_removes_accept_keeps(tmp_path):
    doc = field_doc(tmp_path, "ins", "insfield")
    out, _ = resolve(doc, "reject", ids=[1])
    assert texts(out) == ["Cite  end.", "Last."] and "ZOTERO" not in document_xml(out)
    out, _ = resolve(doc, "accept", ids=[1])
    assert texts(out) == ["Cite (1) end.", "Last."] and document_xml(out).count("ZOTERO_ITEM") == 1


@pytest.mark.parametrize("mode, verb", [("accept", "remove"), ("reject", "restore")])
def test_partial_field_in_del_refused(tmp_path, mode, verb):
    begin, rest = DEL_FIELD.split('<w:r><w:fldChar w:fldCharType="separate"/>')
    live_rest = ('<w:r><w:fldChar w:fldCharType="separate"/>' + rest).replace("delText", "t")
    doc = build(tmp_path, para("", run("Cite "), rev("del", 1, begin), live_rest, run(" end.")), "partial")
    with pytest.raises(RefusedUnsafe, match=rf"w:del id=1 .*would {verb} part of a field"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode=mode, ids=[1])
    assert not (tmp_path / "o.docx").exists()


@windows_only
@pytest.mark.parametrize("mode, expected", [("accept", "Cite  end."), ("reject", "Cite (1) end.")])
def test_edit_document_resolves_deleted_citation_through_zotero_guard(tmp_path, mode, expected):
    doc = field_doc(tmp_path, "del", f"guard-{mode}")
    result = run_batch(doc, {"op_id": "a", "type": f"{mode}_revisions", "ids": [1]})
    assert result.status == "complete", result
    assert texts(doc)[0] == expected


@windows_only
def test_edit_document_resolves_move_holding_citation(tmp_path):
    body = (para("", f'<w:moveFromRangeStart w:id="1" w:author="A" w:date="{DATE}" w:name="m"/>',
                 rev("moveFrom", 2, run("Cite ") + FIELD), '<w:moveFromRangeEnd w:id="1"/>', run(" stays."))
            + para("", run("New place "), f'<w:moveToRangeStart w:id="3" w:author="A" w:date="{DATE}" w:name="m"/>',
                   rev("moveTo", 4, run("Cite ") + FIELD), '<w:moveToRangeEnd w:id="3"/>'))
    doc = build(tmp_path, body, "guardmove")
    result = run_batch(doc, {"op_id": "a", "type": "accept_revisions", "ids": [2]})
    assert result.status == "complete", result
    assert texts(doc) == [" stays.", "New place Cite (1)"]


@windows_only
def test_edit_document_altered_citation_still_refused(tmp_path, monkeypatch):
    doc = field_doc(tmp_path, "del", "guard-mutated")
    before = file_sha256(doc)
    from research_mcp import edit_batch

    def altered(source, output, operation):  # a resolution that renders the restored citation differently
        result = edit_batch._apply_operation_local(source, output, operation)
        out = Path(output)
        with zipfile.ZipFile(out) as archive:
            entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
            for info, data in entries:
                if info.filename == "word/document.xml":
                    data = data.replace(b"<w:t>(1)</w:t>", b"<w:t>(9)</w:t>")
                archive.writestr(info, data)
        return result

    monkeypatch.setattr(edit_batch, "_apply_operation", altered)  # in-process, so the tampering applies
    result = run_batch(doc, {"op_id": "a", "type": "reject_revisions", "ids": [1]})
    assert result.status != "complete"
    assert "Zotero" in str(result.operations[0].error)
    assert file_sha256(doc) == before


SHIFT_BODY = (para(mark(rev("del", 1)), run("A1.")) + para("", run("B2."))
              + para(mark(rev("del", 2)), run("C3.")) + para("", run("D4."))
              + para(mark(rev("del", 3)), run("E5.")) + para("", run("F6."))
              + para("", run("Target seven.")) + para("", run("Last.")))


@windows_only
@needs_word
def test_accept_with_several_merges_then_indexed_ops_hit_the_right_paragraphs(tmp_path):
    doc = build(tmp_path, SHIFT_BODY, "shifts")
    result = run_batch(doc, {"op_id": "a", "type": "accept_revisions", "ids": [1, 2, 3]},
                       {"op_id": "d", "type": "replace_text", "paragraph_index": 4, "find": "D4", "replace": "Dx"},
                       {"op_id": "t", "type": "replace_text", "paragraph_index": 7, "find": "seven", "replace": "7"})
    assert result.status == "complete", result
    assert result.operations[0].details["index_shifts"] == [[2, -1], [4, -1], [6, -1]]
    assert texts(doc) == ["A1.B2.", "C3.Dx.", "E5.F6.", "Target 7.", "Last."]


TWO_MERGES = (para(mark(rev("del", 101)), run("A.")) + para("", run("B.")) + para("", run("Target."))
              + para(mark(rev("del", 102)), run("Target.")) + para("", run(" Tail.")) + para("", run("F."))
              + para("", run("G.")))


@windows_only
def test_two_merge_operations_keep_original_paragraph_identity(tmp_path):
    """sol61 final review P1: the second accept reports its shift in its own (already shifted) coordinates;
    an edit addressed to source paragraph 4 must still land in source paragraph 4, not 3."""
    doc = build(tmp_path, TWO_MERGES, "twomerges")
    result = run_batch(doc, {"op_id": "m1", "type": "accept_revisions", "ids": [101]},
                       {"op_id": "m2", "type": "accept_revisions", "ids": [102]},
                       {"op_id": "edit", "type": "tracked_replace_text", "paragraph_index": 4, "find": "Target.",
                        "replace": "CHANGED."})
    assert result.status == "complete", result.message
    assert [view(x) for x in paragraphs(doc)] == ["A.B.", "Target.", "CHANGED. Tail.", "F.", "G."]


MERGES_AND_MOVE = (para(mark(rev("del", 101)), run("A.")) + para("", run("B.")) + para("", run("C."))
                   + para(mark(rev("del", 102)), run("D.")) + para("", run("E."))
                   + para(mark(rev("moveFrom", 205)),
                          f'<w:moveFromRangeStart w:id="206" w:author="{AUTHOR}" w:date="{DATE}" w:name="mv"/>',
                          rev("moveFrom", 207, run("Moved paragraph.")), '<w:moveFromRangeEnd w:id="206"/>')
                   + para("", run("G."))
                   + para(mark(rev("moveTo", 211)),
                          f'<w:moveToRangeStart w:id="212" w:author="{AUTHOR}" w:date="{DATE}" w:name="mv"/>',
                          rev("moveTo", 213, run("Moved paragraph.")), '<w:moveToRangeEnd w:id="212"/>')
                   + para("", run("I.")) + para("", run("J.")))


@windows_only
@needs_word
def test_several_merge_operations_and_a_move_then_indexed_edits(tmp_path):
    """Three resolution ops, each removing one paragraph (two deleted marks, then a paragraph move whose
    source paragraph goes), then edits addressed in source coordinates (1-based): each lands in its own
    paragraph. Summing the three shifts against the source index sent paragraph 5 to paragraph 2."""
    doc = build(tmp_path, MERGES_AND_MOVE, "mergesmove")
    result = run_batch(doc, {"op_id": "m1", "type": "accept_revisions", "ids": [101]},
                       {"op_id": "m2", "type": "accept_revisions", "ids": [102]},
                       {"op_id": "mv", "type": "accept_revisions", "ids": [205, 207, 211, 213]},
                       {"op_id": "e3", "type": "replace_text", "paragraph_index": 3, "find": "C.", "replace": "C3."},
                       {"op_id": "e5", "type": "replace_text", "paragraph_index": 5, "find": "E.", "replace": "E5."},
                       {"op_id": "e7", "type": "replace_text", "paragraph_index": 7, "find": "G.", "replace": "G7."},
                       {"op_id": "e9", "type": "replace_text", "paragraph_index": 9, "find": "I.", "replace": "I9."},
                       {"op_id": "e10", "type": "replace_text", "paragraph_index": 10, "find": "J.",
                        "replace": "J10."})
    assert result.status == "complete", result
    assert [op.details["index_shifts"] for op in result.operations[:3]] == [[[2, -1]], [[4, -1]], [[5, -1]]]
    assert texts(doc) == ["A.B.", "C3.", "D.E5.", "G7.", "Moved paragraph.", "I9.", "J10."]


MOVE_SECTIONS = (para('<w:outlineLvl w:val="0"/>', run("INTRO"))
                 + para("", run("Stay "), f'<w:moveFromRangeStart w:id="201" w:author="{AUTHOR}" w:date="{DATE}" w:name="m1"/>',
                        rev("moveFrom", 202, run("Moved words. ")), '<w:moveFromRangeEnd w:id="201"/>', run("rest."))
                 + para('<w:outlineLvl w:val="0"/>', run("METHODS"))
                 + para("", run("Arrival "), f'<w:moveToRangeStart w:id="203" w:author="{AUTHOR}" w:date="{DATE}" w:name="m1"/>',
                        rev("moveTo", 204, run("Moved words. ")), '<w:moveToRangeEnd w:id="203"/>', run("end."))
                 + para("", run("Last.")))


MOVE_INTO_ROW = (para("", run("Stay "), f'<w:moveFromRangeStart w:id="201" w:author="{AUTHOR}" w:date="{DATE}" w:name="m1"/>',
                      rev("moveFrom", 202, run("Moved words. ")), '<w:moveFromRangeEnd w:id="201"/>', run("rest."))
                 + "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr>"
                 + "<w:tc>" + para("", run("Arrival "), f'<w:moveToRangeStart w:id="203" w:author="{AUTHOR}" w:date="{DATE}" w:name="m1"/>',
                                   rev("moveTo", 204, run("Moved words. ")), '<w:moveToRangeEnd w:id="203"/>', run("end.")) + "</w:tc>"
                 + "<w:tc>" + para("", run("Neighbour cell.")) + "</w:tc></w:tr></w:tbl>"
                 + para("", run("Last.")))


@windows_only
def test_failed_row_op_rolls_back_the_move_partner_side(tmp_path):
    """The accept names the source side (body paragraph 1) only; its partner sits in a table row. A failing
    edit in that row must roll the move back with it (a Word-routed section op cannot test this: any such op
    is refused outright in a document with tracked revisions)."""
    doc = build(tmp_path, MOVE_INTO_ROW, "groupmove")
    before = file_sha256(doc)
    result = run_batch(doc, {"op_id": "a", "type": "accept_revisions", "ids": [202]},
                       {"op_id": "t", "type": "tracked_replace_text", "paragraph_index": 3, "find": "Absent",
                        "replace": "x", "author": "Tester"}, allow_partial=True)
    statuses = {op.op_id: op.status for op in result.operations}
    assert statuses == {"a": "rolled_back", "t": "failed"}, (statuses, result)
    assert [sorted(group.op_ids) for group in result.groups] == [["a", "t"]]
    assert file_sha256(doc) == before


def test_group_keys_cover_merge_target_and_move_partner(tmp_path):
    from research_mcp.edit_batch import _Groups
    from research_mcp.edit_models import EditDocumentRequest

    doc = build(tmp_path, MOVE_SECTIONS, "keys")
    request = EditDocumentRequest(source=str(doc), source_sha256=file_sha256(doc), output=str(doc),
                                  operations=[{"op_id": "a", "type": "accept_revisions", "ids": [202]}])
    groups = _Groups(doc, list(request.operations))
    keys = sorted(value for (kind, value), members in groups.keyed.items() if kind == "para" and 0 in members)
    assert keys == ["2", "4"]
    shifted = build(tmp_path, SHIFT_BODY, "keys2")
    request = EditDocumentRequest(source=str(shifted), source_sha256=file_sha256(shifted), output=str(shifted),
                                  operations=[{"op_id": "a", "type": "accept_revisions", "ids": [2]}])
    groups = _Groups(shifted, list(request.operations))
    keys = sorted(value for (kind, value), members in groups.keyed.items() if kind == "para" and 0 in members)
    assert keys == ["3", "4"]


def table(tbl_pr: str, grid: str, tr_pr: str, tc_pr: str) -> str:
    return (f"<w:tbl><w:tblPr>{tbl_pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr><w:trPr>{tr_pr}</w:trPr>"
            f"<w:tc><w:tcPr>{tc_pr}</w:tcPr>{para('', run('One'))}</w:tc>"
            f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{para("", run("Two"))}</w:tc></w:tr></w:tbl>'
            + para("", run("After.")))


TABLE_BODY = table(
    '<w:tblW w:w="5000" w:type="pct"/>' + rev("tblPrChange", 501, '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'),
    '<w:gridCol w:w="4000"/><w:gridCol w:w="4000"/>',
    '<w:jc w:val="center"/>' + rev("trPrChange", 503, "<w:trPr/>"),
    '<w:tcW w:w="4000" w:type="dxa"/>' + rev("tcPrChange", 504, '<w:tcPr><w:tcW w:w="4320" w:type="dxa"/>'
                                                              '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:tcPr>'))


def props(path: Path) -> dict:
    tbl = next(document_root(path).iter(W + "tbl"))

    def kids(el):
        return [(c.tag.replace(W, ""), c.get(W + "w") or c.get(W + "fill") or c.get(W + "val")) for c in (el if el is not None else [])]
    return {"tblPr": kids(tbl.find(W + "tblPr")), "grid": kids(tbl.find(W + "tblGrid")),
            "trPr": kids(tbl.find(f"{W}tr/{W}trPr")), "tcPr": kids(tbl.find(f"{W}tr/{W}tc/{W}tcPr"))}


def document_root(path: Path):
    import xml.etree.ElementTree as ET
    return ET.fromstring(document_xml(path))


def test_reject_table_property_changes_restores_old_properties(tmp_path):
    """Word-checked (C:\\tmp\\p562\\rev2\\live_r1\\table_word.py): Word's Reject All gives the same tblPr, trPr
    and tcPr (it adds its default tblLook and drops the empty trPr)."""
    doc = build(tmp_path, TABLE_BODY, "tblprops")
    out, result = resolve(doc, "reject", ids=[501, 503, 504])
    assert result["resolved_count"] == 3
    assert props(out) == {"tblPr": [("tblW", "0")], "grid": [("gridCol", "4000"), ("gridCol", "4000")],
                          "trPr": [], "tcPr": [("tcW", "4320"), ("shd", "FFFF00")]}
    assert texts(out) == ["One", "Two", "After."]
    out, _ = resolve(doc, "accept", ids=[501, 503, 504])
    assert props(out) == {"tblPr": [("tblW", "5000")], "grid": [("gridCol", "4000"), ("gridCol", "4000")],
                          "trPr": [("jc", "center")], "tcPr": [("tcW", "4000")]}


@pytest.mark.parametrize("body, rev_id, message", [
    (table("", '<w:gridCol w:w="4000"/><w:gridCol w:w="4000"/>', "",
           '<w:tcW w:w="4000" w:type="dxa"/>' + rev("tcPrChange", 504, '<w:tcPr><w:tcW w:w="8000" w:type="dxa"/>'
                                                                    '<w:gridSpan w:val="2"/></w:tcPr>')), 504,
     r"would change the table's structure \(w:gridSpan\)"),
    # Word lays the old grid out again against the cells, so its result is not the old XML (table_word.py).
    (table('<w:tblW w:w="8320" w:type="dxa"/><w:tblLayout w:type="fixed"/>',
           '<w:gridCol w:w="4000"/><w:gridCol w:w="4000"/>'
           '<w:tblGridChange w:id="502"><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>'
           '</w:tblGridChange>', "", '<w:tcW w:w="4000" w:type="dxa"/>'), 502,
     r"\(table grid\) is not supported: Word recomputes the grid"),
], ids=["gridSpan", "tblGridChange"])
def test_reject_table_property_change_that_alters_structure_refused(tmp_path, body, rev_id, message):
    doc = build(tmp_path, body, f"tblstruct{rev_id}")
    with pytest.raises(RefusedUnsafe, match=rf"id={rev_id} .*{message}"):
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="reject", ids=[rev_id])


def grid_span_table(rev_id: int) -> str:
    return table("", '<w:gridCol w:w="4000"/><w:gridCol w:w="4000"/>', "",
                 '<w:tcW w:w="4000" w:type="dxa"/>' + rev("tcPrChange", rev_id, '<w:tcPr><w:tcW w:w="8000" '
                                                                       'w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>'))


def grid_change_table(rev_id: int) -> str:
    return table('<w:tblW w:w="8320" w:type="dxa"/>', '<w:gridCol w:w="4000"/><w:gridCol w:w="4000"/>'
                 f'<w:tblGridChange w:id="{rev_id}"><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/>'
                 '</w:tblGrid></w:tblGridChange>', "", '<w:tcW w:w="4000" w:type="dxa"/>')


def test_all_refused_revisions_reported_in_one_error(tmp_path):
    """Both final reviews: a refusal names every refused id at once (count, ids per reason, capped), not the
    first one only, so a Reject All on a document with many grid changes is diagnosed in one call."""
    span_ids = list(range(500, 523))  # 23: more than the 20 listed per reason
    body = ("".join(grid_span_table(i) for i in span_ids) + grid_change_table(600) + grid_change_table(601)
            + para("", run("Fine ")) + para("", rev("ins", 650, run("kept insertion")))
            + para(mark(rev("ins", 700)), run("Last.")))
    doc = build(tmp_path, body, "manyrefusals")
    before = doc.read_bytes()
    with pytest.raises(RefusedUnsafe) as caught:
        revisions.resolve_revisions(doc, tmp_path / "o.docx", mode="reject",
                                    ids=span_ids + [600, 601, 650, 700])
    message = str(caught.value)
    assert message.startswith("reject_revisions: 26 of the 27 selected revisions cannot be rejected here; "
                              "nothing was changed"), message
    assert "(2) rejecting w:tblGridChange (table grid) is not supported" in message
    assert "ids 600, 601" in message
    assert "(23) rejecting w:tcPrChange would change the table's structure (w:gridSpan)" in message
    assert "ids " + ", ".join(str(i) for i in span_ids[:20]) + " and 3 more" in message
    assert "520" not in message and "650" not in message
    assert "(1) rejecting w:ins would merge paragraph" in message and "last paragraph of the document" in message
    assert "ids 700" in message
    assert message.endswith("Deselect these ids to resolve the rest, or resolve them in Word.")
    assert doc.read_bytes() == before and not (tmp_path / "o.docx").exists()


def test_resolved_marker_leaves_no_empty_rpr(moves):
    out, _ = resolve(moves, "reject", author=AUTHOR)
    xml = document_xml(out)
    assert "<w:rPr></w:rPr>" not in xml and "<w:pPr></w:pPr>" not in xml


@windows_only
@needs_word
def test_untracked_edit_after_marker_only_reject_completes(moves):
    reject = run_batch(moves, {"op_id": "r", "type": "reject_revisions", "author": AUTHOR})
    assert reject.status == "complete"
    edit = run_batch(moves, {"op_id": "e", "type": "replace_text", "paragraph_index": 4,
                             "find": "November", "replace": "December"})
    assert edit.status == "complete"
    assert "December." in texts(moves)
