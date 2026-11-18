"""Plan 831 S2: a tracked edit inside the author's own pending insertion is made in place.

Synthetic fixture only. Expectations are read back with this file's own ElementTree readers, never with
revisions.py.

Input -> Expected output -> Checked by
--------------------------------------
* "Claude" w:ins id=10 "three large public " in a paragraph holding comment 0; tracked "large" -> "big", then
  "big public" -> "big" (author "Claude") -> no new revision; w:ins 10 keeps its start tag (id, author, date)
  and reads "three big public ", then "three big "; its runs keep their italic formatting; comment 0 still
  covers "study"; ``revised_in_place`` == [{id "10", author, date, before, after, removed False}]
  -> ``check_in_place`` (Reject All of result == Reject All of input, Accept All == request, bytes outside
  the paragraph identical, every other revision byte-identical) + ``ins_text`` + ``comment_text``.
* "Editor Bot" w:ins id=11 edited by author "Editor Bot" -> in place, same checks.
* the same w:ins edited by author "Claude" -> refused, naming "id=11" and 'Editor Bot'; nothing written.
* w:ins id=17 by "claude" (lower case) edited by "Claude" -> refused, naming "id=17" (exact author match).
* w:ins id=18 by another author: replace inside -> refused naming "id=18"; a pure insertion at its edge ->
  a new "Claude" w:ins, w:ins 18 byte-identical (never joined).
* mixed span "carefully collected" (own "carefully " + original "collected") -> "gathered": "carefully " is
  removed with no w:del, only original characters of "collected" get a new "Claude" w:del (mixed
  formatting keeps the character diff), "gathered" is inserted; w:ins 12 listed.
* pure insertion at the start edge ("The revised" -> "The newly revised") and at the end edge ("revised
  protocol" -> "revised study protocol") of own w:ins 13 -> joins it, no new revision; text that prefixes
  the next word takes that word's formatting (documented rule).
* the same edges with zero-width markers between the new text and the w:ins (fix round 1, F1): comment end
  and reference, bookmark end, comment start after it (end edge); comment start, bookmark start, comment end
  before it, comment start inside it (start edge); both edges at once -> the new text joins the w:ins (one start
  tag, no new revision) and every comment/bookmark keeps its extent mapped through the edit: a comment or
  bookmark ending there and a comment starting there do not take the new text, a bookmark starting there does
  -> ``joined`` (``check_in_place`` + one w:ins with that id) + ``comment_text`` / ``bookmark_text``.
* deleting all of own w:ins 14 ("were enrolled and gave" -> "gave"; the diff could delete " were enrolled
  and" instead) -> the deletion slides onto the own text: the w:ins is gone, no new revision, listed with
  removed True, after "".
* comment 1 starts inside own w:ins 15, comment 2 ends inside own w:ins 16 -> words inside the w:ins are
  changed/deleted in place and the comment ranges map through the diff ("steeply after", "error dropped ");
  deleting all of comment 1's text is refused, naming the comment ("all of its text is deleted").
* own w:ins 19 whose run carries a w:rPrChange -> refused ("formatting change").
* plain replace_text inside own w:ins 10 -> still refused (plain ops never edit inside a w:ins).
* through ``edit_document``: the op detail carries ``revised_in_place``.
* edit_feedback: a model edit, then a second tracked edit inside that edit's insertion -> the first proposal
  is ``superseded`` (no ``reworded``), the chain has 2 rounds, nothing new is reported; after the user
  accepts, the chain is "accepted" with 2 rounds, original "small", kept "moderate".
"""
from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.audit import file_sha256
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.sections import read_layout
from test_edit_feedback import build_doc, events, feedback, feedback_dir, model_edit, records  # noqa: F401
from test_edit_feedback import tracked as fb_tracked
from test_edit_feedback import user_resolve
from test_revisions import W, run_batch

CLAUDE = "Claude"
BOT = "Editor Bot"
OTHER = "Guest User"
OWN_DATE = "2026-11-01T09:00:00Z"
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
SZ = '<w:rPr><w:sz w:val="22"/></w:rPr>'
ITAL = '<w:rPr><w:i/><w:sz w:val="22"/></w:rPr>'
CHANGED = (f'<w:rPr><w:b/><w:rPrChange w:id="20" w:author="{CLAUDE}" w:date="{OWN_DATE}"><w:rPr/></w:rPrChange>'
           '</w:rPr>')


def r(text: str, rpr: str = SZ) -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def ins(rid: int, author: str, *content: str) -> str:
    return f'<w:ins w:id="{rid}" w:author="{author}" w:date="{OWN_DATE}">' + "".join(content) + "</w:ins>"


def cs(cid: int) -> str:
    return f'<w:commentRangeStart w:id="{cid}"/>'


def ce(cid: int) -> str:
    return f'<w:commentRangeEnd w:id="{cid}"/>'


def cref(cid: int) -> str:
    return f'<w:r><w:commentReference w:id="{cid}"/></w:r>'


def p(*content: str) -> str:
    return "<w:p>" + "".join(content) + "</w:p>"


BODY = (
    # 1: the reported case: a comment in the paragraph, the caller's own insertion after it
    p(r("We ran the "), cs(0), r("study"), ce(0), cref(0), r(" on "),
      ins(10, CLAUDE, r("three large public ", ITAL)), r("datasets."))
    # 2: an own insertion by another author name
    + p(r("Results "), ins(11, BOT, r("were strong ", ITAL)), r("overall."))
    # 3: own inserted text next to original text
    + p(r("Data were "), ins(12, CLAUDE, r("carefully ", ITAL)), r("collected daily."))
    # 4: boundaries
    + p(r("The "), ins(13, CLAUDE, r("revised ", ITAL)), r("protocol was approved."))
    # 5: deleted completely
    + p(r("Patients "), ins(14, CLAUDE, r("were enrolled and ", ITAL)), r("gave consent."))
    # 6: comment 1 starts inside the own insertion
    + p(r("Scores "), ins(15, CLAUDE, r("rose ", ITAL), cs(1), r("sharply ", ITAL)), r("after"), ce(1), cref(1),
        r(" training."))
    # 7: comment 2 ends inside the own insertion
    + p(r("Mean "), cs(2), r("error "), ins(16, CLAUDE, r("fell ", ITAL), ce(2), r("quickly ", ITAL)), cref(2),
        r("overall."))
    # 8: author differs only in case
    + p(r("Analysis "), ins(17, "claude", r("was blinded ", ITAL)), r("throughout."))
    # 9: another author's insertion
    + p(r("Existing "), ins(18, OTHER, r("inserted ", ITAL)), r("text."))
    # 10: own insertion whose run carries a tracked formatting change
    + p(r("Some "), ins(19, CLAUDE, r("bold words ", CHANGED)), r("here."))
)

PARA = {"reported": 1, "bot": 2, "mixed": 3, "edges": 4, "whole": 5, "c_start": 6, "c_end": 7, "case": 8,
        "other": 9, "rpr_change": 10}
EDITED_FROM = {1: "10", 2: "11", 3: "12", 4: "13", 5: "14", 6: "15", 7: "16"}

COMMENTS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:comments xmlns:w="{W[1:-1]}">'
    + "".join(f'<w:comment w:id="{n}" w:author="Reviewer" w:date="{OWN_DATE}" w:initials="RV">'
              f'<w:p><w:r><w:t>Note {n}.</w:t></w:r></w:p></w:comment>' for n in range(3))
    + "</w:comments>"
)


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    base = create_tracked_fixture(tmp_path / "base.docx")
    target = tmp_path / "own.docx"
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(target, "x") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                start = xml.index("<w:body>") + len("<w:body>")
                data = (xml[:start] + BODY + SECT + "</w:body></w:document>").encode("utf-8")
            elif info.filename == "word/comments.xml":
                data = COMMENTS.encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)
    return target


# -- independent readers -----------------------------------------------------------------------------------

def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


_P_SPAN = re.compile(r"<w:p(?=[\s>/])[^>]*?/>|<w:p(?=[\s>])[^>]*>.*?</w:p>", re.S)
_REV = re.compile(r'<w:(ins|del) w:id="(\d+)"[^>]*>.*?</w:\1>', re.S)


def element(path: Path, index: int) -> ET.Element:
    return list(ET.fromstring(document_xml(path)).iter(W + "p"))[index - 1]


def _rpr(run: ET.Element) -> str:
    node = run.find(W + "rPr")
    return "" if node is None else ET.tostring(node, encoding="unicode")


def view(paragraph: ET.Element, mode: str) -> list[tuple]:
    """Reject All ("reject") or Accept All ("accept"): characters with their run formatting, plus comment and
    bookmark markers, in order. Markers inside a removed revision are kept (Word keeps them)."""

    out: list[tuple] = []

    def walk(node: ET.Element, removed: bool) -> None:
        for child in node:
            name = child.tag.replace(W, "")
            if name in ("ins", "del"):
                gone = (name == "ins") == (mode == "reject")
                walk(child, removed or gone)
            elif name == "r":
                for part in child:
                    part_name = part.tag.replace(W, "")
                    if part_name in ("t", "delText") and not removed:
                        out.extend((ch, _rpr(child)) for ch in part.text or "")
                    elif part_name == "commentReference":
                        out.append((part_name, part.get(W + "id")))
            elif name in ("commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd"):
                out.append((name, child.get(W + "id")))
            elif name != "pPr":
                walk(child, removed)

    walk(paragraph, False)
    return out


def accepted_text(paragraph: ET.Element) -> str:
    return "".join(item[0] for item in view(paragraph, "accept") if len(item[0]) == 1)


def comment_text(paragraph: ET.Element, cid: str) -> str:
    out, inside = [], False
    for item in view(paragraph, "accept"):
        if item == ("commentRangeStart", cid):
            inside = True
        elif item == ("commentRangeEnd", cid):
            inside = False
        elif inside and len(item[0]) == 1:
            out.append(item[0])
    return "".join(out)


def revisions_raw(xml: str) -> dict[str, str]:
    return {m.group(2): m.group(0) for m in _REV.finditer(xml)}


def ins_element(path: Path, rid: str) -> ET.Element | None:
    return next((el for el in ET.fromstring(document_xml(path)).iter(W + "ins") if el.get(W + "id") == rid), None)


def ins_text(path: Path, rid: str) -> str:
    el = ins_element(path, rid)
    assert el is not None, f"w:ins {rid} is gone"
    return "".join(t.text or "" for t in el.iter(W + "t"))


def ins_rprs(path: Path, rid: str) -> set[str]:
    return {_rpr(run) for run in ins_element(path, rid).iter(W + "r") if run.find(W + "t") is not None}


def head(xml: str, rid: str) -> str | None:
    m = re.search(rf'<w:ins w:id="{rid}"[^>]*>', xml)
    return m.group(0) if m else None


def edit(doc: Path, index: int, find: str, replace: str, author: str = CLAUDE, tracked: bool = True):
    out = doc.parent / f"out-{os.urandom(3).hex()}.docx"
    op = "tracked_replace_text" if tracked else "replace_text"
    kwargs = {"author": author} if tracked else {}
    result = revisions.text_edit(doc, out, op=op, paragraph_index=index, tracked=tracked, find=find,
                                 replace=replace, **kwargs)
    return out, result


def check_in_place(doc: Path, out: Path, index: int, requested: str, edited: set[str]) -> set[str]:
    """Reject All of the result == Reject All of the input; Accept All == request; bytes outside the paragraph
    identical; every revision except the edited own w:ins byte-identical; the edited ones keep their start tag.
    Returns the new revision ids."""

    before, after = element(doc, index), element(out, index)
    assert accepted_text(after) == requested
    assert view(after, "reject") == view(before, "reject")
    old_xml, new_xml = document_xml(doc), document_xml(out)
    (a, b), (c, d) = list(_P_SPAN.finditer(old_xml))[index - 1].span(), list(_P_SPAN.finditer(new_xml))[index - 1].span()
    assert old_xml[:a] == new_xml[:c] and old_xml[b:] == new_xml[d:], "bytes outside the paragraph changed"
    old_revs, new_revs = revisions_raw(old_xml), revisions_raw(new_xml)
    for rid, raw in old_revs.items():
        if rid in edited:
            if rid in new_revs:
                assert head(new_xml, rid) == head(old_xml, rid), f"w:ins {rid} lost its id, author or date"
            continue
        assert new_revs.get(rid) == raw, f"revision {rid} changed"
    fresh = set(new_revs) - set(old_revs)
    for rid in fresh:
        assert f'w:author="{CLAUDE}"' in new_revs[rid].split(">", 1)[0]
    return fresh


def report(result: dict) -> list[tuple]:
    return [(item["id"], item["author"], item["date"], item["before"], item["after"], item["removed"])
            for item in result["revised_in_place"]]


def refused(doc: Path, index: int, find: str, replace: str, author: str = CLAUDE, tracked: bool = True) -> str:
    with pytest.raises(RefusedUnsafe) as caught:
        edit(doc, index, find, replace, author, tracked)
    assert not list(doc.parent.glob("out-*.docx")), "a refused edit wrote a file"
    return str(caught.value)


# -- the fixture -----------------------------------------------------------------------------------------------

def test_fixture_layout(doc):
    layout = read_layout(doc)
    assert layout.para(1).text == "We ran the study on three large public datasets."
    for index in PARA.values():
        assert accepted_text(element(doc, index)) == layout.para(index).text
    assert comment_text(element(doc, 6), "1") == "sharply after"
    assert comment_text(element(doc, 7), "2") == "error fell "


# -- in place --------------------------------------------------------------------------------------------------

def test_reported_case_change_then_delete_a_word_inside_own_insertion(doc):
    one, result = edit(doc, 1, "large", "big")
    assert check_in_place(doc, one, 1, "We ran the study on three big public datasets.", {"10"}) == set()
    assert ins_text(one, "10") == "three big public "
    assert ins_rprs(one, "10") == {_rpr(ET.fromstring(f'<w:r xmlns:w="{W[1:-1]}">{ITAL}</w:r>'))}
    assert comment_text(element(one, 1), "0") == "study"
    assert result["revisions_added"] == []
    assert report(result) == [("10", CLAUDE, OWN_DATE, "three large public ", "three big public ", False)]

    two, result = edit(one, 1, "big public", "big")
    assert check_in_place(one, two, 1, "We ran the study on three big datasets.", {"10"}) == set()
    assert ins_text(two, "10") == "three big "
    assert comment_text(element(two, 1), "0") == "study"
    assert report(result) == [("10", CLAUDE, OWN_DATE, "three big public ", "three big ", False)]
    # against the untouched input as well: Reject All still gives the original paragraph
    assert view(element(two, 1), "reject") == view(element(doc, 1), "reject")


def test_own_insertion_by_another_author_name(doc):
    out, result = edit(doc, 2, "strong", "robust", author=BOT)
    assert check_in_place(doc, out, 2, "Results were robust overall.", {"11"}) == set()
    assert ins_text(out, "11") == "were robust "
    assert report(result) == [("11", BOT, OWN_DATE, "were strong ", "were robust ", False)]


def test_edit_spanning_own_and_original_text(doc):
    out, result = edit(doc, 3, "carefully collected", "gathered")
    fresh = check_in_place(doc, out, 3, "Data were gathered daily.", {"12"})
    xml = document_xml(out)
    deleted = [m.group(0) for m in _REV.finditer(xml) if m.group(1) == "del"]
    # the original word's characters (character diff: mixed formatting is not widened) get the only w:del
    del_text = "".join(re.findall(r"<w:delText[^>]*>([^<]*)</w:delText>", deleted[0])) if deleted else ""
    assert len(deleted) == 1 and del_text and del_text in "collected"
    assert "carefully" not in xml                                   # removed outright, never a w:del
    assert "gathered" in accepted_text(element(out, 3))
    assert {row[0] for row in report(result)} == {"12"}
    assert fresh                                                     # the original word's deletion is new


def test_insertion_at_start_edge_joins_own_insertion(doc):
    out, result = edit(doc, 4, "The revised", "The newly revised")
    assert check_in_place(doc, out, 4, "The newly revised protocol was approved.", {"13"}) == set()
    assert ins_text(out, "13") == "newly revised "
    # an insertion that prefixes the next word takes that word's formatting (documented rule): italic here
    chars = view(element(out, 4), "accept")
    assert {rpr for ch, rpr in chars[4:10]} == {chars[10][1]} != {chars[3][1]}
    assert report(result) == [("13", CLAUDE, OWN_DATE, "revised ", "newly revised ", False)]


def test_insertion_at_end_edge_joins_own_insertion(doc):
    out, result = edit(doc, 4, "revised protocol", "revised study protocol")
    assert check_in_place(doc, out, 4, "The revised study protocol was approved.", {"13"}) == set()
    assert ins_text(out, "13") == "revised study "
    assert report(result) == [("13", CLAUDE, OWN_DATE, "revised ", "revised study ", False)]


def test_deleting_all_of_own_insertion_removes_it(doc):
    out, result = edit(doc, 5, "were enrolled and gave", "gave")
    assert check_in_place(doc, out, 5, "Patients gave consent.", {"14"}) == set()
    assert ins_element(out, "14") is None
    assert report(result) == [("14", CLAUDE, OWN_DATE, "were enrolled and ", "", True)]


def test_comment_starting_inside_own_insertion(doc):
    out, _ = edit(doc, 6, "rose", "climbed")
    check_in_place(doc, out, 6, "Scores climbed sharply after training.", {"15"})
    assert comment_text(element(out, 6), "1") == "sharply after"
    out2, _ = edit(doc, 6, "sharply", "steeply")
    check_in_place(doc, out2, 6, "Scores rose steeply after training.", {"15"})
    assert comment_text(element(out2, 6), "1") == "steeply after"
    with pytest.raises(Exception, match=r"comment \(Reviewer: 'Note 1\.'\).*all of its text is deleted"):
        edit(doc, 6, "sharply after training", "training")


def test_comment_ending_inside_own_insertion(doc):
    out, _ = edit(doc, 7, "fell", "dropped")
    check_in_place(doc, out, 7, "Mean error dropped quickly overall.", {"16"})
    assert comment_text(element(out, 7), "2") == "error dropped "
    out2, _ = edit(doc, 7, "quickly overall", "overall")
    check_in_place(doc, out2, 7, "Mean error fell overall.", {"16"})
    assert comment_text(element(out2, 7), "2") == "error fell "
    assert ins_text(out2, "16") == "fell "


# -- still refused ---------------------------------------------------------------------------------------------

def test_another_authors_insertion_is_refused_naming_it(doc):
    message = refused(doc, 2, "strong", "robust")                   # Editor Bot's insertion, edited by Claude
    assert "id=11" in message and "'Editor Bot'" in message and "another author's insertion" in message
    message = refused(doc, 9, "inserted", "added")
    assert "id=18" in message and repr(OTHER) in message


def test_author_match_is_exact_and_case_sensitive(doc):
    message = refused(doc, 8, "blinded", "masked")
    assert "id=17" in message and "'claude'" in message


def test_insertion_next_to_another_authors_insertion_does_not_join_it(doc):
    out, result = edit(doc, 9, "Existing inserted", "Existing newly inserted")
    fresh = check_in_place(doc, out, 9, "Existing newly inserted text.", set())
    assert len(fresh) == 1
    assert ins_text(out, "18") == "inserted "
    assert "revised_in_place" not in result


def test_own_insertion_with_formatting_change_is_refused(doc):
    message = refused(doc, 10, "bold", "strong")
    assert "formatting change" in message and "id=19" in message


def test_plain_edit_inside_own_insertion_stays_refused(doc):
    message = refused(doc, 1, "large", "big", tracked=False)
    assert "id=10" in message


# -- through edit_document -------------------------------------------------------------------------------------

def test_edit_document_reports_revised_in_place(doc):
    result = run_batch(doc, {"op_id": "own", "type": "tracked_replace_text", "paragraph_index": 1,
                             "find": "large", "replace": "big", "author": CLAUDE})
    assert result.status == "complete", result
    detail = result.operations[0].details["revised_in_place"]
    assert [(d["id"], d["before"], d["after"]) for d in detail] == [("10", "three large public ", "three big public ")]
    assert ins_text(doc, "10") == "three big public "


# -- edit_feedback: the model revising its own pending insertion is a new round ------------------------------

def test_edit_inside_own_pending_insertion_is_a_new_round(tmp_path, feedback_dir):  # noqa: F811
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, fb_tracked(2, "small", "modest"))
    out = model_edit(path, fb_tracked(2, "modest", "moderate"))
    assert out.get("edit_feedback") is None
    assert ins_text(path, next(el.get(W + "id") for el in ET.fromstring(document_xml(path)).iter(W + "ins")
                               if el.get(W + "author") == "Assistant")) == "moderate"
    outcomes = [e["outcome"] for e in events(path) if e["event"] == "outcome"]
    assert outcomes == ["superseded"]
    record = records(path)[2]
    assert (record["rounds"], record["pending"], record["original"], record["last_proposed"]) == \
        (2, True, "small", "moderate")
    assert feedback(path)["edit_feedback"] is None
    user_resolve(path, "accept", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    record = records(path)[2]
    assert (record["outcome"], record["rounds"], record["original"], record["kept"]) == \
        ("accepted", 2, "small", "moderate")


# -- joining across zero-width markers at the edge (fix round 1, F1) ------------------------------------------

def bms(bid: int, name: str) -> str:
    return f'<w:bookmarkStart w:id="{bid}" w:name="{name}"/>'


def bme(bid: int) -> str:
    return f'<w:bookmarkEnd w:id="{bid}"/>'


def remake(doc: Path, paragraph: str) -> Path:
    """`doc` with its body replaced by one paragraph (comments.xml kept)."""

    target = doc.parent / f"one-{os.urandom(3).hex()}.docx"
    with zipfile.ZipFile(doc) as src, zipfile.ZipFile(target, "x") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                start = xml.index("<w:body>") + len("<w:body>")
                data = (xml[:start] + paragraph + SECT + "</w:body></w:document>").encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)
    return target


def bookmark_text(paragraph: ET.Element, bid: str) -> str:
    out, inside = [], False
    for item in view(paragraph, "accept"):
        if item == ("bookmarkStart", bid):
            inside = True
        elif item == ("bookmarkEnd", bid):
            inside = False
        elif inside and len(item[0]) == 1:
            out.append(item[0])
    return "".join(out)


def joined(doc: Path, paragraph: str, find: str, replace: str, rid: str, requested: str) -> tuple[Path, Path]:
    source = remake(doc, paragraph)
    out, result = edit(source, 1, find, replace)
    assert check_in_place(source, out, 1, requested, {rid}) == set(), "the insertion did not join the own w:ins"
    assert result["revisions_added"] == []
    assert [row[0] for row in report(result)] == [rid]
    xml = document_xml(out)
    assert xml.count(f'<w:ins w:id="{rid}"') == 1
    return source, out


def test_end_edge_joins_across_comment_end_and_reference(doc):
    # the reviewer's trigger: comment 0 covers own "revised "; its end and reference sit between the w:ins and
    # "protocol". The new word joins w:ins 80; the comment still ends after "revised " (text typed where an
    # anchor ends never joins it).
    source, out = joined(doc, p(r("The "), cs(0), ins(80, CLAUDE, r("revised ", ITAL)), ce(0), cref(0),
                                r("protocol was approved.")),
                         "revised protocol", "revised study protocol", "80",
                         "The revised study protocol was approved.")
    assert ins_text(out, "80") == "revised study "
    assert comment_text(element(source, 1), "0") == "revised "
    assert comment_text(element(out, 1), "0") == "revised "
    # no marker crossed any text: the comment end still sits right after "revised "
    accept = [item[0] if len(item[0]) == 1 else item for item in view(element(out, 1), "accept")]
    assert "".join(x for x in accept[:accept.index(("commentRangeEnd", "0"))] if isinstance(x, str)) == "The revised "


def test_end_edge_joins_across_bookmark_end(doc):
    source, out = joined(doc, p(r("The "), bms(5, "Mark"), ins(83, CLAUDE, r("revised ", ITAL)), bme(5),
                                r("protocol was approved.")),
                         "revised protocol", "revised study protocol", "83",
                         "The revised study protocol was approved.")
    assert ins_text(out, "83") == "revised study "
    assert bookmark_text(element(source, 1), "5") == bookmark_text(element(out, 1), "5") == "revised "


def test_end_edge_joins_with_comment_start_after_it(doc):
    # a comment starting where the new text goes does not take it: it stays after the w:ins
    source, out = joined(doc, p(r("The "), ins(84, CLAUDE, r("revised ", ITAL)), cs(0), r("protocol"), ce(0),
                                cref(0), r(" was approved.")),
                         "revised protocol", "revised study protocol", "84",
                         "The revised study protocol was approved.")
    assert ins_text(out, "84") == "revised study "
    assert comment_text(element(source, 1), "0") == comment_text(element(out, 1), "0") == "protocol"


def test_start_edge_joins_across_comment_start(doc):
    # comment 0 starts right before own "revised ": the new word joins w:ins 81 but stays outside the comment
    source, out = joined(doc, p(r("The "), cs(0), ins(81, CLAUDE, r("revised ", ITAL)), ce(0), cref(0),
                                r("protocol was approved.")),
                         "The revised", "The newly revised", "81", "The newly revised protocol was approved.")
    assert ins_text(out, "81") == "newly revised "
    assert comment_text(element(source, 1), "0") == comment_text(element(out, 1), "0") == "revised "


def test_start_edge_joins_across_bookmark_start(doc):
    # a bookmark starting there takes the typed text (Word's rule), so it covers the new word too
    source, out = joined(doc, p(r("The "), bms(6, "Start"), ins(82, CLAUDE, r("revised ", ITAL)), bme(6),
                                r("protocol was approved.")),
                         "The revised", "The newly revised", "82", "The newly revised protocol was approved.")
    assert ins_text(out, "82") == "newly revised "
    assert bookmark_text(element(source, 1), "6") == "revised "
    assert bookmark_text(element(out, 1), "6") == "newly revised "


def test_start_edge_joins_across_comment_end_before_it(doc):
    # comment 0 ends right before own "revised ": the new word does not join the comment
    source, out = joined(doc, p(r("The "), cs(0), r("old "), ce(0), cref(0),
                                ins(85, CLAUDE, r("revised ", ITAL)), r("protocol was approved.")),
                         "old revised", "old newly revised", "85", "The old newly revised protocol was approved.")
    assert ins_text(out, "85") == "newly revised "
    assert comment_text(element(source, 1), "0") == comment_text(element(out, 1), "0") == "old "


def test_start_edge_with_comment_start_inside_own_insertion(doc):
    # comment 0 starts inside w:ins 86, before its first run: the new word goes before the comment start
    source, out = joined(doc, p(r("The "), ins(86, CLAUDE, cs(0), r("revised ", ITAL)), ce(0), cref(0),
                                r("protocol was approved.")),
                         "The revised", "The newly revised", "86", "The newly revised protocol was approved.")
    assert ins_text(out, "86") == "newly revised "
    assert comment_text(element(source, 1), "0") == comment_text(element(out, 1), "0") == "revised "


def test_end_edge_between_one_comment_end_and_another_comment_start(doc):
    # comment 0 ends and comment 1 starts between w:ins 87 and "protocol": the new word joins w:ins 87,
    # after comment 0's end and before comment 1's start; neither comment takes it
    source, out = joined(doc, p(r("The "), cs(0), ins(87, CLAUDE, r("revised ", ITAL)), ce(0), cref(0), cs(1),
                                r("protocol"), ce(1), cref(1), r(" was approved.")),
                         "revised protocol", "revised study protocol", "87",
                         "The revised study protocol was approved.")
    assert ins_text(out, "87") == "revised study "
    assert comment_text(element(out, 1), "0") == "revised "
    assert comment_text(element(out, 1), "1") == "protocol"
    assert comment_text(element(source, 1), "1") == "protocol"


def test_both_edges_in_one_edit(doc):
    source, out = joined(doc, p(r("The "), cs(0), ins(88, CLAUDE, r("revised ", ITAL)), ce(0), cref(0),
                                r("protocol was approved.")),
                         "The revised protocol", "The newly revised study protocol", "88",
                         "The newly revised study protocol was approved.")
    assert ins_text(out, "88") == "newly revised study "
    assert comment_text(element(out, 1), "0") == "revised "


def test_point_comment_at_end_edge_keeps_a_separate_insertion(doc):
    # comment 2 has only a reference (no range) right after w:ins 89: it must stay before the new text, and a
    # reference run never enters the w:ins, so the new word gets its own w:ins after it (today's behaviour)
    source = remake(doc, p(r("The "), ins(89, CLAUDE, r("revised ", ITAL)), cref(2), r("protocol was approved.")))
    out, result = edit(source, 1, "revised protocol", "revised study protocol")
    fresh = check_in_place(source, out, 1, "The revised study protocol was approved.", set())
    assert len(fresh) == 1
    assert ins_text(out, "89") == "revised "
    assert "revised_in_place" not in result
    xml = document_xml(out)
    assert xml.index('<w:commentReference w:id="2"/>') < xml.index("study ")
