"""Plan 864 S1b (decision 15) and plan 924: where typed text lands among comment and bookmark markers.

Word's typing rules (AGENTS §4) are the spec: text typed at an anchor's start joins a bookmark but not a
comment; text typed at an anchor's end never joins it (a comment reference stays with its range end); text
replacing anchored characters stays in it; a collapsed bookmark stays before typed text and a point comment
(zero-length range, then its reference) after it (plan 924, measured in Word 2026-11-17). A comment whose text is all deleted is refused, naming its
comment_id and the op that removes it. Fixtures are synthetic; each output is read with this file's own
ElementTree walker (`marked`), not with revisions.py.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.errors import VerificationFailed
from test_revisions import W, document_xml, run_batch, tracked, view, windows_only
from test_revisions_marks_moves import build, para, rev, run

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def comment_start(cid: int) -> str:
    return f'<w:commentRangeStart w:id="{cid}"/>'


def comment_end(cid: int) -> str:
    return (f'<w:commentRangeEnd w:id="{cid}"/>'
            f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>')


def bookmark_start(bid: int, name: str) -> str:
    return f'<w:bookmarkStart w:id="{bid}" w:name="{name}"/>'


def bookmark_end(bid: int) -> str:
    return f'<w:bookmarkEnd w:id="{bid}"/>'


def make(tmp_path: Path, name: str, body: str, comments: list[tuple[int, str, str | None]] = (),
         replies: dict[int, int] | None = None) -> Path:
    """The fixture with `body` and comments (w:id, text, durable id or None for a legacy-only comment);
    `replies` maps a reply's w:id to its parent's (written to commentsExtended.xml)."""

    plain = build(tmp_path, body, f"{name}-plain")
    target = tmp_path / f"{name}.docx"
    entries, ids = [], []
    for cid, text, durable in comments:
        para_id = f"{0x10000000 + cid:08X}"
        entries.append(f'<w:comment w:id="{cid}" w:author="Reviewer B" w:date="2026-11-02T10:00:00Z" w:initials="RB">'
                       f'<w:p w14:paraId="{para_id}"><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>')
        if durable:
            ids.append(f'<w16cid:commentId w16cid:paraId="{para_id}" w16cid:durableId="{durable}"/>')
    parts = {"word/comments.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:comments xmlns:w="{W[1:-1]}" xmlns:w14="{W14}">' + "".join(entries) + "</w:comments>")}
    with zipfile.ZipFile(plain) as src:
        types = src.read("[Content_Types].xml").decode("utf-8")
        rels = src.read("word/_rels/document.xml.rels").decode("utf-8")
        if ids:
            parts["word/commentsIds.xml"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w16cid:commentsIds xmlns:w16cid="{W16CID}">' + "".join(ids) + "</w16cid:commentsIds>")
            parts["[Content_Types].xml"] = types.replace("</Types>", (
                '<Override PartName="/word/commentsIds.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/></Types>'))
            parts["word/_rels/document.xml.rels"] = rels.replace("</Relationships>", (
                '<Relationship Id="rId90" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" '
                'Target="commentsIds.xml"/></Relationships>'))
        if replies:
            extended = "".join(
                f'<w15:commentEx w15:paraId="{0x10000000 + cid:08X}"'
                + (f' w15:paraIdParent="{0x10000000 + replies[cid]:08X}"' if cid in replies else "")
                + ' w15:done="0"/>' for cid, _, _ in comments)
            parts["word/commentsExtended.xml"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w15:commentsEx xmlns:w15="{W15}">' + extended + "</w15:commentsEx>")
            parts["[Content_Types].xml"] = parts.get("[Content_Types].xml", types).replace("</Types>", (
                '<Override PartName="/word/commentsExtended.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>'
                '</Types>'))
            parts["word/_rels/document.xml.rels"] = parts.get("word/_rels/document.xml.rels", rels).replace(
                "</Relationships>", (
                    '<Relationship Id="rId91" Type="http://schemas.microsoft.com/office/2011/relationships/'
                    'commentsExtended" Target="commentsExtended.xml"/></Relationships>'))
        with zipfile.ZipFile(target, "x") as out:
            for info in src.infolist():
                out.writestr(info, parts.pop(info.filename, None) or src.read(info.filename),
                             compress_type=info.compress_type)
            for part, text in parts.items():
                out.writestr(part, text, compress_type=zipfile.ZIP_DEFLATED)
    return target


def marked(path: Path, index: int, mode: str = "raw") -> str:
    """Paragraph `index` with its anchors: [cN .. cN] a comment range, ^N its reference, [Name .. Name] a
    bookmark, / a line break. mode "raw" shows {+inserted+} and {-deleted-} text, "accept" and "reject" the
    text after Accept All / Reject All (markers stay where they are in both)."""

    root = ET.fromstring(document_xml(path))
    names = {item.get(W + "id"): item.get(W + "name") for item in root.iter(W + "bookmarkStart")}
    out: list[str] = []

    def walk(element: ET.Element, hidden: bool) -> None:
        for child in element:
            tag = child.tag[len(W):] if child.tag.startswith(W) else ""
            ident = child.get(W + "id")
            if tag in ("pPr", "rPr"):
                continue
            if tag in ("ins", "del"):
                drop = (mode == "accept" and tag == "del") or (mode == "reject" and tag == "ins")
                if mode == "raw":
                    out.append("{+" if tag == "ins" else "{-")
                walk(child, hidden or drop)
                if mode == "raw":
                    out.append("+}" if tag == "ins" else "-}")
            elif tag in ("t", "delText"):
                out.append("" if hidden else child.text or "")
            elif tag == "br":
                out.append("" if hidden else "/")
            elif tag == "commentRangeStart":
                out.append(f"[c{ident}")
            elif tag == "commentRangeEnd":
                out.append(f"c{ident}]")
            elif tag == "commentReference":
                out.append(f"^{ident}")
            elif tag == "bookmarkStart":
                out.append(f"[{child.get(W + 'name')}")
            elif tag == "bookmarkEnd":
                out.append(f"{names.get(ident, '?')}]")
            else:
                walk(child, hidden)

    walk(list(root.iter(W + "p"))[index - 1], False)
    return "".join(out)


def revision_authors(path: Path) -> set[str]:
    root = ET.fromstring(document_xml(path))
    return {item.get(W + "author") for tag in ("ins", "del") for item in root.iter(W + tag)}


# -- (a) a comment's range end, then a point comment ------------------------------------------------------

CASE_A = para("", run("The tests were "), comment_start(1), run("two-sided permutation tests"), comment_end(1),
              run(" of the gap"), comment_start(2), '<w:commentRangeEnd w:id="2"/>',
              '<w:r><w:commentReference w:id="2"/></w:r>', run("."))
CASE_A_COMMENTS = [(1, "Name the test.", "0A000001"), (2, "Which gap?", "0A000002")]


def test_a_text_typed_at_a_comment_end_and_at_a_point_comment_stays_outside_both(tmp_path):
    source = make(tmp_path, "a", CASE_A, CASE_A_COMMENTS)
    assert marked(source, 1) == "The tests were [c1two-sided permutation testsc1]^1 of the gap[c2c2]^2."
    out = tmp_path / "a-out.docx"
    tracked(source, out, 1, text="The tests were two-sided permutation tests, with draws, of the gap in mean score.")
    # Word types before a point comment (plan 924): the new text precedes its start, end and reference.
    assert marked(out, 1) == ("The tests were [c1two-sided permutation testsc1]^1{+, with draws,+} of the gap"
                              "{+ in mean score+}[c2c2]^2.")
    assert marked(out, 1, "accept") == ("The tests were [c1two-sided permutation testsc1]^1, with draws, of the gap"
                                        " in mean score[c2c2]^2.")
    assert marked(out, 1, "reject") == marked(source, 1)


def test_a_two_edits_in_one_batch_each_typed_outside_the_comments(tmp_path):
    source = make(tmp_path, "a2", CASE_A, CASE_A_COMMENTS)
    first, second = tmp_path / "a2-1.docx", tmp_path / "a2-2.docx"
    tracked(source, first, 1, find="permutation tests", replace="permutation tests, with draws,")
    tracked(first, second, 1, find="the gap", replace="the gap in mean score")
    assert marked(second, 1) == ("The tests were [c1two-sided permutation testsc1]^1{+, with draws,+} of the gap"
                                 "{+ in mean score+}[c2c2]^2.")


# -- (b) a paragraph with line breaks loses its last lines ------------------------------------------------

def test_b_deleting_lines_after_a_bookmark_end_keeps_the_bookmark(tmp_path):
    body = (para("", run("Changes "), bookmark_start(20, "Span"), run("listed below."))
            + para("", run("Revised lines 3-4"), bookmark_end(20),
                   '<w:r><w:br/><w:t xml:space="preserve">Before: old wording</w:t></w:r>',
                   '<w:r><w:br/><w:t xml:space="preserve">After: new wording</w:t></w:r>'))
    source = make(tmp_path, "b", body)
    assert marked(source, 2) == "Revised lines 3-4Span]/Before: old wording/After: new wording"
    out = tmp_path / "b-out.docx"
    tracked(source, out, 2, text="Revised lines 3-4")
    assert marked(out, 1) == "Changes [Spanlisted below."
    assert marked(out, 2) == "Revised lines 3-4Span]{-/Before: old wording-}{-/After: new wording-}"
    assert marked(out, 2, "reject") == marked(source, 2)


def test_b_emptying_the_paragraph_that_ends_a_bookmark_keeps_both_views_and_the_bookmark(tmp_path):
    # The logged shape: tracked_replace_paragraph with text "" on a paragraph with line breaks whose bookmark
    # started in the paragraph before ("paragraphs do not line up with the edit plan" before bc2aec0).
    body = (para("", run("Changes "), bookmark_start(20, "Span"), run("listed below."))
            + para("", run("Revised lines 3-4"), bookmark_end(20),
                   '<w:r><w:br/><w:t xml:space="preserve">Before: old wording</w:t></w:r>',
                   '<w:r><w:br/><w:t xml:space="preserve">After: new wording</w:t></w:r>'))
    source = make(tmp_path, "b0", body)
    out = tmp_path / "b0-out.docx"
    tracked(source, out, 2, text="")
    assert marked(out, 1) == "Changes [Spanlisted below."
    assert marked(out, 2) == "{-Revised lines 3-4-}Span]{-/Before: old wording-}{-/After: new wording-}"
    assert marked(out, 2, "accept") == "Span]"
    assert marked(out, 2, "reject") == marked(source, 2)
    assert marked(out, 1, "accept") == marked(out, 1, "reject") == marked(source, 1)


# -- (c) reject another author's insertion and add a comma at a comment end, one group ------------------

CASE_C = para("", run("In reply, "), comment_start(3), run("we removed the comparison and its reference"),
              comment_end(3), run(" because no source was found"), rev("ins", 501, run(" in the archive")), run("."))


def test_c_comma_typed_at_a_comment_end_stays_outside_it(tmp_path):
    source = make(tmp_path, "c", CASE_C, [(3, "Say why.", "0C000003")])
    out = tmp_path / "c-out.docx"
    tracked(source, out, 1, find="its reference", replace="its reference,")
    assert marked(out, 1) == ("In reply, [c3we removed the comparison and its referencec3]^3{+,+} because no "
                              "source was found{+ in the archive+}.")


@windows_only
def test_c_reject_and_tracked_comma_in_one_group_publish(tmp_path):
    path = make(tmp_path, "c2", CASE_C, [(3, "Say why.", "0C000003")])
    result = run_batch(path,
                       {"op_id": "r", "type": "reject_revisions", "ids": ["501"], "group": "reply"},
                       {"op_id": "t", "type": "tracked_replace_text", "paragraph_index": 1, "find": "its reference",
                        "replace": "its reference,", "group": "reply"})
    assert result.status == "complete", result
    assert marked(path, 1) == ("In reply, [c3we removed the comparison and its referencec3]^3{+,+} because no "
                               "source was found.")


# -- (d) text appended where one bookmark ends and another starts -----------------------------------------

def test_d_text_appended_at_paragraph_end_leaves_the_ending_bookmark_and_joins_the_starting_one(tmp_path):
    body = (para("", run("Results. "), bookmark_start(10, "Spanning"), run("Conclusions"))
            + para("", run("A secondary refinement."), bookmark_start(13, "Following"), bookmark_end(10))
            + para("", run("Next part."), bookmark_end(13)))
    source = make(tmp_path, "d", body)
    out = tmp_path / "d-out.docx"
    tracked(source, out, 2, find="refinement.", replace="refinement. It was minor.")
    assert marked(out, 1) == "Results. [SpanningConclusions"
    assert marked(out, 2) == "A secondary refinement.[FollowingSpanning]{+ It was minor.+}"
    assert marked(out, 2, "reject") == marked(source, 2)


# -- (e) all of a comment's text deleted: refused, naming its comment_id and the op -------------------------

CASE_E = para("", run("Recall rose to "), comment_start(4), run("44%"), comment_end(4), run(" overall"),
              comment_start(5), run(" today"), comment_end(5), run("."))
CASE_E_COMMENTS = [(4, "Same claim as before.", "7F365A60"), (5, "Which day?", None)]


def test_e_deleting_all_of_a_comment_text_names_its_comment_id_and_the_op(tmp_path):
    source = make(tmp_path, "e", CASE_E, CASE_E_COMMENTS)
    with pytest.raises(VerificationFailed) as caught:
        tracked(source, tmp_path / "e-out.docx", 1, find="rose to 44% overall today.", replace="rose overall.")
    message = str(caught.value)
    assert ("comment (Reviewer B: 'Same claim as before.') on '44%' would be left empty because all of its "
            "text is deleted; delete_comment it in this batch, or keep some of its text (put "
            '{"type": "delete_comment", "comment_id": "7F365A60"} before this edit)') in message
    assert ("comment (Reviewer B: 'Which day?') on ' today' would be left empty because all of its text is "
            "deleted; delete_comment it in this batch, or keep some of its text (put "
            '{"type": "delete_comment", "comment_id": "legacy:5"} before this edit)') in message
    assert not (tmp_path / "e-out.docx").exists()


def test_e_two_comments_with_the_same_author_and_text_name_only_the_one_whose_text_is_deleted(tmp_path):
    body = para("", comment_start(6), run("Alpha"), comment_end(6), run(" and "), comment_start(7), run("beta"),
                comment_end(7), run("."))
    source = make(tmp_path, "e3", body, [(6, "Check.", "0E000006"), (7, "Check.", "0E000007")])
    with pytest.raises(VerificationFailed) as caught:
        tracked(source, tmp_path / "e3-out.docx", 1, find="Alpha and ", replace="")
    message = str(caught.value)
    assert '{"type": "delete_comment", "comment_id": "0E000006"}' in message
    assert "0E000007" not in message


def test_e_a_reply_on_deleted_text_points_at_its_thread_root(tmp_path):
    # Word gives a reply its own range markers on the root's text; delete_comment refuses a reply with the
    # default include_replies, so both refusals name the root, whose delete removes the whole thread.
    body = para("", run("Keep "), comment_start(8), comment_start(9), run("gone"), comment_end(8), comment_end(9),
                run(" here."))
    source = make(tmp_path, "e4", body, [(8, "Why?", "0E000008"), (9, "Because.", "0E000009")], replies={9: 8})
    with pytest.raises(VerificationFailed) as caught:
        tracked(source, tmp_path / "e4-out.docx", 1, find="gone here.", replace="here.")
    message = str(caught.value)
    assert message.count('{"type": "delete_comment", "comment_id": "0E000008"}') == 2
    assert "0E000009" not in message


@windows_only
def test_e_the_named_delete_comment_then_the_edit_publish_in_one_batch(tmp_path):
    path = make(tmp_path, "e2", CASE_E, CASE_E_COMMENTS)
    result = run_batch(path,
                       {"op_id": "d4", "type": "delete_comment", "comment_id": "7F365A60"},
                       {"op_id": "d5", "type": "delete_comment", "comment_id": "legacy:5"},
                       {"op_id": "t", "type": "tracked_replace_text", "paragraph_index": 1,
                        "find": "rose to 44% overall today.", "replace": "rose overall."})
    assert result.status == "complete", result
    # The deleted comments' markers are gone; the runs they split stay split, one w:del each.
    assert marked(path, 1) == "Recall rose {-to -}{-44%-}{- -}overall{- today-}."
    assert marked(path, 1, "accept") == "Recall rose overall."
    with zipfile.ZipFile(path) as archive:
        assert "<w:comment " not in archive.read("word/comments.xml").decode("utf-8")


# -- typing at a paragraph's edges ---------------------------------------------------------------------

def test_text_typed_at_a_comment_start_stays_before_it(tmp_path):
    body = para("", comment_start(6), run("Alpha beta."), comment_end(6))
    source = make(tmp_path, "start", body, [(6, "Check.", "0E000006")])
    out = tmp_path / "start-out.docx"
    tracked(source, out, 1, find="Alpha", replace="New Alpha")
    assert marked(out, 1) == "{+New +}[c6Alpha beta.c6]^6"


def test_text_typed_into_an_empty_paragraph_stays_before_a_comment_starting_there(tmp_path):
    body = para("", comment_start(7)) + para("", run("Next line."), comment_end(7))
    source = make(tmp_path, "empty", body, [(7, "Check.", "0E000007")])
    out = tmp_path / "empty-out.docx"
    tracked(source, out, 1, text="Typed")
    assert marked(out, 1) == "{+Typed+}[c7"
    assert marked(out, 2) == "Next line.c7]^7"


# -- sibling: a revision-free CV, research-mcp tracking on --------------------------------------------------

@windows_only
def test_sibling_cv_tracking_on_text_typed_at_a_comment_end_and_a_bookmark_start(tmp_path):
    body = (para("", run("Research Assistant, "), comment_start(8), run("Example Lab"), comment_end(8))
            + para("", run("Skills: "), bookmark_start(30, "Skills"), run("Python, R"), bookmark_end(30)))
    path = make(tmp_path, "cv", body, [(8, "Add the dates.", "0F000008")])
    result = run_batch(path,
                       {"op_id": "on", "type": "set_track_changes", "on": True, "author": "Reviewer B"},
                       {"op_id": "dates", "type": "replace_text", "paragraph_index": 1, "find": "Example Lab",
                        "replace": "Example Lab (2024-2025)"},
                       {"op_id": "level", "type": "replace_text", "paragraph_index": 2, "find": "Python",
                        "replace": "Advanced Python"})
    assert result.status == "complete", result
    assert marked(path, 1) == "Research Assistant, [c8Example Labc8]^8{+ (2024-2025)+}"
    assert marked(path, 2) == "Skills: [Skills{+Advanced +}Python, RSkills]"
    assert revision_authors(path) == {"Reviewer B"}
    assert view(list(ET.fromstring(document_xml(path)).iter(W + "p"))[1]) == "Skills: Advanced Python, R"


# -- plan 924: zero-length anchors, as Word types at them ----------------------------------------------------
# Measured in Word 2026-11-17 (TypeText with Track Changes on, Range.InsertAfter/InsertBefore with it off, on
# either side of the reference mark): text goes before a point comment's start, end and reference, and after a
# collapsed bookmark; a point comment then a collapsed bookmark at one point come out as bookmark, text,
# comment. `test_point_anchor_word_live.py` replays these against Word itself.

def point(cid: int) -> str:
    return comment_start(cid) + comment_end(cid)


def collapsed(bid: int, name: str) -> str:
    return bookmark_start(bid, name) + bookmark_end(bid)


POINT_BODY = "".join([
    para("", run("Led the data team"), point(21), run(". Built the pipeline.")),
    para("", point(22), run("Built the pipeline.")),
    para("", run("Led the data team."), point(23)),
    para("", point(24)),
    para("", run("Led the data team"), collapsed(41, "Ref_X"), run(". Built the pipeline.")),
    para("", collapsed(42, "Ref_S"), run("Built the pipeline.")),
    para("", run("Led the data team."), collapsed(43, "Ref_E")),
    para("", collapsed(44, "Ref_M")),
    para("", run("Led the data team"), point(25), collapsed(45, "Ref_Y"), run(". Built.")),
    para("", run("Led "), comment_start(26), run("the data team"), comment_end(26), point(27), run(". Built.")),
])
POINT_COMMENTS = [(cid, f"Point {cid}.", f"0B0000{cid:02d}") for cid in (21, 22, 23, 24, 25, 26, 27)]
# paragraph: (its text, where "<n>" goes, the result Word gives)
POINT_CASES = {
    1: ("Led the data team. Built the pipeline.", 17, "Led the data team{<1>}[c21c21]^21. Built the pipeline."),
    2: ("Built the pipeline.", 0, "{<2>}[c22c22]^22Built the pipeline."),
    3: ("Led the data team.", 18, "Led the data team.{<3>}[c23c23]^23"),
    4: ("", 0, "{<4>}[c24c24]^24"),
    5: ("Led the data team. Built the pipeline.", 17, "Led the data team[Ref_XRef_X]{<5>}. Built the pipeline."),
    6: ("Built the pipeline.", 0, "[Ref_SRef_S]{<6>}Built the pipeline."),
    7: ("Led the data team.", 18, "Led the data team.[Ref_ERef_E]{<7>}"),
    8: ("", 0, "[Ref_MRef_M]{<8>}"),
    9: ("Led the data team. Built.", 17, "Led the data team[Ref_YRef_Y]{<9>}[c25c25]^25. Built."),
    10: ("Led the data team. Built.", 17, "Led [c26the data teamc26]^26{<10>}[c27c27]^27. Built."),
}


def point_ops() -> list[dict]:
    return [{"op_id": f"p{index}", "type": "replace_paragraph", "paragraph_index": index,
             "text": text[:at] + f"<{index}>" + text[at:]} for index, (text, at, _) in POINT_CASES.items()]


@pytest.mark.parametrize("index", sorted(POINT_CASES))
def test_point_anchor_tracked_text_lands_where_word_types_it(tmp_path, index):
    text, at, word = POINT_CASES[index]
    source = make(tmp_path, f"pt{index}", POINT_BODY, POINT_COMMENTS)
    out = tmp_path / f"pt{index}-out.docx"
    tracked(source, out, index, text=text[:at] + f"<{index}>" + text[at:], author="Reviewer B")
    assert marked(out, index) == word.replace("{", "{+").replace("}", "+}")
    assert revision_authors(out) == {"Reviewer B"}
    rejected = marked(out, index, "reject")
    if index == 9:  # Word's order: the collapsed bookmark moves before the point comment; no text moves
        assert rejected == "Led the data team[Ref_YRef_Y][c25c25]^25. Built."
    else:
        assert rejected == marked(source, index)
    assert marked(out, index, "accept") == word.replace("{", "").replace("}", "")


@pytest.mark.parametrize("index", sorted(POINT_CASES))
def test_point_anchor_plain_edit_on_a_revised_document_lands_where_word_types_it(tmp_path, index):
    text, at, word = POINT_CASES[index]
    source = make(tmp_path, f"pp{index}", POINT_BODY + para("", run("Other "), rev("ins", 501, run("line"))),
                  POINT_COMMENTS)
    out = tmp_path / f"pp{index}-out.docx"
    revisions.text_edit(source, out, op="replace_paragraph", paragraph_index=index, tracked=False,
                        text=text[:at] + f"<{index}>" + text[at:])
    assert marked(out, index) == word.replace("{", "").replace("}", "")
    assert marked(out, 11) == "Other {+line+}"


def test_point_anchor_check_accepts_words_side_and_names_the_point_otherwise(tmp_path):
    # sections' anchor check is what the plain Word path is verified with: Word's own result (text before the
    # point comment) passes, and text after it is refused naming the point by the text around it.
    from research_mcp.sections import StreamEdit, anchor_errors_for_edits, read_layout

    before = make(tmp_path, "chk", para("", run("Led the data team"), point(31), run(". Built.")), [(31, "Why?", None)])
    words = make(tmp_path, "chk-word", para("", run("Led the data team"), run(" of five"), point(31), run(". Built.")),
                 [(31, "Why?", None)])
    other = make(tmp_path, "chk-after", para("", run("Led the data team"), point(31), run(" of five"), run(". Built.")),
                 [(31, "Why?", None)])
    edits = [StreamEdit(17, 0, 8)]
    lengths = [len("Led the data team of five. Built.")]
    assert anchor_errors_for_edits(read_layout(before), read_layout(words), edits, lengths) == []
    problems = anchor_errors_for_edits(read_layout(before), read_layout(other), edits, lengths)
    assert problems == ["comment (Reviewer B: 'Why?') would sit at 'd the data team| of five. Built' instead of "
                        "'ta team of five|. Built.¶' (was 'd the data team|. Built.¶')"]
