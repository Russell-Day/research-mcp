"""Plan 831 S1: tracked (and plain) edits on paragraphs with manual line breaks and on empty paragraphs.

Synthetic fixture only. Every expectation is read back with this file's own ElementTree helpers (``tokens``,
``accepted_text``), never with revisions.py, so a test fails when the engine's fix is removed:

* rewriting text in a run that holds a ``w:br`` (or a tab or page break) splits that run; rejecting the new
  revisions must give back the input paragraph, character for character with its formatting;
* deleting or inserting a line break as a tracked change;
* a tracked insertion into an empty paragraph lands inside the ``w:p`` (also a self-closing ``<w:p/>``) and
  takes the paragraph mark's formatting, without the mark's own revision markers.
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
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.sections import read_layout
from test_revisions import W, run_batch, windows_only
from test_revisions_guards import variant

OTHER = "Guest User"
DATE = "2026-11-02T10:00:00Z"
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
SZ = '<w:rPr><w:sz w:val="22"/></w:rPr>'


def r(*parts: str, rpr: str = SZ) -> str:
    return f"<w:r>{rpr}{''.join(parts)}</w:r>"


def t(text: str) -> str:
    return f'<w:t xml:space="preserve">{text}</w:t>'


BR = "<w:br/>"


def p(*content: str, ppr: str = "") -> str:
    return "<w:p>" + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + "".join(content) + "</w:p>"


def cell(*paras: str) -> str:
    return '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + "".join(paras) + "</w:tc>"


EMPTY_MARK_PPR = '<w:jc w:val="center"/><w:rPr><w:b/><w:sz w:val="22"/></w:rPr>'
MARKED_MARK_PPR = (f'<w:rPr><w:ins w:id="40" w:author="{OTHER}" w:date="{DATE}"/><w:i/>'
                   f'<w:rPrChange w:id="41" w:author="{OTHER}" w:date="{DATE}"><w:rPr/></w:rPrChange></w:rPr>')

BODY = (
    # 1: two manual line breaks, each starting its own run (as Word writes a multi-line response cell)
    p(r(t("Alpha one")), r(BR, t("Alpha two")), r(BR, t("Alpha three")))
    # 2: the response-cell shape: line breaks, then a comment anchored on words of the second line
    + p(r(t("Head line")), r(BR, t("Before: one ")), '<w:commentRangeStart w:id="0"/>', r(t("commented words")),
        '<w:commentRangeEnd w:id="0"/>', '<w:r><w:commentReference w:id="0"/></w:r>', r(t(", tail")),
        r(BR, t("After: end")))
    # 3: a page break inside one run
    + p(r(t("Before break"), '<w:br w:type="page"/>', t("After page")))
    # 4: a tab inside one run
    + p(r(t("Name"), "<w:tab/>", t("Value here")))
    # 5: empty, pPr with a mark rPr
    + p(ppr=EMPTY_MARK_PPR)
    # 6: empty, mark rPr carrying another author's mark insertion and formatting change
    + p(ppr=MARKED_MARK_PPR)
    # 7: empty and self-closing
    + '<w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD"/>'
    # 8: another author's pending insertion
    + p(r(t("Existing ")), f'<w:ins w:id="42" w:author="{OTHER}" w:date="{DATE}">{r(t("inserted"))}</w:ins>',
        r(t(" text.")))
    + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>'
      '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
    # 9, 10, 11 (12 is the row end): a point with a line break, a reply, an empty self-closing right cell
    + "<w:tr>" + cell(p(r(t("Reviewer point")), r(BR, t("second line")))) + cell(p(r(t("Our reply."))))
    + cell('<w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD"/>') + "</w:tr>"
    # 13, 14, 15 (16 is the row end): an empty right cell with a mark rPr
    + "<w:tr>" + cell(p(r(t("Point two")))) + cell(p(r(t("Reply two.")))) + cell(p(ppr=SZ)) + "</w:tr>"
    + "</w:tbl>"
    # 17
    + p(r(t("Closing.")))
    # 18: empty, holding a collapsed bookmark
    + p('<w:bookmarkStart w:id="50" w:name="Spot"/><w:bookmarkEnd w:id="50"/>', ppr=SZ)
)

PARA = {"breaks": 1, "comment": 2, "page": 3, "tab": 4, "empty_mark": 5, "empty_marked": 6, "self_closing": 7,
        "other_rev": 8, "cell_breaks": 9, "cell_empty": 11, "cell_empty_mark": 15, "closing": 17,
        "empty_bookmark": 18}


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    base = create_tracked_fixture(tmp_path / "base.docx")

    def swap(xml: str) -> str:
        start = xml.index("<w:body>") + len("<w:body>")
        return xml[:start] + BODY + SECT + "</w:body></w:document>"

    return variant(base, tmp_path / "breaks.docx", swap)


# -- independent readers -----------------------------------------------------------------------------------

def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


_P_SPAN = re.compile(r"<w:p(?=[\s>/])[^>]*?/>|<w:p(?=[\s>])[^>]*>.*?</w:p>", re.S)


def spans(xml: str) -> list[tuple[int, int]]:
    return [m.span() for m in _P_SPAN.finditer(xml)]


def para_elements(path: Path) -> list[ET.Element]:
    return list(ET.fromstring(document_xml(path)).iter(W + "p"))


def layout_ordinal(index: int) -> int:
    """Layout index -> position among the document's w:p elements (row-end markers have none)."""
    return index - (1 if index > 12 else 0) - (1 if index > 16 else 0) - 1


def element(path: Path, index: int) -> ET.Element:
    return para_elements(path)[layout_ordinal(index)]


def _rpr(run: ET.Element) -> str:
    node = run.find(W + "rPr")
    return "" if node is None else ET.tostring(node, encoding="unicode")


def tokens(paragraph: ET.Element, *, reject: set[str] = frozenset()) -> list[tuple]:
    """Characters and markers with their run formatting; revisions in `reject` are rejected, all others kept
    as they are (their wrapper is part of each token, so another author's revision must survive intact)."""

    out: list[tuple] = []

    def walk(node: ET.Element, context: tuple) -> None:
        for child in node:
            name = child.tag.replace(W, "")
            if name == "pPr":
                out.append(("pPr", ET.tostring(child, encoding="unicode")))
                continue
            if name in ("ins", "del"):
                rid = child.get(W + "id")
                if rid in reject:
                    if name == "ins":
                        continue
                    walk(child, context)
                else:
                    walk(child, context + ((name, rid, child.get(W + "author")),))
                continue
            if name == "r":
                rpr = _rpr(child)
                for part in child:
                    part_name = part.tag.replace(W, "")
                    if part_name in ("t", "delText"):
                        out.extend(("ch", ch, rpr, context) for ch in part.text or "")
                    elif part_name in ("br", "cr", "tab"):
                        out.append(("ch", {"tab": "\t", "cr": "\v"}.get(part_name, "\x0c" if part.get(
                            W + "type") == "page" else "\v"), rpr, context))
                    elif part_name != "rPr":
                        out.append((part_name, part.get(W + "id"), rpr, context))
                continue
            out.append((name, child.get(W + "id"), context))
            walk(child, context)

    walk(paragraph, ())
    return out


def accepted_text(paragraph: ET.Element) -> str:
    """Text with every revision accepted (the view the caller's request describes)."""

    out: list[str] = []

    def walk(node: ET.Element, deleted: bool) -> None:
        for child in node:
            name = child.tag.replace(W, "")
            if name in ("pPr", "rPr"):
                continue
            if name == "t" and not deleted:
                out.append(child.text or "")
            elif name in ("br", "cr") and not deleted:
                out.append("\x0c" if child.get(W + "type") == "page" else "\v")
            elif name == "tab" and not deleted:
                out.append("\t")
            else:
                walk(child, deleted or name == "del")

    walk(paragraph, False)
    return "".join(out)


def new_ids(source: Path, output: Path) -> set[str]:
    def ids(path):
        return {el.get(W + "id") for el in ET.fromstring(document_xml(path)).iter() if el.tag in (W + "ins", W + "del")}
    return ids(output) - ids(source)


def tracked(doc: Path, index: int, **kwargs) -> tuple[Path, dict]:
    out = doc.parent / f"out-{os.urandom(3).hex()}.docx"
    op = "tracked_replace_text" if "find" in kwargs else "tracked_replace_paragraph"
    result = revisions.text_edit(doc, out, op=op, paragraph_index=index, tracked=True, author="Claude", **kwargs)
    return out, result


def check_tracked(doc: Path, out: Path, index: int, requested: str) -> set[str]:
    """Accept All gives the request, Reject All of the new revisions gives the input, nothing else moved."""

    ids = new_ids(doc, out)
    assert ids, "no tracked revision was written"
    before, after = element(doc, index), element(out, index)
    assert accepted_text(after) == requested
    assert tokens(after, reject=ids) == tokens(before)
    old_xml, new_xml = document_xml(doc), document_xml(out)
    (a, b), (c, d) = spans(old_xml)[layout_ordinal(index)], spans(new_xml)[layout_ordinal(index)]
    assert old_xml[:a] == new_xml[:c] and old_xml[b:] == new_xml[d:], "bytes outside the paragraph changed"
    for el in ET.fromstring(new_xml).iter():
        if el.tag in (W + "ins", W + "del") and el.get(W + "id") in ids:
            assert el.get(W + "author") == "Claude"
    return ids


def layout_text(path: Path, index: int) -> str:
    return read_layout(path).para(index).text


# -- the fixture itself --------------------------------------------------------------------------------------

def test_fixture_layout(doc):
    layout = read_layout(doc)
    assert layout.para(PARA["breaks"]).text == "Alpha one\vAlpha two\vAlpha three"
    assert layout.para(PARA["page"]).text == "Before break\x0cAfter page"
    assert layout.para(PARA["cell_breaks"]).text == "Reviewer point\vsecond line"
    for key in ("empty_mark", "empty_marked", "self_closing", "cell_empty", "cell_empty_mark", "empty_bookmark"):
        assert layout.para(PARA[key]).text == "", key
    assert layout.para(PARA["cell_empty"]).cell == 2 and layout.para(PARA["cell_empty_mark"]).cell == 2
    assert layout.para(PARA["closing"]).text == "Closing."
    for index in PARA.values():
        assert accepted_text(element(doc, index)) == layout.para(index).text


# -- (a) line breaks ------------------------------------------------------------------------------------------

@pytest.mark.parametrize("find, replace", [
    ("one\vAlpha two", "uno\vAlpha dos"),     # text on both sides of a break
    ("Alpha two", "Alpha 2"),                  # inside a run that starts with a break
    ("e\vA", "eA"),                            # delete a line break
    ("one\vAlpha", "one Alpha"),               # a line break becomes a space
    ("Alpha two", "Alpha\vtwo"),               # a new line break inside a run holding one
    ("Alpha three", "Alpha three\vAlpha four"),  # a new line at the end
])
def test_tracked_replace_text_on_line_breaks(doc, find, replace):
    old = "Alpha one\vAlpha two\vAlpha three"
    requested = old.replace(find, replace)
    out, result = tracked(doc, PARA["breaks"], find=find, replace=replace)
    assert result["text_after"] == requested
    check_tracked(doc, out, PARA["breaks"], requested)


def test_tracked_replace_paragraph_rewrites_every_line(doc):
    requested = "Alpha first\vBeta two\vAlpha three, now longer"
    out, _ = tracked(doc, PARA["breaks"], text=requested)
    check_tracked(doc, out, PARA["breaks"], requested)


def test_deleted_line_break_stays_a_tracked_break(doc):
    out, _ = tracked(doc, PARA["breaks"], find="e\vA", replace="eA")
    after = element(out, PARA["breaks"])
    deleted = [el for el in after.iter(W + "del") if el.get(W + "author") == "Claude"]
    assert len(deleted) == 1 and [c.tag for c in deleted[0].iter() if c.tag in (W + "br", W + "t", W + "delText")] \
        == [W + "br"]


def test_response_cell_shape_keeps_its_comment(doc):
    """The reported paragraph shape: rewrite a line-break run next to an anchored comment."""
    old = "Head line\vBefore: one commented words, tail\vAfter: end"
    out, _ = tracked(doc, PARA["comment"], find="Before: one", replace="Before: two")
    check_tracked(doc, out, PARA["comment"], old.replace("Before: one", "Before: two"))
    requested = old.replace("After: end", "After: finished")
    out2, _ = tracked(doc, PARA["comment"], text=requested)
    check_tracked(doc, out2, PARA["comment"], requested)
    tree = element(out2, PARA["comment"])
    names = [el.tag.replace(W, "") for el in tree.iter() if el.tag in (W + "commentRangeStart", W + "commentRangeEnd",
                                                                       W + "t")]
    start, end = names.index("commentRangeStart"), names.index("commentRangeEnd")
    covered = [el for el in tree.iter() if el.tag == W + "t"][start:end - 1]
    assert "".join(el.text for el in covered) == "commented words"


def test_page_break_run(doc):
    out, _ = tracked(doc, PARA["page"], find="Before break", replace="Before a break")
    check_tracked(doc, out, PARA["page"], "Before a break\x0cAfter page")
    out2, _ = tracked(doc, PARA["page"], find="After page", replace="After the page")
    check_tracked(doc, out2, PARA["page"], "Before break\x0cAfter the page")


def test_tab_replaced_by_space(doc):
    """A deleted tab stays in the layout text (as a deleted break does); anchors are still measured."""
    out, _ = tracked(doc, PARA["tab"], find="Name\tValue", replace="Name Value")
    check_tracked(doc, out, PARA["tab"], "Name Value here")


def test_table_cell_with_line_break(doc):
    out, _ = tracked(doc, PARA["cell_breaks"], find="second line", replace="next line")
    check_tracked(doc, out, PARA["cell_breaks"], "Reviewer point\vnext line")


def test_other_authors_revision_untouched(doc):
    out, _ = tracked(doc, PARA["other_rev"], find="Existing", replace="Prior")
    check_tracked(doc, out, PARA["other_rev"], "Prior inserted text.")
    raw = re.search(r'<w:ins w:id="42"[^>]*>.*?</w:ins>', document_xml(out), re.S).group(0)
    assert raw == re.search(r'<w:ins w:id="42"[^>]*>.*?</w:ins>', document_xml(doc), re.S).group(0)


# -- (b) empty paragraphs -------------------------------------------------------------------------------------

def _inserted_runs(path: Path, index: int) -> list[ET.Element]:
    return [run for ins in element(path, index).findall(W + "ins") for run in ins.findall(W + "r")]


@pytest.mark.parametrize("key", ["self_closing", "cell_empty"])
def test_fill_self_closing_paragraph(doc, key):
    index = PARA[key]
    out, result = tracked(doc, index, text="Filled reply.")
    assert result["text_before"] == "" and result["text_after"] == "Filled reply."
    check_tracked(doc, out, index, "Filled reply.")
    after = element(out, index)
    assert after.get(W + "rsidR") == "00AB12CD"  # the same paragraph, expanded, not a sibling
    runs = _inserted_runs(out, index)
    assert len(runs) == 1 and runs[0].find(W + "rPr") is None
    assert layout_text(out, index) == "Filled reply."


def test_fill_paragraph_takes_mark_formatting(doc):
    out, _ = tracked(doc, PARA["empty_mark"], text="Centered bold.")
    check_tracked(doc, out, PARA["empty_mark"], "Centered bold.")
    runs = _inserted_runs(out, PARA["empty_mark"])
    assert len(runs) == 1
    assert [c.tag for c in runs[0].find(W + "rPr")] == [W + "b", W + "sz"]
    after = element(out, PARA["empty_mark"])
    assert list(after)[0].tag == W + "pPr" and list(after)[1].tag == W + "ins"


def test_fill_paragraph_drops_mark_revision_markers(doc):
    out, _ = tracked(doc, PARA["empty_marked"], text="New words.")
    check_tracked(doc, out, PARA["empty_marked"], "New words.")
    runs = _inserted_runs(out, PARA["empty_marked"])
    assert len(runs) == 1 and [c.tag for c in runs[0].find(W + "rPr")] == [W + "i"]
    # the mark's own revisions are still there, unchanged
    xml = document_xml(out)
    assert f'<w:ins w:id="40" w:author="{OTHER}" w:date="{DATE}"/>' in xml
    assert f'<w:rPrChange w:id="41" w:author="{OTHER}" w:date="{DATE}"><w:rPr/></w:rPrChange>' in xml


def test_fill_empty_cell_with_mark_rpr(doc):
    out, _ = tracked(doc, PARA["cell_empty_mark"], text="Reply three.")
    check_tracked(doc, out, PARA["cell_empty_mark"], "Reply three.")
    runs = _inserted_runs(out, PARA["cell_empty_mark"])
    assert [c.tag for c in runs[0].find(W + "rPr")] == [W + "sz"]


def test_fill_paragraph_keeps_collapsed_bookmark_before_text(doc):
    """Typed text goes after the paragraph's markers: a collapsed bookmark stays at the start, as in Word."""
    out, _ = tracked(doc, PARA["empty_bookmark"], text="After the mark.")
    check_tracked(doc, out, PARA["empty_bookmark"], "After the mark.")
    names = [el.tag.replace(W, "") for el in element(out, PARA["empty_bookmark"])]
    assert names == ["pPr", "bookmarkStart", "bookmarkEnd", "ins"]


def test_fill_empty_paragraph_with_line_break(doc):
    out, _ = tracked(doc, PARA["self_closing"], text="First line\vsecond line")
    check_tracked(doc, out, PARA["self_closing"], "First line\vsecond line")


# -- plain edits on a document with revisions (revisions.plain_replace) ------------------------------------------

def _plain(doc: Path, index: int, **kwargs) -> Path:
    out = doc.parent / f"plain-{os.urandom(3).hex()}.docx"
    revisions.plain_replace(doc, out, op="replace_text" if "find" in kwargs else "replace_paragraph",
                            paragraph_index=index, **kwargs)
    return out


@pytest.mark.parametrize("find, replace", [("Alpha two", "Alpha 2"), ("e\vA", "e A"), ("Alpha two", "Alpha\vtwo")])
def test_plain_replace_on_line_breaks(doc, find, replace):
    requested = "Alpha one\vAlpha two\vAlpha three".replace(find, replace)
    out = _plain(doc, PARA["breaks"], find=find, replace=replace)
    assert new_ids(doc, out) == set()
    assert accepted_text(element(out, PARA["breaks"])) == requested
    assert layout_text(out, PARA["breaks"]) == requested


def test_plain_fill_self_closing_paragraph(doc):
    out = _plain(doc, PARA["cell_empty"], text="Plain reply.")
    assert new_ids(doc, out) == set()
    assert layout_text(out, PARA["cell_empty"]) == "Plain reply."
    assert layout_text(out, PARA["closing"]) == "Closing."


# -- through edit_document (in place, one undo) -------------------------------------------------------------------

@windows_only
def test_edit_document_fills_empty_cell_and_rewrites_multiline_cell(doc):
    from research_mcp.publication import undo_document

    before = file_sha256(doc)
    result = run_batch(
        doc,
        {"op_id": "a", "type": "tracked_replace_paragraph", "paragraph_index": PARA["cell_empty"],
         "text": "Filled reply."},
        {"op_id": "b", "type": "tracked_replace_text", "paragraph_index": PARA["cell_breaks"],
         "find": "second line", "replace": "next line"})
    assert result.status == "complete"
    layout = read_layout(doc)
    assert layout.para(PARA["cell_empty"]).text == "Filled reply."
    assert layout.para(PARA["cell_breaks"]).text == "Reviewer point\vnext line"
    assert undo_document(str(doc))["restored"] is True and file_sha256(doc) == before


# -- review round 1: deleted breaks leave the accepted view (F1, F2) ------------------------------------------

def one_paragraph(doc: Path, name: str, para: str) -> Path:
    def swap(xml: str) -> str:
        start = xml.index("<w:body>") + len("<w:body>")
        return xml[:start] + para + SECT + "</w:body></w:document>"

    return variant(doc, doc.parent / f"{name}.docx", swap)


BOOKMARKED = {
    "break": (BR, "\v"),
    "tab": ("<w:tab/>", "\t"),
    "page": ('<w:br w:type="page"/>', "\x0c"),
}


@pytest.mark.parametrize("key", sorted(BOOKMARKED))
def test_tracked_deletion_may_not_empty_a_bookmark_over_a_break(doc, key):
    """F1: a bookmark holding only a break (or tab) and text, all of it deleted, would be empty after Accept."""
    mark, ch = BOOKMARKED[key]
    source = one_paragraph(doc, key, p(r(t("ab")), '<w:bookmarkStart w:id="51" w:name="B"/>', r(mark, t("cd")),
                                       '<w:bookmarkEnd w:id="51"/>', r(t("ef"))))
    with pytest.raises(Exception, match="bookmark 'B'.*all of its text is deleted"):
        tracked(source, 1, find=f"ab{ch}cd", replace="X")


def test_tracked_deletion_may_not_empty_a_bookmark_over_text(doc):
    """The same rule for plain text (control for F1)."""
    source = one_paragraph(doc, "text", p(r(t("ab")), '<w:bookmarkStart w:id="51" w:name="B"/>', r(t("cd")),
                                          '<w:bookmarkEnd w:id="51"/>', r(t("ef"))))
    with pytest.raises(Exception, match="bookmark 'B'.*all of its text is deleted"):
        tracked(source, 1, find="abcd", replace="X")


def test_tracked_deletion_may_not_empty_a_comment_over_a_break(doc):
    source = one_paragraph(doc, "comment", p(r(t("ab")), '<w:commentRangeStart w:id="0"/>', r(BR, t("cd")),
                                             '<w:commentRangeEnd w:id="0"/>',
                                             '<w:r><w:commentReference w:id="0"/></w:r>', r(t("ef"))))
    with pytest.raises(Exception, match="comment.*all of its text is deleted"):
        tracked(source, 1, find="ab\vcd", replace="X")


def test_bookmark_keeping_some_text_after_break_deletion(doc):
    """Deleting only the break inside a bookmark keeps the bookmark on the remaining text."""
    source = one_paragraph(doc, "keep", p(r(t("ab")), '<w:bookmarkStart w:id="51" w:name="B"/>', r(BR, t("cd")),
                                          '<w:bookmarkEnd w:id="51"/>', r(t("ef"))))
    out, _ = tracked(source, 1, find="ab\vcd", replace="ab cd")
    check_tracked(source, out, 1, "ab cdef")
    starts = read_layout(out)
    from research_mcp.sections import anchor_extents
    assert anchor_extents(starts)["bookmarks"]["B"] == (2, 5)  # " cd" in the accepted text "ab cdef"


@pytest.mark.parametrize("ch, mark", [("\v", BR), ("\t", "<w:tab/>"), ("\x0c", '<w:br w:type="page"/>')])
def test_inspected_text_after_tracked_break_deletion(doc, ch, mark):
    """F2: after a tracked deletion of a break, inspect and the layout show the accepted text, and a second
    edit written from that text applies."""
    from research_mcp.paragraphs import read_paragraphs

    source = one_paragraph(doc, "chain", p(r(t("Alpha one")), r(mark, t("Alpha two")), r(BR, t("Alpha three"))))
    out, _ = tracked(source, 1, find=f"e{ch}A", replace="eA")
    shown = "Alpha oneAlpha two\vAlpha three"
    assert accepted_text(element(out, 1)) == shown
    assert layout_text(out, 1) == shown
    assert read_paragraphs(out)[0].text == shown
    out2, result = tracked(out, 1, find="oneAlpha two", replace="one Beta two")
    assert result["text_before"] == shown
    assert accepted_text(element(out2, 1)) == "Alpha one Beta two\vAlpha three"
    # Rejecting only the second edit gives back the first edit's result; rejecting both gives the input.
    second = new_ids(out, out2)
    assert tokens(element(out2, 1), reject=second) == tokens(element(out, 1))
    assert tokens(element(out2, 1), reject=new_ids(source, out2)) == tokens(element(source, 1))


def test_paragraph_rewrite_after_tracked_break_deletion(doc):
    """F2 with a whole-paragraph rewrite from the re-inspected text."""
    source = one_paragraph(doc, "rewrite", p(r(t("Alpha one")), r(BR, t("Alpha two")), r(BR, t("Alpha three"))))
    out, _ = tracked(source, 1, find="e\vA", replace="e A")
    shown = layout_text(out, 1)
    assert shown == "Alpha one Alpha two\vAlpha three"
    requested = shown.replace("three", "3")
    out2, _ = tracked(out, 1, text=requested)
    assert accepted_text(element(out2, 1)) == requested
    assert tokens(element(out2, 1), reject=new_ids(source, out2)) == tokens(element(source, 1))


def test_comment_anchor_text_after_tracked_break_deletion(doc):
    """F2 for inspect's comment list: the anchored text is the accepted text, without the deleted break."""
    from research_mcp.comments import inspect_comments
    from research_mcp.paragraphs import read_snapshot

    source = one_paragraph(doc, "scope", p(r(t("ab")), '<w:commentRangeStart w:id="0"/>', r(t("cd")),
                                           r(BR, t("gh")), '<w:commentRangeEnd w:id="0"/>',
                                           '<w:r><w:commentReference w:id="0"/></w:r>', r(t("ij"))))
    assert [c.anchor_text for c in inspect_comments(read_snapshot(source))] == ["cd\vgh"]
    out, _ = tracked(source, 1, find="cd\vgh", replace="cd gh")
    assert [c.anchor_text for c in inspect_comments(read_snapshot(out))] == ["cd gh"]
