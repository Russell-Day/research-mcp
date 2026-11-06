"""Plan 501 S2: tracked changes on the direct-XML engine (revisions.py), synthetic fixture only.

Every expectation is checked with this file's own ElementTree reader (``view``), not with the engine's
resolver, so a test still fails if revisions.py's own verification were removed.
"""
from __future__ import annotations

import datetime as dt
import os
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.audit import file_sha256
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.fixtures import TRACKED_AUTHOR, create_tracked_fixture

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W16DU = "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}"
REV = re.compile(r"<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|tcPrChange)\b[^>]*?(?:/>|>.*?</w:\1>)", re.S)
FIRST_NEW_ID = 31  # fixture's highest w:id is wp:docPr 30


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    return create_tracked_fixture(tmp_path / "tracked.docx")


def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def paragraphs(path: Path) -> list[ET.Element]:
    """w:p elements in document order; ordinal n == layout index n for paragraphs 1-12 of the fixture."""
    return list(ET.fromstring(document_xml(path)).iter(W + "p"))


def view(paragraph: ET.Element, reject=()) -> str:
    """Paragraph text with every revision accepted except the ids in ``reject``, which are rejected."""
    reject = {str(item) for item in reject}
    out: list[str] = []

    def walk(element, ins_id, del_id):
        for child in element:
            if child.tag in (W + "pPr", W + "rPr", W + "instrText"):
                continue
            if child.tag in (W + "ins", W + "moveTo"):
                walk(child, child.get(W + "id"), del_id)
            elif child.tag in (W + "del", W + "moveFrom"):
                walk(child, ins_id, child.get(W + "id"))
            elif child.tag in (W + "t", W + "delText"):
                if (ins_id is None or ins_id not in reject) and (del_id is None or del_id in reject):
                    out.append(child.text or "")
            else:
                walk(child, ins_id, del_id)

    walk(paragraph, None, None)
    return "".join(out)


def revision_elements(path: Path) -> list[ET.Element]:
    kinds = {W + k for k in ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "tcPrChange")}
    return [el for el in ET.fromstring(document_xml(path)).iter() if el.tag in kinds]


def new_revisions(source: Path, output: Path) -> list[ET.Element]:
    before = {el.get(W + "id") for el in revision_elements(source)}
    return [el for el in revision_elements(output) if el.get(W + "id") not in before]


def raw_revisions(path: Path) -> list[str]:
    return [m.group(0) for m in REV.finditer(document_xml(path))]


def paragraph_spans(xml: str) -> list[str]:
    return re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)


def other_parts_identical(source: Path, output: Path) -> bool:
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
        if a.namelist() != b.namelist():
            return False
        return all(a.read(n) == b.read(n) for n in a.namelist() if n != "word/document.xml")


def tracked(source, output, paragraph_index, find=None, replace=None, text=None, **kw):
    return revisions.text_edit(source, output, op="tracked_replace_text" if find else "tracked_replace_paragraph",
                               paragraph_index=paragraph_index, tracked=True, find=find, replace=replace,
                               text=text, **kw)


# -- tracked_replace_text / tracked_replace_paragraph (D3, D4, D5) ----------------------------------------

def test_tracked_replace_text_next_to_other_authors_revisions(doc, tmp_path):
    out = tmp_path / "out.docx"
    result = tracked(doc, out, 2, find="accurate", replace="precise")

    added = new_revisions(doc, out)
    assert [(el.tag.replace(W, ""), el.get(W + "author")) for el in added] == [("del", "Claude"), ("ins", "Claude")]
    ids = [int(el.get(W + "id")) for el in added]
    assert ids == sorted(ids) and ids[0] >= FIRST_NEW_ID and len(set(ids)) == len(ids)
    p = paragraphs(out)[1]
    assert view(p) == "The model was highly precise on the test set."  # D5 accept-new == requested
    assert view(p, reject=ids) == view(paragraphs(doc)[1])  # D5 reject-new == input
    assert view(p, reject=ids + [1, 2]) == "The model was accurate overall on the test set."
    # D4: all existing revision elements byte-identical; other paragraphs and parts untouched
    assert all(raw in document_xml(out) for raw in raw_revisions(doc))
    before, after = paragraph_spans(document_xml(doc)), paragraph_spans(document_xml(out))
    assert before[:1] + before[2:] == after[:1] + after[2:]
    assert other_parts_identical(doc, out)
    assert result["revisions_added"] == [{"id": str(ids[0]), "type": "del", "author": "Claude", "text": "accurate"},
                                         {"id": str(ids[1]), "type": "ins", "author": "Claude", "text": "precise"}]


def test_tracked_dates_follow_word_convention(doc, tmp_path):
    out = tmp_path / "out.docx"
    now = dt.datetime(2026, 10, 30, 12, 0, 0, tzinfo=dt.timezone.utc)
    tracked(doc, out, 1, find="opening", replace="opened", now=now)
    for el in new_revisions(doc, out):
        assert el.get(W + "date") == now.astimezone().strftime("%Y-%m-%dT%H:%M:%SZ")  # local wall time + Z
        assert el.get(W16DU + "dateUtc") == "2026-10-30T12:00:00Z"  # true UTC


def test_revision_stamp_omits_date_utc_without_w16du_namespace():
    now = dt.datetime(2026, 10, 30, 12, 0, 0, tzinfo=dt.timezone.utc)
    plain = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>'
    assert "dateUtc" not in revisions.revision_stamp(plain, now)
    assert revisions.revision_stamp(plain, now) == f'w:date="{now.astimezone().strftime("%Y-%m-%dT%H:%M:%SZ")}"'


def test_custom_author_and_ascending_ids_for_several_changes(doc, tmp_path):
    out = tmp_path / "out.docx"
    tracked(doc, out, 2, text="The model was highly precise on the full test set.", author="Reviewer Two")

    added = new_revisions(doc, out)
    assert {el.get(W + "author") for el in added} == {"Reviewer Two"}
    assert [el.tag.replace(W, "") for el in added] == ["del", "ins", "ins"]
    ids = [int(el.get(W + "id")) for el in added]
    assert ids == list(range(FIRST_NEW_ID, FIRST_NEW_ID + 3))  # unique, assigned left to right
    assert view(paragraphs(out)[1]) == "The model was highly precise on the full test set."
    assert view(paragraphs(out)[1], reject=ids) == "The model was highly accurate on the test set."


def test_tracked_replace_paragraph_widens_a_partial_word_and_keeps_bookmark(doc, tmp_path):
    out = tmp_path / "out.docx"
    tracked(doc, out, 1, text="Plain opener paragraph with no revisions.")

    added = new_revisions(doc, out)
    assert [(el.tag.replace(W, ""), "".join(el.itertext())) for el in added] == [("del", "opening"), ("ins", "opener")]
    p = paragraphs(out)[0]
    assert view(p) == "Plain opener paragraph with no revisions."
    assert p.find(W + "bookmarkStart") is not None and p.find(W + "bookmarkEnd") is not None


def test_tracked_edit_beside_a_field_keeps_the_field(doc, tmp_path):
    out = tmp_path / "out.docx"
    tracked(doc, out, 3, find=". The same", replace="; the same")

    p_in, p_out = paragraphs(doc)[2], paragraphs(out)[2]
    assert view(p_out) == "Prior work (1); the same method applies."
    assert view(p_out, reject=[el.get(W + "id") for el in new_revisions(doc, out)]) == view(p_in)
    assert [el.text for el in p_out.iter(W + "instrText")] == [el.text for el in p_in.iter(W + "instrText")]
    assert len(list(p_out.iter(W + "fldChar"))) == 3


def test_tracked_edit_in_comment_range_keeps_the_anchor(doc, tmp_path):
    out = tmp_path / "out.docx"
    tracked(doc, out, 4, find="Commented phrase", replace="Commented words")

    p = paragraphs(out)[3]
    assert view(p) == "Commented words stays anchored."
    for tag in ("commentRangeStart", "commentRangeEnd"):
        assert p.find(W + tag) is not None
    assert p.find(f".//{W}commentReference") is not None


@pytest.mark.parametrize("index, find, replace, expected", [
    (6, "Stay", "Remain", "Remain text."),            # next to another author's moveFrom
    (9, "Inserted", "Added", "Added paragraph mark here."),
    (10, "Centered", "Aligned", "Aligned by the other author."),
    (11, "A", "B", "Cell added B"),                    # table cell, beside an insertion
])
def test_tracked_edit_shapes_that_work(doc, tmp_path, index, find, replace, expected):
    out = tmp_path / "out.docx"
    tracked(doc, out, index, find=find, replace=replace)
    ids = [el.get(W + "id") for el in new_revisions(doc, out)]
    p_in, p_out = paragraphs(doc)[index - 1], paragraphs(out)[index - 1]
    assert view(p_out) == view(p_in).replace(find, replace) and expected in view(p_out)
    assert view(p_out, reject=ids) == view(p_in)
    assert all(raw in document_xml(out) for raw in raw_revisions(doc))


def test_whole_formatted_run_with_rpr_change_can_be_deleted(doc, tmp_path):
    out = tmp_path / "out.docx"
    tracked(doc, out, 5, text="Then plain.")
    ids = [el.get(W + "id") for el in new_revisions(doc, out)]
    assert view(paragraphs(out)[4]) == "Then plain."
    assert view(paragraphs(out)[4], reject=ids) == "Bold words then plain."


@pytest.mark.parametrize("index, find, replace, message", [
    (2, "highly", "very", "inside w:ins id=1 by 'Other Author'"),
    (3, "(1)", "(2)", "would delete field text"),
    (5, "Bold", "Strong", "w:rPrChange id=3 by 'Other Author'"),
    (7, "Moved", "Shifted", "inside w:moveTo id=7 by 'Other Author'"),
    (8, "caption", "legend", "holds a drawing (w:drawing)"),
])
def test_unsafe_tracked_edits_are_refused_and_write_nothing(doc, tmp_path, index, find, replace, message):
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=re.escape(message)):
        tracked(doc, out, index, find=find, replace=replace)
    assert not out.exists()


def test_insertion_strictly_inside_another_authors_insertion_is_refused(doc, tmp_path):
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="w:ins id=1"):
        tracked(doc, out, 2, text="The model was highXly accurate on the test set.")
    assert not out.exists()


def test_empty_author_is_refused(doc, tmp_path):
    with pytest.raises(RefusedUnsafe, match="author must not be empty"):
        tracked(doc, tmp_path / "out.docx", 1, find="opening", replace="opened", author="  ")


def test_wrong_insertion_text_is_rolled_back(doc, tmp_path, monkeypatch):
    """D5 in code: if accepting the new revisions would not give the requested text, nothing is saved."""
    original = revisions._Writer.insertion

    def corrupt(self, item):
        item.text = item.text + "X"
        return original(self, item)

    monkeypatch.setattr(revisions._Writer, "insertion", corrupt)
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed):
        tracked(doc, out, 1, find="opening", replace="opened")
    assert not out.exists()


def test_untracked_deletion_is_rolled_back(doc, tmp_path, monkeypatch):
    """D5 in code: if rejecting the new revisions would not give back the input, nothing is saved."""
    monkeypatch.setattr(revisions._Writer, "deletion", lambda self, run_xml, text: "")
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed):
        tracked(doc, out, 1, find="opening", replace="opened")
    assert not out.exists()


# -- plain edits on a document with revisions (D4) -------------------------------------------------------

def test_plain_replace_on_revised_document_adds_no_revisions(doc, tmp_path):
    from research_mcp.editing import replace_text_to_version

    out = tmp_path / "out.docx"
    replace_text_to_version(doc, out, paragraph_index=1, find="opening", replace="opened")
    assert raw_revisions(out) == raw_revisions(doc)
    assert view(paragraphs(out)[0]) == "Plain opened paragraph with no revisions."


def test_plain_replace_touching_a_revision_is_refused_naming_it(doc, tmp_path):
    from research_mcp.editing import replace_text_to_version

    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=r"w:ins id=1 by 'Other Author'"):
        replace_text_to_version(doc, out, paragraph_index=2, find="highly accurate", replace="precise")
    assert not out.exists()


# -- accept_revisions / reject_revisions -----------------------------------------------------------------

def resolve(doc, tmp_path, mode, **selectors):
    out = tmp_path / f"{mode}-{os.urandom(3).hex()}.docx"
    return out, revisions.resolve_revisions(doc, out, mode=mode, **selectors)


def unselected_identical(source: Path, output: Path, resolved_ids) -> bool:
    keep = [raw for raw in raw_revisions(source) if re.search(r'w:id="(\d+)"', raw).group(1) not in resolved_ids]
    return all(raw in document_xml(output) for raw in keep)


def test_accept_by_paragraph(doc, tmp_path):
    out, result = resolve(doc, tmp_path, "accept", paragraph_index=2)
    p = paragraphs(out)[1]
    assert result["resolved_count"] == 2
    assert p.find(f".//{W}ins") is None and p.find(f".//{W}del") is None
    assert view(p) == "The model was highly accurate on the test set."
    assert unselected_identical(doc, out, {"1", "2"})


@pytest.mark.parametrize("rev_id, expected", [
    (2, "The model was highly accurate overall on the test set."),
    (1, "The model was accurate on the test set."),
])
def test_reject_one_run_revision_by_id(doc, tmp_path, rev_id, expected):
    out, result = resolve(doc, tmp_path, "reject", ids=[rev_id])
    assert result["resolved_count"] == 1
    assert view(paragraphs(out)[1]) == expected
    assert unselected_identical(doc, out, {str(rev_id)})
    assert f'w:id="{rev_id}"' not in document_xml(out)


def test_formatting_changes_accept_and_reject(doc, tmp_path):
    accepted, _ = resolve(doc, tmp_path, "accept", ids=[3])
    rejected, _ = resolve(doc, tmp_path, "reject", ids=[3])
    run_a = paragraphs(accepted)[4].find(W + "r")
    run_r = paragraphs(rejected)[4].find(W + "r")
    assert run_a.find(f"{W}rPr/{W}b") is not None and run_a.find(f".//{W}rPrChange") is None
    assert run_r.find(f"{W}rPr/{W}b") is None and run_r.find(f".//{W}rPrChange") is None

    rejected_p, _ = resolve(doc, tmp_path, "reject", ids=[9])
    ppr = paragraphs(rejected_p)[9].find(W + "pPr")
    assert ppr.find(W + "jc") is None and ppr.find(W + "pPrChange") is None
    assert ppr.find(f"{W}rPr/{W}del") is not None  # the paragraph-mark deletion id 12 is kept

    accepted_cell, _ = resolve(doc, tmp_path, "accept", ids=[11])
    assert "tcPrChange" not in document_xml(accepted_cell) and 'w:w="4500"' in document_xml(accepted_cell)
    assert unselected_identical(doc, accepted_cell, {"11"})
    # fix round 1 of plan 562 S2: rejecting a cell-property change restores the old tcPr (as Word does)
    rejected_cell, _ = resolve(doc, tmp_path, "reject", ids=[11])
    assert "tcPrChange" not in document_xml(rejected_cell) and 'w:w="4000"' in document_xml(rejected_cell)
    assert unselected_identical(doc, rejected_cell, {"11"})


def test_accept_inserted_mark_and_reject_deleted_mark(doc, tmp_path):
    accepted, _ = resolve(doc, tmp_path, "accept", ids=[8])
    rejected, _ = resolve(doc, tmp_path, "reject", ids=[12])
    assert len(paragraphs(accepted)) == len(paragraphs(rejected)) == len(paragraphs(doc))
    assert paragraphs(accepted)[8].find(f"{W}pPr/{W}rPr/{W}ins") is None
    assert paragraphs(rejected)[9].find(f"{W}pPr/{W}rPr/{W}del") is None


@pytest.mark.parametrize("mode, selectors, message", [
    # plan 562 S2: moves and paragraph merges resolve now; a merge into the following table is still refused
    ("accept", {"author": TRACKED_AUTHOR}, "w:del id=12 (paragraph mark) by 'Other Author' (2026-10-10T17:37:00Z) "
                                          "would merge paragraph 10 with a table"),
    ("accept", {"ids": [12]}, "would merge paragraph 10 with a table"),
    ("accept", {"ids": [1], "paragraph_index": 3}, "no revision with id 1 matches the other selectors"),
    ("reject", {"paragraph_index": 1}, "no revision matches the selection; nothing to reject"),
    ("accept", {}, "give ids, author, a paragraph range or a section heading"),
    ("accept", {"ids": [999]}, "no revision with id 999"),
])
def test_unsafe_resolutions_are_refused_and_write_nothing(doc, tmp_path, mode, selectors, message):
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=re.escape(message)):
        revisions.resolve_revisions(doc, out, mode=mode, **selectors)
    assert not out.exists()


def test_selectors_combine_with_and(doc, tmp_path):
    out, result = resolve(doc, tmp_path, "accept", author=TRACKED_AUTHOR, paragraph_index=2,
                          through_paragraph_index=5)
    assert sorted(r["id"] for r in result["resolved"]) == ["1", "2", "3"]
    assert unselected_identical(doc, out, {"1", "2", "3"})


def test_resolve_then_new_tracked_edit_round_trip(doc, tmp_path):
    """New revisions can themselves be listed and rejected, giving back the input bytes' text."""
    edited = tmp_path / "edited.docx"
    tracked(doc, edited, 1, find="opening", replace="opened")
    claude = revisions.list_revisions(edited, author="Claude")
    assert claude["matched"] == 2
    back = tmp_path / "back.docx"
    revisions.resolve_revisions(edited, back, mode="reject", author="Claude")
    assert paragraph_spans(document_xml(back))[1:] == paragraph_spans(document_xml(doc))[1:]
    assert view(paragraphs(back)[0]) == view(paragraphs(doc)[0])


# -- list_revisions ---------------------------------------------------------------------------------------

def test_list_revisions_matches_independent_parse(doc):
    listing = revisions.list_revisions(doc)
    parsed = [(el.get(W + "id"), el.tag.replace(W, "")) for el in revision_elements(doc)]
    assert [(r["id"], r["type"]) for r in listing["revisions"]] == parsed
    assert listing["total_in_document"] == len(parsed) == 10
    assert listing["authors"] == {TRACKED_AUTHOR: 10}


def test_list_revisions_by_paragraph_shows_both_texts(doc):
    listing = revisions.list_revisions(doc, paragraph_index=2)
    assert [(r["id"], r["type"], r["text"]) for r in listing["revisions"]] == [("1", "ins", "highly "),
                                                                              ("2", "del", " overall")]
    assert listing["scope"] == [2, 2]
    assert listing["paragraphs"] == [{"paragraph_index": 2,
                                      "accepted_text": "The model was highly accurate on the test set.",
                                      "original_text": "The model was accurate overall on the test set."}]


def test_list_revisions_limit_truncates(doc):
    listing = revisions.list_revisions(doc, limit=3)
    assert listing["matched"] == 10 and listing["truncated"] is True and len(listing["revisions"]) == 3


def test_list_revisions_operation_is_read_only(doc, tmp_path):
    from research_mcp.edit_models import ListRevisionsOperation

    before = file_sha256(doc)
    out = tmp_path / "out.docx"
    result = revisions.apply_operation(doc, out, ListRevisionsOperation(op_id="l", type="list_revisions", author=TRACKED_AUTHOR))
    assert result["read_only"] is True and result["matched"] == 10
    assert file_sha256(doc) == before and file_sha256(out) == before


# -- server envelope defaults ------------------------------------------------------------------------------

def test_server_payload_defaults_for_revision_ops():
    from research_mcp.server import MCPEditOperation, _strict_operation_payload

    edit = _strict_operation_payload(MCPEditOperation(op_id="e", type="tracked_replace_text", paragraph_index=2,
                                                      find="a", replace="b"))
    assert edit.author == "Claude"  # not the envelope's comment author default
    named = _strict_operation_payload(MCPEditOperation(op_id="e", type="tracked_replace_text", paragraph_index=2,
                                                       find="a", replace="b", author="Reviewer Two"))
    assert named.author == "Reviewer Two"
    accept = _strict_operation_payload(MCPEditOperation(op_id="a", type="accept_revisions", ids=[5]))
    assert (accept.ids, accept.author, accept.paragraph_index, accept.heading) == ([5], None, None, None)
    ranged = _strict_operation_payload(MCPEditOperation(op_id="r", type="reject_revisions", paragraph_index=2,
                                                        through_paragraph_index=4))
    assert (ranged.paragraph_index, ranged.through_paragraph_index) == (2, 4)
    with pytest.raises(ValueError):
        _strict_operation_payload(MCPEditOperation(op_id="a", type="accept_revisions"))


# -- edit_document: checkpoints, one-call undo, atomic batches ---------------------------------------------

windows_only = pytest.mark.skipif(os.name != "nt", reason="undo store writes need Windows Python")


def run_batch(path: Path, *ops: dict, allow_partial: bool = False):
    from research_mcp.edit_batch import edit_document
    from research_mcp.edit_models import EditDocumentRequest

    return edit_document(EditDocumentRequest(source=str(path), source_sha256=file_sha256(path),
                                             output=str(path), operations=list(ops), allow_partial=allow_partial))


@windows_only
def test_edit_document_tracked_edit_in_place_then_one_call_undo(doc):
    from research_mcp.publication import undo_document

    before = file_sha256(doc)
    result = run_batch(doc, {"op_id": "e1", "type": "tracked_replace_text", "paragraph_index": 2,
                             "find": "accurate", "replace": "precise"},
                       {"op_id": "a1", "type": "reject_revisions", "ids": [2]})
    assert result.status == "complete" and result.undo_token
    assert view(paragraphs(doc)[1]) == "The model was highly precise overall on the test set."
    undone = undo_document(str(doc))
    assert undone["restored"] is True and file_sha256(doc) == before


@windows_only
def test_edit_document_list_revisions_writes_nothing(doc):
    from research_mcp.publication import document_history

    before = file_sha256(doc)
    result = run_batch(doc, {"op_id": "l", "type": "list_revisions", "paragraph_index": 2})
    assert result.status == "complete" and result.output is None and result.undo_token is None
    assert result.operations[0].details["matched"] == 2
    assert file_sha256(doc) == before
    assert not document_history(str(doc)).get("entries")


@windows_only
def test_edit_document_refused_op_publishes_nothing(doc):
    before = file_sha256(doc)
    result = run_batch(doc, {"op_id": "ok", "type": "tracked_replace_text", "paragraph_index": 1,
                             "find": "opening", "replace": "opened"},
                       {"op_id": "bad", "type": "tracked_replace_text", "paragraph_index": 2,
                        "find": "highly", "replace": "very"})
    assert result.status != "complete" and result.undo_token is None
    assert "w:ins id=1" in (result.operations[1].error.message if result.operations[1].error else "")
    assert file_sha256(doc) == before


def test_list_revisions_standalone_tool(doc):
    """Plan 501 S3: Haiku kept calling list_revisions as its own tool; it is one now (read only)."""
    import asyncio

    from research_mcp import server

    before = file_sha256(doc)
    previous_root = server.ROOT_DIR
    server.set_root_dir(doc.parent)
    try:
        result = asyncio.run(server.server.call_tool("list_revisions", {"path": str(doc), "paragraph_index": 2}))
    finally:
        server.set_root_dir(previous_root)
    payload = result.structured_content
    assert [(r["id"], r["type"], r["text"]) for r in payload["revisions"]] == [("1", "ins", "highly "),
                                                                              ("2", "del", " overall")]
    assert file_sha256(doc) == before
