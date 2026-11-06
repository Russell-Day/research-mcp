"""Plan 501 S2 review r1 regressions (F1-F4 and the resolution guard), synthetic fixture only.

Each test failed on the round-1 engine: F1 published an accept/reject whose selected revision shares a table
row with a failed edit; F2 let property-only revisions (*PrChange) reach Word; F3 trusted the planner's own
expected text; F4 reused a w:id written with single quotes or another prefix; the resolution guard had no
test that needed resolve_revisions' independent ElementTree comparison.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.audit import audit_docx, file_sha256
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.fixtures import create_tracked_fixture
from test_revisions import W, document_xml, paragraph_spans, paragraphs, raw_revisions, run_batch, view, windows_only

W_NS = W[1:-1]


def variant(source: Path, target: Path, transform=lambda xml: xml, extra: dict[str, str | bytes] | None = None) -> Path:
    """Copy `source` with word/document.xml passed through `transform` and `extra` parts added."""
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, "x") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                data = transform(data.decode("utf-8")).encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)
        for name, text in (extra or {}).items():
            data = text if isinstance(text, bytes) else text.encode("utf-8")
            out.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
    return target


def drop_paragraphs(xml: str, *keys: str) -> str:
    for span in paragraph_spans(xml):
        if any(key in span for key in keys):
            xml = xml.replace(span, "", 1)
    return xml


def placement(path: Path) -> dict[str, int | None]:
    _, doc = revisions.load(path)
    return {revision.id: revision.paragraph_index for revision in doc.revisions}


# -- F1: an accept/reject joins the row of every revision it selects ------------------------------------

HEADING = "<w:p><w:r><w:rPr><w:caps/></w:rPr><w:t>RESULTS</w:t></w:r></w:p>"


@pytest.fixture
def grouped(tmp_path: Path) -> Path:
    """Moves and the deleted-mark paragraph removed (accept by author would refuse them); a caps RESULTS
    heading before the table. Revision 10 sits in cell A (paragraph 9), 11 on cell B (paragraph 10)."""
    base = create_tracked_fixture(tmp_path / "base.docx")
    path = variant(base, tmp_path / "grouped.docx", lambda xml: drop_paragraphs(
        xml, "moveFromRangeStart", "moveToRangeStart", "Centered by the other author").replace(
        "<w:tbl>", HEADING + "<w:tbl>", 1))
    found = placement(path)
    assert found["10"] == 9 and found["11"] == 10 and found["1"] == 2
    assert view(paragraphs(path)[7]) == "RESULTS"
    return path


SELECTORS = {
    "ids": {"ids": [10]},
    "author": {"author": "Other Author"},
    "range": {"paragraph_index": 1, "through_paragraph_index": 9},
    "heading": {"heading": "RESULTS"},
}


@windows_only
@pytest.mark.parametrize("selector", list(SELECTORS))
def test_accept_selecting_a_row_rolls_back_with_a_failed_edit_in_that_row(grouped, selector):
    before = document_xml(grouped)
    result = run_batch(
        grouped,
        {"op_id": "a", "type": "accept_revisions", **SELECTORS[selector]},
        {"op_id": "b", "type": "tracked_replace_text", "paragraph_index": 10, "find": "absent", "replace": "X"},
        {"op_id": "c", "type": "tracked_replace_text", "paragraph_index": 1, "find": "opening", "replace": "opened"},
        allow_partial=True,
    )
    status = {item.op_id: item.status for item in result.operations}
    assert status == {"a": "rolled_back", "b": "failed", "c": "applied"}, result.message
    assert result.status == "partial"
    after = document_xml(grouped)
    assert all(raw in after for raw in raw_revisions(grouped))  # nothing of op a published: ins 10 is still there
    assert re.search(r'<w:ins w:id="10"', after) and re.search(r'<w:ins w:id="10"', before)
    assert view(paragraphs(grouped)[0]) == "Plain opened paragraph with no revisions."  # unrelated op published


# -- F2: property-only revisions are revisions -----------------------------------------------------------

@pytest.fixture
def property_only(tmp_path: Path) -> Path:
    """Only rPrChange 3, pPrChange 9 and tcPrChange 11 remain: the audit's ins/del/move counter sees none."""
    base = create_tracked_fixture(tmp_path / "base.docx")

    def strip(xml: str) -> str:
        xml = drop_paragraphs(xml, "The model was ", "moveFromRangeStart", "moveToRangeStart",
                              "Inserted paragraph mark here.")
        xml = re.sub(r"<w:rPr><w:del [^>]*/></w:rPr>", "", xml)
        return re.sub(r'<w:ins w:id="10"[^>]*>(.*?)</w:ins>', r"\1", xml)

    path = variant(base, tmp_path / "property_only.docx", strip)
    assert audit_docx(path).revision_count == 0
    assert revisions.revision_count(path) == 3
    return path


@pytest.fixture
def in_process(monkeypatch):
    """edit_document runs each op in a worker process; run it here so this test's monkeypatches apply."""
    from research_mcp import edit_batch

    monkeypatch.setattr(edit_batch, "_apply_operation", edit_batch._apply_operation_local)


@pytest.fixture
def no_word(monkeypatch):
    from research_mcp import editing, operation_process, server

    def started(*args, **kwargs):
        raise AssertionError("Word was started on a document with tracked revisions")

    for module, name in ((editing, "IsolatedWordDocument"), (editing, "create_versioned_copy"),
                         (operation_process, "run_operation"), (server, "apply_journal_style")):
        monkeypatch.setattr(module, name, started)


def test_revision_count_sees_every_kind_in_every_part(tmp_path):
    base = create_tracked_fixture(tmp_path / "base.docx")
    header = (f"<x:hdr xmlns:x='{W_NS}'><x:p><x:pPr><x:pPrChange x:id='40' x:author='A'><x:pPr/></x:pPrChange>"
              "</x:pPr></x:p><x:tbl><x:tblPr><x:tblPrChange x:id='41' x:author='A'><x:tblPr/></x:tblPrChange>"
              "</x:tblPr></x:tbl><x:sectPr><x:sectPrChange x:id='42' x:author='A'/></x:sectPr></x:hdr>")
    numbering = (f'<w:numbering xmlns:w="{W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pPr>'
                 '<w:numberingChange w:id="43" w:author="A"/></w:pPr></w:lvl></w:abstractNum></w:numbering>')
    path = variant(base, tmp_path / "parts.docx",
                   extra={"word/header1.xml": header, "word/numbering.xml": numbering})
    assert revisions.revision_count(base) == 10
    assert revisions.revision_count(path) == 14


def test_property_only_replace_goes_direct_xml(property_only, tmp_path, no_word):
    from research_mcp.editing import replace_paragraph_to_version, replace_text_to_version

    out = tmp_path / "out.docx"
    result = replace_text_to_version(property_only, out, paragraph_index=1, find="opening", replace="opened")
    assert result["engine"] == "direct_xml"
    assert raw_revisions(out) == raw_revisions(property_only)  # every property change byte-identical
    assert view(paragraphs(out)[0]) == "Plain opened paragraph with no revisions."
    out2 = tmp_path / "out2.docx"
    result = replace_paragraph_to_version(property_only, out2, paragraph_index=1,
                                          text="Plain opening paragraph with one revision.")
    assert result["engine"] == "direct_xml" and raw_revisions(out2) == raw_revisions(property_only)


@pytest.mark.parametrize("call", [
    lambda e, s, o: e.insert_text_to_version(s, o, paragraph_index=1, position="after", anchor="opening",
                                             text=" new"),
    lambda e, s, o: e.insert_paragraph_to_version(s, o, after_paragraph_index=1, text="New paragraph."),
    lambda e, s, o: e.insert_table_to_version(s, o, after_paragraph_index=1, headers=("A", "B"),
                                              rows=(("1", "2"),), descriptive_columns=(1,)),
    lambda e, s, o: e.replace_table_row_to_version(s, o, match="Cell B", cells=["Cell added A", "Cell C"]),
], ids=["insert_text", "insert_paragraph", "insert_table", "replace_table_row"])
def test_word_ops_refuse_property_only_revisions_before_word(property_only, tmp_path, no_word, call):
    from research_mcp import editing

    before = file_sha256(property_only)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="3 tracked revision elements"):
        call(editing, property_only, out)
    assert not out.exists() and file_sha256(property_only) == before


def test_add_comment_keeps_property_only_revisions(property_only, tmp_path, no_word):
    """Plan 562 S1: comments are direct XML now, so add_comment works on revised documents and keeps
    every revision byte-identical (it used to refuse here)."""
    from research_mcp import editing

    out = tmp_path / "out.docx"
    editing.add_comment_to_version(property_only, out, paragraph_index=1, anchor="opening", text="Note",
                                   author="Claude", initials="C")
    assert raw_revisions(out) == raw_revisions(property_only)


def test_style_tools_refuse_property_only_revisions_before_word(property_only, tmp_path, no_word):
    from research_mcp import server

    sha = file_sha256(property_only)
    with pytest.raises(RefusedUnsafe, match="style_text: source contains 3 tracked revision elements"):
        server.style_text(str(property_only), sha, 1, "opening", bold=True, preview_only=False)
    out = tmp_path / "styled.docx"
    with pytest.raises(RefusedUnsafe, match="apply_document_style: source contains 3 tracked revision"):
        server.apply_document_style(str(property_only), str(out), sha)
    assert not out.exists() and file_sha256(property_only) == sha


@pytest.mark.parametrize("tool, call", [
    ("clean_bookmarks", lambda server, src, out: server.clean_bookmarks(src, out, ["h"])),
    ("format_tables", lambda server, src, out: server.format_tables(src, out, [])),
    ("insert_zotero_citation", lambda server, src, out: server.insert_zotero_citation(
        src, out, 1, "opening", ["ABCD1234"], src)),
])
def test_other_word_tools_refuse_revisions_before_word(property_only, tmp_path, no_word, tool, call):
    """review s2r2 F5: these tools save through Word, which rewrites other authors' revision records."""
    from research_mcp import server

    sha = file_sha256(property_only)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=f"{tool}: source contains 3 tracked revision elements"):
        call(server, str(property_only), str(out))
    assert not out.exists() and file_sha256(property_only) == sha


@windows_only
def test_edit_document_insert_on_property_only_revisions_publishes_nothing(property_only, in_process, no_word):
    before = file_sha256(property_only)
    result = run_batch(property_only, {"op_id": "i", "type": "insert_text", "paragraph_index": 1,
                                       "position": "after", "anchor": "opening", "text": " new"})
    assert result.status in ("refused", "no_changes") and result.undo_token is None and result.output is None
    assert result.operations[0].status == "failed"
    assert "3 tracked revision elements" in result.operations[0].error.message
    assert file_sha256(property_only) == before


# -- F3: D5 is checked against the caller's request, not the planner's expectation --------------------------

@pytest.fixture
def faulty_planner(monkeypatch):
    import dataclasses

    original = revisions.TextEditor.plan

    def plan(self, new_text):
        edits = original(self, new_text)
        return edits[:-1] + [dataclasses.replace(edits[-1], text=edits[-1].text + "X")]

    monkeypatch.setattr(revisions.TextEditor, "plan", plan)


@pytest.mark.parametrize("tracked", [True, False])
def test_faulty_planner_is_caught_against_the_request(tmp_path, faulty_planner, tracked):
    doc = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="requested text"):
        revisions.text_edit(doc, out, op="tracked_replace_text", paragraph_index=1, tracked=tracked,
                            find="opening", replace="opened")
    assert not out.exists()


@windows_only
def test_edit_document_faulty_planner_publishes_nothing(tmp_path, in_process, faulty_planner):
    doc = create_tracked_fixture(tmp_path / "tracked.docx")
    before = file_sha256(doc)
    result = run_batch(doc, {"op_id": "e", "type": "tracked_replace_text", "paragraph_index": 1,
                             "find": "opening", "replace": "opened"})
    assert result.status in ("refused", "no_changes") and result.undo_token is None and result.output is None
    assert result.operations[0].status == "failed"
    assert file_sha256(doc) == before
    assert "openedX" not in document_xml(doc)


# -- F4: new ids avoid every id in the package, whatever its quoting or prefix -------------------------------

def test_new_ids_avoid_single_quoted_and_other_prefix_ids(tmp_path):
    base = create_tracked_fixture(tmp_path / "base.docx")
    header = (f"<w:hdr xmlns:w='{W_NS}'><w:p><w:ins w:id='31' w:author='A' w:date='2026-01-01T00:00:00Z'>"
              "<w:r><w:t>Header</w:t></w:r></w:ins></w:p></w:hdr>")
    footer = (f'<x:ftr xmlns:x="{W_NS}"><x:p><x:bookmarkStart x:id="32" x:name="f"/><x:bookmarkEnd x:id="32"/>'
              "</x:p></x:ftr>")
    doc = variant(base, tmp_path / "ids.docx", extra={"word/header1.xml": header, "word/footer1.xml": footer})
    out = tmp_path / "out.docx"
    result = revisions.text_edit(doc, out, op="tracked_replace_text", paragraph_index=1, tracked=True,
                                 find="opening", replace="opened")
    new = [int(item["id"]) for item in result["revisions_added"]]
    assert new and min(new) >= 33 and not {31, 32} & set(new)


def utf16_header(body: str) -> bytes:
    return ('<?xml version="1.0" encoding="UTF-16"?>' + f'<w:hdr xmlns:w="{W_NS}"><w:p>{body}</w:p></w:hdr>').encode("utf-16")


def test_new_ids_avoid_ids_in_a_utf16_part(tmp_path):
    """review s2r2 F4: a UTF-16 part never contains the UTF-8 namespace bytes, so a byte prefilter skipped it."""
    base = create_tracked_fixture(tmp_path / "base.docx")
    header = utf16_header('<w:ins w:id="40" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>H</w:t></w:r></w:ins>')
    doc = variant(base, tmp_path / "ids.docx", extra={"word/header1.xml": header})
    assert 40 in revisions._package_ids(revisions.read_parts(doc))["word/header1.xml"]
    result = revisions.text_edit(doc, tmp_path / "out.docx", op="tracked_replace_text", paragraph_index=1,
                                 tracked=True, find="opening", replace="opened")
    new = [int(item["id"]) for item in result["revisions_added"]]
    assert new and min(new) >= 41


def test_post_write_check_catches_a_collision_with_a_utf16_part(tmp_path, monkeypatch):
    base = create_tracked_fixture(tmp_path / "base.docx")
    doc = variant(base, tmp_path / "ids.docx",
                  extra={"word/header1.xml": utf16_header('<w:bookmarkStart w:id="31" w:name="h"/>')})
    monkeypatch.setattr(revisions, "next_revision_id", lambda parts: 31)
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="collide"):
        revisions.text_edit(doc, out, op="tracked_replace_text", paragraph_index=1, tracked=True,
                            find="opening", replace="opened")
    assert not out.exists()


def test_post_write_check_catches_an_id_collision(tmp_path, monkeypatch):
    base = create_tracked_fixture(tmp_path / "base.docx")
    header = f"<w:hdr xmlns:w='{W_NS}'><w:p><w:bookmarkStart w:id='31' w:name='h'/></w:p></w:hdr>"
    doc = variant(base, tmp_path / "ids.docx", extra={"word/header1.xml": header})
    monkeypatch.setattr(revisions, "next_revision_id", lambda parts: 31)  # a scan that missed the header id
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="collide"):
        revisions.text_edit(doc, out, op="tracked_replace_text", paragraph_index=1, tracked=True,
                            find="opening", replace="opened")
    assert not out.exists()


# -- resolve_revisions: the independent ElementTree comparison is load-bearing -------------------------------

@pytest.mark.parametrize("ids, corrupt", [
    ([3], lambda body: body.replace("opening", "openinX", 1)),  # unrelated plain text, no revision touched
    ([1], lambda body: body.replace("highly ", "", 1)),  # the accepted insertion's text dropped
], ids=["plain-text", "selected-insertion"])
def test_resolution_writer_corruption_is_caught(tmp_path, monkeypatch, ids, corrupt):
    doc = create_tracked_fixture(tmp_path / "tracked.docx")
    original = revisions._render

    def render(xml, node, actions, dirty):
        out = original(xml, node, actions, dirty)
        return corrupt(out) if node.name == "w:body" else out

    monkeypatch.setattr(revisions, "_render", render)
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="independent resolution"):
        revisions.resolve_revisions(doc, out, mode="accept", ids=ids)
    assert not out.exists()
