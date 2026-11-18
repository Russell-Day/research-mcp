"""Plan 864 S1: research-mcp's tracking switch (custom document property) and how edit_document routes ops.

Fixtures are synthetic: the mixed-run résumé (no revisions) and the tracked-changes fixture (other authors'
revisions). Every expectation names what lands in the .docx, not just "no exception"."""

from __future__ import annotations

import asyncio
import os
import re
import zipfile
from pathlib import Path

import pytest

from research_mcp.audit import file_sha256
from research_mcp.edit_models import (
    DEFAULT_MAX_OPERATIONS,
    DeleteCitationOperation,
    EditDocumentRequest,
    InsertTextOperation,
    ReplaceParagraphOperation,
    ReplaceTextOperation,
    SweepCharactersOperation,
)
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_mixed_run_fixture, create_tracked_fixture
from research_mcp.paragraphs import read_paragraphs
from research_mcp.track_setting import (
    CUSTOM_CONTENT_TYPE,
    CUSTOM_PART,
    CUSTOM_REL_TYPE,
    DEFAULT_AUTHOR,
    NOT_TRACKABLE_MESSAGE,
    NOT_TRACKABLE_OPS,
    NOT_TRACKABLE_TOOLS,
    OFF,
    OFF_WITH_REVISIONS_NOTE,
    PROPERTY_NAME,
    TrackSetting,
    describe,
    read_setting,
    route_operation,
    write_setting,
)

windows_only = pytest.mark.skipif(os.name != "nt", reason="undo store writes need Windows Python")

ZOTERO_CUSTOM = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ZOTERO_PREF_1">'
    '<vt:lpwstr>&lt;data data-version="3" zotero-version="7.0"&gt;&lt;session id="abc"/&gt;</vt:lpwstr></property>'
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="ZOTERO_PREF_2">'
    '<vt:lpwstr>&lt;pref name="fieldType" value="Field"/&gt;&lt;/data&gt;</vt:lpwstr></property>'
    "</Properties>"
)
ZOTERO_PROPS = re.findall(r"<property\b.*?</property>", ZOTERO_CUSTOM)


def entries(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def document_xml(path: Path) -> str:
    return entries(path)["word/document.xml"].decode("utf-8")


def with_custom(source: Path, output: Path, custom: str) -> Path:
    """Copy of ``source`` that already has a docProps/custom.xml (as Zotero leaves one)."""

    parts = entries(source)
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>", f'<Override PartName="/{CUSTOM_PART}" ContentType="{CUSTOM_CONTENT_TYPE}"/></Types>'.encode())
    parts["_rels/.rels"] = parts["_rels/.rels"].replace(
        b"</Relationships>",
        f'<Relationship Id="rId9" Type="{CUSTOM_REL_TYPE}" Target="{CUSTOM_PART}"/></Relationships>'.encode())
    parts[CUSTOM_PART] = custom.encode("utf-8")
    with zipfile.ZipFile(output, "x", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output


@pytest.fixture
def cv(tmp_path: Path) -> Path:
    return create_mixed_run_fixture(tmp_path / "cv.docx")


@pytest.fixture
def tracked_doc(tmp_path: Path) -> Path:
    return create_tracked_fixture(tmp_path / "tracked.docx")


# -- the property ------------------------------------------------------------------------------------------

def test_unset_document_reads_off_and_describes_null_author(cv):
    assert read_setting(cv) == OFF
    assert describe(cv) == {"on": False, "author": None}


def test_write_creates_custom_part_content_type_and_relationship(cv, tmp_path):
    out = tmp_path / "on.docx"
    result = write_setting(cv, out, on=True, author="Reviewer B")

    assert result["changed"] is True and result["created_part"] is True
    assert read_setting(out) == TrackSetting(on=True, author="Reviewer B", is_set=True)
    assert describe(out) == {"on": True, "author": "Reviewer B"}
    before, after = entries(cv), entries(out)
    assert set(after) == set(before) | {CUSTOM_PART}
    assert f'PartName="/{CUSTOM_PART}" ContentType="{CUSTOM_CONTENT_TYPE}"'.encode() in after["[Content_Types].xml"]
    assert f'Type="{CUSTOM_REL_TYPE}" Target="{CUSTOM_PART}"'.encode() in after["_rels/.rels"]
    assert PROPERTY_NAME.encode() in after[CUSTOM_PART] and b'"author": "Reviewer B"' in after[CUSTOM_PART]
    for name in set(before) - {"[Content_Types].xml", "_rels/.rels"}:
        assert after[name] == before[name], name


def test_write_appends_beside_zotero_prefs_byte_identical(cv, tmp_path):
    src = with_custom(cv, tmp_path / "zotero.docx", ZOTERO_CUSTOM)
    out = tmp_path / "on.docx"
    write_setting(src, out, on=True)

    before, after = entries(src), entries(out)
    custom = after[CUSTOM_PART].decode("utf-8")
    for prop in ZOTERO_PROPS:  # every other property, byte for byte
        assert prop in custom
    assert custom.startswith(ZOTERO_CUSTOM[: ZOTERO_CUSTOM.index("</Properties>")])
    assert f'pid="4" name="{PROPERTY_NAME}"' in custom
    assert read_setting(out) == TrackSetting(on=True, author=DEFAULT_AUTHOR, is_set=True)
    assert {name for name in before if before[name] != after[name]} == {CUSTOM_PART}


def test_rewrite_replaces_only_its_own_property_and_keeps_author(cv, tmp_path):
    src = with_custom(cv, tmp_path / "zotero.docx", ZOTERO_CUSTOM)
    on = tmp_path / "on.docx"
    write_setting(src, on, on=True, author="Reviewer B")
    off = tmp_path / "off.docx"
    write_setting(on, off, on=False)  # author None keeps the author already set

    assert read_setting(off) == TrackSetting(on=False, author="Reviewer B", is_set=True)
    custom = entries(off)[CUSTOM_PART].decode("utf-8")
    assert custom.count(PROPERTY_NAME) == 1
    for prop in ZOTERO_PROPS:
        assert prop in custom
    again = tmp_path / "again.docx"
    unchanged = write_setting(off, again, on=False)
    assert unchanged["read_only"] is True and entries(again) == entries(off)


def test_broken_property_is_reported_by_describe_and_refused_by_read(cv, tmp_path):
    src = with_custom(cv, tmp_path / "bad.docx", ZOTERO_CUSTOM.replace(
        "</Properties>", f'<property fmtid="x" pid="9" name="{PROPERTY_NAME}"><vt:lpwstr>yes</vt:lpwstr>'
        "</property></Properties>"))
    with pytest.raises(RefusedUnsafe, match="not valid JSON"):
        read_setting(src)
    shown = describe(src)
    assert shown["on"] is False and "not valid JSON" in shown["error"]
    fixed = tmp_path / "fixed.docx"
    write_setting(src, fixed, on=True)  # set_track_changes rewrites it
    assert read_setting(fixed).on is True


def test_copy_under_another_name_keeps_the_switch(cv, tmp_path):
    on = tmp_path / "v19.docx"
    write_setting(cv, on, on=True, author="Reviewer B")
    copy = tmp_path / "v20.docx"
    copy.write_bytes(on.read_bytes())
    assert read_setting(copy) == TrackSetting(on=True, author="Reviewer B", is_set=True)


# -- routing -----------------------------------------------------------------------------------------------

ON = TrackSetting(on=True, author="Reviewer B", is_set=True)
SET_OFF = TrackSetting(on=False, author="Reviewer B", is_set=True)


def _route(op, setting):
    return route_operation(op, op, setting)


def test_replace_text_and_paragraph_route_to_tracked_ops_with_switch_author():
    op = ReplaceTextOperation(op_id="a", type="replace_text", paragraph_index=3, find="old", replace="new")
    routed, info = _route(op, ON)
    assert (routed.type, routed.find, routed.replace, routed.author, routed.paragraph_index) == \
        ("tracked_replace_text", "old", "new", "Reviewer B", 3)
    assert info == {"tracking": "on", "tracking_author": "Reviewer B", "routed_as": "tracked_replace_text"}

    para = ReplaceParagraphOperation(op_id="b", type="replace_paragraph", paragraph_index=2, text="All new.",
                                     author="Editor")
    routed, info = _route(para, ON)
    assert (routed.type, routed.text, routed.author) == ("tracked_replace_paragraph", "All new.", "Editor")
    assert info["tracking_author"] == "Editor"


@pytest.mark.parametrize(("position", "expected"), [("after", "anchor text"), ("before", " textanchor")])
def test_insert_text_routes_to_tracked_replace_of_its_anchor(position, expected):
    text = " text"
    op = InsertTextOperation(op_id="i", type="insert_text", paragraph_index=1, position=position,
                             anchor="anchor", text=text)
    routed, info = _route(op, ON)
    assert (routed.type, routed.find, routed.replace) == ("tracked_replace_text", "anchor", expected)
    assert info["routed_as"] == "tracked_replace_text"


def test_tracked_false_or_switch_off_leaves_the_op_alone():
    op = ReplaceTextOperation(op_id="a", type="replace_text", paragraph_index=1, find="x", replace="y",
                              tracked=False)
    assert _route(op, ON) == (op, {"tracking": "off"})
    plain = ReplaceTextOperation(op_id="a", type="replace_text", paragraph_index=1, find="x", replace="y")
    assert _route(plain, OFF) == (plain, {"tracking": "off"})
    forced = plain.model_copy(update={"tracked": True})
    routed, info = _route(forced, OFF)  # tracked: true without the switch tracks by "Claude"
    assert routed.type == "tracked_replace_text" and routed.author == DEFAULT_AUTHOR


def test_insert_table_is_tracked_in_place_while_on():
    """Plan 916 S5: insert_table no longer refuses while on; it keeps its type and takes tracked/author."""
    from research_mcp.edit_models import InsertTableOperation

    assert NOT_TRACKABLE_OPS == ()
    op = InsertTableOperation(op_id="x", type="insert_table", after_paragraph_index=1, headers=["a"],
                              rows=[["b"]], descriptive_columns=[1])
    routed, info = _route(op, ON)
    assert (routed.type, routed.tracked, routed.author) == ("insert_table", True, ON.author)
    assert info["routed_as"] == "insert_table" and info["tracking"] == "on"
    untracked = op.model_copy(update={"tracked": False})
    assert _route(untracked, ON) == (untracked, {"tracking": "off"})
    assert _route(op, OFF)[0] is op


def test_delete_citation_default_follows_the_switch_only_when_set():
    op = DeleteCitationOperation(op_id="d", type="delete_citation", paragraph_index=1, current_text="8")
    routed, _ = _route(op, OFF)  # never set: today's default, tracked by Claude
    assert (routed.tracked, routed.author) == (True, "Claude")
    routed, info = _route(op, SET_OFF)  # set off: removed outright
    assert routed.tracked is False and info == {"tracking": "off"}
    routed, info = _route(op, ON)
    assert (routed.tracked, routed.author) == (True, "Reviewer B")
    explicit = DeleteCitationOperation(op_id="d", type="delete_citation", paragraph_index=1, current_text="8",
                                       tracked=True, author="Editor")
    routed, _ = _route(explicit, SET_OFF)
    assert (routed.tracked, routed.author) == (True, "Editor")


def test_sweep_characters_stays_tracked_and_takes_the_switch_author():
    op = SweepCharactersOperation(op_id="s", type="sweep_characters")
    routed, info = _route(op, ON)
    assert routed.author == "Reviewer B" and info["tracking"] == "on"
    routed, info = _route(op, SET_OFF)
    assert routed.author == "Claude" and info["tracking"] == "on"


def test_text_edit_regex_counts_routed_ops_as_tracked_proposals():
    from research_mcp.edit_feedback_word import _TEXT_EDIT_OP

    assert _TEXT_EDIT_OP.search("replace_text(tracked) e1, insert_text(tracked) e2")
    assert _TEXT_EDIT_OP.search("replace_paragraph(tracked) p")
    assert not _TEXT_EDIT_OP.search("replace_text e1, insert_text e2, replace_paragraph p")


def test_batch_cap_leaves_room_for_a_leading_switch(cv):
    edits = [{"op_id": f"e{n}", "type": "replace_text", "paragraph_index": 29, "find": "x", "replace": "y"}
             for n in range(DEFAULT_MAX_OPERATIONS)]
    switch = {"op_id": "track_changes", "type": "set_track_changes", "on": True}
    EditDocumentRequest(source=str(cv), source_sha256="0" * 64, operations=[switch, *edits])
    with pytest.raises(ValueError, match=f"at most {DEFAULT_MAX_OPERATIONS} operations"):
        EditDocumentRequest(source=str(cv), source_sha256="0" * 64,
                            operations=[*edits, {**edits[0], "op_id": "extra"}])


def test_whole_document_tools_refuse_while_on(cv, tmp_path):
    from research_mcp.server import _refuse_while_tracking

    on = tmp_path / "on.docx"
    write_setting(cv, on, on=True)
    for tool in NOT_TRACKABLE_TOOLS:
        with pytest.raises(RefusedUnsafe) as caught:
            _refuse_while_tracking(str(on), tool, None)
        assert str(caught.value) == NOT_TRACKABLE_MESSAGE.format(op=tool)
        _refuse_while_tracking(str(on), tool, False)  # tracked=False: allowed
        _refuse_while_tracking(str(cv), tool, None)  # switch never set: allowed


# -- through edit_document (direct XML only; Word is not started) -------------------------------------------

def run_batch(path: Path, *ops: dict):
    from research_mcp.edit_batch import edit_document

    return edit_document(EditDocumentRequest(source=str(path), source_sha256=file_sha256(path), output=str(path),
                                             operations=list(ops)))


def revisions_by(xml: str, author: str) -> list[tuple[str, str]]:
    found = []
    for kind, body in re.findall(rf'<w:(ins|del) w:id="\d+" w:author="{re.escape(author)}"[^>]*>(.*?)</w:\1>',
                                 xml, re.S):
        found.append((kind, "".join(re.findall(r"<w:(?:t|delText)\b[^>]*>([^<]*)<", body))))
    return found


@windows_only
def test_cv_switch_on_reviewer_b_replace_text_lands_as_del_ins(cv):
    """Sibling case: revision-free résumé, tracking on with author "Reviewer B", a plain replace_text."""

    result = run_batch(cv, {"op_id": "on", "type": "set_track_changes", "on": True, "author": "Reviewer B"},
                       {"op_id": "e", "type": "replace_text", "paragraph_index": 29, "find": "closing",
                        "replace": "final"})
    assert result.status == "complete", result
    assert sorted(revisions_by(document_xml(cv), "Reviewer B")) == [("del", "closing"), ("ins", "final")]
    assert (result.tracking, result.tracking_author, result.tracking_note) == ("on", "Reviewer B", None)
    edit = result.operations[1].details
    assert (edit["tracking"], edit["tracking_author"], edit["routed_as"]) == ("on", "Reviewer B",
                                                                              "tracked_replace_text")
    assert read_setting(cv) == TrackSetting(on=True, author="Reviewer B", is_set=True)


@windows_only
def test_switch_tracks_insert_text_and_insert_paragraph_and_undo_restores(cv):
    from research_mcp.publication import undo_document

    original = file_sha256(cv)
    switched = run_batch(cv, {"op_id": "on", "type": "set_track_changes", "on": True})
    assert switched.status == "complete" and switched.undo_token
    assert switched.operations[0].details["track_changes"] == {"on": True, "author": "Claude"}

    inserted = run_batch(cv, {"op_id": "i", "type": "insert_text", "paragraph_index": 29, "position": "after",
                              "anchor": "closing", "text": " and last"})
    assert inserted.status == "complete", inserted
    [(kind, text)] = revisions_by(document_xml(cv), "Claude")
    assert kind == "ins" and text.strip() == "and last"
    assert read_paragraphs(cv)[28].text == "Plain closing and last paragraph."
    assert inserted.tracking == "on" and inserted.tracking_author == "Claude"

    # plan 864 S2: insert_paragraph is tracked while the switch is on (no longer refused)
    added = run_batch(cv, {"op_id": "p", "type": "insert_paragraph", "after_paragraph_index": 29, "text": "New"})
    assert added.status == "complete", added
    assert added.tracking == "on" and added.operations[0].details["routed_as"] == "insert_paragraph"
    assert read_paragraphs(cv)[29].text == "New"
    xml = document_xml(cv)
    runs = re.findall(r'<w:ins w:id="\d+" w:author="Claude"[^>/]*>(.*?)</w:ins>', xml, re.S)
    assert sorted("".join(re.findall(r"<w:t\b[^>]*>([^<]*)<", body)).strip() for body in runs) == ["New", "and last"]
    assert len(re.findall(r'<w:ins w:id="\d+" w:author="Claude"[^>]*/>', xml)) == 1  # the inserted paragraph mark

    undo_document(str(cv), undo_token=switched.undo_token)  # back to before the switch
    assert file_sha256(cv) == original and read_setting(cv) == OFF


@windows_only
def test_tracked_false_writes_plainly_and_off_with_revisions_notes_it(tracked_doc):
    from research_mcp.revisions import revision_count

    count = revision_count(tracked_doc)
    off = run_batch(tracked_doc, {"op_id": "i", "type": "insert_text", "paragraph_index": 1, "position": "after",
                                  "anchor": "opening", "text": " first"})
    assert off.status == "complete", off  # insert_text on a revised document: direct XML, untracked
    assert revision_count(tracked_doc) == count
    assert (off.tracking, off.tracking_note) == ("off", OFF_WITH_REVISIONS_NOTE)

    on = run_batch(tracked_doc, {"op_id": "on", "type": "set_track_changes", "on": True},
                   {"op_id": "e", "type": "replace_text", "paragraph_index": 1, "find": "Plain", "replace": "Simple",
                    "tracked": False},
                   {"op_id": "f", "type": "replace_text", "paragraph_index": 14, "find": "Closing",
                    "replace": "Final"})
    assert on.status == "complete", on
    xml = document_xml(tracked_doc)
    assert read_paragraphs(tracked_doc)[0].text == \
        "Simple opening first paragraph with no revisions."
    assert revisions_by(xml, "Claude") == [("del", "Closing"), ("ins", "Final")]
    assert on.operations[1].details["tracking"] == "off" and on.operations[2].details["tracking"] == "on"
    assert on.tracking == "on" and on.tracking_note is None


@windows_only
def test_switch_on_then_off_in_one_batch_leaves_the_following_edit_untracked(cv):
    """A later op in the batch sees the latest switch: on, one tracked edit, off, then a plain edit."""

    result = run_batch(cv, {"op_id": "on", "type": "set_track_changes", "on": True, "author": "Reviewer B"},
                       {"op_id": "t", "type": "replace_text", "paragraph_index": 29, "find": "closing",
                        "replace": "final"},
                       {"op_id": "off", "type": "set_track_changes", "on": False},
                       {"op_id": "p", "type": "replace_text", "paragraph_index": 29, "find": "paragraph",
                        "replace": "line"})
    assert result.status == "complete", result
    assert sorted(revisions_by(document_xml(cv), "Reviewer B")) == [("del", "closing"), ("ins", "final")]
    assert read_paragraphs(cv)[28].text == "Plain final line."
    tracked, plain = result.operations[1].details, result.operations[3].details
    assert (tracked["tracking"], tracked["routed_as"]) == ("on", "tracked_replace_text")
    assert plain["tracking"] == "off" and plain.get("routed_as") in (None, "replace_text")
    assert read_setting(cv) == TrackSetting(on=False, author="Reviewer B", is_set=True)
    assert result.tracking == "off"


def test_server_track_changes_sugar_prepends_the_switch(cv, monkeypatch):
    """edit_document(track_changes=true) runs set_track_changes first; the caller's op ids stay free."""

    from research_mcp import server as srv

    seen = {}

    def fake(request):
        seen["ops"] = [(op.op_id, op.type, getattr(op, "on", None)) for op in request.operations]
        raise RefusedUnsafe("stop here")

    monkeypatch.setattr(srv, "_edit_document", fake)

    async def call():
        return await srv.server.call_tool("edit_document", {
            "source": str(cv), "source_sha256": "0" * 64, "output": str(cv), "track_changes": True,
            "operations": [{"op_id": "track_changes", "type": "replace_text", "paragraph_index": 29,
                            "find": "closing", "replace": "final"}]})

    with pytest.raises(Exception, match="stop here"):
        asyncio.run(call())
    assert seen["ops"] == [("track_changes_2", "set_track_changes", True),
                           ("track_changes", "replace_text", None)]


@windows_only
def test_plain_replace_text_under_the_switch_is_a_tracked_feedback_proposal(tmp_path, monkeypatch):
    """Edit feedback sees the routed op as a tracked proposal by the switch's author (accepted in Word)."""

    from test_edit_feedback import ASSISTANT, build_doc, call, model_edit, records, user_resolve

    monkeypatch.setenv("RESEARCH_MCP_FEEDBACK_DIR", str(tmp_path / "feedback"))
    doc = build_doc(tmp_path / "manuscript.docx")
    model_edit(doc, {"type": "set_track_changes", "on": True, "author": ASSISTANT},
               {"type": "replace_text", "paragraph_index": 2, "find": "small", "replace": "modest"})
    user_resolve(doc, "accept", 2)  # accepts the switch author's revisions in that paragraph
    line = call("audit_document", path=str(doc))["edit_feedback"]  # a later call sees the outcome
    assert line == "Since your last edit: 1 accepted. Call edit_feedback for details."
    record = records(doc)[2]
    assert (record["outcome"], record["tracked"], record["author"]) == ("accepted", True, ASSISTANT)
    assert (record["original"], record["proposed"], record["kept"]) == ("small", "modest", "modest")
