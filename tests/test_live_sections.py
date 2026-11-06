"""Live Word tests for minimal-diff editing, row and section rewrites, and groups (plan 328)."""

from __future__ import annotations

from pathlib import Path
import os
import sys

import pytest

from research_mcp import sections as S
from research_mcp.audit import file_sha256
from research_mcp.edit_batch import edit_document
from research_mcp.edit_models import EditDocumentRequest
from research_mcp.editing import (
    replace_paragraph_to_version,
    replace_section_to_version,
    replace_table_row_to_version,
    replace_text_to_version,
)
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_mixed_run_fixture
from research_mcp.publication import undo_document


pytestmark = [
    pytest.mark.word_live,
    pytest.mark.skipif(
        os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1" or sys.platform != "win32",
        reason="set RESEARCH_MCP_WORD_LIVE=1 to run disposable live Word tests",
    ),
]

TITLE = "Delta Park Museum, Department of Collections, Outreach Volunteer"
NEW_TITLE = "Delta Park Museum, Department of Colonial History, Seasonal Outreach Volunteer"


@pytest.fixture
def mixed(tmp_path: Path) -> Path:
    return create_mixed_run_fixture(tmp_path / "mixed.docx")


def _runs(item: S.Para) -> list[tuple[str, str]]:
    runs: list[list[str]] = []
    for ch, fp in zip(item.text, item.fps):
        if runs and runs[-1][1] == fp:
            runs[-1][0] += ch
        else:
            runs.append([ch, fp])
    return [(text, S._short(fp)) for text, fp in runs]


def test_live_whole_title_replace_keeps_run_formatting(mixed, tmp_path):
    output = tmp_path / "title.docx"
    result = replace_text_to_version(mixed, output, paragraph_index=17, find=TITLE, replace=NEW_TITLE)
    assert result["text_after"] == NEW_TITLE
    assert result["verification"]["consistency_check"] == "passed"
    assert _runs(S.read_layout(output).para(17)) == [
        ("Delta Park Museum", "[b]"),
        (", Department of Colonial History, ", "[plain]"),
        ("Seasonal Outreach Volunteer", "[i]"),
    ]
    assert file_sha256(mixed) == result["source_sha256"]


def test_live_replace_paragraph_inside_hyperlink_and_field_refusal(mixed, tmp_path):
    item = S.read_layout(mixed).para(28)
    result = replace_paragraph_to_version(mixed, tmp_path / "hyper.docx", paragraph_index=28,
                                          text=item.text.replace("see the top", "see top"))
    assert result["verification"]["consistency_check"] == "passed"
    after = S.read_layout(tmp_path / "hyper.docx").para(28)
    assert ("top", "[u=single in hyperlink]") in _runs(after)
    with pytest.raises(RefusedUnsafe, match="cut through field QUOTE"):
        replace_paragraph_to_version(mixed, tmp_path / "field.docx", paragraph_index=28,
                                     text=item.text.replace("Example 2024", "Example 2025"))


def _single_character_citation_fixture(path: Path) -> Path:
    """A paragraph with two Zotero-like fields whose results are one character each."""
    import docx
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = docx.Document()
    paragraph = document.add_paragraph("Keeping pace with the exponential growth of information by 2020.")

    def field(result: str) -> None:
        def run(child) -> None:
            r = OxmlElement("w:r")
            r.append(child)
            paragraph._p.append(r)
        begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin"); run(begin)
        code = OxmlElement("w:instrText"); code.set(qn("xml:space"), "preserve")
        code.text = ' ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"x","note":"' + "y" * 1500 + '"} '
        run(code)
        separate = OxmlElement("w:fldChar"); separate.set(qn("w:fldCharType"), "separate"); run(separate)
        r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr"); sup = OxmlElement("w:vertAlign")
        sup.set(qn("w:val"), "superscript"); rpr.append(sup); r.append(rpr)
        t = OxmlElement("w:t"); t.text = result; r.append(t); paragraph._p.append(r)
        end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end"); run(end)

    field("1")
    paragraph.add_run(" This gap has implications for outcomes and resource allocation.")
    field("2")
    document.save(path)
    return path


def test_live_replace_text_beside_single_character_citation(tmp_path):
    # Word reports a one-character field result as spanning the whole field
    # (begin marker to past the end marker); seen on three paragraphs of a real document.
    source = _single_character_citation_fixture(tmp_path / "cite.docx")
    output = tmp_path / "cite_out.docx"
    result = replace_text_to_version(source, output, paragraph_index=1,
                                     find="the exponential growth of information",
                                     replace="the growth of published information")
    assert result["verification"]["consistency_check"] == "passed"
    after = S.read_layout(output).para(1).text
    assert "the growth of published information by 2020.1 This gap" in after
    assert after.endswith("resource allocation.2")
    tail = replace_text_to_version(output, tmp_path / "cite_tail.docx", paragraph_index=1,
                                   find="and resource allocation.", replace="and costs.")
    assert tail["verification"]["consistency_check"] == "passed"
    assert S.read_layout(tmp_path / "cite_tail.docx").para(1).text.endswith("and costs.2")


def test_live_replace_table_row_and_date_refusal(mixed, tmp_path):
    cells = [NEW_TITLE + "\nContributed to collections review and donor records\n"
             "Focused on: Visitor Programs, Place-Based Exhibition\nAdded third line", "03/2020 – 08/2024"]
    result = replace_table_row_to_version(mixed, tmp_path / "row.docx", match=TITLE, cells=cells)
    assert result["index_delta"] == 1 and result["verification"]["consistency_check"] == "passed"
    row = S.find_row(S.read_layout(tmp_path / "row.docx"), NEW_TITLE)
    layout = S.read_layout(tmp_path / "row.docx")
    assert [S._cell_text(layout, cell) for cell in row.cells] == cells
    with pytest.raises(RefusedUnsafe, match="does not read as a date"):
        replace_table_row_to_version(mixed, tmp_path / "bad.docx", match=TITLE, cells=[TITLE, "Ongoing work"])


def test_live_replace_section_by_displayed_caps_heading(mixed, tmp_path):
    rows = [
        ["Lakeside University, College of Arts, Rivera Lab (Soil), Teaching Assistant\n"
         "Field research on soil microbes and crop yield", "01/2024 – Present "],
        [NEW_TITLE + "\nContributed to collections review", "03/2020 – Present "],
        ["New Org, Intern\nDid things", "06/2024 – 08/2024"],
    ]
    output = tmp_path / "section.docx"
    result = replace_section_to_version(mixed, output, heading="WORK EXPERIENCE", rows=rows)
    assert result["verification"]["consistency_check"] == "passed"
    layout = S.read_layout(output)
    section = S.locate_section(layout, "WORK EXPERIENCE")
    assert [[S._cell_text(layout, cell) for cell in row.cells] for row in section.rows] == rows
    assert layout.para(section.heading.index).text == "WorK Experience"  # stored case kept


def test_live_allow_partial_group_rollback_and_one_call_undo(mixed):
    original = file_sha256(mixed)
    request = EditDocumentRequest(
        source=str(mixed), source_sha256=original, allow_partial=True,
        operations=[
            {"op_id": "title", "type": "replace_text", "paragraph_index": 17, "find": TITLE, "replace": NEW_TITLE},
            {"op_id": "date", "type": "replace_text", "paragraph_index": 20, "find": "Present", "replace": "Ongoing work"},
            {"op_id": "closing", "type": "replace_text", "paragraph_index": 29, "find": "Plain", "replace": "Simple"},
        ],
    )
    result = edit_document(request)
    statuses = {item.op_id: item.status for item in result.operations}
    assert statuses == {"title": "rolled_back", "date": "failed", "closing": "applied"}
    assert [(g.group, g.status) for g in result.groups] == [("table 1 row 5", "rolled_back")]
    layout = S.read_layout(mixed)
    assert layout.para(17).text == TITLE and layout.para(29).text == "Simple closing paragraph."
    assert result.undo_call and result.undo_token in result.undo_call
    undone = undo_document(str(mixed), undo_token=result.undo_token)
    assert undone["restored"] and file_sha256(mixed) == original


def _with_closing(mixed: Path, target: Path, runs: str) -> Path:
    import zipfile

    from research_mcp.fixtures import _para, _run

    with zipfile.ZipFile(mixed) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    old = _para(_run("Plain closing paragraph.")).encode()
    assert parts["word/document.xml"].count(old) == 1
    parts["word/document.xml"] = parts["word/document.xml"].replace(old, runs.encode())
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return target


def test_live_rewritten_word_keeps_formatting_of_its_unchanged_characters(mixed, tmp_path):
    """Review r1 finding 1: AB**CD**_EF_ -> AXCDYF keeps CD bold and F italic."""

    from research_mcp.fixtures import _para, _run

    source = _with_closing(mixed, tmp_path / "inner.docx", _para(_run("AB"), _run("CD", "b"), _run("EF", "i")))
    output = tmp_path / "inner_out.docx"
    result = replace_paragraph_to_version(source, output, paragraph_index=29, text="AXCDYF")
    assert result["verification"]["consistency_check"] == "passed"
    assert _runs(S.read_layout(output).para(29)) == [("AX", "[plain]"), ("CD", "[b]"), ("YF", "[i]")]

    source = _with_closing(mixed, tmp_path / "stretch.docx",
                           _para(_run("AB"), _run("CD", "b"), _run("EF", "i"), _run("GH")))
    output = tmp_path / "stretch_out.docx"
    result = replace_paragraph_to_version(source, output, paragraph_index=29, text="AxCDyEFzH")
    assert result["verification"]["consistency_check"] == "passed"
    assert _runs(S.read_layout(output).para(29)) == [("Ax", "[plain]"), ("CDy", "[b]"), ("EF", "[i]"), ("zH", "[plain]")]


def test_live_row_named_by_a_staged_title_rolls_back_with_it(mixed):
    """Review r1 finding 2: the date op names the row by the title op 1 wrote; both roll back."""

    original = file_sha256(mixed)
    renamed = TITLE.replace("Delta Park Museum", "Delta Park Center")
    request = EditDocumentRequest(
        source=str(mixed), source_sha256=original, allow_partial=True,
        operations=[
            {"op_id": "title", "type": "replace_text", "paragraph_index": 17,
             "find": "Delta Park Museum", "replace": "Delta Park Center"},
            {"op_id": "date", "type": "replace_table_row", "match": renamed, "cells": [renamed, "Ongoing work"]},
            {"op_id": "closing", "type": "replace_text", "paragraph_index": 29, "find": "Plain", "replace": "Simple"},
        ],
    )
    result = edit_document(request)
    statuses = {item.op_id: item.status for item in result.operations}
    assert statuses == {"title": "rolled_back", "date": "failed", "closing": "applied"}
    assert [(g.group, g.op_ids, g.status) for g in result.groups] == [("table 1 row 5", ["title", "date"], "rolled_back")]
    layout = S.read_layout(mixed)
    assert layout.para(17).text == TITLE and layout.para(29).text == "Simple closing paragraph."
    undone = undo_document(str(mixed), undo_token=result.undo_token)
    assert undone["restored"] and file_sha256(mixed) == original


# Bookmarks and comments (review r2 P1) ---------------------------------------

def _word_anchors(source: Path, target: Path, find: str, offset: int, length: int, kind: str) -> Path:
    """Copy `source` with Word's own bookmark and/or comment on `find`[offset:offset + length]."""

    from research_mcp.word import IsolatedWordDocument

    with IsolatedWordDocument(source) as session:
        document = session.document
        found = document.Content
        assert found.Find.Execute(find, True), find  # MatchCase
        start = int(found.Start) + offset
        anchor = document.Range(start, start + length)
        if kind in ("bookmark", "both"):
            document.Bookmarks.Add("KeepClosing", anchor)
        if kind in ("comment", "both"):
            document.Comments.Add(anchor, "Keep this word.")
        session.save_as(target)
    return target


def _saved_anchors(path: Path) -> dict:
    """Bookmark and comment text read straight from the saved XML, independent of the planner."""

    import xml.etree.ElementTree as ET
    import zipfile

    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    open_marks: dict[str, str] = {}
    texts: dict[str, str] = {}
    for node in root.iter():
        mark = node.get(S.W + "id")
        if node.tag == S.W + "bookmarkStart" and node.get(S.W + "name") != "_GoBack":
            open_marks[mark] = "bookmark"
            texts["bookmark"] = ""
        elif node.tag == S.W + "commentRangeStart":
            open_marks["c" + mark] = "comment"
            texts["comment"] = ""
        elif node.tag == S.W + "bookmarkEnd":
            open_marks.pop(mark, None)
        elif node.tag == S.W + "commentRangeEnd":
            open_marks.pop("c" + mark, None)
        elif node.tag == S.W + "t":
            for kind in open_marks.values():
                texts[kind] += node.text or ""
    return texts


def _expected(kind: str, bookmark: str, comment: str | None = None) -> dict:
    wanted = {"bookmark": bookmark, "comment": comment if comment is not None else bookmark}
    return {key: value for key, value in wanted.items() if kind in (key, "both")}


@pytest.mark.parametrize("kind", ["bookmark", "comment", "both"])
def test_live_typing_after_an_anchored_word_keeps_the_whole_word_anchored(mixed, tmp_path, kind):
    source = _word_anchors(mixed, tmp_path / "anchored.docx", "Plain closing paragraph.", 6, 7, kind)
    assert _saved_anchors(source) == _expected(kind, "closing")
    output = tmp_path / "typed.docx"
    result = replace_paragraph_to_version(source, output, paragraph_index=29, text="Plain closing! paragraph.")
    assert result["verification"]["consistency_check"] == "passed"
    assert result["verification"]["anchors_verified"] == (2 if kind == "both" else 1)
    assert S.read_layout(output).para(29).text == "Plain closing! paragraph."
    assert _saved_anchors(output) == _expected(kind, "closing")


@pytest.mark.parametrize(("text", "bookmark", "comment"), [
    ("Plain !closing paragraph.", "!closing", "closing"),  # Word: a bookmark takes text typed at its start
    ("Plain closink paragraph.", "closink", "closink"),
    ("Plain Closing paragraph.", "Closing", "Closing"),
    ("Plain clo-sing paragraph.", "clo-sing", "clo-sing"),
])
def test_live_edits_at_anchor_boundaries_follow_word_typing(mixed, tmp_path, text, bookmark, comment):
    source = _word_anchors(mixed, tmp_path / "anchored.docx", "Plain closing paragraph.", 6, 7, "both")
    output = tmp_path / "edited.docx"
    result = replace_paragraph_to_version(source, output, paragraph_index=29, text=text)
    assert result["verification"]["anchors_verified"] == 2
    assert S.read_layout(output).para(29).text == text
    assert _saved_anchors(output) == {"bookmark": bookmark, "comment": comment}


@pytest.mark.parametrize("kind", ["bookmark", "comment", "both"])
def test_live_deleting_all_anchored_text_is_refused_by_name(mixed, tmp_path, kind):
    source = _word_anchors(mixed, tmp_path / "anchored.docx", "Plain closing paragraph.", 6, 7, kind)
    original = file_sha256(source)
    request = EditDocumentRequest(
        source=str(source), source_sha256=original,
        operations=[{"op_id": "drop", "type": "replace_paragraph", "paragraph_index": 29, "text": "Plain  paragraph."}],
    )
    result = edit_document(request)
    assert result.status == "no_changes" and file_sha256(source) == original
    (op,) = result.operations
    assert op.status == "failed"
    message = op.error.model_dump_json()
    assert "on 'closing' would be deleted because all of its text is deleted" in message
    assert "protected structure" not in message


def test_live_partial_deletion_keeps_the_remaining_anchored_text(mixed, tmp_path):
    source = _word_anchors(mixed, tmp_path / "anchored.docx", "Plain closing paragraph.", 6, 7, "both")
    output = tmp_path / "edited.docx"
    result = replace_paragraph_to_version(source, output, paragraph_index=29, text="Plain clos paragraph.")
    assert result["verification"]["anchors_verified"] == 2
    assert _saved_anchors(output) == {"bookmark": "clos", "comment": "clos"}


def test_live_typing_at_a_paragraph_start_anchor(mixed, tmp_path):
    source = _word_anchors(mixed, tmp_path / "anchored.docx", "Plain closing paragraph.", 0, 5, "both")
    output = tmp_path / "edited.docx"
    result = replace_paragraph_to_version(source, output, paragraph_index=29, text="Very Plain closing paragraph.")
    assert result["verification"]["anchors_verified"] == 2
    assert _saved_anchors(output) == {"bookmark": "Very Plain", "comment": "Plain"}


def test_live_cell_edits_keep_anchored_words(mixed, tmp_path):
    source = _word_anchors(mixed, tmp_path / "anchored.docx", "Collections,", 0, 11, "both")
    layout = S.read_layout(source)
    row = S.find_row(layout, TITLE)
    cells = [S._cell_text(layout, cell) for cell in row.cells]
    cells[0] = cells[0].replace("Collections,", "Collections!,", 1) + "\nAdded third line"
    output = tmp_path / "row.docx"
    result = replace_table_row_to_version(source, output, match=TITLE, cells=cells)
    assert result["verification"]["anchors_verified"] == 2
    assert _saved_anchors(output) == {"bookmark": "Collections", "comment": "Collections"}
    section = S.locate_section(layout, "WORK EXPERIENCE")
    rows = [[S._cell_text(layout, cell) for cell in item.cells] for item in section.rows]
    rows = [[cell.replace("Collections,", "Collections!,", 1) for cell in item] for item in rows]
    output = tmp_path / "section.docx"
    result = replace_section_to_version(source, output, heading="WORK EXPERIENCE", rows=rows)
    assert result["verification"]["anchors_verified"] == 2
    assert _saved_anchors(output) == {"bookmark": "Collections", "comment": "Collections"}
