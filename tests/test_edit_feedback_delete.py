"""Edit feedback for a paragraph the model deleted with replace_section / replace_table_row (plan 864 S3a).

A deleted paragraph has no "after" text, so it is identified by its place, as an added one is (S2): the
paragraphs that stay on each side of it (``nb_prev``/``nb_next``). It is ``accepted`` when it is gone and those
two are adjacent again, ``rejected`` when its text reads again between them, ``reworded`` for a near match there,
and pending while its own ``w:del`` is in the document or the place is ambiguous (a repeated neighbour that lost
or gained a copy, two paragraphs between the neighbours reading as it, an empty paragraph between them). The last
line of a table cell is deleted Word's way (the mark of the line before is deleted). Everything runs through the
real tools; the user resolves outside the MCP.
"""
from __future__ import annotations

from pathlib import Path

from test_edit_feedback import (  # noqa: F401 - feedback_dir is a fixture
    ASSISTANT, events, feedback, feedback_dir, model_edit,
)
from test_edit_feedback_insert import user_deletes_nth, user_replaces_paragraph, user_resolve_all
from test_tracked_sections import LINE_PPR, OTHER_INS, TEXTS, make, para, run

from research_mcp import edit_feedback_word as fbw
from research_mcp.sections import read_layout

H1 = '<w:pStyle w:val="Heading1"/>'


def section(*paragraphs: str, tracked: bool = True) -> dict:
    return {"type": "replace_section", "heading": "Methods", "paragraphs": list(paragraphs), "tracked": tracked,
            "author": ASSISTANT}


def row(first_cell: str, tracked: bool = True) -> dict:
    return {"type": "replace_table_row", "match": "Role A", "cells": [first_cell, "2024"], "tracked": tracked,
            "author": ASSISTANT}


def texts(path: Path) -> list[str]:
    return [item.text for item in read_layout(path).paragraphs]


def outcomes(path: Path) -> list[tuple[str, str]]:
    return [(e["outcome"], e["kept"]) for e in events(path) if e["event"] == "outcome"]


def deleted(path: Path) -> list[dict]:
    return [e for e in events(path) if e["event"] == "proposal" and e.get("deleted")]


def user_types(path: Path, find: str, replace: str) -> None:
    """An unrelated untracked edit by the user, in the paragraph holding ``find``."""
    user_replaces_paragraph(path, find, replace)


ONE = "Since your last edit: 1 {}. Call edit_feedback for details."


# ---------------------------------------------------------------- capture

def test_capture_records_deleted_and_the_neighbours(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body."))
    [prop] = deleted(path)
    assert (prop["before"], prop["after"], prop["tracked"], prop["block"]) == ("Second body.", "", True, 1)
    assert (prop["nb_prev"], prop["nb_next"], prop["nb_prev_n"], prop["nb_next_n"]) == (
        "First body.", "Third body.", 1, 1)
    assert texts(path)[7:10] == ["First body.", "", "Third body."]  # pending: reads empty until accepted


def test_capture_untracked_delete_has_no_placeholder(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    [prop] = deleted(path)
    assert (prop["before"], prop["tracked"], prop["nb_prev"], prop["nb_next"]) == (
        "Second body.", False, "First body.", "Third body.")
    assert texts(path)[7:9] == ["First body.", "Third body."]


def test_an_emptied_paragraph_is_a_text_proposal_not_a_deletion(tmp_path, feedback_dir):
    """replace_section keeping a blank line in place of the text: the mark stays, so not a deleted paragraph."""
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "", "Third body."))
    assert deleted(path) == []


# ---------------------------------------------------------------- tracked deletion of a body paragraph

def test_tracked_delete_pending_then_user_rejects(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body."))
    assert feedback(path)["edit_feedback"] is None  # pending: its w:del is still there
    user_types(path, "Results", "Results and discussion")  # an unrelated change: still pending
    assert feedback(path)["edit_feedback"] is None
    user_resolve_all(path, "reject")
    assert texts(path)[7:10] == ["First body.", "Second body.", "Third body."]
    assert feedback(path)["edit_feedback"] == ONE.format("rejected")
    assert outcomes(path) == [("rejected", "Second body.")]


def test_tracked_delete_user_accepts(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body."))
    user_resolve_all(path, "accept")
    assert texts(path)[7:9] == ["First body.", "Third body."]
    assert feedback(path)["edit_feedback"] == ONE.format("accepted")
    assert outcomes(path) == [("accepted", "")]


def test_two_deleted_paragraphs_are_one_block(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("Third body."))
    props = deleted(path)
    assert [(p["before"], p["block"], p["nb_prev"], p["nb_next"]) for p in props] == [
        ("First body.", 2, "Methods", "Third body."), ("Second body.", 2, "Methods", "Third body.")]
    user_resolve_all(path, "reject")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 2 rejected. Call edit_feedback for details."


# ---------------------------------------------------------------- the last line of a cell (Word's way)

def test_trailing_cell_line_tracked_delete_rejected(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, row("Role A\nLine two"))
    [prop] = deleted(path)
    assert (prop["before"], prop["nb_prev"], prop["nb_next"]) == ("Line three", "Line two", "2024")
    assert feedback(path)["edit_feedback"] is None
    user_resolve_all(path, "reject")
    assert texts(path) == TEXTS
    assert feedback(path)["edit_feedback"] == ONE.format("rejected")


def test_trailing_cell_line_tracked_delete_accepted(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, row("Role A\nLine two"))
    user_resolve_all(path, "accept")
    assert texts(path)[1:4] == ["Role A", "Line two", "2024"]
    assert feedback(path)["edit_feedback"] == ONE.format("accepted")


# ---------------------------------------------------------------- untracked

def test_plain_delete_is_accepted_at_the_next_look(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    user_types(path, "Results", "Results and discussion")
    assert feedback(path)["edit_feedback"] == ONE.format("accepted")
    assert outcomes(path) == [("accepted", "")]


def test_plain_delete_typed_back_is_rejected_then_reworded(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    user_replaces_paragraph(path, "First body.", "First body.", "Second body.")
    assert feedback(path)["edit_feedback"] == ONE.format("rejected")
    assert outcomes(path) == [("rejected", "Second body.")]


def test_plain_delete_reworded_between_the_neighbours(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    user_replaces_paragraph(path, "First body.", "First body.", "Second body, reworded.")
    feedback(path)
    assert outcomes(path) == [("reworded", "Second body, reworded.")]


# ---------------------------------------------------------------- traps: pending, never a guess

def test_empty_paragraph_between_the_neighbours_stays_pending(tmp_path, feedback_dir):
    """The user leaves a blank line where the paragraph was: neither its neighbours adjacent nor its text back."""
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    user_replaces_paragraph(path, "First body.", "First body.", "")
    assert texts(path)[7:10] == ["First body.", "", "Third body."]
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def test_two_copies_between_the_neighbours_stay_pending(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body."))
    user_resolve_all(path, "reject")
    user_replaces_paragraph(path, "Second body.", "Second body.", "Second body.")
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def test_exact_copy_beside_a_reworded_one_stays_pending(tmp_path, feedback_dir):
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body."))
    user_resolve_all(path, "reject")
    user_replaces_paragraph(path, "Second body.", "Second body.", "Second body, reworded.")
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def repeated_body() -> str:
    """Fixture T with "Same line." both before the deleted paragraph and under Results."""
    return (para("", run("Intro paragraph"), OTHER_INS, run("."))
            + para(H1, run("Methods"))
            + para("", run("Same line.")) + para("", run("Second body.")) + para("", run("Third body."))
            + para(H1, run("Results")) + para("", run("Same line.")) + para(LINE_PPR, run("Closing.")))


def test_repeated_neighbour_rejected_when_every_copy_stays(tmp_path, feedback_dir):
    path = make(tmp_path, "rep", repeated_body())
    model_edit(path, section("Same line.", "Third body."))
    [prop] = deleted(path)
    assert (prop["nb_prev"], prop["nb_prev_n"], prop["nb_next_n"]) == ("Same line.", 2, 1)
    user_resolve_all(path, "reject")
    assert feedback(path)["edit_feedback"] == ONE.format("rejected")


def test_repeated_neighbour_that_lost_a_copy_stays_pending(tmp_path, feedback_dir):
    path = make(tmp_path, "rep", repeated_body())
    model_edit(path, section("Same line.", "Third body."))
    user_resolve_all(path, "reject")
    user_deletes_nth(path, "Same line.", 1)
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def test_repeated_neighbour_that_gained_a_copy_stays_pending(tmp_path, feedback_dir):
    path = make(tmp_path, "rep", repeated_body())
    model_edit(path, section("Same line.", "Third body.", tracked=False))
    user_replaces_paragraph(path, "Closing.", "Same line.", "Closing.")
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def test_stray_copy_beside_one_neighbour_stays_pending(tmp_path, feedback_dir):
    """The user deletes the next neighbour, types the text back after the previous one and types a copy of the
    neighbour pair further down. The pair is adjacent, but the text stands beside the original previous neighbour,
    so the deletion is not known to be accepted."""
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    user_replaces_paragraph(path, "Third body.", "Second body.")
    user_replaces_paragraph(path, "Results", "Results", "a a.", "b b.", "c c.", "First body.", "Third body.")
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def test_rejoined_pair_elsewhere_with_the_text_back_stays_pending(tmp_path, feedback_dir):
    """The text reads again between the neighbours, and a copy of the neighbour pair stands adjacent further
    down: two places fit, so it stays open."""
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    user_replaces_paragraph(path, "First body.", "First body.", "Second body.")
    user_replaces_paragraph(path, "Results", "Results", "First body.", "Third body.")
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def test_accepted_deletion_whose_text_is_typed_back_later_is_rejected(tmp_path, feedback_dir):
    """A deletion already accepted is watched: its text typed back between its neighbours reverts it."""
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "Third body.", tracked=False))
    user_types(path, "Results", "Results and discussion")
    assert feedback(path)["edit_feedback"] == ONE.format("accepted")
    user_replaces_paragraph(path, "First body.", "First body.", "Second body.")
    feedback(path)
    assert outcomes(path) == [("accepted", ""), ("rejected", "Second body.")]


def test_judge_never_reads_an_empty_paragraph_as_a_deleted_one_gone():
    prop = {"before": "Second body.", "after": "", "anchor_l": "", "anchor_r": "", "start": True, "end": True,
            "deleted": True}
    assert fbw.judge(prop, "") is None
    assert fbw.judge(prop, "Second body.") == ("rejected", "Second body.")
    plain = dict(prop, deleted=False)  # a paragraph the model emptied: empty is accepted
    assert fbw.judge(plain, "") == ("accepted", "")


# ---------------------------------------------------------------- a delete and an add in the same gap (review r1)
# The planner pairs the lines of one gap first-to-first (a TextStep: the paragraph keeps its mark, its text is
# replaced), then inserts or deletes the rest, so the document's revisions record a rewording plus an added or a
# deleted paragraph. Feedback pairs them the same way: one text proposal, plus one place proposal per extra line.

def text_proposals(path: Path) -> list[dict]:
    return [e for e in events(path) if e["event"] == "proposal" and not e.get("deleted") and not e.get("added")]


def added(path: Path) -> list[dict]:
    return [e for e in events(path) if e["event"] == "proposal" and e.get("added")]


def mixed_gap_with_an_extra_line(tmp_path) -> Path:
    path = make(tmp_path, "m")
    model_edit(path, section("First body.", "New second body.", "Extra line.", "Third body."))
    assert [e["after"] for e in added(path)] == ["Extra line."]
    assert len(text_proposals(path)) == 1 and not deleted(path)
    return path


def test_mixed_gap_reword_and_added_line_accepted(tmp_path, feedback_dir):
    path = mixed_gap_with_an_extra_line(tmp_path)
    user_resolve_all(path, "accept")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 2 accepted. Call edit_feedback for details."


def test_mixed_gap_reword_and_added_line_rejected(tmp_path, feedback_dir):
    path = mixed_gap_with_an_extra_line(tmp_path)
    user_resolve_all(path, "reject")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 2 rejected. Call edit_feedback for details."


def mixed_gap_with_deleted_lines(tmp_path) -> Path:
    path = make(tmp_path, "m")
    model_edit(path, section("Only body."))
    assert sorted(e["before"] for e in deleted(path)) == ["Second body.", "Third body."]
    assert len(text_proposals(path)) == 1 and not added(path)
    return path


def test_mixed_gap_reword_and_deleted_lines_accepted(tmp_path, feedback_dir):
    path = mixed_gap_with_deleted_lines(tmp_path)
    user_resolve_all(path, "accept")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 3 accepted. Call edit_feedback for details."


def test_mixed_gap_reword_and_deleted_lines_rejected(tmp_path, feedback_dir):
    path = mixed_gap_with_deleted_lines(tmp_path)
    user_resolve_all(path, "reject")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 3 rejected. Call edit_feedback for details."


# ---------------------------------------------------------------- table rows added or removed (plan 864 S3b)
# replace_section(rows=...) adding or removing a row: each of its paragraphs with text is an added or deleted
# paragraph proposal judged by place, one block per row. Row-end markers are not paragraphs here (no w:p, no
# text): a row's neighbours are the nearest paragraphs that stay, the last line of the row above (or the paragraph
# before the table) and the first line of the row below (or the paragraph after the table).

import pytest  # noqa: E402

from test_tracked_sections import ROWS, rows_doc, table_row  # noqa: E402


def rows(*cells: list[str], tracked: bool = True) -> dict:
    return {"type": "replace_section", "heading": "Experience", "rows": list(cells), "tracked": tracked,
            "author": ASSISTANT}


def places(path: Path) -> list[tuple]:
    return [(e["before"] or e["after"], e["block"], e["nb_prev"], e["nb_next"])
            for e in events(path) if e["event"] == "proposal" and (e.get("added") or e.get("deleted"))]


def counted(n: int, outcome: str) -> str:
    return f"Since your last edit: {n} {outcome}. Call edit_feedback for details."


D = ["Role D", "2021"]


@pytest.mark.parametrize("mode, outcome", [("accept", "accepted"), ("reject", "rejected")])
@pytest.mark.parametrize("where", ["end", "middle"])
def test_rows_feedback_tracked_added_row(tmp_path, feedback_dir, where, mode, outcome):
    """F1/F2: a tracked row added after the last row or between two rows."""
    path = rows_doc(tmp_path, f"F-add-{where}-{mode}")
    if where == "end":
        model_edit(path, rows(*ROWS, D))
        assert places(path) == [("Role D", 2, "2022", "Methods"), ("2021", 2, "2022", "Methods")]
    else:
        model_edit(path, rows(ROWS[0], ["Role N", "2025"], ROWS[1], ROWS[2]))
        assert places(path) == [("Role N", 2, "2024", "Role B"), ("2025", 2, "2024", "Role B")]
    assert feedback(path)["edit_feedback"] is None  # pending while its w:ins is there
    user_types(path, "Results", "Results and discussion")
    assert feedback(path)["edit_feedback"] is None
    user_resolve_all(path, mode)
    assert feedback(path)["edit_feedback"] == counted(2, outcome)


@pytest.mark.parametrize("mode, outcome, kept", [("accept", "accepted", ["", ""]),
                                                 ("reject", "rejected", ["2023", "Role B"])])
def test_rows_feedback_tracked_removed_row(tmp_path, feedback_dir, mode, outcome, kept):
    """F3: a tracked row deletion."""
    path = rows_doc(tmp_path, f"F-del-{mode}")
    model_edit(path, rows(ROWS[0], ROWS[2]))
    assert places(path) == [("Role B", 2, "2024", "Role C"), ("2023", 2, "2024", "Role C")]
    assert feedback(path)["edit_feedback"] is None  # pending while its w:del is there
    user_resolve_all(path, mode)
    assert feedback(path)["edit_feedback"] == counted(2, outcome)
    assert sorted(kept for _, kept in outcomes(path)) == kept


def test_rows_feedback_plain_added_row(tmp_path, feedback_dir):
    """F4: switch off on a revised document: the row is there at the next look (accepted); taken out again by the
    user, a second time, it is rejected."""
    path = rows_doc(tmp_path, "F-plain-add")
    original = path.read_bytes()
    model_edit(path, rows(*ROWS, D, tracked=False))
    assert places(path) == [("Role D", 2, "2022", "Methods"), ("2021", 2, "2022", "Methods")]
    user_types(path, "Results", "Results and discussion")
    assert feedback(path)["edit_feedback"] == counted(2, "accepted")

    again = rows_doc(tmp_path, "F-plain-add-out")
    model_edit(again, rows(*ROWS, D, tracked=False))
    again.write_bytes(original)  # the user removes the row
    assert feedback(again)["edit_feedback"] == counted(2, "rejected")


def test_rows_feedback_plain_removed_row(tmp_path, feedback_dir):
    """F5: a plain row removal: accepted at the next look; the row put back by the user is rejected."""
    path = rows_doc(tmp_path, "F-plain-del")
    original = path.read_bytes()
    model_edit(path, rows(ROWS[0], ROWS[2], tracked=False))
    assert places(path) == [("Role B", 2, "2024", "Role C"), ("2023", 2, "2024", "Role C")]
    user_types(path, "Results", "Results and discussion")
    assert feedback(path)["edit_feedback"] == counted(2, "accepted")

    again = rows_doc(tmp_path, "F-plain-del-back")
    model_edit(again, rows(ROWS[0], ROWS[2], tracked=False))
    again.write_bytes(original)  # the user puts the row back
    assert feedback(again)["edit_feedback"] == counted(2, "rejected")
    assert sorted(kept for _, kept in outcomes(again)) == ["2023", "Role B"]


BLANK_ROWS = [ROWS[0], ["Role B", ""], ["Role C", ""]]


def blank_rows_doc(tmp_path: Path, name: str) -> Path:
    return rows_doc(tmp_path, name, [table_row(cells) for cells in BLANK_ROWS])


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
@pytest.mark.parametrize("mode, outcome", [("accept", "accepted"), ("reject", "rejected")])
def test_rows_feedback_blank_cell_removed(tmp_path, feedback_dir, tracked, mode, outcome):
    """F6: the removed row's blank cell paragraph is no proposal and no neighbour: tracked, it reads "" until
    Accept removes it with the row, so a neighbour "" would never be found again (the deletion pending forever)."""
    path = blank_rows_doc(tmp_path, f"F-blank-del-{tracked}-{mode}")
    original = path.read_bytes()
    model_edit(path, rows(BLANK_ROWS[0], BLANK_ROWS[2], tracked=tracked))
    assert places(path) == [("Role B", 1, "2024", "Role C")]
    if tracked:
        user_resolve_all(path, mode)
    elif mode == "accept":
        user_types(path, "Results", "Results and discussion")
    else:
        path.write_bytes(original)
    assert feedback(path)["edit_feedback"] == counted(1, outcome)


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
@pytest.mark.parametrize("mode, outcome", [("accept", "accepted"), ("reject", "rejected")])
def test_rows_feedback_blank_cell_added(tmp_path, feedback_dir, tracked, mode, outcome):
    """F7: an added row with a blank cell: one proposal, for the line with text."""
    path = blank_rows_doc(tmp_path, f"F-blank-add-{tracked}-{mode}")
    original = path.read_bytes()
    model_edit(path, rows(*BLANK_ROWS, ["Role D", ""], tracked=tracked))
    assert places(path) == [("Role D", 2, "Role C", "Methods")]  # row C's blank last cell is not a neighbour
    [prop] = [e for e in events(path) if e["event"] == "proposal"]
    assert (prop.get("gap_prev"), prop.get("gap_next")) == (1, None)
    if tracked:
        user_resolve_all(path, mode)
    elif mode == "accept":
        user_types(path, "Results", "Results and discussion")
    else:
        path.write_bytes(original)
    assert feedback(path)["edit_feedback"] == counted(1, outcome)


@pytest.mark.parametrize("mode", ["accept", "reject"])
def test_rows_feedback_identical_row_added_stays_pending(tmp_path, feedback_dir, mode):
    """F8: a copy of row B added after B. Accept: "2023" / "Role C" adjacent at the second copy while "Role B"
    stands after the first; Reject: the neighbour "2023", read twice when recorded, lost a copy. Both pending."""
    path = rows_doc(tmp_path, f"F-dup-add-{mode}")
    model_edit(path, rows(ROWS[0], ROWS[1], ROWS[1], ROWS[2]))
    assert places(path) == [("Role B", 2, "2023", "Role C"), ("2023", 2, "2023", "Role C")]
    user_resolve_all(path, mode)
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


def test_rows_feedback_identical_row_removed(tmp_path, feedback_dir):
    """F9: one of two identical rows removed. Reject: the text is back after the first "2023" while the second
    "2023" / "Role C" pair is adjacent: both pending. Accept: "Role B" accepted; "2023" pending, since its text is
    its own previous neighbour's and the "2023" left may be it."""
    start = [table_row(cells) for cells in (ROWS[0], ROWS[1], ROWS[1], ROWS[2])]
    rejected = rows_doc(tmp_path, "F-dup-del-reject", start)
    model_edit(rejected, rows(*ROWS))
    assert places(rejected) == [("Role B", 2, "2023", "Role C"), ("2023", 2, "2023", "Role C")]
    user_resolve_all(rejected, "reject")
    assert feedback(rejected)["edit_feedback"] is None and outcomes(rejected) == []

    accepted = rows_doc(tmp_path, "F-dup-del-accept", start)
    model_edit(accepted, rows(*ROWS))
    user_resolve_all(accepted, "accept")
    assert feedback(accepted)["edit_feedback"] == (
        "Since your last edit: 1 accepted (1 pending). Call edit_feedback for details.")
    assert outcomes(accepted) == [("accepted", "")]


# -- plan 864 S3b early review P1: a blank neighbour line that holds text again leaves the outcome open -----------

from test_edit_feedback_insert import user_fills_blank  # noqa: E402


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
def test_rows_feedback_removed_row_after_a_blank_cell(tmp_path, feedback_dir, tracked):
    """Row C removed after row B's blank cell: the neighbour is "Role B" past the blank (``gap_prev`` 1). Gone,
    with the blank still blank: accepted."""
    path = blank_rows_doc(tmp_path, f"F-gap-del-{tracked}")
    model_edit(path, rows(*BLANK_ROWS[:2], tracked=tracked))
    assert places(path) == [("Role C", 1, "Role B", "Methods")]
    [prop] = [e for e in events(path) if e["event"] == "proposal"]
    assert prop.get("gap_prev") == 1
    if tracked:
        user_resolve_all(path, "accept")
    else:
        user_types(path, "Results", "Results and discussion")
    assert feedback(path)["edit_feedback"] == counted(1, "accepted")


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
def test_rows_feedback_removed_row_blank_cell_filled_stays_pending(tmp_path, feedback_dir, tracked):
    """The same removal, then the user types into row B's blank cell: Role B / 2019 / Methods. "2019" may be the
    removed row's text put back and rewritten, or a new value in the blank: pending."""
    path = blank_rows_doc(tmp_path, f"F-gap-fill-{tracked}")
    model_edit(path, rows(*BLANK_ROWS[:2], tracked=tracked))
    if tracked:
        user_resolve_all(path, "accept")
    user_fills_blank(path, "2019")
    assert texts(path)[-8:-5] == ["Role B", "2019", ""]
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
def test_rows_feedback_added_row_blank_cell_filled_stays_pending(tmp_path, feedback_dir, tracked):
    """F7's added row D (neighbour "Role C" past C's blank cell), then the user removes the row and types into C's
    blank cell: pending, not rejected."""
    path = blank_rows_doc(tmp_path, f"F-gap-add-fill-{tracked}")
    original = path.read_bytes()
    model_edit(path, rows(*BLANK_ROWS, ["Role D", ""], tracked=tracked))
    if tracked:
        user_resolve_all(path, "reject")
    else:
        path.write_bytes(original)
    user_fills_blank(path, "2020", 1)
    assert feedback(path)["edit_feedback"] is None
    assert outcomes(path) == []
