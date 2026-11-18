"""Edit feedback for a paragraph the model added with insert_paragraph (plan 864 S2, Case J).

An added paragraph has no "before" text, so "rejected" cannot mean "the span reads as before". It is recorded
with ``added`` and the paragraphs already on each side of the added block (``nb_prev``/``nb_next``); it is
``rejected`` when it is gone and those two are adjacent again (Reject in Word, an MCP reject, or the user deleting
it), ``accepted`` when its text is there and its revision is not. An empty paragraph elsewhere never reads as a
rejection. Everything runs through the real tools; the user resolves outside the MCP.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from test_edit_feedback import (  # noqa: F401 - feedback_dir is a fixture
    ASSISTANT, PARAGRAPHS, _swap, accepted_text, build_doc, events, feedback, feedback_dir, model_edit,
)

from research_mcp import edit_feedback_word as fbw
from research_mcp.revisions import resolve_revisions


def add(paragraph: int, text: str, *, tracked: bool = True) -> dict:
    return {"type": "insert_paragraph", "after_paragraph_index": paragraph, "text": text, "tracked": tracked,
            "author": ASSISTANT}


def outcomes(path: Path) -> list[tuple[str, str, str]]:
    return [(e["outcome"], e["via"], e["kept"]) for e in events(path) if e["event"] == "outcome"]


def user_resolve_all(path: Path, mode: str) -> None:
    staged = path.with_name(f"user-{mode}-all.docx")
    resolve_revisions(path, staged, mode=mode, author=ASSISTANT)
    _swap(path, staged)


def user_deletes_paragraph(path: Path, text: str) -> None:
    """The user deletes, untracked, the one paragraph holding ``text``."""
    staged = path.with_name("user-delete.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                found = [m for m in re.finditer(r"<w:p>(?:(?!<w:p>).)*?</w:p>", xml, re.S) if text in m.group(0)]
                assert len(found) == 1
                xml = xml[:found[0].start()] + xml[found[0].end():]
                data = xml.encode("utf-8")
            out.writestr(info, data)
    _swap(path, staged)


def test_capture_records_added_and_the_neighbours(tmp_path):
    a = build_doc(tmp_path / "a.docx")
    b = tmp_path / "b.docx"
    from research_mcp.revisions import insert_paragraphs

    insert_paragraphs(a, b, after_paragraph_index=3, text="First added.\nSecond added.", style_id=None,
                      tracked=True, author=ASSISTANT)
    proposals, lost, _ = fbw.capture(a, b)
    assert [(p["para"], p["before"], p["after"], p["tracked"], p["added"], p["nb_prev"], p["nb_next"])
            for p in proposals] == [
        (4, "", "First added.", True, True, PARAGRAPHS[1], PARAGRAPHS[2]),
        (5, "", "Second added.", True, True, PARAGRAPHS[1], PARAGRAPHS[2]),
    ]
    assert not lost


def test_tracked_insert_pending_then_user_rejects(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "manuscript.docx")
    model_edit(path, add(3, "An added sentence."))
    assert accepted_text(path, 4) == "An added sentence."
    assert feedback(path)["edit_feedback"] is None  # pending: its revision is still there
    user_resolve_all(path, "reject")
    assert accepted_text(path, 4) == PARAGRAPHS[2]
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert outcomes(path) == [("rejected", "user", "")]


def test_tracked_insert_user_accepts(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "manuscript.docx")
    model_edit(path, add(3, "An added sentence."))
    user_resolve_all(path, "accept")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert outcomes(path) == [("accepted", "user", "An added sentence.")]


def test_two_added_paragraphs_rejected(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "manuscript.docx")
    model_edit(path, add(3, "First added.\nSecond added."))
    user_resolve_all(path, "reject")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 2 rejected. Call edit_feedback for details."


def test_insert_at_end_of_body_rejected(tmp_path, feedback_dir):
    """Word's way (no following paragraph): the anchor's mark is the inserted one; Reject merges back."""
    path = build_doc(tmp_path / "manuscript.docx")
    last = len(PARAGRAPHS) + 1
    model_edit(path, add(last, "Closing words."))
    assert accepted_text(path, last + 1) == "Closing words."
    user_resolve_all(path, "reject")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


def test_untracked_insert_deleted_by_the_user_is_rejected(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "manuscript.docx")
    model_edit(path, add(3, "A plain added sentence.", tracked=False))
    user_deletes_paragraph(path, "A plain added sentence.")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


def test_an_empty_paragraph_is_not_a_rejection():
    prop = {"before": "", "after": "Added.", "anchor_l": "", "anchor_r": "", "start": True, "end": True,
            "added": True}
    assert fbw.judge(prop, "") is None
    assert fbw.judge(prop, "Added.") == ("accepted", "Added.")
    plain = dict(prop, added=False)  # an empty paragraph the model filled: empty again is a rejection
    assert fbw.judge(plain, "") == ("rejected", "")


def test_user_rewords_the_added_paragraph(tmp_path, feedback_dir):
    """Neighbours not adjacent (the reworded paragraph sits between them): not rejected."""
    path = build_doc(tmp_path / "manuscript.docx")
    model_edit(path, add(3, "An added sentence about sites.", tracked=False))
    staged = path.with_name("user-type.docx")
    from research_mcp.revisions import text_edit

    text_edit(path, staged, op="replace_text", paragraph_index=4, find="about sites", replace="about every site",
              tracked=False)
    _swap(path, staged)
    feedback(path)
    assert [e["outcome"] for e in events(path) if e["event"] == "outcome"] == ["reworded"]


def test_reject_all_also_reverts_the_neighbours_own_changes(tmp_path, feedback_dir):
    """Both neighbours hold another author's tracked change; Word's Reject All rejects those too, so the
    neighbours read as their original text afterwards: still a rejection of the added paragraph."""
    from research_mcp.revisions import resolve_revisions, text_edit

    path = build_doc(tmp_path / "manuscript.docx")
    for paragraph, find, replace in ((3, "across all sites", "at every site"), (4, "0.001", "0.01")):
        staged = path.with_name(f"other-{paragraph}.docx")
        text_edit(path, staged, op="replace_text", paragraph_index=paragraph, find=find, replace=replace,
                  tracked=True, author="Reviewer B")
        _swap(path, staged)
    model_edit(path, add(3, "An added sentence."))
    proposal = next(e for e in events(path) if e.get("added"))
    assert proposal["nb_prev"] == "Results were good at every site."
    assert proposal["nb_prev_orig"] == PARAGRAPHS[1] and proposal["nb_next_orig"] == PARAGRAPHS[2]
    staged = path.with_name("user-reject-everything.docx")
    resolve_revisions(path, staged, mode="reject", paragraph_index=1, through_paragraph_index=len(PARAGRAPHS) + 2)
    _swap(path, staged)
    assert accepted_text(path, 3) == PARAGRAPHS[1] and accepted_text(path, 4) == PARAGRAPHS[2]
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


# -- S2 review r1: F2 (identity by place, not text) and F4 (neighbours resolved only partly) ---------------

def outcome_events(path: Path) -> list[dict]:
    return [e for e in events(path) if e["event"] == "outcome"]


def test_rejected_insertion_with_the_same_text_elsewhere_is_rejected(tmp_path, feedback_dir):
    """F2 repro: the added text also stands as an unrelated paragraph further down. Rejecting the added one
    rejoins its neighbours; the unrelated paragraph is not it, so this is not 'accepted' there."""
    path = build_doc(tmp_path / "manuscript.docx")
    same = PARAGRAPHS[3]  # paragraph 5
    model_edit(path, add(3, same))
    user_resolve_all(path, "reject")
    assert accepted_text(path, 5) == same
    assert [(e["outcome"], e["para"]) for e in outcome_events(path)] == []  # nothing until the next look
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert [(e["outcome"], e["kept"]) for e in outcome_events(path)] == [("rejected", "")]


def test_accepted_insertion_with_the_same_text_elsewhere_is_accepted_in_place(tmp_path, feedback_dir):
    """The added paragraph is still there and the same text stands further down: accepted at its own place."""
    path = build_doc(tmp_path / "manuscript.docx")
    same = PARAGRAPHS[3]  # paragraph 5, paragraph 6 after the insert
    model_edit(path, add(3, same))
    user_resolve_all(path, "accept")
    assert accepted_text(path, 4) == same and accepted_text(path, 6) == same
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert [(e["outcome"], e["para"]) for e in outcome_events(path)] == [("accepted", 4)]


def test_insertion_gone_with_a_neighbour_stays_pending(tmp_path, feedback_dir):
    """The added paragraph and its next neighbour are both deleted: its place is gone, so neither rejected nor
    removed; it stays pending."""
    path = build_doc(tmp_path / "manuscript.docx")
    model_edit(path, add(3, "A plain added sentence.", tracked=False))
    user_deletes_paragraph(path, "A plain added sentence.")
    user_deletes_paragraph(path, PARAGRAPHS[2])
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_neighbour_resolved_only_partly_still_reads_as_rejected(tmp_path, feedback_dir):
    """F4: the previous neighbour holds Reviewer B's replacement (a deletion and an insertion). The user rejects
    the added paragraph and only Reviewer B's insertion, leaving the deletion pending: the neighbour reads as
    neither its accepted nor its Reject All text, yet is one of its own readings, so the insertion is rejected
    (not 'removed')."""
    from research_mcp.revisions import text_edit

    path = build_doc(tmp_path / "manuscript.docx")
    staged = path.with_name("other.docx")
    text_edit(path, staged, op="replace_text", paragraph_index=3, find="good", replace="strong",
              tracked=True, author="Reviewer B")
    _swap(path, staged)
    model_edit(path, add(3, "An added sentence."))
    proposal = next(e for e in events(path) if e.get("added"))
    assert proposal["nb_prev"] == "Results were strong across all sites."
    assert proposal["nb_prev_orig"] == PARAGRAPHS[1]
    assert fbw._reading_hash("Results were  across all sites.") in proposal["nb_prev_alts"]  # 1 of 4 readings
    staged = path.with_name("user-reject-mine.docx")
    resolve_revisions(path, staged, mode="reject", author=ASSISTANT)
    _swap(path, staged)
    reviewer_ins = [r.id for r in fbw.Snapshot(path).revs[3] if r.type == "ins"]
    assert len(reviewer_ins) == 1
    staged = path.with_name("user-reject-theirs.docx")
    resolve_revisions(path, staged, mode="reject", ids=reviewer_ins)
    _swap(path, staged)
    assert accepted_text(path, 3) == "Results were  across all sites." and accepted_text(path, 4) == PARAGRAPHS[2]
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert [e["outcome"] for e in outcome_events(path)] == ["rejected"]


# -- S2 review r2: the original place must be the only one that fits ------------------------------------------

PREV, NEXT = "Prev paragraph text here.", "Next paragraph text here."


def user_deletes_nth(path: Path, text: str, nth: int) -> None:
    """The user deletes, untracked, the ``nth`` (0-based) paragraph holding ``text``."""
    staged = path.with_name("user-delete-nth.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                found = [m for m in re.finditer(r"<w:p>(?:(?!<w:p>).)*?</w:p>", xml, re.S) if text in m.group(0)]
                xml = xml[:found[nth].start()] + xml[found[nth].end():]
                data = xml.encode("utf-8")
            out.writestr(info, data)
    _swap(path, staged)


def test_next_neighbour_deleted_with_a_duplicate_pair_elsewhere_stays_pending(tmp_path, feedback_dir):
    """Reviewer repro: Prev, Added., a, b, c, d, Prev, Next after the original Next was deleted. The rejoined
    Prev, Next further down is a duplicate, not the original place, and the insertion still stands: pending."""
    path = build_doc(tmp_path / "m.docx", [PREV, NEXT, "a a.", "b b.", "c c.", "d d.", PREV, NEXT])
    model_edit(path, add(2, "Added.", tracked=False))
    proposal = next(e for e in events(path) if e.get("added"))
    assert (proposal["nb_prev_n"], proposal["nb_next_n"]) == (2, 2)
    user_deletes_nth(path, NEXT, 0)
    assert [accepted_text(path, i) for i in (2, 3, 4)] == [PREV, "Added.", "a a."]
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_insertion_and_its_neighbour_deleted_with_a_duplicate_triple_elsewhere_stays_pending(tmp_path, feedback_dir):
    """A Prev, Added., Next triple already stood elsewhere; the user deletes the added paragraph and its
    previous neighbour. The triple left is not the original place: pending, not accepted there."""
    path = build_doc(tmp_path / "m.docx", [PREV, NEXT, "a a.", PREV, "Added.", NEXT])
    model_edit(path, add(2, "Added.", tracked=False))
    user_deletes_nth(path, "Added.", 0)
    user_deletes_nth(path, PREV, 0)
    assert [accepted_text(path, i) for i in (2, 3, 4, 5, 6)] == [NEXT, "a a.", PREV, "Added.", NEXT]
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_repeated_neighbour_intact_still_decides(tmp_path, feedback_dir):
    """A neighbour's text repeats elsewhere but every copy is still there and only one place fits: rejected."""
    path = build_doc(tmp_path / "m.docx", [PREV, NEXT, "a a.", PREV, "b b."])
    model_edit(path, add(2, "Added.", tracked=False))
    user_deletes_nth(path, "Added.", 0)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


def user_appends(path: Path, *texts: str) -> None:
    """The user types new paragraphs at the end of the body."""
    staged = path.with_name("user-append.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                paragraphs = "".join(f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>" for t in texts)
                data = data.decode("utf-8").replace("<w:sectPr/>", paragraphs + "<w:sectPr/>").encode("utf-8")
            out.writestr(info, data)
    _swap(path, staged)


def test_insertion_left_beside_one_neighbour_is_not_rejected_by_a_new_pair(tmp_path, feedback_dir):
    """Unique neighbours at the edit; the user deletes the original Next and later types a Prev, Next pair at the
    end. That pair is rejoined, but the insertion still stands after the original Prev: pending."""
    path = build_doc(tmp_path / "m.docx", [PREV, NEXT, "a a.", "b b.", "c c.", "d d."])  # far apart
    model_edit(path, add(2, "Added.", tracked=False))
    user_deletes_nth(path, NEXT, 0)
    user_appends(path, PREV, NEXT)
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_repeated_neighbour_pair_both_rejoined_is_rejected(tmp_path, feedback_dir):
    """The Prev, Next pair repeats; the user deletes the insertion, so both pairs are adjacent and the proposal
    reads at neither: rejected (at the nearer pair)."""
    path = build_doc(tmp_path / "m.docx", [PREV, NEXT, "a a.", PREV, NEXT])
    model_edit(path, add(2, "Added.", tracked=False))
    user_deletes_nth(path, "Added.", 0)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert [(e["outcome"], e["para"]) for e in outcome_events(path)] == [("rejected", 3)]


def user_replaces_paragraph(path: Path, text: str, *new_texts: str) -> None:
    """The user replaces, untracked, the one paragraph holding ``text`` with plain paragraphs ``new_texts``."""
    staged = path.with_name("user-replace.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                found = [m for m in re.finditer(r"<w:p>(?:(?!<w:p>).)*?</w:p>", xml, re.S) if text in m.group(0)]
                assert len(found) == 1
                paragraphs = "".join(f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>" for t in new_texts)
                data = (xml[:found[0].start()] + paragraphs + xml[found[0].end():]).encode("utf-8")
            out.writestr(info, data)
    _swap(path, staged)


def test_reworded_insertion_beside_a_new_exact_copy_stays_pending(tmp_path, feedback_dir):
    """Reviewer repro: the user rewords the added "Added." to "Added!" and types a separate "Added." between the
    same neighbours. Either could be the insertion: pending, not accepted at the copy."""
    path = build_doc(tmp_path / "m.docx", [PREV, NEXT, "a a."])
    model_edit(path, add(2, "Added.", tracked=False))
    user_replaces_paragraph(path, "Added.", "Added!", "Added.")
    assert [accepted_text(path, i) for i in (2, 3, 4, 5)] == [PREV, "Added!", "Added.", NEXT]
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_two_added_paragraphs_accepted(tmp_path, feedback_dir):
    """The block's other added paragraph sits between the same neighbours and reads close to this one; it is
    not a rival copy, so both are accepted."""
    path = build_doc(tmp_path / "manuscript.docx")
    model_edit(path, add(3, "First added.\nSecond added."))
    user_resolve_all(path, "accept")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 2 accepted. Call edit_feedback for details."


# -- plan 864 S3b: a blank paragraph is no neighbour ----------------------------------------------------------

def user_deletes_blank(path: Path) -> None:
    """The user deletes the one blank paragraph of the body."""
    staged = path.with_name("user-delete-blank.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                found = [m for m in re.finditer(r"<w:p>(?:(?!<w:p>).)*?</w:p>", xml, re.S)
                         if not "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", m.group(0))).strip()]
                assert len(found) == 1
                xml = xml[:found[0].start()] + xml[found[0].end():]
                data = xml.encode("utf-8")
            out.writestr(info, data)
    _swap(path, staged)


def test_blank_line_before_the_insertion_is_passed_over(tmp_path, feedback_dir):
    """Added after a blank line: the neighbour is the paragraph before the blank (``gap_prev`` 1); deleting the
    added paragraph leaves that blank line next to Next, which reads as rejected."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a."])
    model_edit(path, add(3, "Added.", tracked=False))
    proposal = next(e for e in events(path) if e.get("added"))
    assert (proposal["nb_prev"], proposal.get("gap_prev"), proposal["nb_next"]) == (PREV, 1, NEXT)
    user_deletes_nth(path, "Added.", 0)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


def test_blank_line_deleted_beside_a_kept_insertion_is_accepted(tmp_path, feedback_dir):
    """The user deletes the blank line and keeps the added paragraph: Prev, Added., Next. The added paragraph
    now stands where the blank was; it is accepted, not read as Prev and Next rejoined."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a."])
    model_edit(path, add(3, "Added.", tracked=False))
    user_deletes_blank(path)
    assert [accepted_text(path, i) for i in (2, 3, 4)] == [PREV, "Added.", NEXT]
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."


# -- plan 864 S3b early review P1: a blank line that holds text again is no blank line ---------------------------

def user_fills_blank(path: Path, text: str, nth: int = 0) -> None:
    """The user types ``text`` into the ``nth`` (0-based) blank paragraph of the body."""
    staged = path.with_name("user-fill-blank.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                found = [m for m in re.finditer(r"<w:p>(?:(?!<w:p>).)*?</w:p>", xml, re.S)
                         if not "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", m.group(0))).strip()]
                filled = f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"
                xml = xml[:found[nth].start()] + filled + xml[found[nth].end():]
                data = xml.encode("utf-8")
            out.writestr(info, data)
    _swap(path, staged)


def test_blank_line_filled_and_insertion_gone_stays_pending(tmp_path, feedback_dir):
    """Prev, blank, Added., Next becomes Prev, Other words., Next. The blank line may have been filled and the
    insertion deleted (rejected), or the blank line deleted and the insertion rewritten: pending, not rejected."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a."])
    model_edit(path, add(3, "Added.", tracked=False))
    user_deletes_nth(path, "Added.", 0)
    user_fills_blank(path, "Other words entirely, nothing alike.")
    assert [accepted_text(path, i) for i in (2, 3, 4)] == [PREV, "Other words entirely, nothing alike.", NEXT]
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_blank_line_deleted_and_insertion_rewritten_stays_pending(tmp_path, feedback_dir):
    """The same end state reached the other way (blank deleted, the insertion rewritten past recognition)."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a."])
    model_edit(path, add(3, "Added.", tracked=False))
    user_replaces_paragraph(path, "Added.", "Other words entirely, nothing alike.")
    user_deletes_blank(path)
    assert [accepted_text(path, i) for i in (2, 3, 4)] == [PREV, "Other words entirely, nothing alike.", NEXT]
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_blank_line_deleted_with_the_insertion_is_rejected(tmp_path, feedback_dir):
    """The user deletes both the blank line and the insertion: Prev and Next adjacent, nothing between: rejected."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a."])
    model_edit(path, add(3, "Added.", tracked=False))
    user_deletes_nth(path, "Added.", 0)
    user_deletes_blank(path)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


def blanks_both_sides(tmp_path) -> Path:
    """Prev, blank, Added., blank, Next: a gap on each side."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", "", NEXT, "a a."])
    model_edit(path, add(3, "Added.", tracked=False))
    proposal = next(e for e in events(path) if e.get("added"))
    assert (proposal["nb_prev"], proposal.get("gap_prev"), proposal["nb_next"], proposal.get("gap_next")) == (
        PREV, 1, NEXT, 1)
    return path


def test_blanks_on_both_sides_kept_is_accepted(tmp_path, feedback_dir):
    path = blanks_both_sides(tmp_path)
    user_replaces_paragraph(path, "a a.", "a a b.")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."


def test_blanks_on_both_sides_deleted_is_rejected(tmp_path, feedback_dir):
    path = blanks_both_sides(tmp_path)
    user_deletes_nth(path, "Added.", 0)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


def test_blanks_on_both_sides_one_filled_stays_pending(tmp_path, feedback_dir):
    """Prev, Other., blank, Next: the filled line may be the insertion rewritten: pending."""
    path = blanks_both_sides(tmp_path)
    user_deletes_nth(path, "Added.", 0)
    user_fills_blank(path, "Other words entirely, nothing alike.", 0)
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_blanks_on_both_sides_next_filled_stays_pending(tmp_path, feedback_dir):
    path = blanks_both_sides(tmp_path)
    user_deletes_nth(path, "Added.", 0)
    user_fills_blank(path, "Other words entirely, nothing alike.", 1)
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_repeated_neighbour_elsewhere_without_a_blank_line_is_no_place(tmp_path, feedback_dir):
    """Prev repeats; its other copy is followed by text, not a blank line, and has no Next in reach: that copy is
    no place, so deleting the insertion still reads as rejected at the original one."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a.", PREV, "b b."])
    model_edit(path, add(3, "Added.", tracked=False))
    proposal = next(e for e in events(path) if e.get("added"))
    assert (proposal["nb_prev_n"], proposal.get("gap_prev")) == (2, 1)
    user_deletes_nth(path, "Added.", 0)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."


# -- plan 864 S3b recheck r2 P1: a dropped match must not hide the proposal standing past a filled blank ---------

def test_filled_blank_before_a_kept_insertion_with_a_repeated_pair_stays_pending(tmp_path, feedback_dir):
    """Reviewer repro: Prev and Next repeat. At the original place the blank line is filled and Next moved away
    (Prev, Other, Added.); the other copy still reads Prev, blank, Next. The insertion still stands: pending, not
    rejected at the other copy."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a.", PREV, "", NEXT, "c c."])
    model_edit(path, add(3, "Added.", tracked=False))
    proposal = next(e for e in events(path) if e.get("added"))
    assert (proposal["nb_prev_n"], proposal["nb_next_n"], proposal.get("gap_prev")) == (2, 2, 1)
    user_fills_blank(path, "Other words entirely, nothing alike.", 0)
    user_deletes_nth(path, NEXT, 0)
    user_appends(path, NEXT)
    assert [accepted_text(path, i) for i in (2, 3, 4)] == [PREV, "Other words entirely, nothing alike.", "Added."]
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []


def test_filled_blank_after_a_kept_insertion_with_a_repeated_pair_stays_pending(tmp_path, feedback_dir):
    """The mirror: Added., Other, Next at the original place (Prev moved away), Prev, blank, Next elsewhere."""
    path = build_doc(tmp_path / "m.docx", [PREV, "", NEXT, "a a.", PREV, "", NEXT, "c c."])
    model_edit(path, add(2, "Added.", tracked=False))
    proposal = next(e for e in events(path) if e.get("added"))
    assert (proposal["nb_prev_n"], proposal["nb_next_n"], proposal.get("gap_next")) == (2, 2, 1)
    user_fills_blank(path, "Other words entirely, nothing alike.", 0)
    user_deletes_nth(path, PREV, 0)
    user_appends(path, PREV)
    assert [accepted_text(path, i) for i in (2, 3, 4)] == ["Added.", "Other words entirely, nothing alike.", NEXT]
    assert feedback(path)["edit_feedback"] is None
    assert outcome_events(path) == []
