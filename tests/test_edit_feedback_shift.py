"""Edit feedback (plan 916 S2): filled-in blank paragraphs after an outside save shifts paragraphs.

The model fills blank answer lines paragraph-wide (``before`` is ""). An outside save then inserts paragraphs
above them, so every answer sits one or two paragraphs further down. A blank line near the recorded place reads
as the old text, but it is not the answer's place: the fill is still there, kept. These tests check that such
fills are judged ``accepted`` where they now stand, and that a fill the user really blanked is still
``rejected`` at its own place (its recorded neighbours around it).

Every document holds one earlier tracked insertion by another author, so the model's edits run as direct XML
(no Word). The outside changes are plain XML rewrites, as another program's save would leave the file.
"""

from __future__ import annotations

import re
import zipfile
from collections import Counter
from pathlib import Path

from test_edit_feedback import (  # noqa: F401 - feedback_dir is a fixture
    W, _swap, _t, call, events, feedback, feedback_dir, model_edit,
)

from research_mcp import edit_feedback_word as fbw

QUESTIONS = [
    "Which statement reports the cash position at year end?",
    "Name two operating activities in the cash flow statement.",
    "How is gross margin computed from the income statement?",
    "What does the current ratio measure?",
    "Where are dividends paid shown?",
    "Which line shows depreciation expense?",
    "Why might net income differ from operating cash flow?",
    "What is the effect of a stock split on equity?",
    "Which account records prepaid rent?",
    "How are accrued wages classified?",
    "What does the auditor's opinion cover?",
    "When is revenue recognised under the five-step model?",
    "Which note explains the lease obligations?",
    "What is the quick ratio's numerator?",
    "How did inventory turnover change?",
    "Which segment reported the highest operating income?",
    "What drives the change in retained earnings?",
    "How is diluted earnings per share calculated?",
    "Which liabilities are due within one year?",
    "What is the purpose of the statement of comprehensive income?",
]
ANSWERS = [
    "The balance sheet.",
    "Collections from customers and payments to suppliers.",
    "Revenue minus cost of goods sold, divided by revenue.",
    "Short-term liquidity: current assets over current liabilities.",
    "In financing activities.",
    "It sits within operating expenses.",
    "Accruals and non-cash charges such as depreciation.",
    "None; only the share count and par value change.",
    "Prepaid expenses, a current asset.",
    "As current liabilities.",
    "Whether the statements are fairly presented.",
    "When control of the goods passes to the customer.",
    "Note 9 on commitments.",
    "Cash, marketable securities and receivables.",
    "It slowed from eight times to six times a year.",
    "The industrial segment.",
    "Net income less dividends declared.",
    "Net income over weighted shares plus dilutive options.",
    "Accounts payable, accrued wages and the current debt portion.",
    "To report gains and losses that bypass net income.",
]
ROLES = [
    "Research assistant, Harbour Institute, 2012-2013",
    "Data analyst, Northfield Clinic, 2013-2014",
    "Graduate fellow, Lakeside University, 2014-2016",
    "Teaching assistant, Lakeside University, 2015-2016",
    "Visiting scholar, Coastal Lab, 2016",
    "Postdoctoral fellow, Ridge Center, 2016-2018",
    "Statistician, Valley Health, 2018-2019",
    "Senior analyst, Valley Health, 2019-2020",
    "Lecturer, Eastgate College, 2020-2021",
    "Assistant professor, Eastgate College, 2021-2023",
    "Associate professor, Eastgate College, 2023-",
    "Consultant, Open Data Group, 2022-2023",
]
DUTIES = [
    "Cleaned survey data and kept the codebook.",
    "Built weekly reports on clinic wait times.",
    "Studied missing-data methods for registries.",
    "Ran two sections of introductory statistics.",
    "Prototyped a model of coastal flooding.",
    "Led a cohort study of readmissions.",
    "Designed sampling plans for audits.",
    "Managed a team of four analysts.",
    "Taught regression and study design.",
    "Started a lab on causal inference.",
    "Chair of the graduate curriculum committee.",
    "Advised a city on open data releases.",
]
TOP = "Name: ________   Section: ____"
P_RE = re.compile(r"<w:p(?:\s[^>]*)?(?:/>|>.*?</w:p>)", re.S)


def build(path: Path, title: str, author: str, paragraphs: list[str]) -> Path:
    """A document whose first paragraph holds ``author``'s tracked insertion, then ``paragraphs`` ("" = blank)."""
    first = ('<w:p>' + _t(title + " ")
             + f'<w:ins w:id="1" w:author="{author}" w:date="2026-10-01T09:00:00Z">' + _t("(draft)") + "</w:ins>"
             + "</w:p>")
    body = first + "".join(f"<w:p>{_t(text)}</w:p>" if text else "<w:p/>" for text in paragraphs)
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f"<w:body>{body}<w:sectPr/></w:body></w:document>")
    from research_mcp.fixtures import _MIXED_CONTENT_TYPES, _MIXED_DOCUMENT_RELS, _MIXED_RELS, _MIXED_STYLES

    with zipfile.ZipFile(path, "x", zipfile.ZIP_DEFLATED) as archive:
        for name, data in (("[Content_Types].xml", _MIXED_CONTENT_TYPES), ("_rels/.rels", _MIXED_RELS),
                           ("word/document.xml", document), ("word/_rels/document.xml.rels", _MIXED_DOCUMENT_RELS),
                           ("word/styles.xml", _MIXED_STYLES)):
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 11, 1, 0, 0, 0)), data)
    return path


def outside_save(path: Path, change) -> None:
    """Another program saves the file: ``change`` gets the body's paragraph XML list (the first is the title)
    and returns the new list."""
    staged = path.with_name("outside-save.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                start, end = xml.index("<w:body>") + len("<w:body>"), xml.index("<w:sectPr")
                paragraphs = P_RE.findall(xml[start:end])
                assert "".join(paragraphs) == xml[start:end]
                data = (xml[:start] + "".join(change(paragraphs)) + xml[end:]).encode("utf-8")
            out.writestr(info, data)
    _swap(path, staged)


def plain_p(text: str) -> str:
    return f"<w:p>{_t(text)}</w:p>" if text else "<w:p/>"


def insert_at(position: int, *texts: str):
    return lambda paragraphs: paragraphs[:position] + [plain_p(t) for t in texts] + paragraphs[position:]


def set_at(position: int, text: str):
    return lambda paragraphs: paragraphs[:position] + [plain_p(text)] + paragraphs[position + 1:]


def fill_all(path: Path, indices: list[int], texts: list[str]) -> list[dict]:
    model_edit(path, *({"type": "replace_paragraph", "paragraph_index": index, "text": text, "tracked": False}
                       for index, text in zip(indices, texts)))
    proposals = [e for e in events(path) if e["event"] == "proposal"]
    assert [(p["para"], p["before"], p["after"], p["start"], p["end"]) for p in proposals] == [
        (index, "", text, True, True) for index, text in zip(indices, texts)]
    return proposals


def results(path: Path) -> dict[str, tuple[str, str | None, int | None]]:
    """Each proposal's last outcome: (outcome, kept, paragraph)."""
    feedback(path)  # the next tool call sees the outside save
    return {e["id"]: (e["outcome"], e["kept"], e["para"]) for e in events(path) if e["event"] == "outcome"}


def worksheet(tmp_path: Path) -> tuple[Path, list[int]]:
    """Title, then per question: a blank spacer line, the question, a blank answer line."""
    lines = [text for question in QUESTIONS for text in ("", question, "")]
    path = build(tmp_path / "worksheet.docx", "Financial statements worksheet", "Reviewer A", lines)
    first = min(fbw.Snapshot(path).indices)
    return path, [first + 1 + 3 * k + 2 for k in range(len(QUESTIONS))]


# ---------------------------------------------------------------- shifted fills are kept

def test_fills_shifted_by_a_paragraph_at_the_top_are_accepted(tmp_path, feedback_dir):
    path, answers = worksheet(tmp_path)
    proposals = fill_all(path, answers, ANSWERS)
    outside_save(path, insert_at(0, TOP))
    got = results(path)
    assert Counter(outcome for outcome, _, _ in got.values()) == {"accepted": 20}
    assert [got[p["id"]][2] for p in proposals] == [index + 1 for index in answers]


def test_accepted_fills_stay_accepted_after_a_second_shift(tmp_path, feedback_dir):
    path, answers = worksheet(tmp_path)
    fill_all(path, answers, ANSWERS)
    outside_save(path, insert_at(0, TOP))
    assert Counter(outcome for outcome, _, _ in results(path).values()) == {"accepted": 20}
    outside_save(path, insert_at(0, "Due Friday."))
    outcomes = [e for e in events(path) if e["event"] == "outcome"]
    feedback(path)
    assert [e for e in events(path) if e["event"] == "outcome"] == outcomes  # nothing flips to rejected


def test_cv_fills_with_two_paragraphs_inserted_mid_document_are_accepted(tmp_path, feedback_dir):
    lines = [text for role in ROLES for text in (role, "", "")]  # each role, its blank duty line, a spacer
    path = build(tmp_path / "cv.docx", "Curriculum vitae", "Reviewer B", lines)
    first = min(fbw.Snapshot(path).indices)
    duties = [first + 1 + 3 * k + 1 for k in range(len(ROLES))]
    proposals = fill_all(path, duties, DUTIES)
    # two paragraphs between entry 8 and entry 9 (after entry 8's spacer: body position 1 + 3 * 8)
    outside_save(path, insert_at(1 + 3 * 8, "Awards", "Best poster prize, 2019"))
    got = results(path)
    assert Counter(outcome for outcome, _, _ in got.values()) == {"accepted": len(ROLES)}
    assert [got[p["id"]][2] for p in proposals] == [index + (2 if k >= 8 else 0) for k, index in enumerate(duties)]


# ---------------------------------------------------------------- real rejections still count

def test_fill_blanked_in_place_is_rejected(tmp_path, feedback_dir):
    path, answers = worksheet(tmp_path)
    proposals = fill_all(path, answers, ANSWERS)
    outside_save(path, set_at(1 + 3 * 4 + 2, ""))  # the user deletes answer 5's text
    got = results(path)
    assert got[proposals[4]["id"]] == ("rejected", "", answers[4])
    assert Counter(outcome for outcome, _, _ in got.values()) == {"accepted": 19, "rejected": 1}


def test_fill_blanked_after_a_shift_is_rejected_at_its_new_place(tmp_path, feedback_dir):
    path, answers = worksheet(tmp_path)
    proposals = fill_all(path, answers, ANSWERS)
    outside_save(path, lambda ps: insert_at(0, TOP)(set_at(1 + 3 * 4 + 2, "")(ps)))
    got = results(path)
    assert got[proposals[4]["id"]] == ("rejected", "", answers[4] + 1)
    assert Counter(outcome for outcome, _, _ in got.values()) == {"accepted": 19, "rejected": 1}
    assert all(got[p["id"]] == ("accepted", ANSWERS[k], answers[k] + 1) for k, p in enumerate(proposals) if k != 4)


def test_fill_reworded_after_a_shift_is_reworded(tmp_path, feedback_dir):
    path, answers = worksheet(tmp_path)
    proposals = fill_all(path, answers, ANSWERS)
    mine = "Accruals and non-cash charges, mainly depreciation."
    outside_save(path, lambda ps: insert_at(0, TOP)(set_at(1 + 3 * 6 + 2, mine)(ps)))
    got = results(path)
    assert got[proposals[6]["id"]] == ("reworded", mine, answers[6] + 1)
    assert Counter(outcome for outcome, _, _ in got.values()) == {"accepted": 19, "reworded": 1}


def test_identical_fill_blanked_is_rejected_not_taken_from_its_twin(tmp_path, feedback_dir):
    path, answers = worksheet(tmp_path)
    texts = list(ANSWERS)
    texts[3] = texts[4] = "Not covered in this chapter."
    proposals = fill_all(path, answers, texts)
    outside_save(path, set_at(1 + 3 * 3 + 2, ""))  # the user blanks the first of the twins
    got = results(path)
    assert got[proposals[3]["id"]] == ("rejected", "", answers[3])
    assert got[proposals[4]["id"]] == ("accepted", texts[4], answers[4])
    assert Counter(outcome for outcome, _, _ in got.values()) == {"accepted": 19, "rejected": 1}


# ---------------------------------------------------------------- ambiguity stays open

def test_blank_and_answer_equally_placed_stay_pending(tmp_path, feedback_dir):
    """Identical lines all around: after the outside save a line reading as the fill and a blank line both stand
    exactly among the recorded neighbours, so neither place can be told from the other."""
    same = "See the attached table."
    path = build(tmp_path / "repeated.docx", "Attachment list", "Reviewer A", [same] * 5 + [""] + [same] * 3)
    first = min(fbw.Snapshot(path).indices)
    (proposal,) = fill_all(path, [first + 6], ["Not applicable."])
    assert proposal["nbt"] == [same] * 6
    outside_save(path, lambda ps: [ps[0]] + [plain_p(t) for t in [same] * 3 + ["Not applicable."] + [same] * 3
                                             + [""] + [same] * 3])
    assert results(path) == {}
    assert [e["id"] for e in fbw.doc_log(path).events() if e["event"] == "proposal"] == [proposal["id"]]


TWIN = "Not applicable."
TWIN_LINES = ["Alpha", "Beta", "Gamma", "Delta", "Question one", TWIN, "Question two", "", "Next question",
              "Other text", "End"]
FOUR = ("New 1", "New 2", "New 3", "New 4")


def twin_document(tmp_path: Path) -> tuple[Path, int, dict]:
    """A blank answer line two paragraphs below an identical answer; the model fills it with that answer."""
    path = build(tmp_path / "twin.docx", "Attachment list", "Reviewer A", TWIN_LINES)
    first = min(fbw.Snapshot(path).indices)
    (proposal,) = fill_all(path, [first + 8], [TWIN])
    return path, first, proposal


def test_fill_blanked_after_a_long_shift_is_rejected_not_taken_from_its_twin(tmp_path, feedback_dir):
    """Four paragraphs above the fill move it past the nearby window while the user blanks it: the blank line
    standing exactly among the recorded neighbours wins over the identical answer near the recorded place."""
    path, first, proposal = twin_document(tmp_path)
    outside_save(path, lambda ps: insert_at(1, *FOUR)(set_at(8, "")(ps)))
    assert results(path) == {proposal["id"]: ("rejected", "", first + 12)}


def test_fill_kept_then_blanked_after_a_long_shift_is_rejected_not_taken_from_its_twin(tmp_path, feedback_dir):
    """The watched path: the fill was accepted after the long shift, then the user blanks it."""
    path, first, proposal = twin_document(tmp_path)
    outside_save(path, insert_at(1, *FOUR))
    assert results(path)[proposal["id"]][:2] == ("accepted", TWIN)
    outside_save(path, set_at(12, ""))
    assert results(path) == {proposal["id"]: ("rejected", "", first + 12)}


def test_fill_kept_after_a_long_shift_is_accepted_at_its_own_line_not_its_twin(tmp_path, feedback_dir):
    path, first, proposal = twin_document(tmp_path)
    outside_save(path, insert_at(1, *FOUR))
    assert results(path) == {proposal["id"]: ("accepted", TWIN, first + 12)}



# ---------------------------------------------------------------- a repeated block further away

BLOCK = ["Line one", "Line two", "Line three", "", "Line five", "Line six", "Line seven"]
REPEATED = BLOCK + ["Part two", "Notes", "Continued"] + BLOCK + ["End"]  # body positions 4 and 14 are blank
REVISED = (7, "Line seven, revised")  # the user changes one line around the fill: 5 of its 6 neighbours remain
OPEN = (None, ("unresolved", None))


def repeated(tmp_path: Path, second: str) -> tuple[Path, int, dict]:
    """Two blocks with the same six lines around a blank line, ten paragraphs apart; the second block's line is
    ``second``. The model fills the first block's blank line."""
    lines = REPEATED[:13] + [second] + REPEATED[14:]
    path = build(tmp_path / "repeated.docx", "Form A", "Reviewer A", lines)
    first = min(fbw.Snapshot(path).indices)
    (proposal,) = fill_all(path, [first + 4], [TWIN])
    return path, first, proposal


def verdict(path: Path, proposal: dict, *, watched: bool):
    """What the chooser finds: the open path reads rewordings too, the watched path full readings only."""
    return fbw._blank_fill_verdict(proposal, fbw.Snapshot(path), short=not watched)


def test_fill_kept_beside_a_repeated_block_with_a_blank_line_stays_pending(tmp_path, feedback_dir):
    """Open path. The fill keeps 5 of its 6 recorded neighbours; the other block's blank line keeps all 6. Either
    could be the proposal's place, and they disagree (kept, blanked): left open, never rejected over there."""
    path, first, proposal = repeated(tmp_path, "")
    outside_save(path, set_at(*REVISED))
    assert results(path) == {}
    assert verdict(path, proposal, watched=False) == OPEN


def test_fill_kept_beside_a_repeated_block_with_a_blank_line_stays_pending_when_watched(tmp_path, feedback_dir):
    path, first, proposal = repeated(tmp_path, "")
    outside_save(path, set_at(*REVISED))
    assert fbw.reconcile_props([], fbw.Snapshot(path), via="user", sha=None, watched=[proposal]) == []
    assert verdict(path, proposal, watched=True) == OPEN


def test_fill_blanked_beside_a_repeated_block_with_its_twin_stays_pending(tmp_path, feedback_dir):
    """Open path. The user blanks the fill and changes one line around it; the other block holds an identical
    answer among all 6 recorded neighbours. Never accepted at that twin: left open."""
    path, first, proposal = repeated(tmp_path, TWIN)
    outside_save(path, lambda ps: set_at(*REVISED)(set_at(4, "")(ps)))
    assert results(path) == {}
    assert verdict(path, proposal, watched=False) == OPEN


def test_fill_blanked_beside_a_repeated_block_with_its_twin_stays_pending_when_watched(tmp_path, feedback_dir):
    path, first, proposal = repeated(tmp_path, TWIN)
    outside_save(path, lambda ps: set_at(*REVISED)(set_at(4, "")(ps)))
    assert verdict(path, proposal, watched=True) == OPEN


def test_fill_kept_in_place_beside_a_repeated_block_is_accepted(tmp_path, feedback_dir):
    """The user edits only the last line: the fill stands exactly among its recorded neighbours at its own place,
    so the other block's blank line (exactly placed too) does not compete with it."""
    path, first, proposal = repeated(tmp_path, "")
    outside_save(path, set_at(18, "End of form."))
    assert results(path) == {proposal["id"]: ("accepted", TWIN, first + 4)}
    assert verdict(path, proposal, watched=True) == (first + 4, ("accepted", TWIN))


# ---------------------------------------------------------------- a copy of the block added above the fill

SINGLE = BLOCK + ["Part two", "End"]  # one block: body position 4 is blank
SEAM = ["Copied block", "Original follows"]  # the copy of the block plus these: the original fill moves 4 -> 13


def copied(tmp_path: Path, copy_line: str, original_line: str) -> tuple[Path, int, dict]:
    """The model fills the only block's blank line; the user then copies the block above it (its line reading
    ``copy_line``), so the copy stands at the recorded place and the original, past NEAR, reads ``original_line``.
    Both stand exactly among the recorded neighbours."""
    path = build(tmp_path / "copied.docx", "Form A", "Reviewer A", SINGLE)
    first = min(fbw.Snapshot(path).indices)
    (proposal,) = fill_all(path, [first + 4], [TWIN])
    copy = BLOCK[:3] + [copy_line] + BLOCK[4:] + SEAM
    outside_save(path, lambda ps: insert_at(1, *copy)(set_at(4, original_line)(ps)))
    return path, first, proposal


def test_fill_records_how_many_places_stand_among_its_neighbours(tmp_path, feedback_dir):
    path = build(tmp_path / "single.docx", "Form A", "Reviewer A", SINGLE)
    (single,) = fill_all(path, [min(fbw.Snapshot(path).indices) + 4], [TWIN])
    assert single["nbt_n"] == 1
    _, _, two = repeated(tmp_path, "")
    assert two["nbt_n"] == 2


def test_fill_kept_below_a_copied_block_with_a_blank_line_stays_pending(tmp_path, feedback_dir):
    """Open path. The answer is kept at the original (now 13); the copy at the recorded place is blank. One more
    place stands among the recorded neighbours than when the fill was recorded: which is the fill is not known."""
    path, first, proposal = copied(tmp_path, "", TWIN)
    assert results(path) == {}
    assert verdict(path, proposal, watched=False) == OPEN


def test_fill_kept_below_a_copied_block_with_a_blank_line_stays_pending_when_watched(tmp_path, feedback_dir):
    path, first, proposal = copied(tmp_path, "", TWIN)
    assert fbw.reconcile_props([], fbw.Snapshot(path), via="user", sha=None, watched=[proposal]) == []
    assert verdict(path, proposal, watched=True) == OPEN


def test_fill_blanked_below_a_copied_block_with_its_answer_stays_pending(tmp_path, feedback_dir):
    """Open path. The user blanks the original (now 13); the copy at the recorded place holds the same answer.
    Never accepted at the copy: left open."""
    path, first, proposal = copied(tmp_path, TWIN, "")
    assert results(path) == {}
    assert verdict(path, proposal, watched=False) == OPEN


def test_fill_blanked_below_a_copied_block_with_its_answer_stays_pending_when_watched(tmp_path, feedback_dir):
    path, first, proposal = copied(tmp_path, TWIN, "")
    assert verdict(path, proposal, watched=True) == OPEN


def test_fill_recorded_without_its_count_is_decided_near_beside_a_copied_block(tmp_path, feedback_dir):
    """A proposal recorded before ``nbt_n`` existed cannot tell a copy from a block that was already there: the
    paragraph standing exactly among its neighbours at the recorded place decides, as for a repeated block."""
    path, first, proposal = copied(tmp_path, "", TWIN)
    older = {k: v for k, v in proposal.items() if k != "nbt_n"}
    for watched in (False, True):
        assert verdict(path, older, watched=watched) == (first + 4, ("rejected", ""))


# ---------------------------------------------------------------- older proposals keep the plain order

def test_older_fill_without_neighbours_blanked_beside_its_twin_is_rejected(tmp_path, feedback_dir):
    """A proposal recorded before neighbours were stored (no ``nbt``) has no place to check, so it keeps the
    plain candidate order: blanked in place, it is rejected there, never accepted at an identical answer nearby
    (open and watched paths alike)."""
    path, answers = worksheet(tmp_path)
    texts = list(ANSWERS)
    texts[3] = texts[4] = "Not covered in this chapter."
    proposals = fill_all(path, answers, texts)
    older = {k: v for k, v in proposals[3].items() if k != "nbt"}
    outside_save(path, set_at(1 + 3 * 3 + 2, ""))  # the user blanks the first of the twins
    snap = fbw.Snapshot(path)
    for decided in (fbw.reconcile_props([older], snap, via="user", sha=None),
                    fbw.reconcile_props([], snap, via="user", sha=None, watched=[older])):
        assert [(e["outcome"], e["kept"], e["para"]) for e in decided] == [("rejected", "", answers[3])]
