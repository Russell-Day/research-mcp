"""Edit feedback when the model edits its own pending insertion in place (plan 831 S2, fix round 2).

A model edit that empties its own pending insertion is a model re-edit, not a user outcome: with the paired
deletion of that tracked replacement still there, the chain gets a new round proposing the deletion alone,
pending until the user accepts or rejects it; with nothing of the proposal left (a pure insertion), the model
withdrew it (``removed`` via model). ``undo_document`` of such an edit returns the chain to the earlier round, and
undoing that undo makes the later round current again. Everything runs through the real tools
(``edit_document``, ``undo_document``, ``edit_feedback``); the user resolves outside the MCP.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from test_edit_feedback import (  # noqa: F401 - feedback_dir is a fixture
    ASSISTANT, PARAGRAPHS, _swap, accepted_text, build_doc, call, events, feedback, feedback_dir, model_edit,
    records, sha, tracked, user_resolve,
)

from research_mcp import edit_feedback as fb
from research_mcp import edit_feedback_word as fbw
from research_mcp.checkpoints import file_sha256
from research_mcp.revisions import load, resolve_revisions, text_edit

SENTENCE = PARAGRAPHS[0]  # "The model was trained on a small dataset of clinical notes." (paragraph 2)


def outcomes(path: Path) -> list[tuple[str, str]]:
    return [(e["outcome"], e["via"]) for e in events(path) if e["event"] == "outcome"]


def chain(path: Path) -> dict:
    return records(path)[2]


def net(record: dict) -> tuple:
    """rounds, pending, outcome, via, original, last proposed and kept, whitespace collapsed as the view shows."""
    norm = lambda text: None if text is None else " ".join(text.split())  # noqa: E731
    return (record["rounds"], record["pending"], record["outcome"], record["via"], norm(record["original"]),
            norm(record["last_proposed"]), norm(record["kept"]))


def own_ins_ids(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
    return re.findall(rf'<w:ins w:id="(\d+)" w:author="{ASSISTANT}"', xml)


def emptied(path: Path) -> dict:
    """small -> modest (tracked), then the model empties its own pending "modest" in place."""
    model_edit(path, tracked(2, "small", "modest"))
    out = model_edit(path, tracked(2, "modest", ""))
    assert out["operations"][0]["details"]["revised_in_place"][0]["removed"] is True
    return out


# ---------------------------------------------------------------- capture alone

def test_capture_emptied_own_insertion_proposes_its_deletion(tmp_path):
    a = build_doc(tmp_path / "a.docx")
    b, c = tmp_path / "b.docx", tmp_path / "c.docx"
    text_edit(a, b, op="tracked_replace_text", paragraph_index=2, tracked=True, author=ASSISTANT, find="small",
              replace="modest")
    text_edit(b, c, op="tracked_replace_text", paragraph_index=2, tracked=True, author=ASSISTANT, find="modest",
              replace="")
    proposals, lost, _ = fbw.capture(b, c)
    assert [(p["para"], p["before"].strip(), p["after"].strip(), p["tracked"]) for p in proposals] == \
        [(2, "small", "", True)]
    assert [(r["type"], r["author"], r["text"]) for r in proposals[0]["revs"]] == [("del", ASSISTANT, "small")]
    assert [key[0] for key in lost] == ["ins"]


# ---------------------------------------------------------------- the emptied replacement: round 2

def test_emptied_own_replacement_is_round_two_pending(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    out = emptied(path)
    assert out.get("edit_feedback") is None
    assert outcomes(path) == [("superseded", "model")]                    # no "reworded" via model
    assert net(chain(path)) == (2, True, None, None, "small", "", "")
    assert feedback(path)["pending"] == 1
    assert feedback(path).get("edit_feedback") is None


def test_emptied_own_replacement_then_user_accepts(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied(path)
    user_resolve(path, "accept", 2)
    assert accepted_text(path, 2) == SENTENCE.replace("small", "")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "accepted", "user", "small", "", "")


def test_emptied_own_replacement_then_user_rejects(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied(path)
    user_resolve(path, "reject", 2)
    assert accepted_text(path, 2) == SENTENCE
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "rejected", "user", "small", "", "small")


def test_emptied_own_replacement_then_model_rejects(tmp_path, feedback_dir):
    """The reviewer's sequence: the model rejects the deletion that is left (MCP reject, via model)."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied(path)
    model_edit(path, {"type": "reject_revisions", "author": ASSISTANT, "paragraph_index": 2})
    assert accepted_text(path, 2) == SENTENCE
    assert net(chain(path)) == (2, False, "rejected", "model", "small", "", "small")


def test_partly_emptied_own_insertion_stays_a_two_round_chain(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    out = model_edit(path, tracked(2, "modest", "mod"))
    assert out["operations"][0]["details"]["revised_in_place"][0]["removed"] is False
    assert net(chain(path)) == (2, True, None, None, "small", "mod", "mod")
    user_resolve(path, "accept", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "accepted", "user", "small", "mod", "mod")


def test_model_rejecting_only_its_insertion_also_leaves_the_deletion_pending(tmp_path, feedback_dir):
    """The same document reached by an MCP reject of the insertion half alone: the deletion is still the model's
    pending proposal, so the chain stays open on it (it used to close as reworded via model)."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    model_edit(path, {"type": "reject_revisions", "ids": own_ins_ids(path)})
    assert net(chain(path)) == (2, True, None, None, "small", "", "")
    user_resolve(path, "reject", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "rejected", "user", "small", "", "small")


# ---------------------------------------------------------------- a pure insertion withdrawn

def test_emptied_own_pure_insertion_is_removed_via_model(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small dataset", "small curated dataset"))
    assert outcomes(path) == []
    out = model_edit(path, {"type": "tracked_replace_paragraph", "paragraph_index": 2, "text": SENTENCE,
                            "author": ASSISTANT})
    assert out["operations"][0]["details"]["revised_in_place"][0]["removed"] is True
    assert accepted_text(path, 2) == SENTENCE and own_ins_ids(path) == []
    assert out.get("edit_feedback") is None
    assert outcomes(path) == [("removed", "model")]
    assert net(chain(path))[:4] == (1, False, "removed", "model")
    assert feedback(path)["pending"] == 0


def test_op_ids_never_classify_the_write(tmp_path, feedback_dir):
    """Edit feedback classifies a write by its applied op types, never by op ids: a withdrawal whose op_id reads
    like another op type (here reject_revisions, also past character 200) is still removed via model."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small dataset", "small curated dataset"))
    for op_id in ("reject_revisions", "x" * 210 + " reject_revisions"):
        assert fbw._RESOLVE_OP.search(op_id)
    op = {"op_id": "x" * 210 + " reject_revisions", "type": "tracked_replace_paragraph", "paragraph_index": 2,
          "text": SENTENCE, "author": ASSISTANT}
    out = call("edit_document", source=str(path), source_sha256=sha(path), output=str(path), operations=[op])
    assert out["status"] == "complete", out
    assert outcomes(path) == [("removed", "model")]


def test_model_reject_of_pure_insertion_stays_rejected_via_model(tmp_path, feedback_dir):
    """Same document as the withdrawal above, reached by an MCP reject: that stays a rejection via model."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small dataset", "small curated dataset"))
    model_edit(path, {"type": "reject_revisions", "author": ASSISTANT, "paragraph_index": 2})
    assert accepted_text(path, 2) == SENTENCE
    assert outcomes(path) == [("rejected", "model")]


def test_withdrawal_batched_with_a_reject_reads_as_model_rejection(tmp_path, feedback_dir):
    """One write that both empties an own pure insertion and rejects another revision cannot say which op removed
    the insertion: it falls back to a rejection via model (never a user outcome); the batch's other reject is
    recorded as usual."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small dataset", "small curated dataset"), tracked(3, "good", "strong"))
    model_edit(path, {"type": "tracked_replace_paragraph", "paragraph_index": 2, "text": SENTENCE,
                      "author": ASSISTANT},
               {"type": "reject_revisions", "author": ASSISTANT, "paragraph_index": 3})
    assert accepted_text(path, 2) == SENTENCE and accepted_text(path, 3) == PARAGRAPHS[1]
    assert sorted(outcomes(path)) == [("rejected", "model"), ("rejected", "model")]


# ---------------------------------------------------------------- undo and redo

def test_undo_of_the_emptying_edit_returns_the_chain_to_round_one(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    after_first = file_sha256(path)
    model_edit(path, tracked(2, "modest", ""))
    after_second = file_sha256(path)

    out = call("undo_document", path=str(path))
    assert file_sha256(path) == after_first
    assert out.get("edit_feedback") is None
    assert outcomes(path) == [("superseded", "model"), ("undone", "model"), ("reopened", "model")]
    assert net(chain(path)) == (1, True, None, None, "small", "modest", "modest")

    call("undo_document", path=str(path))                                    # undo the undo
    assert file_sha256(path) == after_second
    assert net(chain(path)) == (2, True, None, None, "small", "", "")

    user_resolve(path, "reject", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "rejected", "user", "small", "", "small")


def test_undo_to_round_one_then_user_accepts_round_one(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied(path)
    call("undo_document", path=str(path))
    user_resolve(path, "accept", 2)
    assert accepted_text(path, 2) == SENTENCE.replace("small", "modest")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (1, False, "accepted", "user", "small", "modest", "modest")


def test_undo_of_an_in_place_rewording_returns_to_round_one(tmp_path, feedback_dir):
    """The same rule for any later round: modest -> moderate in place, undone, is round 1 pending again; a new
    edit then supersedes round 1 and counts as round 2."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    model_edit(path, tracked(2, "modest", "moderate"))
    call("undo_document", path=str(path))
    assert net(chain(path)) == (1, True, None, None, "small", "modest", "modest")
    model_edit(path, tracked(2, "modest", "limited"))
    assert net(chain(path)) == (2, True, None, None, "small", "limited", "limited")
    user_resolve(path, "accept", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "accepted", "user", "small", "limited", "limited")


def test_undo_of_a_first_round_is_still_a_model_rejection(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    call("undo_document", path=str(path))
    assert outcomes(path) == [("rejected", "model")]
    assert sha(path) and net(chain(path))[:4] == (1, False, "rejected", "model")
    call("undo_document", path=str(path))                                    # undo the undo: pending again
    assert net(chain(path)) == (1, True, None, None, "small", "modest", "modest")


# ---------------------------------------------------------------- several changes in one tracked edit (fix round 3)
# One tracked edit writes every change of a paragraph with the same author and date, so the deletion left of an
# emptied insertion is the one beside it in the XML, never just the first deletion by that author and date.

TWO = SENTENCE.replace("small", "modest").replace("clinical", "medical")


def by_original(path: Path) -> dict[str, dict]:
    return {" ".join(r["original"].split()): r for r in fb.chains(events(path))}


def two_replacements(path: Path, empty: str) -> None:
    """small -> modest and clinical -> medical in one tracked edit, then the model empties ``empty``."""
    model_edit(path, {"type": "tracked_replace_paragraph", "paragraph_index": 2, "text": TWO, "author": ASSISTANT})
    assert len(fb.chains(events(path))) == 2
    model_edit(path, tracked(2, empty, ""))


def user_resolves_one(path: Path, mode: str, kind: str, text: str) -> None:
    _, doc = load(path)
    rid = next(r.id for r in doc.revisions if r.type == kind and r.text == text)
    staged = path.with_name(f"user-{mode}-{rid}.docx")
    resolve_revisions(path, staged, mode=mode, ids=[rid])
    _swap(path, staged)


def test_capture_pairs_the_emptied_insertion_with_its_own_deletion(tmp_path):
    a = build_doc(tmp_path / "a.docx")
    b, c = tmp_path / "b.docx", tmp_path / "c.docx"
    text_edit(a, b, op="tracked_replace_paragraph", paragraph_index=2, tracked=True, author=ASSISTANT, text=TWO)
    text_edit(b, c, op="tracked_replace_text", paragraph_index=2, tracked=True, author=ASSISTANT, find="medical",
              replace="")
    proposals, _, _ = fbw.capture(b, c)
    assert [(p["before"], p["after"], p["anchor_l"][-3:], p["anchor_r"][:6]) for p in proposals] == \
        [("clinical", "", "of ", " notes")]
    assert [(r["type"], r["text"]) for r in proposals[0]["revs"]] == [("del", "clinical")]


def test_second_of_two_replacements_emptied_then_user_accepts_its_deletion(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    two_replacements(path, "medical")
    chains = by_original(path)
    assert net(chains["clinical"]) == (2, True, None, None, "clinical", "", "")
    assert net(chains["small"]) == (1, True, None, None, "small", "modest", "modest")
    user_resolves_one(path, "accept", "del", "clinical")
    assert feedback(path)["edit_feedback"] == \
        "Since your last edit: 1 accepted (1 pending). Call edit_feedback for details."
    chains = by_original(path)
    assert net(chains["clinical"]) == (2, False, "accepted", "user", "clinical", "", "")
    assert net(chains["small"]) == (1, True, None, None, "small", "modest", "modest")


def test_second_of_two_replacements_emptied_then_user_rejects_all(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    two_replacements(path, "medical")
    user_resolve(path, "reject", 2)
    assert accepted_text(path, 2) == SENTENCE
    assert feedback(path)["edit_feedback"] == "Since your last edit: 2 rejected. Call edit_feedback for details."
    chains = by_original(path)
    assert chains["clinical"]["kept"] == "clinical"                          # exact, not " clinical "
    assert net(chains["clinical"]) == (2, False, "rejected", "user", "clinical", "", "clinical")
    assert net(chains["small"]) == (1, False, "rejected", "user", "small", "modest", "small")


def test_first_of_two_replacements_emptied_then_user_accepts_then_rejects_the_rest(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    two_replacements(path, "modest")
    chains = by_original(path)
    assert net(chains["small"]) == (2, True, None, None, "small", "", "")
    assert net(chains["clinical"]) == (1, True, None, None, "clinical", "medical", "medical")
    user_resolves_one(path, "accept", "del", "small")
    feedback(path)
    assert net(by_original(path)["small"]) == (2, False, "accepted", "user", "small", "", "")
    assert by_original(path)["clinical"]["pending"]
    user_resolve(path, "reject", 2)
    assert accepted_text(path, 2) == SENTENCE.replace("small", "")
    feedback(path)
    chains = by_original(path)
    assert net(chains["small"]) == (2, False, "accepted", "user", "small", "", "")
    assert net(chains["clinical"]) == (1, False, "rejected", "user", "clinical", "medical", "clinical")


def test_pure_insertion_beside_an_unrelated_replacement_is_withdrawn(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    new = SENTENCE.replace("small", "modest").replace("clinical", "vetted clinical")
    model_edit(path, {"type": "tracked_replace_paragraph", "paragraph_index": 2, "text": new, "author": ASSISTANT})
    model_edit(path, tracked(2, "vetted ", ""))
    assert accepted_text(path, 2) == SENTENCE.replace("small", "modest")
    chains = by_original(path)
    assert net(chains[""])[:4] == (1, False, "removed", "model")
    assert net(chains["small"]) == (1, True, None, None, "small", "modest", "modest")
    user_resolve(path, "reject", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(by_original(path)["small"]) == (1, False, "rejected", "user", "small", "modest", "small")


# ---------------------------------------------------------------- undo of a model decision (fix round 3)

def test_undo_of_a_pure_insertion_withdrawal_reopens_it(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small dataset", "small curated dataset"))
    model_edit(path, tracked(2, "curated ", ""))
    assert outcomes(path) == [("removed", "model")]
    call("undo_document", path=str(path))
    assert outcomes(path)[-1] == ("reopened", "model")
    assert net(chain(path))[:4] == (1, True, None, None)
    assert chain(path)["last_proposed"].strip() == "curated"
    call("undo_document", path=str(path))                                    # undo the undo: withdrawn again
    assert net(chain(path))[:4] == (1, False, "rejected", "model")
    call("undo_document", path=str(path))                                    # and back: pending, user accepts
    assert net(chain(path))[:2] == (1, True)
    user_resolve(path, "accept", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path))[:4] == (1, False, "accepted", "user")


def test_undo_of_a_model_reject_reopens_the_round(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied(path)
    model_edit(path, {"type": "reject_revisions", "author": ASSISTANT, "paragraph_index": 2})
    assert net(chain(path)) == (2, False, "rejected", "model", "small", "", "small")
    call("undo_document", path=str(path))
    assert net(chain(path)) == (2, True, None, None, "small", "", "")
    user_resolve(path, "accept", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "accepted", "user", "small", "", "")


def test_undo_of_a_model_reject_of_a_first_round_reopens_it(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    model_edit(path, {"type": "reject_revisions", "author": ASSISTANT, "paragraph_index": 2})
    assert net(chain(path))[:4] == (1, False, "rejected", "model")
    call("undo_document", path=str(path))
    assert net(chain(path)) == (1, True, None, None, "small", "modest", "modest")


def test_switching_undo_branches_counts_only_the_current_branch(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    model_edit(path, tracked(2, "modest", "moderate"))
    undo = call("undo_document", path=str(path))
    model_edit(path, tracked(2, "modest", "limited"))
    assert net(chain(path)) == (2, True, None, None, "small", "limited", "limited")
    call("undo_document", path=str(path), undo_token=undo["undo_token"])     # back to the "moderate" branch
    assert net(chain(path)) == (2, True, None, None, "small", "moderate", "moderate")
    call("undo_document", path=str(path))                                    # and to the "limited" branch
    assert net(chain(path)) == (2, True, None, None, "small", "limited", "limited")
    user_resolve(path, "reject", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "rejected", "user", "small", "limited", "small")



def test_restoring_a_deep_old_branch_counts_its_rounds(tmp_path, feedback_dir):
    """Three rounds, back to round 1, a new round 2 there, then the old round 3 again: all three of its rounds
    count (round 2 of the old branch, undone by the first switch, comes back with it)."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    model_edit(path, tracked(2, "small", "modest"))
    model_edit(path, tracked(2, "modest", "moderate"))
    model_edit(path, tracked(2, "moderate", "limited"))
    second = next(e for e in events(path) if e["event"] == "proposal" and e["after"] == "moderate")
    undo = call("undo_document", path=str(path), undo_token=second["undo_token"])   # back to round 1
    assert net(chain(path))[:2] == (1, True)
    model_edit(path, tracked(2, "modest", "tiny"))
    assert net(chain(path)) == (2, True, None, None, "small", "tiny", "tiny")
    call("undo_document", path=str(path), undo_token=undo["undo_token"])             # the old round 3 again
    assert net(chain(path)) == (3, True, None, None, "small", "limited", "limited")
    user_resolve(path, "accept", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (3, False, "accepted", "user", "small", "limited", "limited")


# ---------------------------------------------------------------- a word stored in several runs (fix round 4)

RUN_BREAK = '</w:t></w:r><w:r><w:t xml:space="preserve">'


def rewrite(path: Path, old: str, new: str) -> None:
    """Rewrite the document XML as Word may store it (one word in several identically formatted runs)."""
    staged = path.with_name("rewritten.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w") as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                assert xml.count(old) == 1
                data = xml.replace(old, new).encode("utf-8")
            dst.writestr(info, data)
    _swap(path, staged)


def split_clinical(path: Path, *parts: str) -> None:
    rewrite(path, "clinical notes.</w:t></w:r>", RUN_BREAK.join(parts[:-1] + (parts[-1] + " notes.",)) + "</w:t></w:r>")
    assert accepted_text(path, 2) == SENTENCE


def emptied_split(path: Path, *parts: str) -> None:
    """clinical (stored as ``parts``) -> medical, one tracked deletion per run; then the model empties "medical"."""
    split_clinical(path, *parts)
    model_edit(path, tracked(2, "clinical", "medical"))
    assert [r.text for r in load(path)[1].revisions if r.type == "del"] == list(parts)
    model_edit(path, tracked(2, "medical", ""))
    proposal = [e for e in events(path) if e["event"] == "proposal"][-1]
    assert proposal["before"] == "clinical" and proposal["after"] == ""
    assert [(r["type"], r["text"], r.get("part")) for r in proposal["revs"]] == [("del", p, True) for p in parts]
    assert net(chain(path)) == (2, True, None, None, "clinical", "", "")


def test_split_deletion_user_rejects_all(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied_split(path, "clini", "cal")
    user_resolve(path, "reject", 2)
    assert accepted_text(path, 2) == SENTENCE
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "rejected", "user", "clinical", "", "clinical")


def test_split_deletion_user_accepts_all(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied_split(path, "clini", "cal")
    user_resolve(path, "accept", 2)
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "accepted", "user", "clinical", "", "")


def test_split_deletion_pending_until_every_fragment_is_accepted(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied_split(path, "clini", "cal")
    user_resolves_one(path, "accept", "del", "cal")                         # "clini" is still pending
    assert feedback(path).get("edit_feedback") is None
    assert net(chain(path))[:4] == (2, True, None, None)
    user_resolves_one(path, "accept", "del", "clini")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "accepted", "user", "clinical", "", "")


def test_split_deletion_one_fragment_rejected_is_decided_when_all_are(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied_split(path, "clini", "cal")
    user_resolves_one(path, "reject", "del", "cal")                         # reads "cal"; "clini" still pending
    assert feedback(path).get("edit_feedback") is None
    assert net(chain(path))[:4] == (2, True, None, None)
    user_resolves_one(path, "accept", "del", "clini")
    assert accepted_text(path, 2) == SENTENCE.replace("clinical", "cal")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 reworded. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "reworded", "user", "clinical", "", "cal")


def test_split_deletion_in_three_runs(tmp_path, feedback_dir):
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    emptied_split(path, "cli", "ni", "cal")
    user_resolves_one(path, "reject", "del", "ni")
    user_resolves_one(path, "reject", "del", "cli")
    assert net(chain(path))[:4] == (2, True, None, None)                    # "cal" is still pending
    user_resolves_one(path, "reject", "del", "cal")
    assert accepted_text(path, 2) == SENTENCE
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, "rejected", "user", "clinical", "", "clinical")


@pytest.mark.parametrize(("first", "second", "outcome", "kept"), [
    ("reject", "reject", "rejected", "murmur"),
    ("reject", "accept", "reworded", "mur"),
    ("accept", "accept", "accepted", ""),
])
def test_identical_fragments_stay_pending_until_both_are_resolved(tmp_path, feedback_dir, first, second, outcome,
                                                                  kept):
    """A word stored as two identical runs ("mur|mur"): resolving one fragment leaves the other pending."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    rewrite(path, "clinical notes.</w:t></w:r>", "mur" + RUN_BREAK + "mur notes.</w:t></w:r>")
    sentence = SENTENCE.replace("clinical", "murmur")
    assert accepted_text(path, 2) == sentence
    model_edit(path, tracked(2, "murmur", "medical"))
    assert [r.text for r in load(path)[1].revisions if r.type == "del"] == ["mur", "mur"]
    model_edit(path, tracked(2, "medical", ""))
    assert net(chain(path)) == (2, True, None, None, "murmur", "", "")
    user_resolves_one(path, first, "del", "mur")
    assert feedback(path).get("edit_feedback") is None                      # one "mur" deletion is still pending
    assert net(chain(path))[:4] == (2, True, None, None)
    user_resolves_one(path, second, "del", "mur")
    assert feedback(path)["edit_feedback"] == f"Since your last edit: 1 {outcome}. Call edit_feedback for details."
    assert net(chain(path)) == (2, False, outcome, "user", "murmur", "", kept)


def test_insertion_between_two_split_deletions_keeps_them_apart(tmp_path, feedback_dir):
    """``clini|cal`` -> medical and `` no|tes`` -> `` records`` stored with nothing between the second deletion and
    the first insertion: emptying "medical" proposes "clinical" alone, not "clinical notes"."""
    path = build_doc(tmp_path / "synthetic_manuscript.docx")
    when = f'w:author="{ASSISTANT}" w:date="2026-11-01T10:00:00Z"'

    def dele(n: int, text: str) -> str:
        return f'<w:del w:id="{n}" {when}><w:r><w:delText xml:space="preserve">{text}</w:delText></w:r></w:del>'

    def ins(n: int, text: str) -> str:
        return f'<w:ins w:id="{n}" {when}><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:ins>'

    rewrite(path, "clinical notes.</w:t></w:r>",
            "</w:t></w:r>" + dele(21, "clini") + dele(22, "cal") + ins(23, "medical") + dele(24, " no")
            + dele(25, "tes") + ins(26, " records") + '<w:r><w:t xml:space="preserve">.</w:t></w:r>')
    assert accepted_text(path, 2) == SENTENCE.replace("clinical notes", "medical records")
    model_edit(path, tracked(2, "medical", ""))
    proposal = next(e for e in events(path) if e["event"] == "proposal")
    assert proposal["before"] == "clinical" and proposal["after"] == ""
    assert [r["text"] for r in proposal["revs"]] == ["clini", "cal"]
    assert net(chain(path))[:2] == (1, True)
    user_resolves_one(path, "reject", "del", "clini")
    user_resolves_one(path, "reject", "del", "cal")                         # " no|tes" -> " records" still pending
    assert accepted_text(path, 2) == SENTENCE.replace("notes", "records")
    assert feedback(path)["edit_feedback"] == "Since your last edit: 1 rejected. Call edit_feedback for details."
    assert net(chain(path)) == (1, False, "rejected", "user", "clinical", "", "clinical")
