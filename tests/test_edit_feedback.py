"""Edit feedback (plan 829 S1): what the user did with the model's edits.

Every test builds a synthetic manuscript by "A. Rivera" with one earlier revision of hers, so every edit runs
as direct XML (no Word). The model's edits go through the real ``edit_document`` tool (author "Assistant");
the user's actions are made outside the MCP (library accept/reject, plain text edits, XML deletion), as
Word would leave the file, and the next tool call sees them.
"""

from __future__ import annotations

import asyncio
import os
import re
import sys
import zipfile
from pathlib import Path

import pytest

from research_mcp import edit_feedback as fb
from research_mcp import edit_feedback_word as fbw
from research_mcp.checkpoints import file_sha256
from research_mcp.revisions import load, resolve_revisions, text_edit
from research_mcp.server import server

ASSISTANT = "Assistant"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

PARAGRAPHS = [
    "The model was trained on a small dataset of clinical notes.",          # 2 accept
    "Results were good across all sites.",                                    # 3 reject
    "We used a learning rate of 0.001 for training.",                         # 4 reword
    "Further work will examine external validity.",                           # 5 pending
    "Annotators reviewed every note twice before release.",                   # 6 untracked, reverted
    "The pipeline runs nightly on the shared servers.",                       # 7 untracked, kept
    "Errors were logged for manual review by the team.",                      # 8 untracked, reworded
    "This paragraph is removed by the author later.",                         # 9 deleted paragraph
    "Limitations include the single-center design of the study.",             # 10 five-round chain
    "Data sharing is restricted by institutional policy.",                    # 11 MCP accept
]


def _t(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def build_doc(path: Path, paragraphs=PARAGRAPHS) -> Path:
    """A. Rivera's synthetic manuscript: paragraph 1 holds her own tracked insertion."""
    first = ('<w:p>' + _t("Methods for the ")
             + '<w:ins w:id="1" w:author="A. Rivera" w:date="2026-10-01T09:00:00Z">' + _t("pilot ") + "</w:ins>"
             + _t("cohort study by A. Rivera.") + "</w:p>")
    body = first + "".join(f"<w:p>{_t(text)}</w:p>" for text in paragraphs)
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


# ---------------------------------------------------------------- tools, as the model calls them

def call(name: str, **arguments) -> dict:
    result = asyncio.run(server.call_tool(name, arguments))
    assert result.is_error is False, result
    return result.structured_content


def sha(path: Path) -> str:
    return file_sha256(path).upper()


def model_edit(path: Path, *ops: dict) -> dict:
    operations = [dict(op, op_id=f"op-{n}") for n, op in enumerate(ops, 1)]
    out = call("edit_document", source=str(path), source_sha256=sha(path), output=str(path), operations=operations)
    assert out["status"] == "complete", out
    return out


def tracked(paragraph: int, find: str, replace: str) -> dict:
    return {"type": "tracked_replace_text", "paragraph_index": paragraph, "find": find, "replace": replace,
            "author": ASSISTANT}


def plain(paragraph: int, find: str, replace: str) -> dict:
    return {"type": "replace_text", "paragraph_index": paragraph, "find": find, "replace": replace}


# ---------------------------------------------------------------- the user, outside the MCP

def _swap(path: Path, staged: Path) -> None:
    os.replace(staged, path)


def user_resolve(path: Path, mode: str, paragraph: int) -> None:
    staged = path.with_name(f"user-{mode}-{paragraph}.docx")
    resolve_revisions(path, staged, mode=mode, author=ASSISTANT, paragraph_index=paragraph)
    _swap(path, staged)


def user_type(path: Path, paragraph: int, find: str, replace: str) -> None:
    """An untracked edit typed by the user."""
    staged = path.with_name(f"user-type-{paragraph}.docx")
    text_edit(path, staged, op="replace_text", paragraph_index=paragraph, tracked=False, find=find, replace=replace)
    _swap(path, staged)


def user_delete_paragraph(path: Path, contains: str) -> None:
    staged = path.with_name("user-delete.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                paragraphs = re.findall(r"<w:p>.*?</w:p>", xml)
                victim = next(p for p in paragraphs if contains in re.sub(r"<[^>]+>", "", p))
                data = xml.replace(victim, "", 1).encode("utf-8")
            dst.writestr(info, data)
    _swap(path, staged)


def accepted_text(path: Path, paragraph: int) -> str:
    from research_mcp.revisions import ParagraphModel

    _, doc = load(path)
    return ParagraphModel(doc.xml, doc.para_nodes[paragraph]).text


@pytest.fixture()
def feedback_dir(tmp_path, monkeypatch):
    folder = tmp_path / "feedback"
    monkeypatch.setenv("RESEARCH_MCP_FEEDBACK_DIR", str(folder))
    return folder


@pytest.fixture()
def doc(tmp_path, feedback_dir):
    return build_doc(tmp_path / "rivera_manuscript.docx")


def events(path: Path) -> list[dict]:
    return fbw.doc_log(path).events()


def feedback(path: Path, **kw) -> dict:
    return call("edit_feedback", path=str(path), **kw)


def records(path: Path) -> dict[int, dict]:
    return {r["para"]: r for r in fb.chains(events(path))}


# ---------------------------------------------------------------- the cases

def test_fixture_layout(doc):
    assert accepted_text(doc, 1) == "Methods for the pilot cohort study by A. Rivera."
    assert accepted_text(doc, 2) == PARAGRAPHS[0]
    assert accepted_text(doc, 11) == PARAGRAPHS[-1]


def test_tracked_edit_accepted_in_word(doc):
    model_edit(doc, tracked(2, "small", "modest"))
    user_resolve(doc, "accept", 2)
    line = call("inspect_document", path=str(doc), first_paragraph=1, last_paragraph=3)["edit_feedback"]
    assert line == "Since your last edit: 1 accepted. Call edit_feedback for details."
    record = records(doc)[2]
    assert (record["outcome"], record["via"], record["tracked"], record["rounds"]) == ("accepted", "user", True, 1)
    assert (record["original"], record["proposed"], record["kept"]) == ("small", "modest", "modest")
    assert record["author"] == ASSISTANT


def test_tracked_edit_rejected_in_word(doc):
    model_edit(doc, tracked(3, "good", "consistent"))
    user_resolve(doc, "reject", 3)
    line = call("audit_document", path=str(doc))["edit_feedback"]
    assert line == "Since your last edit: 1 rejected. Call edit_feedback for details."
    record = records(doc)[3]
    assert (record["outcome"], record["kept"]) == ("rejected", "good")


def test_tracked_edit_reworded_by_user(doc):
    model_edit(doc, tracked(4, "used", "selected"))
    user_resolve(doc, "accept", 4)
    user_type(doc, 4, "selected", "chose")
    out = feedback(doc)
    assert out["edit_feedback"] == "Since your last edit: 1 reworded. Call edit_feedback for details."
    assert len(out["items"]) == 1
    assert re.fullmatch(r'\[c-[0-9a-f]{8}\] ¶4 reworded \(tracked\): "…We " "used" → proposed "selected" '
                        r'→ kept "chose"', out["items"][0]), out["items"][0]


def test_tracked_edit_still_pending_is_counted_not_listed(doc):
    model_edit(doc, tracked(5, "examine", "assess"), tracked(2, "small", "modest"))
    user_resolve(doc, "accept", 2)
    out = feedback(doc)
    assert out["edit_feedback"] == "Since your last edit: 1 accepted (1 pending). Call edit_feedback for details."
    assert out["pending"] == 1
    assert len(out["items"]) == 1 and "examine" not in out["items"][0] and "accepted" in out["items"][0]
    assert records(doc)[5]["pending"] is True


def test_untracked_replace_kept_reverted_reworded(doc):
    model_edit(doc, plain(6, "twice", "two times"), plain(7, "nightly", "every night"),
               plain(8, "manual review", "later review"))
    user_type(doc, 6, "two times", "twice")                # reverted
    user_type(doc, 8, "later review", "expert review")     # reworded
    out = feedback(doc)                                    # 7 kept as written
    assert out["edit_feedback"] == ("Since your last edit: 1 accepted, 1 rejected, 1 reworded. "
                                    "Call edit_feedback for details.")
    got = records(doc)
    assert (got[6]["outcome"], got[6]["tracked"], got[6]["kept"]) == ("rejected", False, "twice")
    assert (got[7]["outcome"], got[7]["kept"]) == ("accepted", "every night")
    assert (got[8]["outcome"], got[8]["kept"]) == ("reworded", "expert")
    assert out["counts"] == {"accepted": 1, "rejected": 1, "reworded": 1, "removed": 0}


@pytest.mark.parametrize("how", ["tracked", "untracked"])
def test_rewrite_of_span_and_its_anchor_is_reworded_not_removed(doc, how):
    """The user rewrote the model's words and the text right after them, so no exact span fits: the paragraph is
    still the edited one, and only the user's words where the span was are kept."""
    if how == "tracked":
        model_edit(doc, tracked(7, "nightly", "every night"))
        user_resolve(doc, "accept", 7)
    else:
        model_edit(doc, plain(7, "nightly", "every night"))
    user_type(doc, 7, "every night on the shared", "each evening on our")
    feedback(doc)
    got = records(doc)[7]
    assert (got["outcome"], got["kept"]) == ("reworded", "each evening")


def test_only_the_anchor_rewritten_is_accepted(doc):
    model_edit(doc, plain(7, "nightly", "every night"))
    user_type(doc, 7, " on the shared", " on our")
    feedback(doc)
    got = records(doc)[7]
    assert (got["outcome"], got["kept"]) == ("accepted", "every night")


@pytest.mark.parametrize("neighbour", ["unchanged", "edited too"])
def test_deleted_paragraph_with_a_parallel_neighbour_stays_removed(tmp_path, feedback_dir, neighbour):
    """A deleted paragraph is not 'reworded' into the similar sentence that takes its index; when the user edited
    that sentence as well, nothing is decided."""
    path = build_doc(tmp_path / "parallel.docx", [
        "In the north cohort, the effect of feeding on brood size was strong after rain.",
        "In the south cohort, the effect of shading on hive weight was mild before rain.",
    ])
    model_edit(path, plain(2, "strong", "large"))
    user_delete_paragraph(path, "north cohort")
    if neighbour == "edited too":
        user_type(path, 2, "mild", "modest")
    feedback(path)
    outs = [(e["outcome"], e["kept"]) for e in events(path) if e["event"] == "outcome"]
    if neighbour == "unchanged":
        assert outs == [("removed", None)]
    else:  # which paragraph survived is not known: nothing is decided
        assert outs == [], outs


def test_pending_tracked_revision_stays_pending_when_its_anchor_changes(doc):
    model_edit(doc, tracked(7, "nightly", "every night"))
    user_type(doc, 7, " on the shared", " on our")       # the revision itself is left for the user to decide
    out = feedback(doc)
    assert out["pending"] == 1 and sum(out["counts"].values()) == 0, out
    assert records(doc)[7]["pending"] is True



def test_pending_tracked_revision_beside_a_lookalike_paragraph_stays_pending(tmp_path, feedback_dir):
    """A neighbour that already reads like the changed paragraph does not make its pending revision 'removed'."""
    path = build_doc(tmp_path / "lookalike.docx", [
        "The pipeline runs every night on our servers.",
        "The pipeline runs nightly on the shared servers.",
    ])
    model_edit(path, tracked(3, "nightly", "every night"))
    user_type(path, 3, " on the shared", " on our")
    out = feedback(path)
    assert out["pending"] == 1 and sum(out["counts"].values()) == 0, out


def test_pending_tracked_revision_stays_pending_when_its_lookalike_neighbour_changes_too(tmp_path, feedback_dir):
    """The model's revision is still in the document: another paragraph that now reads like the edited one does
    not decide it."""
    path = build_doc(tmp_path / "lookalike2.docx", [
        "The pipeline runs every night on our servers.",
        "The pipeline runs nightly on the shared servers.",
    ])
    model_edit(path, tracked(3, "nightly", "every night"))
    user_type(path, 2, "on our servers", "on remote servers")
    user_type(path, 3, " on the shared", " on our")
    out = feedback(path)
    assert out["pending"] == 1 and sum(out["counts"].values()) == 0, out

@pytest.mark.parametrize("op, expected", [
    (plain(6, "reviewed", "carefully reviewed"), "carefully"),  # an insertion: the span is "carefully "
    (plain(6, " twice", ""), ""),                                # a deletion: the span is empty
])
def test_inserted_or_deleted_span_kept_while_its_anchor_changes_is_accepted(doc, op, expected):
    model_edit(doc, op)
    if expected:
        user_type(doc, 6, "every note", "each note")
    else:
        user_type(doc, 6, "before release", "prior to release")
    feedback(doc)
    got = records(doc)[6]
    assert (got["outcome"], (got["kept"] or "").strip()) == ("accepted", expected), got


def test_paragraph_deleted_is_removed_and_counted(doc):
    model_edit(doc, tracked(9, "removed", "deleted"))
    user_delete_paragraph(doc, "This paragraph is")
    out = feedback(doc)
    assert out["edit_feedback"] == "Since your last edit: 1 removed. Call edit_feedback for details."
    assert out["counts"]["removed"] == 1
    assert out["items"] == []                         # counted, not listed
    assert records(doc)[9]["outcome"] == "removed"
    asked = feedback(doc, outcome="removed")["items"]
    assert len(asked) == 1 and "removed" in asked[0]


def test_mcp_accept_and_reject_of_own_proposal_is_via_model(doc):
    model_edit(doc, tracked(11, "restricted", "limited"), tracked(3, "good", "consistent"))
    ids = [r.id for r in load(doc)[1].revisions if r.author == ASSISTANT and r.paragraph_index == 11]
    out = model_edit(doc, {"type": "accept_revisions", "ids": [int(i) for i in ids]},
                     {"type": "reject_revisions", "author": ASSISTANT, "paragraph_index": 3})
    assert out.get("edit_feedback") is None          # the model's own decisions are not news to the model
    got = records(doc)
    assert (got[11]["outcome"], got[11]["via"]) == ("accepted", "model")
    assert (got[3]["outcome"], got[3]["via"]) == ("rejected", "model")
    item = feedback(doc)["items"]
    assert any("via model" in line and "restricted" in line for line in item)


def back_and_forth(path: Path) -> None:
    """Five tracked rounds on one span. The library refuses to edit inside a pending insertion (even the
    model's own), so the user decides each round before the model proposes again: accept, reword, accept,
    accept, and a final rewording."""
    model_edit(path, tracked(10, "single-center", "one-site"))
    user_resolve(path, "accept", 10)
    model_edit(path, tracked(10, "one-site", "single-site"))
    user_resolve(path, "accept", 10)
    user_type(path, 10, "single-site", "monocentric")
    model_edit(path, tracked(10, "monocentric", "single-hospital"))
    user_resolve(path, "accept", 10)
    model_edit(path, tracked(10, "single-hospital", "one-hospital"))
    user_resolve(path, "accept", 10)
    model_edit(path, tracked(10, "one-hospital", "single-institution"))
    user_resolve(path, "accept", 10)
    user_type(path, 10, "single-institution", "single-centre")


def test_five_round_chain_collapses_to_one_record(doc):
    back_and_forth(doc)
    out = feedback(doc)
    assert out["edit_feedback"] == "Since your last edit: 1 reworded. Call edit_feedback for details."
    assert len(out["items"]) == 1                                  # one record, not five
    item = out["items"][0]
    assert re.fullmatch(r'\[c-[0-9a-f]{8}\] ¶10 reworded \(5 rounds, tracked\): "…Limitations include the " '
                        r'"single-center" → proposed "one-site" → kept "single-centre"', item), item
    proposals = [e for e in events(doc) if e["event"] == "proposal"]
    assert len(proposals) == 5 and len({e["chain"] for e in proposals}) == 1   # the log keeps every round
    assert [e["round"] for e in proposals] == [1, 2, 3, 4, 5]
    rounds = feedback(doc, record_id=proposals[0]["chain"], detail="rounds")
    assert len(rounds["rounds"]) == 5 and rounds["items"] == [item.replace("] ¶", "] rivera_manuscript.docx ¶")]
    assert rounds["rounds"][1].startswith("round 2 ") and "reworded" in rounds["rounds"][1]
    assert '"one-site" → "single-site" → kept "monocentric"' in rounds["rounds"][1]
    assert '"one-hospital" → "single-institution" → kept "single-centre"' in rounds["rounds"][4]
    assert all(len(line) <= fb.ITEM_MAX for line in rounds["rounds"])
    plain_view = feedback(doc, record_id=proposals[0]["chain"])            # detail="net": no rounds
    assert plain_view["rounds"] is None and len(plain_view["items"]) == 1


def test_model_rewriting_its_own_pending_edit_supersedes_it(doc):
    words = ["single-center", "one-site", "single-site", "monocentric", "single-hospital", "one-hospital"]
    for old, new in zip(words, words[1:]):
        model_edit(doc, plain(10, old, new))
    user_type(doc, 10, "one-hospital", "single-centre")
    out = feedback(doc)
    assert out["edit_feedback"] == "Since your last edit: 1 reworded. Call edit_feedback for details."
    assert len(out["items"]) == 1 and "(5 rounds)" in out["items"][0]
    assert '"single-center" → proposed "one-site" → kept "single-centre"' in out["items"][0]
    outcomes = [e["outcome"] for e in events(doc) if e["event"] == "outcome"]
    assert outcomes == ["superseded"] * 4 + ["reworded"]


def test_summary_line_only_when_something_is_new(doc):
    model_edit(doc, tracked(2, "small", "modest"))
    assert call("inspect_document", path=str(doc), search_text="model").get("edit_feedback") is None
    user_resolve(doc, "accept", 2)
    first = call("inspect_document", path=str(doc), search_text="model")
    assert first["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    again = call("inspect_document", path=str(doc), search_text="model")
    assert again.get("edit_feedback") is None                   # nothing new since the last call
    assert call("list_bookmarks", path=str(doc)).get("edit_feedback") is None


def test_summary_line_on_the_next_write_after_user_changes(doc):
    model_edit(doc, tracked(2, "small", "modest"), tracked(3, "good", "consistent"))
    user_resolve(doc, "reject", 3)
    out = model_edit(doc, plain(7, "nightly", "every night"))
    # pending counts every open proposal: paragraph 2's and the plain edit this call just made
    assert out["edit_feedback"] == "Since your last edit: 1 rejected (2 pending). Call edit_feedback for details."


def test_default_view_cap_newest_first_and_item_length(doc, tmp_path):
    long_doc = build_doc(tmp_path / "long.docx", [f"Sentence {n} " + "word " * 120 + f"tail{n}." for n in range(14)])
    for n in range(14):
        model_edit(long_doc, plain(n + 2, f"tail{n}", f"end{n}"))
    for n in range(14):
        user_type(long_doc, n + 2, f"end{n}", f"close{n}")
    out = feedback(long_doc)
    assert len(out["items"]) == 10 and out["more"] == 4
    assert all(len(item) <= fb.ITEM_MAX for item in out["items"])
    assert out["counts"]["reworded"] == 14
    nxt = feedback(long_doc)                        # the first 10 are now seen: the other 4 follow
    assert len(nxt["items"]) == 4 and nxt["more"] == 0
    assert set(nxt["items"]).isdisjoint(out["items"])
    last = feedback(long_doc)                       # nothing unseen: the most recent, with a note
    assert last["note"].startswith("Nothing new") and len(last["items"]) == 10


def test_outcomes_settled_together_list_the_later_proposal_first(doc):
    """One sync settles several outcomes at the same time; "newest first" then means the later proposal."""
    model_edit(doc, plain(6, "twice", "two times"), plain(7, "nightly", "every night"))
    model_edit(doc, plain(8, "manual", "careful"))
    user_type(doc, 6, "two times", "twice")
    user_type(doc, 7, "every night", "nightly")
    user_type(doc, 8, "careful", "manual")
    out = feedback(doc)
    assert [item.split()[1] for item in out["items"]] == ["¶8", "¶7", "¶6"]
    md_lines = [line for line in Path(out["md_path"]).read_text(encoding="utf-8").splitlines() if line.startswith("- ")]
    assert [line.split()[5] for line in md_lines] == ["¶8", "¶7", "¶6"]
    assert all(re.match(r"- \d{4}-\d\d-\d\d \d\d:\d\d UTC \[c-", line) for line in md_lines), md_lines


def test_long_text_items_stay_under_300(doc, tmp_path):
    big = build_doc(tmp_path / "big.docx", ["Start " + "alpha beta gamma " * 40 + "end."])
    model_edit(big, {"type": "replace_paragraph", "paragraph_index": 2,
                     "text": "Start " + "delta epsilon " * 60 + "fin."})
    user_type(big, 2, "fin.", "finish.")
    out = feedback(big)
    assert out["items"] and all(len(i) <= fb.ITEM_MAX for i in out["items"])


def test_scope_all_query_across_two_documents(doc, tmp_path):
    other = build_doc(tmp_path / "rivera_letter.docx")
    model_edit(doc, tracked(3, "good", "consistent"))
    model_edit(other, tracked(3, "good", "strong"))
    user_resolve(doc, "reject", 3)
    user_resolve(other, "accept", 3)
    call("inspect_document", path=str(doc), search_text="Results")
    call("inspect_document", path=str(other), search_text="Results")
    out = call("edit_feedback", scope="all", query="good")
    assert len(out["items"]) == 2
    assert any("rivera_manuscript.docx" in i and "rejected" in i for i in out["items"])
    assert any("rivera_letter.docx" in i and "accepted" in i for i in out["items"])
    assert call("edit_feedback", scope="all", query="consistent", outcome="rejected")["items"].__len__() == 1


def test_seen_marks(doc):
    model_edit(doc, tracked(2, "small", "modest"))
    user_resolve(doc, "accept", 2)
    first = feedback(doc)
    assert len(first["items"]) == 1 and first.get("note") is None
    seen = [e for e in events(doc) if e["event"] == "seen"]
    assert len(seen) == 1 and seen[0]["chains"] == [records(doc)[2]["id"]]
    second = feedback(doc)
    assert second["note"].startswith("Nothing new") and second["items"] == first["items"]
    assert len([e for e in events(doc) if e["event"] == "seen"]) == 1     # already-seen items are not re-marked


def test_md_mirror_written(doc):
    model_edit(doc, tracked(3, "good", "consistent"))
    user_resolve(doc, "reject", 3)
    out = feedback(doc)
    md = Path(out["md_path"]).read_text(encoding="utf-8")
    assert md.startswith("# Edit feedback: rivera_manuscript.docx")
    assert "1 rejected" in md and '"good" → proposed "consistent"' in md


def test_folder_refused_inside_onedrive(doc, monkeypatch, tmp_path):
    monkeypatch.setenv("RESEARCH_MCP_FEEDBACK_DIR", str(tmp_path / "OneDrive" / "feedback"))
    with pytest.raises(RuntimeError, match="OneDrive"):
        fbw.feedback_root()
    out = model_edit(doc, tracked(2, "small", "modest"))           # the edit itself still succeeds
    assert "OneDrive" in out["feedback_warning"]
    assert not (tmp_path / "OneDrive").exists()
    assert "OneDrive" in feedback(doc)["feedback_warning"]


def test_default_folder(monkeypatch, tmp_path):
    monkeypatch.delenv("RESEARCH_MCP_FEEDBACK_DIR", raising=False)
    monkeypatch.setenv("RESEARCH_MCP_HOME", str(tmp_path / "home"))
    assert fbw.feedback_root() == (tmp_path / "home" / "edit_feedback").resolve()
    monkeypatch.delenv("RESEARCH_MCP_HOME")
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: tmp_path / "user"))
    assert fbw.feedback_root() == (tmp_path / "user" / "research_mcp" / "edit_feedback").resolve()


def test_reconcile_failure_is_a_warning_and_the_tool_succeeds(doc, monkeypatch):
    model_edit(doc, tracked(2, "small", "modest"))
    user_resolve(doc, "accept", 2)

    def broken(*args, **kwargs):
        raise ValueError("synthetic reconcile failure")

    monkeypatch.setattr(fbw, "reconcile_props", broken)
    out = call("inspect_document", path=str(doc), search_text="model")
    assert out["matches"]                                           # the tool's own result is intact
    assert out["feedback_warning"].startswith("edit feedback not recorded (reconcile): ValueError")
    monkeypatch.setattr(fbw, "capture", broken)
    edit = model_edit(doc, plain(7, "nightly", "every night"))
    assert edit["status"] == "complete" and edit["feedback_warning"].startswith("edit feedback not recorded (capture)")


def test_undo_records_model_rejection(doc):
    model_edit(doc, plain(7, "nightly", "every night"))
    out = call("undo_document", path=str(doc))
    assert out.get("edit_feedback") is None
    record = records(doc)[7]
    assert (record["outcome"], record["via"]) == ("rejected", "model")


def test_rivera_sibling_case(doc):
    """The required sibling case: accept, reject, reword, a 5-round chain and an untracked edit the user
    reverts, all on one document, seen in one default view."""
    model_edit(doc, tracked(2, "small", "modest"), tracked(3, "good", "consistent"), tracked(4, "used", "selected"),
               plain(6, "twice", "two times"))
    back_and_forth(doc)
    user_resolve(doc, "accept", 2)
    user_resolve(doc, "reject", 3)
    user_resolve(doc, "accept", 4)
    user_type(doc, 4, "selected", "picked")
    user_type(doc, 6, "two times", "twice")
    line = call("audit_document", path=str(doc))["edit_feedback"]
    # the chain's last round is news too; paragraph 6 was already reported as kept during the chain, and
    # its later revert is recorded as a new outcome of the same proposal
    assert line == "Since your last edit: 1 accepted, 2 rejected, 2 reworded. Call edit_feedback for details."
    out = feedback(doc)
    assert out["pending"] == 0 and len(out["items"]) == 5 and out["counts"] == {
        "accepted": 1, "rejected": 2, "reworded": 2, "removed": 0}
    got = records(doc)
    assert got[2]["outcome"] == "accepted" and got[3]["outcome"] == "rejected"
    assert got[4]["outcome"] == "reworded" and got[4]["kept"] == "picked"
    assert got[6]["outcome"] == "rejected" and got[6]["tracked"] is False
    assert got[10]["rounds"] == 5 and got[10]["outcome"] == "reworded" and got[10]["kept"] == "single-centre"


def test_document_untouched_by_reconcile(doc):
    model_edit(doc, tracked(2, "small", "modest"))
    user_resolve(doc, "accept", 2)
    before = sha(doc)
    feedback(doc)
    call("inspect_document", path=str(doc), search_text="model")
    assert sha(doc) == before


def test_generic_module_has_no_word_imports():
    source = Path(fb.__file__).read_text(encoding="utf-8")
    imports = re.findall(r"^\s*(?:from|import)\s+([\w.]+)", source, re.M)
    assert set(imports) <= {"__future__", "difflib", "json", "os", "re", "uuid", "contextlib", "datetime",
                            "pathlib", "hashlib", "collections", ".checkpoints", ".textclean"}


# ---------------------------------------------------------------- review r1 (each fails without its fix)

def test_right_anchor_inside_the_users_text_is_not_called_accepted(tmp_path, feedback_dir):
    """r1 finding 1: "The small end" -> proposed "The large end"; the user's text "The large end result end"
    fits "large" (accepted) and "large end result" (reworded) equally well. Nothing is asserted: the proposal
    stays open (pending) and is judged again when the document next changes."""
    prop = fbw.hunks("The small end", "The large end")[0]
    assert fbw.judge(prop, "The large end result end") == ("unresolved", None)
    assert fbw.judge(prop, "The large end") == ("accepted", "large")
    assert fbw.judge(prop, "The huge end") == ("reworded", "huge")
    path = build_doc(tmp_path / "anchor.docx", ["The small end", "Other text here."])
    model_edit(path, plain(2, "small", "large"))
    user_type(path, 2, "large end", "large end result end")
    out = feedback(path)
    assert out.get("edit_feedback") is None and out["items"] == [] and out["pending"] == 1
    assert not [e for e in events(path) if e["event"] == "outcome"]
    user_type(path, 2, "large end result end", "large end")       # now only one reading fits
    out = feedback(path)
    assert out["edit_feedback"] == "Since your last edit: 1 accepted. Call edit_feedback for details."
    assert out["pending"] == 0


def test_repeated_word_in_one_paragraph_gives_two_records(tmp_path, feedback_dir):
    """r1 finding 2: two separate occurrences in one paragraph, with different anchors, are two chains."""
    p1 = fb.make_proposal(para=1, before="small", after="large", anchor_l="First ", anchor_r=" item", tracked=False)
    assert fb.assign_chain([], p1) == []
    p2 = fb.make_proposal(para=1, before="large", after="huge", anchor_l="Second ", anchor_r=" item", tracked=False)
    assert fb.assign_chain([p1], p2) == [] and p1["chain"] != p2["chain"]          # the reviewer's probe
    path = build_doc(tmp_path / "repeat.docx", ["First small item and second large item.", "Other text here."])
    model_edit(path, plain(2, "First small", "First large"))
    model_edit(path, plain(2, "second large", "second huge"))
    user_type(path, 3, "Other", "Another")
    out = feedback(path)
    assert out["edit_feedback"] == "Since your last edit: 2 accepted. Call edit_feedback for details."
    assert len(out["items"]) == 2
    proposals = [e for e in events(path) if e["event"] == "proposal"]
    assert len({e["chain"] for e in proposals}) == 2 and [e["round"] for e in proposals] == [1, 1]
    assert "superseded" not in [e["outcome"] for e in events(path) if e["event"] == "outcome"]


OLD_DATE = "2026-10-02T08:00:00Z"


def with_older_assistant_insertion(path: Path, text: str = "large ") -> Path:
    """Paragraph 2 gets an earlier, still pending "Assistant" insertion of ``text`` (not an edit the log knows)."""
    staged = path.with_name("older.docx")
    plain_p = f"<w:p>{_t(PARAGRAPHS[0])}</w:p>"
    marked = (f'<w:p>{_t("The model was trained on a ")}<w:ins w:id="7" w:author="{ASSISTANT}" w:date="{OLD_DATE}">'
              f'{_t(text)}</w:ins>{_t("small dataset of clinical notes.")}</w:p>')
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                assert plain_p in xml
                data = xml.replace(plain_p, marked).encode("utf-8")
            dst.writestr(info, data)
    _swap(path, staged)
    return path


@pytest.mark.parametrize("mode", ["reject", "accept"])
def test_older_same_author_revision_does_not_keep_a_decided_proposal_pending(doc, mode):
    """r1 finding 3: an older pending "Assistant" insertion ("large ") stays in the paragraph; the new proposal
    ("clinical" -> "large-scale": same author, same type, overlapping text) is decided on its own and must be
    recorded. Rejecting is told apart by the span; accepting by the date (to the minute, as Word keeps it)."""
    with_older_assistant_insertion(doc)
    assert accepted_text(doc, 2) == "The model was trained on a large small dataset of clinical notes."
    model_edit(doc, tracked(2, "clinical", "large-scale"))
    new_ids = [int(r.id) for r in load(doc)[1].revisions if r.author == ASSISTANT and r.date != OLD_DATE]
    assert len(new_ids) == 2
    staged = doc.with_name("user-only-new.docx")
    resolve_revisions(doc, staged, mode=mode, ids=new_ids)
    _swap(doc, staged)
    assert [r.date for r in load(doc)[1].revisions if r.author == ASSISTANT] == [OLD_DATE]   # the older one stays
    out = feedback(doc)
    word = "rejected" if mode == "reject" else "accepted"
    assert out["edit_feedback"] == f"Since your last edit: 1 {word}. Call edit_feedback for details."
    assert len(out["items"]) == 1 and '"clinical" → proposed "large-scale"' in out["items"][0]
    assert out["pending"] == 0



@pytest.mark.parametrize("older, n", [("large ", 1), ("large-scale", 2)])
def test_rejected_revision_beside_an_older_same_minute_one_is_not_kept_pending(doc, monkeypatch, older, n):
    """An older revision by the same author, from the same minute (overlapping text, or identical text), does not
    stand in for the model's rejected one once the user has also rewritten the words around it."""
    import time

    from research_mcp.revisions import revision_stamp

    if time.localtime().tm_sec >= 50:
        time.sleep(61 - time.localtime().tm_sec)  # start of a minute: the edit below lands in the same one
    now = re.search(r'w:date="([^"]+)"', revision_stamp("")).group(1)  # Word's convention, as the edit writes it
    monkeypatch.setattr(sys.modules[__name__], "OLD_DATE", now)
    with_older_assistant_insertion(doc, older)
    model_edit(doc, tracked(2, "clinical", "large-scale"))
    revs = next(e for e in events(doc) if e["event"] == "proposal")["revs"]
    ins = next(r for r in revs if r["type"] == "ins")
    assert (ins["text"], ins["date"][:16], ins["n"]) == ("large-scale", now[:16], n)  # really the same minute
    new_ids = [int(r.id) for r in load(doc)[1].revisions if r.author == ASSISTANT and r.id != "7"]
    assert len(new_ids) == 2
    staged = doc.with_name("reject-new.docx")
    resolve_revisions(doc, staged, mode="reject", ids=new_ids)
    _swap(doc, staged)
    user_type(doc, 2, "The model was trained on a ", "Here are ")
    user_type(doc, 2, "small dataset of clinical notes.", "my results.")
    assert [r.id for r in load(doc)[1].revisions if r.author == ASSISTANT] == ["7"]  # only the older one is left
    assert feedback(doc)["pending"] == 0


def test_edit_feedback_root_failure_is_a_warning(doc, monkeypatch):
    def broken():
        raise OSError("synthetic root failure")
    monkeypatch.setattr(fbw, "feedback_root", broken)
    out = feedback(doc)
    assert "synthetic root failure" in out["feedback_warning"] and not out.get("items")


def test_feedback_lock_wait_is_short():
    """A busy log costs a warning after at most a second, never a long stall on a tool call."""
    assert fb.LOCK_TIMEOUT <= 1.0


def test_edit_feedback_store_read_failure_is_a_warning(doc):
    """r1 finding 4: an unreadable log (here: bytes that are not UTF-8) gives feedback_warning, not an error."""
    model_edit(doc, tracked(2, "small", "modest"))
    with fbw.doc_log(doc).jsonl.open("ab") as handle:
        handle.write(b"\xff\xfe not utf-8\n")
    out = feedback(doc)
    assert "edit feedback store unavailable: UnicodeDecodeError" in out["feedback_warning"]
    assert out["items"] == []
    chain = "c-00000000"
    out = feedback(doc, record_id=chain, detail="rounds")
    assert "edit feedback store unavailable" in out["feedback_warning"]
    out = call("edit_feedback", scope="all", query="small")
    assert "edit feedback store unavailable: UnicodeDecodeError" in out["feedback_warning"]
    with pytest.raises(Exception, match="needs path"):                     # request errors stay errors
        asyncio.run(server.call_tool("edit_feedback", {}))


def test_edit_feedback_lock_and_seen_mark_failures_are_warnings(doc, monkeypatch):
    """r1 finding 4: a busy lock gives a warning; a failed seen mark still returns the items, with a warning."""
    from research_mcp.checkpoints import Refusal

    model_edit(doc, tracked(2, "small", "modest"))
    user_resolve(doc, "accept", 2)
    call("inspect_document", path=str(doc), search_text="model")      # settles the outcome (a log write)

    def failing_append(self, events, md=True):
        raise OSError("synthetic seen-mark failure")

    append = fb.DocLog.append
    monkeypatch.setattr(fb.DocLog, "append", failing_append)
    out = feedback(doc)
    assert len(out["items"]) == 1 and "accepted" in out["items"][0]
    assert out["feedback_warning"] == "seen marks not saved: OSError: synthetic seen-mark failure"
    monkeypatch.setattr(fb.DocLog, "append", append)   # (not monkeypatch.undo(): that drops the feedback folder)

    def busy(*args, **kwargs):
        raise Refusal("doc_busy", "edit feedback is busy", {})

    monkeypatch.setattr(fb, "_hold_file", busy)
    out = feedback(doc)
    assert out["feedback_warning"].startswith("edit feedback store unavailable: Refusal") and not out["items"]


def test_versioned_write_capture_failure_reaches_the_result(doc, tmp_path, monkeypatch):
    """r1 finding 5: a versioned write records feedback on the output; its warning must not be dropped."""
    def broken(*args, **kwargs):
        raise ValueError("synthetic capture failure")

    monkeypatch.setattr(fbw, "capture", broken)
    output = tmp_path / "rivera_manuscript_v2.docx"
    out = call("edit_document", source=str(doc), source_sha256=sha(doc), output=str(output),
               operations=[dict(plain(7, "nightly", "every night"), op_id="op-1")])
    assert out["status"] == "complete" and output.exists()
    assert out["feedback_warning"].startswith("edit feedback not recorded (capture): ValueError: synthetic")


def test_twenty_round_back_and_forth_is_one_record(doc):
    """r1 finding 6: twenty rounds on one span (tracked and plain, the user accepting, rewording, or leaving the
    model to rewrite its own text) are one default record with rounds 20 and no intermediate round in it."""
    current = "single-center"
    for n in range(1, 21):
        proposed = f"variant-{n:02d}"
        if n % 2:
            model_edit(doc, tracked(10, current, proposed))
            user_resolve(doc, "accept", 10)
        else:
            model_edit(doc, plain(10, current, proposed))
        current = proposed
        if n % 4 == 0:
            current = f"user-{n:02d}"
            user_type(doc, 10, proposed, current)
    out = feedback(doc)
    assert len(out["items"]) == 1
    item = out["items"][0]
    assert re.fullmatch(r'\[c-[0-9a-f]{8}\] ¶10 reworded \(20 rounds\): "…Limitations include the " '
                        r'"single-center" → proposed "variant-01" → kept "user-20"', item), item
    assert len(item) <= fb.ITEM_MAX
    assert not any(f"variant-{n:02d}" in line or f"user-{n:02d}" in line
                   for n in range(2, 20) for line in out["items"])
    record = records(doc)[10]
    assert record["rounds"] == 20 and record["kept"] == "user-20"
    proposals = [e for e in events(doc) if e["event"] == "proposal"]
    assert len(proposals) == 20 and len({e["chain"] for e in proposals}) == 1
    assert [e["round"] for e in proposals] == list(range(1, 21))
    rounds = feedback(doc, record_id=proposals[0]["chain"], detail="rounds")["rounds"]
    assert len(rounds) == 20 and all(len(line) <= fb.ITEM_MAX for line in rounds)


def test_torn_last_line_does_not_swallow_the_next_event(tmp_path):
    """r1 finding 7: a crash mid-write leaves a last line without its newline; the next append ends it first."""
    log = fb.DocLog(tmp_path / "fb", "torn-00000000")
    with log.locked():
        log.append([{"v": 1, "event": "sync", "sha": "a" * 64}], md=False)
    with log.jsonl.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write('{"v":1,"event":"pro')                              # torn: no newline
    with log.locked():
        log.append([{"v": 1, "event": "sync", "sha": "b" * 64}], md=False)
    assert [e["sha"][0] for e in log.events()] == ["a", "b"]
