"""Word side of edit feedback (plan 829): capture model proposals from a .docx write, reconcile them later.

The format-neutral store, chain collapsing and views are in ``edit_feedback``; this module is the seam that
knows .docx. It reads documents with ``revisions.load`` / ``revisions.ParagraphModel`` (accepted-view text,
tracked revisions with author/date/text), takes the doc-id normalisation, sha and OneDrive refusal from
``checkpoints`` and the folder setting from ``config``. It never writes a document; it only writes the feedback
files. Every entry point returns ``{"edit_feedback": ..., "feedback_warning": ...}`` and never raises.

Capture (after an MCP publish): diff the accepted-view paragraph text of the state before the write with the
state after it. Paragraphs whose XML is unchanged are skipped without building their text. Each changed span is
one proposal (word-level hunks; more than ``MAX_HUNKS`` in one paragraph become one paragraph-wide proposal) with
the unchanged text on each side as anchors. A paragraph that gained revisions gives tracked proposals (the new
revisions' type/author/date/text are kept); one that only lost revisions was an MCP accept/reject: its open
proposals are judged at once, ``via: model``. Two exceptions, both the model re-editing its own pending proposal:
an own insertion removed while the deletion of its tracked replacement stays (the ``w:del`` beside it, same author
and date; one edit gives all its changes one author and date, so position decides) is a new round of that chain
proposing the deletion alone ("small" -> "modest", then "modest" emptied: round 2 "small" -> "", pending on that
deletion). A word stored in several runs is deleted as one ``w:del`` per run; the deletion is then the whole
group of touching ``w:del`` (same author and date, no text between), every fragment kept, and the round stays
pending while any fragment does (decided when all are). An own insertion emptied by a text edit with nothing of
its proposal left (no deletion beside it) ends its chain as ``removed`` via model (the model withdrew it). An
MCP reject of a pure insertion leaves the same document, so the write's ``op`` decides: only a write whose ops
are tracked text edits withdraws; one that also names accept_revisions/reject_revisions, or any other writer,
keeps it ``rejected`` via model.

Undo (``on_sight`` via model): when ``undo_document`` returns the document to the state right after an earlier
MCP write, the chains that write left open go back to its round (later rounds ``undone``, that round
``reopened``), including a round a later MCP write decided (a withdrawal, an MCP accept/reject); undoing the
undo makes the later round current again. Rounds of an undo branch the restored round did not build on (its
``prev`` lineage) are ``undone``, so switching branches does not count them, and the rounds it did build on
count again, however deep the old branch.

Reconcile (a later tool call sees a sha the log has not accounted for): only open proposals are checked. Each is
looked for first at its recorded paragraph and nearby (``NEAR``), then in the rest of the document by its full
anchors, building paragraph text lazily. A tracked proposal is pending while its span still reads as proposed
and a revision of the same type and author, with matching text and the same date to the minute, stays in that
paragraph (Word truncates dates to the minute and renumbers w:id on save; the w:id is never used). Otherwise:
after-text between the anchors → accepted; before-text → rejected; other text between the anchors → reworded
(``kept`` is that text); not found → removed. When the anchors fit two readings equally well (an anchor repeated
inside the user's new text), nothing is asserted: the proposal stays open until the document next changes.
"""

from __future__ import annotations

import difflib
import os
import re
import time
from collections import Counter
from pathlib import Path

from . import checkpoints as ck
from . import config
from . import edit_feedback as fb

NEAR = 3            # paragraphs either side of the recorded index searched with shortened anchors too
SHORT_ANCHOR = 12   # characters of each anchor used for the nearby, forgiving search
MAX_HUNKS = 3       # more changed spans than this in one paragraph: one paragraph-wide proposal
CTX_MAX = 200       # uncut context kept each side of a span, to rank readings when its anchors repeat
READINGS_MAX = 8    # own tracked changes of a neighbour of an added paragraph enumerated either way (<= 2**8)
ADDED_SLACK = 2     # paragraphs the user may have added between an added block's neighbours
_TOKEN = re.compile(r"\w+(?:[-'\u2019]\w+)*|\s+|[^\w\s]")  # a hyphenated word is one token
_TEXT_REVS = ("ins", "del", "moveFrom", "moveTo")
_INS = ("ins", "moveTo")
# plan 864: a text op the tracking switch made tracked is named "<type>(tracked)" in edit_document's op string
_TEXT_EDIT_OP = re.compile(r"\b(?:tracked_replace_text|tracked_replace_paragraph|sweep_characters)\b"
                           r"|\b(?:replace_text|replace_paragraph|insert_text)\(tracked\)")
_RESOLVE_OP = re.compile(r"\b(?:accept|reject)_revisions\b")
# plan 864 S6 (L5): writes that change formatting only (the style_text tool, or an edit_document batch whose applied
# ops are style_text and the switch) propose no text and decide no proposal
_FORMAT_ONLY_OPS = frozenset({"style_text", "format_paragraph", "set_track_changes"})
# content that reads as text between two revision elements (markers such as bookmarks and comments do not)
_GAP_TEXT = re.compile(r"<w:(?:t|delText|tab|br|cr|sym|noBreakHyphen|softHyphen|instrText|delInstrText|drawing)\b")


def _format_only(tool: str | None, op_types: str | None) -> bool:
    """L5: the write changed run formatting only, so it records no proposal (text unchanged) and decides none.
    ``op_types``: the applied op types, comma-separated (edit_document's ``op_types``; never op ids)."""
    if tool == "style_text":
        return True
    kinds = {part.strip().split(" ", 1)[0] for part in (op_types or "").split(",") if part.strip()}
    return bool(kinds & {"style_text", "format_paragraph"}) and kinds <= _FORMAT_ONLY_OPS


def _none() -> dict:
    return {"edit_feedback": None, "feedback_warning": None}


# ---------------------------------------------------------------- where the files go

def feedback_root() -> Path:
    """RESEARCH_MCP_FEEDBACK_DIR (environment or config file), else ``RESEARCH_MCP_HOME/edit_feedback`` when that
    is set, else ``~/research_mcp/edit_feedback``. Refused inside OneDrive (RuntimeError), as the undo store is."""
    raw = (config.get("RESEARCH_MCP_FEEDBACK_DIR", "") or "").strip().strip('"')
    if raw:
        root = Path(os.path.expandvars(os.path.expanduser(raw)))
    elif os.environ.get("RESEARCH_MCP_HOME"):
        root = ck.home() / "edit_feedback"
    else:
        root = Path.home() / "research_mcp" / "edit_feedback"
    return ck.assert_outside_onedrive(root.resolve())


def doc_log(path) -> fb.DocLog:
    p = Path(path)
    return fb.DocLog(feedback_root(), fb.file_stem(p.stem, ck.norm_path(p)), str(p.resolve()))


# ---------------------------------------------------------------- reading a .docx

class Snapshot:
    """One state of a .docx: paragraph XML for cheap comparison, accepted-view text and revisions on demand."""

    def __init__(self, path) -> None:
        from .revisions import load

        _, self.doc = load(path)
        self.indices = sorted(self.doc.para_nodes)
        self.raw = [self.doc.xml[self.doc.para_nodes[i].start:self.doc.para_nodes[i].end] for i in self.indices]
        self._text: dict[int, str] = {}
        self.revs: dict[int, list] = {}
        self.mark_dels: dict[int, set[tuple]] = {}  # (author, date) of each deleted paragraph mark, by paragraph
        self.row_dels: set[int] = set()  # paragraphs of table rows tracked as deleted (w:del in the row's trPr)
        for revision in self.doc.revisions:
            if revision.type in _TEXT_REVS and revision.paragraph_index is not None and revision.text:
                self.revs.setdefault(revision.paragraph_index, []).append(revision)
            elif revision.type == "del" and revision.mark == "paragraph" and revision.paragraph_index is not None:
                self.mark_dels.setdefault(revision.paragraph_index, set()).add((revision.author, revision.date))
            elif revision.type == "del" and revision.mark == "row" and revision.paragraph_index is not None:
                row = self.doc.layout.row_of(revision.paragraph_index)  # the row mark sits before its first cell
                if row is not None:
                    self.row_dels.update(index for cell in row.cells for index in cell)
        self.built = 0
        self._readings: dict[int, set[str] | None] = {}

    def text(self, index: int | None) -> str:
        if index is None or index not in self.doc.para_nodes:
            return ""
        if index not in self._text:
            from .revisions import ParagraphModel

            self._text[index] = ParagraphModel(self.doc.xml, self.doc.para_nodes[index]).text
            self.built += 1
        return self._text[index]

    def original(self, index: int | None) -> str:
        """The paragraph as Reject All would leave it (its tracked insertions out, deletions back)."""
        if index is None or index not in self.doc.para_nodes:
            return ""
        from .revisions import ParagraphModel

        return ParagraphModel(self.doc.xml, self.doc.para_nodes[index]).original

    def readings(self, index: int | None) -> set[str] | None:
        """Every text the paragraph can read as when each of its own tracked text changes is accepted or
        rejected on its own (READINGS_MAX of them at most; None beyond, or when the enumeration does not
        reproduce the accepted text), so a neighbour the user resolved only partly is still recognised."""
        if index is None or index not in self.doc.para_nodes:
            return None
        if index not in self._readings:
            self._readings[index] = self._enumerate(index)
        return self._readings[index]

    def _enumerate(self, index: int) -> set[str] | None:
        import copy
        import itertools

        from .ooxml import W
        from .revisions import _et_final_chars, _fragment, et_resolve

        node = self.doc.para_nodes[index]
        base = _fragment(self.doc.xml, self.doc.xml[node.start:node.end])
        tags = {W + name for name in _TEXT_REVS}

        def wrappers(paragraph) -> list:
            marks = {child for rpr in paragraph.iter(W + "rPr") for child in rpr}  # mark and run-property markers
            return [el for el in paragraph.iter() if el.tag in tags and el not in marks]

        count = len(wrappers(base))
        if count > READINGS_MAX:
            return None
        out: set[str] = set()
        try:
            for choice in itertools.product((True, False), repeat=count):
                clone = copy.deepcopy(base)
                found = wrappers(clone)
                take = {el for el, accept in zip(found, choice) if accept}
                et_resolve(clone, lambda el: el in take, "accept")
                et_resolve(clone, lambda el: el in found and el not in take, "reject")
                text = "".join(ch for ch, _ in _et_final_chars(clone))
                if all(choice) and text != self.text(index):
                    return None  # this reader and ParagraphModel disagree: do not trust the other readings
                out.add(text)
        except Exception:  # noqa: BLE001 - a structure the resolver refuses: fall back to the two readings
            return None
        return out

    def rev_keys(self, index: int | None) -> Counter:
        return Counter((r.type, r.author, r.date, r.text) for r in self.revs.get(index, [])) if index else Counter()


# ---------------------------------------------------------------- capture

def _align(before: Snapshot, after: Snapshot) -> list[tuple[int | None, int | None, bool]]:
    """Every paragraph of both states in document order as (before index, after index, changed); None for an
    added or removed paragraph. The after indices appear in ascending order, so a removed paragraph sits where
    the after paragraphs listed before it end."""
    out: list[tuple[int | None, int | None, bool]] = []
    matcher = difflib.SequenceMatcher(None, before.raw, after.raw, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        bi, ai = before.indices[i1:i2], after.indices[j1:j2]
        if tag == "equal" or len(bi) == len(ai):
            out.extend((b, a, tag != "equal") for b, a in zip(bi, ai))
            continue
        inner = difflib.SequenceMatcher(None, [before.text(i) for i in bi], [after.text(i) for i in ai],
                                        autojunk=False)
        for tag2, k1, k2, l1, l2 in inner.get_opcodes():
            b, a = bi[k1:k2], ai[l1:l2]
            common = min(len(b), len(a))
            out.extend((x, y, True) for x, y in zip(b[:common], a[:common]))
            out.extend((x, None, True) for x in b[common:])
            out.extend((None, y, True) for y in a[common:])
    return out


def _pairs(before: Snapshot, after: Snapshot) -> list[tuple[int | None, int | None]]:
    """Changed paragraphs as (before index, after index); None for an added or removed paragraph."""
    return [(b, a) for b, a, changed in _align(before, after) if changed]


def _tracked_deletion(before: Snapshot, after: Snapshot, b: int, a: int) -> bool:
    """Paragraph ``b`` -> ``a`` is a tracked paragraph deletion (replace_section / replace_table_row removing a
    line): its text is now all in new ``w:del`` revisions and a paragraph mark was deleted by the same author at
    the same time, its own, or (Word's way for the last line of a cell) the mark of the paragraph before it. A
    paragraph emptied with its mark kept is a text proposal, not this."""
    if not before.text(b) or after.text(a):
        return False
    new, lost = after.rev_keys(a) - before.rev_keys(b), before.rev_keys(b) - after.rev_keys(a)
    if not new or lost or any(key[0] != "del" for key in new):
        return False
    stamps = {(key[1], key[2]) for key in new}
    pos = after.indices.index(a)
    marks = after.mark_dels.get(a, set()) | (after.mark_dels.get(after.indices[pos - 1], set()) if pos else set())
    return bool(stamps & marks)


def _deleted_places(before: Snapshot, after: Snapshot,
                    aligned: list[tuple[int | None, int | None, bool]]) -> dict[int, dict]:
    """Place fields (``_place_fields``) of every paragraph the write deleted, keyed by its before index: removed
    outright (untracked) or tracked (``_tracked_deletion``; it reads as an empty paragraph until accepted). Its
    neighbours are the nearest paragraphs in the after state that were not deleted; deleted paragraphs between
    the same neighbours form one block (``block``, ``block_others``). A blank paragraph of a table row this write
    tracked as deleted is no proposal (no text) but a placeholder too: Accept removes it with the row."""
    gone: list[tuple[int, int, int | None]] = []  # (before index, after position or gap, after index)
    blanks: set[int] = set()  # after positions of blank paragraphs in rows this write tracked as deleted
    seen = 0
    for b, a, changed in aligned:
        if a is not None:
            if changed and b is not None and _tracked_deletion(before, after, b, a):
                gone.append((b, seen, a))
            elif (changed and b is not None and a in after.row_dels and b not in before.row_dels
                  and not before.text(b) and not after.text(a)):
                blanks.add(seen)
            seen += 1
        elif changed and b is not None and before.text(b):
            gone.append((b, seen, None))
    if not gone:
        return {}
    tracked = {pos for _, pos, a in gone if a is not None} | blanks
    n = len(after.indices)
    blocks: dict[tuple[int, int], list[int]] = {}
    for b, pos, a in gone:
        lo = pos - 1
        while lo >= 0 and lo in tracked:
            lo -= 1
        hi = pos + 1 if a is not None else pos
        while hi < n and hi in tracked:
            hi += 1
        blocks.setdefault((lo, hi), []).append(b)
    out = {}
    for (lo, hi), members in blocks.items():
        for b in members:
            others = [before.text(x) for x in members if x != b]
            out[b] = {"deleted": True, "block": len(members),
                      **_place_fields(after, lo, hi, others, skip=tracked)}
            out[b]["_spot"] = after.indices[hi] if hi < n else (after.indices[-1] + 1 if n else b)
    return out


def hunks(before: str, after: str) -> list[dict]:
    """Word-level changed spans of one paragraph, with the unchanged text on each side as anchors."""
    if before == after:
        return []
    if not before or not after:
        return [{"before": before, "after": after, "anchor_l": "", "anchor_r": "", "start": True, "end": True,
                 "b_start": 0, "b_end": len(before)}]
    bt, at = _TOKEN.findall(before), _TOKEN.findall(after)
    bo = [0]
    for token in bt:
        bo.append(bo[-1] + len(token))
    ao = [0]
    for token in at:
        ao.append(ao[-1] + len(token))
    spans: list[list[int]] = []  # [b1, b2, a1, a2] in tokens
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, bt, at, autojunk=False).get_opcodes():
        if tag == "equal":
            continue
        if spans and i1 - spans[-1][1] <= 2:  # one or two unchanged tokens between changes: one span
            spans[-1][1], spans[-1][3] = i2, j2
        else:
            spans.append([i1, i2, j1, j2])
    if len(spans) > MAX_HUNKS:
        spans = [[spans[0][0], spans[-1][1], spans[0][2], spans[-1][3]]]
    out = []
    for n, (b1, b2, a1, a2) in enumerate(spans):
        left_from = ao[spans[n - 1][3]] if n else 0
        right_to = ao[spans[n + 1][2]] if n + 1 < len(spans) else len(after)
        out.append({
            "before": before[bo[b1]:bo[b2]], "after": after[ao[a1]:ao[a2]],
            "anchor_l": after[left_from:ao[a1]][-fb.ANCHOR_MAX:],
            "anchor_r": after[ao[a2]:right_to][:fb.ANCHOR_MAX],
            "start": a1 == 0 and b1 == 0, "end": a2 == len(at) and b2 == len(bt),
            "b_start": bo[b1], "b_end": bo[b2],                      # where the span was, in the old text
            "ctx_l": after[:ao[a1]][-CTX_MAX:], "ctx_r": after[ao[a2]:][:CTX_MAX],  # uncut context
        })
    return out


def _assign_revs(hunk: dict, new_revs: list[tuple]) -> list[dict]:
    """The new revisions whose text lies in this span (all of them when none can be placed)."""
    mine = [key for key in new_revs
            if (key[0] in _INS and key[3] and key[3] in hunk["after"])
            or (key[0] not in _INS and key[3] and key[3] in hunk["before"])] or list(new_revs)
    return [{"type": t, "author": a, "date": d, "text": x} for t, a, d, x in mine]



def _neighbours(snap: Snapshot, index: int) -> list[str]:
    """The non-empty paragraphs NEAR ``index`` in ``snap`` as the write left them, cut (``fb.neighbour_texts``)."""
    if index not in snap.doc.para_nodes:
        return []
    pos = snap.indices.index(index)
    near = snap.indices[max(0, pos - NEAR):pos] + snap.indices[pos + 1:pos + 1 + NEAR]
    return fb.neighbour_texts(map(snap.text, near))


def _window_counts(snap: Snapshot) -> Counter:
    """How many paragraphs of ``snap`` stand among each list of neighbours (``_neighbours`` of every paragraph): a
    filled blank line records its count (``nbt_n``), so a copy of its block added later shows as one more."""
    indices = snap.indices
    return Counter(tuple(fb.neighbour_texts(map(snap.text, indices[max(0, k - NEAR):k] + indices[k + 1:k + 1 + NEAR])))
                   for k in range(len(indices)))


def _around(snap: Snapshot, index: int) -> list[str]:
    """The paragraphs within 2 * NEAR of ``index`` (itself excluded): where its recorded neighbours should be."""
    pos = snap.indices.index(index)
    span = 2 * NEAR
    return [snap.text(i) for i in snap.indices[max(0, pos - span):pos] + snap.indices[pos + 1:pos + 1 + span]]


def _tighten(hunk: dict) -> dict:
    """Move whitespace both sides of a span share into its anchors: emptying "medical" in "of medical notes"
    leaves "of  notes", whose word diff is " medical " -> "  "; the proposal is "clinical" -> "" between
    "of " and " notes", so a Reject reads back exactly "clinical"."""
    before, after = hunk["before"], hunk["after"]
    p = 0
    while p < min(len(before), len(after)) and before[p] == after[p] and before[p].isspace():
        p += 1
    s = 0
    while (s < min(len(before), len(after)) - p and before[-1 - s] == after[-1 - s]
           and before[-1 - s].isspace()):
        s += 1
    if not p and not s:
        return hunk
    lead, trail = after[:p], after[len(after) - s:]
    return dict(hunk, before=before[p:len(before) - s], after=after[p:len(after) - s],
                anchor_l=(hunk["anchor_l"] + lead)[-fb.ANCHOR_MAX:],
                anchor_r=(trail + hunk["anchor_r"])[:fb.ANCHOR_MAX],
                ctx_l=(hunk.get("ctx_l", "") + lead)[-CTX_MAX:], ctx_r=(trail + hunk.get("ctx_r", ""))[:CTX_MAX],
                start=hunk["start"] and not p, end=hunk["end"] and not s,
                b_start=hunk["b_start"] + p, b_end=hunk["b_end"] - s)


def _key(revision) -> tuple:
    return (revision.type, revision.author, revision.date, revision.text)


def _touch(snap: Snapshot, first, second) -> bool:
    """``first`` ends where ``second`` starts in the XML with no text of its own between them (markers such as
    bookmarks and comments do not count)."""
    return first.node.end <= second.node.start and not _GAP_TEXT.search(snap.doc.xml, first.node.end,
                                                                         second.node.start)


def _replaced(snap: Snapshot, index: int, ins) -> list:
    """The deletions of the tracked replacement ``ins`` belongs to, in document order: the ``w:del`` by the same
    author with the same date that touches it in the XML (``<w:del>small</w:del><w:ins>modest</w:ins>``), plus
    every such ``w:del`` touching that one in turn. A word stored in several runs (``clini`` + ``cal``) is
    deleted as one ``w:del`` per run, so the replacement's deletion is that whole group. One tracked edit
    writes every change of a paragraph with one author and date, so author and date alone do not tell which
    deletion goes with which insertion; the position does, and any text between (another replacement's
    insertion) ends the group."""
    dels = sorted((r for r in snap.revs.get(index, []) if r.type == "del" and r.author == ins.author
                   and r.date == ins.date and r.text), key=lambda r: r.node.start)
    for n, other in enumerate(dels):
        if _touch(snap, other, ins):        # the group ends right before the insertion
            lo = n
            while lo and _touch(snap, dels[lo - 1], dels[lo]):
                lo -= 1
            return dels[lo:n + 1]
        if _touch(snap, ins, other):        # or starts right after it
            hi = n
            while hi + 1 < len(dels) and _touch(snap, dels[hi], dels[hi + 1]):
                hi += 1
            return dels[n:hi + 1]
    return []


def _withdrawn_pairs(before: Snapshot, after: Snapshot, b: int | None, a: int | None,
                     lost: Counter) -> list[tuple[tuple, list[tuple]]]:
    """(insertion key, deletion keys) for each own insertion this write removed whose deletion of the same
    tracked replacement (``_replaced``: adjacent, same author and date; every fragment, in document order) is
    still in the paragraph: the model emptied its pending insertion (or rejected only that half), so it now
    proposes the deletion alone. An insertion with no adjacent deletion was a pure insertion (withdrawn,
    ``_withdrawn``), whatever other deletion the same edit made elsewhere in the paragraph."""
    if a is None or b is None:
        return []
    kept = after.rev_keys(a) & before.rev_keys(b)
    out, used = [], Counter()
    for revision in before.revs.get(b, []):
        key = _key(revision)
        if key[0] != "ins" or not key[3] or not key[2] or key not in lost:
            continue
        group = []
        for partner in map(_key, _replaced(before, b, revision)):
            if kept[partner] > used[partner]:
                used[partner] += 1
                group.append(partner)
        if group:
            out.append((key, group))
    return out


def capture(before_path, after_path) -> tuple[list[dict], set[tuple], Snapshot]:
    """Proposal fields for every changed span, the revision keys the write removed, and the after snapshot."""
    before, after = Snapshot(before_path), Snapshot(after_path)
    proposals: list[dict] = []
    lost_all: set[tuple] = set()
    aligned = _align(before, after)
    pairs = [(b, a) for b, a, changed in aligned if changed]
    added = {a for b, a in pairs if b is None and a is not None}
    deleted = _deleted_places(before, after, aligned)
    windows = None  # _window_counts(after), built on first use
    for b, a in pairs:
        btext, atext = before.text(b), after.text(a)
        new = after.rev_keys(a) - before.rev_keys(b)
        lost = before.rev_keys(b) - after.rev_keys(a)
        lost_all.update(lost)
        paired = _withdrawn_pairs(before, after, b, a, lost) if lost and btext != atext else []
        if btext == atext or (lost and not new and not paired):
            continue  # formatting only, or an MCP accept/reject (judged against its proposals instead)
        found = hunks(btext, atext)
        for hunk in found:
            revs = _assign_revs(hunk, list(new)) if new else []
            for pair in list(paired):  # the deletion left of an emptied own replacement: propose it
                ins_key, del_keys = pair
                if ins_key[3] in hunk["before"]:
                    paired.remove(pair)
                    whole = "".join(key[3] for key in del_keys)
                    hunk = _tighten(dict(hunk, before=hunk["before"].replace(ins_key[3], whole, 1)))
                    part = len(del_keys) > 1  # one deletion in several runs: pending while any fragment stays
                    for kind, who, when, words in del_keys:
                        rev = {"type": kind, "author": who, "date": when, "text": words}
                        revs.append(dict(rev, part=True) if part else rev)
            left = Counter((r.type, r.author, _minute(r.date), r.text) for r in after.revs.get(a, []))
            for rev in revs:  # identical revisions the write left in the paragraph, this one included
                rev["n"] = left[(rev["type"], rev["author"], _minute(rev["date"]), rev["text"])]
            para = a if a is not None else b
            nbt = _neighbours(after, a) if a is not None else []
            if a is not None and nbt and hunk["start"] and hunk["end"] and not (hunk["before"] or "").strip():
                windows = windows if windows is not None else _window_counts(after)
                nbt_n = {"nbt_n": windows[tuple(nbt)]}  # a filled blank line: places then exactly among nbt
            else:
                nbt_n = {}
            extra = _added_fields(after, a, added) if b is None and a is not None else dict(deleted.get(b, {}))
            spot = extra.pop("_spot", None)
            if a is None and spot is not None:
                para = spot  # a paragraph removed outright: where it stood, the paragraph that took its place
            proposals.append(dict(hunk, para=para, where=f"¶{para}", nbt=nbt, **nbt_n, tracked=bool(new or revs),
                                  author=revs[0]["author"] if revs else None,
                                  date=revs[0]["date"] if revs else None, revs=revs, _btext=btext, **extra))
    return proposals, lost_all, after


def _added_fields(snap: Snapshot, index: int, added: set[int]) -> dict:
    """A paragraph the write added (insert_paragraph): ``added``, the size of its added block (``block``) and
    the paragraphs that were already there on each side of that block (``nb_prev``/``nb_next``, cut by
    ``fb.neighbour_key``; None at the start or end of the document). Its identity is its place between those
    two, not its text (the same text elsewhere is another paragraph). A neighbour holding tracked changes also
    keeps its other readings, since the user may resolve them with the added paragraph: the Reject All reading
    (``nb_prev_orig``/``nb_next_orig``, when it differs) and, for at most READINGS_MAX own changes, every reading
    with each change accepted or rejected (``nb_prev_alts``/``nb_next_alts``)."""
    pos = snap.indices.index(index)
    before = [k for k in range(pos) if snap.indices[k] not in added]
    later = [k for k in range(pos + 1, len(snap.indices)) if snap.indices[k] not in added]
    lo, hi = (before[-1] if before else -1), (later[0] if later else len(snap.indices))
    others = [snap.text(snap.indices[k]) for k in range(lo + 1, hi) if k != pos]
    return {"added": True, "block": hi - lo - 1, **_place_fields(snap, lo, hi, others)}


def _place_fields(snap: Snapshot, lo: int, hi: int, others, skip=frozenset()) -> dict:
    """The place of a paragraph the write added or deleted, between positions ``lo`` and ``hi`` of ``snap``
    (-1 / len at the document's edges): the texts of the block's other members (``block_others``, hashed),
    each neighbour's key and other readings (``nb_prev``/``nb_next``, ``_orig``, ``_alts``) and how many
    paragraphs read as it now (``nb_*_n``; positions in ``skip``, a tracked deletion's empty placeholders,
    are not counted). A blank paragraph is no neighbour (it reads as every other blank one, and a table row's
    blank cells are common): the neighbour is the nearest paragraph with text beyond it, and the blank ones passed
    on the way are counted (``gap_prev``/``gap_next``, when any; ``_place_outcome`` passes over them)."""
    out: dict = {}
    hashes = {_reading_hash(fb.neighbour_key(text)) for text in others}
    if hashes:  # the block's other members: between the same neighbours, they are not rivals
        out["block_others"] = sorted(hashes)
    n = len(snap.indices)
    for name, step in (("gap_prev", -1), ("gap_next", 1)):
        gap = 0
        while 0 <= (k := lo if step < 0 else hi) < n and (k in skip or _blank(snap.text(snap.indices[k]))):
            gap += k not in skip
            if step < 0:
                lo -= 1
            else:
                hi += 1
        if gap:
            out[name] = gap
    for name, k in (("nb_prev", lo), ("nb_next", hi)):
        at = snap.indices[k] if 0 <= k < len(snap.indices) else None
        out[name] = None if at is None else fb.neighbour_key(snap.text(at))
        if at is None or not snap.revs.get(at):
            continue
        if (orig := fb.neighbour_key(snap.original(at))) != out[name]:
            out[name + "_orig"] = orig
        alts = {_reading_hash(fb.neighbour_key(text)) for text in snap.readings(at) or ()} - {
            _reading_hash(out[name]), _reading_hash(orig)}
        if alts:
            out[name + "_alts"] = sorted(alts)
    keys = [fb.neighbour_key(snap.text(i)) for k, i in enumerate(snap.indices) if k not in skip]
    for side in ("prev", "next"):  # how many paragraphs read as that neighbour now (repeated text, blank lines)
        if (test := _neighbour_test(out, side)) is not None:
            out[f"nb_{side}_n"] = sum(1 for key in keys if test(key))
    return out


def _blank(text: str) -> bool:
    return not text.strip()


def _reading_hash(key: str) -> str:
    """A neighbour reading as stored in ``nb_*_alts``: a short hash (up to 2**READINGS_MAX of them per side)."""
    import hashlib

    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:12]


def _neighbour_test(prop: dict, side: str):
    """Whether a paragraph key is one of the readings neighbour ``side`` ("prev"/"next") of an added paragraph
    may have now; None at the document edge."""
    key = prop.get(f"nb_{side}")
    if key is None:
        return None
    keys = {key, prop.get(f"nb_{side}_orig", key)}
    hashes = set(prop.get(f"nb_{side}_alts") or ())
    return lambda text: text in keys or (bool(hashes) and _reading_hash(text) in hashes)


def _place_outcome(prop: dict, snap: Snapshot) -> tuple[str, str | None, int] | None:
    """(outcome, kept, paragraph) for a paragraph the model added or deleted, found by its place, else None
    (pending).

    Its recorded neighbours adjacent again: an added paragraph is ``rejected``, a deleted one ``accepted``. A
    paragraph between them (at most ``block`` + ADDED_SLACK apart) that reads as the proposal: an added one is
    ``accepted`` (or ``reworded``) there, a deleted one ``rejected`` (its text is back; ``reworded`` for a near
    match). An empty paragraph between them decides nothing (``judge`` gives it no verdict). The same text
    anywhere else is never it. Decided only when the places that fit agree: one paragraph between the neighbours
    reading as the proposal, exactly or reworded (the block's other members, ``block_others``, aside), or
    rejoined pairs only, and no paragraph reading as the proposal right after a previous-neighbour match or right
    before a next-neighbour match elsewhere (it may stand where one of its neighbours was deleted).
    A neighbour that read the same as other paragraphs when the edit was made (``nb_*_n`` > 1) must still have
    that many matches, else a lost copy could be the original and a duplicate pair elsewhere taken for it.
    Otherwise None, so it stays pending and is judged again when the document next changes (never a guess,
    and never ``removed``)."""
    keys = [fb.neighbour_key(snap.text(i)) for i in snap.indices]
    n = len(keys)
    prev, nxt = _neighbour_test(prop, "prev"), _neighbour_test(prop, "next")
    lefts = [-1] if prev is None else [k for k in range(n) if prev(keys[k])]
    rights = [n] if nxt is None else [k for k in range(n) if nxt(keys[k])]
    span = max(int(prop.get("block") or 1), 1) + ADDED_SLACK
    for side, found in (("prev", lefts), ("next", rights)):
        recorded = prop.get(f"nb_{side}_n")
        if recorded is not None and recorded > 1 and len(found) != recorded:
            return None  # a repeated neighbour lost or gained a copy: the original place is not known
    ends = set(lefts) | set(rights)

    def reads_as_proposal(k: int) -> bool:
        text = snap.text(snap.indices[k])
        return judge(prop, text) is not None or judge(prop, text, short=True) is not None

    # The paragraphs that were blank between a neighbour and the block (``gap_*``) are passed over, as if the last
    # blank one were the neighbour. Walking from the neighbour, the walk stops early at the proposal (the user
    # deleted a blank line beside it) or at the other neighbour (the user deleted blank lines and the block is
    # gone). A gap paragraph that now holds other text is no blank line any more and may be the proposal
    # rewritten: that match cannot be placed, so when the other neighbour or a paragraph reading as the proposal is
    # within reach of it the outcome stays pending; a match with neither in reach (a repeated neighbour elsewhere)
    # is just dropped.
    gp, gn = int(prop.get("gap_prev") or 0), int(prop.get("gap_next") or 0)
    raw_lefts, raw_rights = set(lefts), set(rights)
    span_reach = span + gp + gn

    def walk(start: int, gap: int, step: int, other: set[int]) -> int | None:
        """The shifted match, or None when a gap paragraph holds other text."""
        at = start
        for _ in range(gap):
            k = at + step
            if not 0 <= k < n or k in other or reads_as_proposal(k):
                break
            if not _blank(snap.text(snap.indices[k])):
                return None
            at = k
        return at

    def proposal_near(ks: range) -> bool:
        """Whether a paragraph in ``ks`` (clipped to the document) reads as the proposal."""
        return any(reads_as_proposal(k) for k in ks if 0 <= k < n)

    shifted_lefts, shifted_rights = [], []
    for lo in lefts:
        if not gp:
            shifted_lefts.append(lo)
        elif (at := walk(lo, gp, 1, raw_rights)) is not None:
            shifted_lefts.append(at)
        elif (any(lo < hi <= lo + 1 + span_reach for hi in raw_rights)
              or proposal_near(range(lo + 1, lo + 1 + span_reach))):
            return None  # the other neighbour, or the proposal itself, is past the filled line: may be its place
    for hi in rights:
        if not gn:
            shifted_rights.append(hi)
        elif (at := walk(hi, gn, -1, raw_lefts)) is not None:
            shifted_rights.append(at)
        elif (any(hi - 1 - span_reach <= lo < hi for lo in raw_lefts)
              or proposal_near(range(hi - span_reach, hi))):
            return None
    lefts, rights = shifted_lefts, shifted_rights
    rejoined = [lo for lo in lefts if lo + 1 in rights]
    inside = sorted({k for lo in lefts for hi in rights if lo + 1 < hi <= lo + 1 + span for k in range(lo + 1, hi)})
    verdicts: list[tuple[int, tuple[str, str | None]]] = []
    for short in (False, True):
        verdicts = [(snap.indices[k], v) for k in inside if (v := judge(prop, snap.text(snap.indices[k]),
                                                                        short=short)) is not None]
        if verdicts:
            break
    inside_set = set(inside)

    # The proposal standing next to only one of its neighbours (the other deleted): another place it may be. A
    # paragraph that reads as a neighbour ("Third body." beside a deleted "Second body.") counts only exactly.
    stray = {k for k in [lo + 1 for lo in lefts if lo >= 0] + [hi - 1 for hi in rights if hi < n]
             if 0 <= k < n and k not in inside_set
             and (judge(prop, snap.text(snap.indices[k])) is not None if k in ends else reads_as_proposal(k))}
    if stray:
        return None
    if verdicts:
        if rejoined or len({index for index, _ in verdicts}) > 1 or verdicts[0][1][0] == "unresolved":
            return None  # two places fit (repeated neighbours, or two readings): leave it open
        siblings = set(prop.get("block_others") or ())
        chosen = verdicts[0][0]
        if any(snap.indices[k] != chosen and reads_as_proposal(k)
               and _reading_hash(keys[k]) not in siblings for k in inside):
            return None  # an exact copy beside a reworded one (the user's own, or the original reworded): open
        index, (outcome, kept) = verdicts[0]
        return outcome, kept, index
    if rejoined:
        # Several rejoined pairs (a repeated neighbour pair, every copy still there) all mean the same: the
        # proposal reads at none of them, so it is gone wherever it stood; only the spot reported is the nearest.
        para = prop.get("para") or 0
        best = min(rejoined, key=lambda k: abs(k + 1 - para))  # keys[k + 1] is where the paragraph stood
        spot = snap.indices[min(best + 1, n - 1)] if n else None
        return ("accepted" if prop.get("deleted") else "rejected"), "", spot
    return None


# ---------------------------------------------------------------- reconcile

def _between(text: str, left: str, right: str, start: bool, end: bool) -> list[tuple[int, int]]:
    """Every (lo, hi) with an occurrence of ``left`` ending at lo and an occurrence of ``right`` starting at
    hi >= lo (paragraph start/end when the flag is set). Anchors can repeat ("word word word", or the right
    anchor inside the user's own new text), so every pairing is returned and ``judge`` decides."""
    if start:
        los = [0]
    elif left:
        los, at = [], text.find(left)
        while at >= 0:
            los.append(at + len(left))
            at = text.find(left, at + 1)
    else:
        return []
    if end:
        his = [len(text)]
    elif right:
        his, at = [], text.find(right)
        while at >= 0:
            his.append(at)
            at = text.find(right, at + 1)
    else:
        return []
    return [(lo, hi) for lo in los for hi in his if hi >= lo]


def judge(prop: dict, text: str, *, short: bool = False) -> tuple[str, str | None] | None:
    """(outcome, kept) when this paragraph text holds the proposal's span, else None.

    ``("unresolved", None)`` when the span is there but its edges are not: two readings fit the anchors
    equally well (e.g. proposed "The large end" from "The small end", now "The large end result end": "large"
    accepted, or "large end result" typed by the user). The proposal then stays open (counted as pending) and
    is judged again when the document next changes. Asserting "accepted" there could teach the model a style
    the user did not choose."""
    before, after = prop.get("before", ""), prop.get("after", "")
    left, right = prop.get("anchor_l", ""), prop.get("anchor_r", "")
    start, end = bool(prop.get("start")), bool(prop.get("end"))
    if start and end:  # paragraph-wide
        if text == after and not prop.get("deleted"):  # an empty paragraph is not a deleted one gone: by place
            return "accepted", text
        if text == before and not prop.get("added"):  # an added paragraph had no text before: _place_outcome
            return "rejected", text
        if not short:
            return None
        ref = after or before
        if ref and difflib.SequenceMatcher(None, text, ref, autojunk=False).ratio() >= 0.5:
            return "reworded", text
        return None
    if short:
        left, right = left[-SHORT_ANCHOR:], right[:SHORT_ANCHOR]
    if not start and not left and not end and not right:
        return None
    room = 3 * max(len(before), len(after)) + 40  # further apart, the anchors cannot be this span's
    pairs = [(lo, hi) for lo, hi in _between(text, left, right, start, end) if hi - lo <= room]
    if not pairs:
        return None
    if len({text[lo:hi] for lo, hi in pairs}) > 1:
        # rank the readings by how far the text around them still matches the context the edit had
        ctx_l, ctx_r = prop.get("ctx_l") or prop.get("anchor_l", ""), prop.get("ctx_r") or prop.get("anchor_r", "")
        scored = sorted(((fb._common_suffix(text[:lo], ctx_l) + fb._common_prefix(text[hi:], ctx_r), lo, hi)
                         for lo, hi in pairs), reverse=True)
        top = sorted({text[lo:hi] for score, lo, hi in scored if score == scored[0][0]}, key=len)
        if len(top) > 1 and (after in top or before in top):
            return "unresolved", None  # an "accepted"/"rejected" reading competes with another
        span = top[0]  # every tied reading is a rewording: the shortest, as a non-greedy match would take
    else:
        span = text[pairs[0][0]:pairs[0][1]]
    if span == after:
        return "accepted", after
    if span == before:
        return "rejected", before
    return "reworded", span


def _minute(date: str | None) -> str | None:
    """Word keeps a revision's date only to the minute when it saves (17:42:49Z becomes 17:42:00Z)."""
    return date[:16] if date and len(date) >= 16 else None


def _still_tracked(prop: dict, snap: Snapshot, index: int) -> bool:
    """The proposal's own revision is still in paragraph ``index``: its span does not read as something other
    than proposed (a pending revision shows its proposed text in the accepted view; a span that cannot be placed
    after the user changed its surroundings counts as not contradicting it), and a revision of the same
    type, by the same author, with matching text and the same date to the minute (when both have one) is
    there. An older revision by the same author elsewhere in the paragraph, or one with another date, does not
    keep a decided proposal pending."""
    if not prop.get("tracked") or not snap.revs.get(index):
        return False
    keys = [(r.get("type"), r.get("text"), _minute(r.get("date"))) for r in prop.get("revs") or [] if r.get("text")]
    if not keys:
        return False
    text = snap.text(index)
    seen = judge(prop, text) or judge(prop, text, short=True)
    if seen is None and (fit := fb.rewritten(prop, text, _around(snap, index))):
        seen = fit[1]  # the user changed the words around the pending revision
    author = prop.get("author")
    if seen is not None and seen[0] not in ("accepted", "unresolved"):
        # the span reads as something else: decided, whatever revision is still there, except a deletion stored
        # in several runs while a fragment of it is pending (one fragment rejected: decided when all are).
        # `n` counts every identical revision the write left, this proposal's own identical fragments included
        # ("mur|mur"), so one of them is still pending while more than the others (n - own) remain.
        here = Counter((r.type, r.author, _minute(r.date), r.text) for r in snap.revs[index])
        parts = [r for r in prop.get("revs") or [] if r.get("part") and r.get("text")]
        own = Counter((r.get("type"), _minute(r.get("date")), r.get("text")) for r in parts)
        return any(here[(r.get("type"), author, _minute(r.get("date")), r.get("text"))]
                   > r.get("n", 1) - own[(r.get("type"), _minute(r.get("date")), r.get("text"))]
                   for r in parts)
    if seen is None:
        # not placed at all (the user changed its surroundings): only the revision itself can say it is still
        # there, so it must match exactly, and as many copies as the write left must remain (an older identical
        # revision from the same minute is not this one)
        here = Counter((r.type, r.author, _minute(r.date), r.text) for r in snap.revs[index])
        return any(here[(r.get("type"), author, _minute(r.get("date")), r.get("text"))] >= r.get("n", 1)
                   for r in prop.get("revs") or [] if r.get("text"))
    for revision in snap.revs[index]:  # placed and read as proposed: a matching revision keeps it pending
        if revision.author != author:
            continue
        minute = _minute(revision.date)
        for kind, words, when in keys:
            if revision.type != kind or not (revision.text == words or fb.overlaps(revision.text, words)):
                continue
            if when and minute and when != minute:
                continue
            return True
    return False


def _near_positions(prop: dict, snap: Snapshot) -> tuple[list[int], list[int]]:
    """Positions in ``snap.indices``: the recorded paragraph and NEAR either side (nearest first), then every
    other paragraph by distance. With no recorded paragraph every position is far, in document order."""
    indices = snap.indices
    para = prop.get("para")
    if para is None or not indices:
        return [], list(range(len(indices)))
    pos = min(range(len(indices)), key=lambda k: abs(indices[k] - para))
    near = [pos]
    for step in range(1, NEAR + 1):
        near.extend(k for k in (pos - step, pos + step) if 0 <= k < len(indices))
    nearset = set(near)
    return near, sorted(set(range(len(indices))) - nearset, key=lambda k: abs(k - pos))


def _candidates(prop: dict, snap: Snapshot):
    """(index, short) in search order: the recorded paragraph and NEAR either side with forgiving anchors
    allowed, then every other paragraph by full anchors."""
    near, far = _near_positions(prop, snap)
    for k in near:
        yield snap.indices[k], False
    for k in near:
        yield snap.indices[k], True
    for k in far:
        yield snap.indices[k], False


def _blank_fill(prop: dict) -> bool:
    """A paragraph-wide proposal whose old text was empty (a blank answer line filled in), with its recorded
    neighbours (``nbt``). Its ``before`` reads as every blank paragraph, so a blank one near its recorded place
    is not evidence that it was rejected unless those neighbours still sit around it. An older proposal with no
    ``nbt`` keeps the plain candidate order: without its place, an identical answer elsewhere must not be taken
    for it (a real rejection would read as accepted)."""
    return (bool(prop.get("start")) and bool(prop.get("end")) and not (prop.get("before") or "").strip()
            and not prop.get("added") and not prop.get("deleted") and bool(prop.get("nbt")))


def _in_order(a: list[str], b: list[str]) -> int:
    """How many items ``a`` and ``b`` share in the same order (longest common subsequence)."""
    row = [0] * (len(b) + 1)
    for x in a:
        prev = 0
        for j, y in enumerate(b, 1):
            prev, row[j] = row[j], prev + 1 if x == y else max(row[j], row[j - 1])
    return row[-1]


_PENDING = ("unresolved", None)
_UNPLACED = "unplaced"  # only blank paragraphs away from the recorded neighbours: no verdict, not removed either


def _blank_fill_verdict(prop: dict, snap: Snapshot, *, short: bool):
    """(index, verdict) for a blank-fill proposal (``_blank_fill``); ``(None, _PENDING)`` when it must stay open;
    ``_UNPLACED`` when the only readings near are blank paragraphs that are not placed (the caller may still find
    the user's rewrite, else it stays open); None when nothing near or far reads as it (the caller's fallback).

    Every verdict in the NEAR window is collected before choosing. A blank paragraph counts as the proposal rejected
    only where the recorded neighbours (``nbt``) still sit around it: exactly them ("exactly"), or most of them in
    their order. A paragraph reading as the proposal (accepted, or reworded with ``short``) beats a blank one unless
    the blank one sits better among those neighbours. Both exactly placed, or a blank one better placed but not
    exactly while the proposed text itself is there: left open. A paragraph further away that stands exactly among
    those neighbours (shifted past NEAR, or a repeated block) decides only when no near paragraph is placed; against
    a near one placed but not exactly, a different verdict is left open; against a near one placed exactly, too, when
    more paragraphs stand exactly among them now than when the proposal was recorded (``nbt_n``: a copy of the
    block was added). Without a placed blank paragraph the best-placed reading is taken. Nothing decisive near: the
    same choice over every other paragraph by full readings, where a blank one counts only exactly placed. Blank
    paragraphs that are not placed decide nothing: unless the caller finds the user's rewrite, the proposal stays
    open rather than being called rejected or removed."""
    context = prop.get("nbt") or None
    indices = snap.indices
    windows: dict[int, list[str]] = {}

    def window(k: int) -> list[str]:
        if k not in windows:
            near = indices[max(0, k - NEAR):k] + indices[k + 1:k + 1 + NEAR]
            windows[k] = fb.neighbour_texts(map(snap.text, near))
        return windows[k]

    def rank(k: int) -> tuple[bool, int]:
        """(exactly the recorded neighbours, how many of them are around in their order)."""
        if context is None:
            return False, 0
        here = window(k)
        return here == context, _in_order(here, context)

    def placed(k: int) -> bool:
        if context is None:
            return True  # no context recorded: nothing contradicts the place
        exact, score = rank(k)
        return exact or 2 * score > len(context)

    def choose(pool: list[tuple[int, int, tuple]]):
        """pool: (order, position, verdict). The pick, _PENDING, or None (only unplaced blank readings)."""
        reads = [item for item in pool if item[2][0] != "rejected"]
        blanks = [item for item in pool if item[2][0] == "rejected" and placed(item[1])]
        if not reads and not blanks:
            return None
        best_n = max(reads, key=lambda item: (rank(item[1]), item[2][0] == "accepted", -item[0]), default=None)
        if not blanks:
            return indices[best_n[1]], best_n[2]
        best_r = max(blanks, key=lambda item: (rank(item[1]), -item[0]))
        if not reads:
            return indices[best_r[1]], best_r[2]
        kn, kr = rank(best_n[1]), rank(best_r[1])
        if kn[0] and kr[0]:
            return None, _PENDING  # both stand exactly among the recorded neighbours
        if kn[0] or kr <= kn:
            return indices[best_n[1]], best_n[2]
        if kr[0] or not any(item[2][0] == "accepted" for item in reads):
            return indices[best_r[1]], best_r[2]
        return None, _PENDING  # the proposed text is there, a blank paragraph sits better but not exactly: open

    near, far = _near_positions(prop, snap)
    full = [(order, k, v) for order, k in enumerate(near) if (v := judge(prop, snap.text(indices[k]))) is not None]
    pool = list(full)
    if short:
        seen = {k for _, k, _ in full}
        pool += [(len(near) + order, k, v) for order, k in enumerate(near)
                 if k not in seen and (v := judge(prop, snap.text(indices[k]), short=True)) is not None]
    far_pool = [(order, k, v) for order, k in enumerate(far) if (v := judge(prop, snap.text(indices[k]))) is not None]
    picked = choose(pool)
    # A paragraph further away that stands exactly among the recorded neighbours: an outside edit shifted it past
    # NEAR, or it belongs to a repeated block. It decides only when no near candidate can be the proposal's place
    # (a twin of the answer nearby must not hide its rejection). A near candidate keeping most of them in order
    # against a far one keeping all leaves a different verdict open: which of the two is the proposal's place is
    # not known. A near candidate standing exactly among them decides, unless more places stand exactly among them
    # now than when the proposal was recorded (``nbt_n``): a copy of the block was added, the near one may be that
    # copy (the original shifted past NEAR), so a different verdict is left open too. An older record without
    # ``nbt_n``: the near one decides.
    far_exact = [(2 * len(near) + order, k, v) for order, k, v in far_pool if rank(k)[0]]
    if far_exact and not any(rank(k)[0] for _, k, _ in pool):
        widened = choose(pool + far_exact)
        if not any(placed(k) for _, k, _ in pool):
            picked = widened
        elif widened[1] != picked[1]:
            picked = None, _PENDING
    elif (far_exact and isinstance(prop.get("nbt_n"), int) and any(v != picked[1] for _, _, v in far_exact)
          and sum(rank(k)[0] for k in range(len(indices))) > prop["nbt_n"]):
        picked = None, _PENDING
    if picked is not None:
        return picked
    # Far from the recorded place a blank paragraph counts only with exactly the recorded neighbours around it.
    picked = choose([item for item in far_pool if item[2][0] != "rejected" or rank(item[1])[0]])
    if picked is not None:
        return picked
    return _UNPLACED if pool or far_pool else None


def reconcile_props(props: list[dict], snap: Snapshot, *, via: str, sha: str | None,
                    watched: list[dict] = ()) -> list[dict]:
    """Outcome events for the open proposals that are decided in ``snap``; pending ones give none.

    ``watched`` proposals were already accepted: they get a new outcome only when their span is found, by its
    full anchors, reading something else (the user reverted or reworded kept text). Not finding them records
    nothing, since later edits nearby may have changed their anchors."""
    out = []
    for prop in watched:
        if prop.get("deleted"):  # by place: its text typed back between its neighbours reverts the deletion
            decided = _place_outcome(prop, snap)
            if decided is not None and decided[0] != "accepted":
                out.append(fb.make_outcome(prop, decided[0], kept=decided[1], via=via, sha=sha, para=decided[2]))
            continue
        if _blank_fill(prop):
            found = _blank_fill_verdict(prop, snap, short=False)
            if found not in (None, _UNPLACED) and found[1][0] not in ("accepted", "unresolved"):
                out.append(fb.make_outcome(prop, found[1][0], kept=found[1][1], via=via, sha=sha, para=found[0]))
            continue
        for index, short in _candidates(prop, snap):
            if short:
                continue
            verdict = judge(prop, snap.text(index))
            if verdict is not None:
                if verdict[0] not in ("accepted", "unresolved"):
                    out.append(fb.make_outcome(prop, verdict[0], kept=verdict[1], via=via, sha=sha, para=index))
                break
    for prop in props:
        if prop.get("tracked") and any(_still_tracked(prop, snap, index) for index, short in _candidates(prop, snap)
                                       if not short and abs(index - (prop.get("para") or 0)) <= NEAR):
            continue  # pending: the model's revision is still in the document
        if prop.get("added") or prop.get("deleted"):  # identity by place between its neighbours, never by text
            if (decided := _place_outcome(prop, snap)) is not None:
                outcome, kept, index = decided
                out.append(fb.make_outcome(prop, outcome, kept=kept, via=via, sha=sha, para=index))
            continue
        found, unplaced = None, False
        if _blank_fill(prop):
            found = _blank_fill_verdict(prop, snap, short=True)
            if found == _UNPLACED:
                found, unplaced = None, True
        else:
            for index, short in _candidates(prop, snap):
                verdict = judge(prop, snap.text(index), short=short)
                if verdict is not None:
                    found = (index, verdict)
                    break
        if found is None:  # the user's rewrite took the anchors too: the paragraph nearby that still reads like it
            near = [index for index, short in _candidates(prop, snap) if short]
            fits = [(fit[0], -abs(index - (prop.get("para") or 0)), index, fit[1])
                    for index in near if (fit := fb.rewritten(prop, snap.text(index), _around(snap, index)))]
            if fits:
                _, _, index, verdict = max(fits)
                found = (index, verdict)
        if found is None and unplaced:
            continue  # a blank paragraph near, away from the recorded neighbours: stays open
        if found is None:
            out.append(fb.make_outcome(prop, "removed", kept=None, via=via, sha=sha))
            continue
        index, (outcome, kept) = found
        if outcome == "unresolved":
            continue  # the text fits two readings: stays open, judged again when the document next changes
        out.append(fb.make_outcome(prop, outcome, kept=kept, via=via, sha=sha, para=index))
    return out


def _withdrawn(prop: dict, lost: set[tuple]) -> bool:
    """A text edit (not an accept/reject) removed this proposal's own insertion and nothing of the proposal is
    left: the model withdrew it (``removed`` via model, not a user outcome)."""
    keys = [(r.get("type"), r.get("author"), r.get("date"), r.get("text")) for r in prop.get("revs") or []
            if r.get("text")]
    return bool(keys) and all(key in lost for key in keys) and any(key[0] == "ins" for key in keys)


def _lineage(upto: list[dict]) -> set[str]:
    """Ids of the rounds the last of ``upto`` was built on, itself included: each round's ``prev`` (the chain's
    open round when it was proposed). A round on an undo branch the chain later left is not among them. A log
    written before ``prev`` was recorded counts every round as the lineage (the old behaviour)."""
    by_id = {e["id"]: e for e in upto}
    if any("prev" not in e for e in upto[1:]):
        return set(by_id)
    out, cur = set(), upto[-1]
    while cur is not None and cur["id"] not in out:
        out.add(cur["id"])
        cur = by_id.get(cur.get("prev"))
    return out


def _restore(events: list[dict], sha: str) -> list[dict]:
    """``undo_document`` (or undoing that undo) returned the document to ``sha``, the state right after an earlier
    MCP write W: each chain W left open goes back to W's round. Its later rounds (undecided, or decided by a
    later MCP write: a withdrawal, an MCP accept/reject) are ``undone`` and W's round, if a later round
    superseded it, an undo undid it or a later MCP write decided it, is ``reopened``; an open round before it is
    superseded again (an undone one too, when W's round built on it), or ``undone`` when it lies on an undo branch
W's round did not build on. Chains first
    proposed after W, and chains the user decided since, are left to the usual reconcile (an undone first round
    reads as rejected via model)."""
    sha = sha.lower()
    props = [(n, e) for n, e in enumerate(events) if e.get("event") == "proposal"]
    hits = [n for n, e in props if (e.get("sha_after") or "").lower() == sha]
    if not hits:
        return []
    mark = hits[-1]
    last = fb.last_outcomes(events)
    at = {e["id"]: n for n, e in enumerate(events) if e.get("event") == "outcome"}
    state = lambda e: (last.get(e["id"]) or {}).get("outcome")  # noqa: E731

    def taken_back(e: dict) -> bool:  # decided by a later MCP write, which the restore takes back
        outcome = last.get(e["id"])
        return (bool(outcome) and outcome["outcome"] in fb.OUTCOMES and outcome.get("via") == "model"
                and at[e["id"]] > mark and (outcome.get("sha") or "").lower() != sha)

    by_chain: dict[str, list[tuple[int, dict]]] = {}
    for n, e in props:
        if e.get("chain"):
            by_chain.setdefault(e["chain"], []).append((n, e))
    out: list[dict] = []
    for items in by_chain.values():
        upto = [e for n, e in items if n <= mark]
        later = [e for n, e in items if n > mark]
        if not upto:
            continue  # first proposed after W
        target = upto[-1]
        if not later and state(target) != "undone" and not taken_back(target):
            continue  # nothing of this chain happened after W
        if any(state(e) not in (None, "superseded", "undone", "reopened") and not taken_back(e)
               for e in later + [target]):
            continue  # the user decided since: the usual reconcile judges it
        out.extend(fb.make_outcome(e, "undone", kept=None, via="model", sha=sha) for e in later
                   if state(e) != "undone")
        if state(target) in ("superseded", "undone") or taken_back(target):
            out.append(fb.make_outcome(target, "reopened", kept=None, via="model", sha=sha))
            line = _lineage(upto)
            # with ``prev`` recorded, a round W's round built on comes back even when an earlier branch switch
            # undid it (old logs count every round as the lineage, so an undone one stays undone there)
            back = (None, "reopened", "undone") if all("prev" in e for e in upto[1:]) else (None, "reopened")
            for e in upto[:-1]:
                if e["id"] not in line and state(e) in (None, "superseded", "reopened"):
                    out.append(fb.make_outcome(e, "undone", kept=None, via="model", sha=sha))
                elif e["id"] in line and state(e) in back:
                    out.append(fb.make_outcome(e, "superseded", kept=None, via="model", sha=sha))
    return out


def _sync(sha: str, started: float, **extra) -> dict:
    event = {"v": fb.VERSION, "event": "sync", "time": fb.now_iso(), "sha": sha.lower(),
             "ms": round((time.perf_counter() - started) * 1000, 1)}
    event.update({k: v for k, v in extra.items() if v is not None})
    return event


def _pending(events: list[dict]) -> int:
    return sum(1 for record in fb.chains(events) if record["pending"])


def _warn(stage: str, exc: Exception) -> dict:
    return {"edit_feedback": None,
            "feedback_warning": f"edit feedback not recorded ({stage}): {type(exc).__name__}: {exc}"[:300]}


# ---------------------------------------------------------------- entry points (never raise)

def after_publish(target, *, before_file, sha_before: str, sha_after: str, undo_token: str | None,
                  tool: str | None, op: str | None, op_types: str | None = None) -> dict:
    """After an MCP write of ``target`` (``before_file`` holds the state it replaced): first judge open proposals
    against that earlier state if the user changed the document since the log last saw it, then record the
    write's own accept/reject of earlier proposals (via model) and its new proposals. ``op_types`` (the applied op
    types without op ids, from edit_document) classifies the write; without it ``op`` is read as before."""
    kinds = op_types if op_types is not None else op
    try:
        started = time.perf_counter()
        log = doc_log(target)
        sha_before, sha_after = sha_before.lower(), sha_after.lower()
        with log.locked():
            events = log.events()
            new: list[dict] = []
            user_outcomes: list[dict] = []
            opens, watched = fb.open_proposals(events), fb.watched_proposals(events)
            if (opens or watched) and fb.last_sha(events) not in (None, sha_before):
                user_outcomes = reconcile_props(opens, Snapshot(before_file), via="user", sha=sha_before,
                                                watched=watched)
                new.extend(user_outcomes)
                new.append(_sync(sha_before, started, checked=len(opens) + len(watched)))
            if ck.file_sha256(target) != sha_after:
                log.append(new)
                return {"edit_feedback": None,
                        "feedback_warning": "edit feedback: the document changed again right after this write; "
                                            "this write's proposals were not recorded"}
            fields, lost, after = capture(before_file, target)
            if _format_only(tool, kinds):
                fields, lost = [], set()
            for item in fields:  # new proposals first: one that rewrites an open proposal supersedes it
                para_text = item.pop("_btext", None)
                proposal = fb.make_proposal(path=str(Path(target).resolve()), tool=tool, op=(op or "")[:200] or None,
                                            undo_token=undo_token, sha_before=sha_before, sha_after=sha_after, **item)
                new.extend(fb.assign_chain(events + new, proposal, para_text))
                new.append(proposal)
            if lost:  # the write accepted or rejected tracked proposals of its own (via model)
                fresh = {e["id"] for e in new if e.get("event") == "proposal"}
                touched = [p for p in fb.open_proposals(events + new) if p["id"] not in fresh and p.get("tracked")
                           and any((r["type"], r["author"], r["date"], r["text"]) in lost for r in p.get("revs") or [])]
                # edit_document's op names its applied op types; a withdrawal needs a tracked text edit and no
                # accept/reject in the same write (the two leave the same document)
                text_edit = bool(kinds and _TEXT_EDIT_OP.search(kinds)) and not _RESOLVE_OP.search(kinds)
                gone = [p for p in touched if text_edit and _withdrawn(p, lost)]
                new.extend(fb.make_outcome(p, "removed", kept=None, via="model", sha=sha_after) for p in gone)
                new.extend(reconcile_props([p for p in touched if p not in gone], after, via="model", sha=sha_after))
            if new or log.exists():
                new.append(_sync(sha_after, started, checked=after.built, proposals=len(fields)))
                log.append(new)
            pending = _pending(events + new)
        return {"edit_feedback": fb.summary_line(fb.chain_outcomes(user_outcomes, events), pending),
                "feedback_warning": None}
    except Exception as exc:  # noqa: BLE001 - feedback never fails the tool
        return _warn("capture", exc)


def on_sight(path, *, via: str = "user") -> dict:
    """A tool read or changed ``path`` without an MCP publish: if the document is not the state the log last
    accounted for, judge its open proposals now. Cheap when there is nothing to do (no log, nothing open, or
    the same sha)."""
    try:
        if path is None:
            return _none()
        p = Path(path)
        if p.suffix.lower() != ".docx" or not p.exists():
            return _none()
        log = doc_log(p)
        if not log.exists():
            return _none()
        events = log.events()
        # an undo can reopen a chain a later MCP write decided, so it looks even when nothing is open
        if not fb.open_proposals(events) and not fb.watched_proposals(events) and not (
                via == "model" and any(e.get("event") == "proposal" for e in events)):
            return _none()
        sha = ck.file_sha256(p)
        if sha == fb.last_sha(events):
            return _none()
        started = time.perf_counter()
        with log.locked():
            events = log.events()
            if sha == fb.last_sha(events):
                return _none()
            restored = _restore(events, sha) if via == "model" else []
            opens, watched = fb.open_proposals(events + restored), fb.watched_proposals(events + restored)
            if not (opens or watched or restored):
                return _none()
            snap = Snapshot(p)
            outcomes = reconcile_props(opens, snap, via=via, sha=sha, watched=watched)
            new = restored + outcomes + [_sync(sha, started, checked=snap.built, open=len(opens),
                                               watched=len(watched))]
            log.append(new)
            pending = _pending(events + new)
        return {"edit_feedback": fb.summary_line(fb.chain_outcomes(outcomes, events), pending),
                "feedback_warning": None}
    except Exception as exc:  # noqa: BLE001 - feedback never fails the tool
        return _warn("reconcile", exc)


def tool_fields(path, records: list[dict], *, via: str = "user") -> dict:
    """The feedback fields for a per-document tool result. From this call's publications of ``path`` when it
    wrote one, else a reconcile on sight of ``path``; warnings from every other publication of the call (a
    versioned write records its feedback on the output) are added. Only non-empty fields are returned."""
    try:
        target = ck.norm_path(path) if path else None
    except Exception:  # noqa: BLE001
        target = None

    def same(record) -> bool:
        try:
            return bool(target) and ck.norm_path(record["path"]) == target
        except Exception:  # noqa: BLE001
            return False

    noted = [r for r in records if r.get("path") and ("edit_feedback" in r or "feedback_warning" in r)]
    mine = [r for r in noted if same(r)]
    others = [r for r in noted if not same(r)]
    if mine:
        fields = {"edit_feedback": next((r["edit_feedback"] for r in mine if r.get("edit_feedback")), None),
                  "feedback_warning": None}
    else:  # no write of this document (or one that changed nothing): settle what the user did
        fields = dict(on_sight(path, via=via))
    if not fields.get("edit_feedback"):
        fields["edit_feedback"] = next((r["edit_feedback"] for r in others if r.get("edit_feedback")), None)
    warnings = [fields.get("feedback_warning")] + [r.get("feedback_warning") for r in mine + others]
    fields["feedback_warning"] = " ".join(dict.fromkeys(w for w in warnings if w)) or None
    return {k: v for k, v in fields.items() if v}
