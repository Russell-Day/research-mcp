"""Edit feedback store: which model edits the user accepted, rejected or reworded (plan 829).

Every edit an MCP server makes to a user's document is a *proposal*. When the server next sees the file, the
user's outcome for each open proposal is recorded. This module is the format-neutral part: the folder layout,
the event log, chain collapsing, the default view, the summary line and the ``.md`` mirror. It knows nothing
about Word or PowerPoint (``edit_feedback_word`` holds the Word capture and reconcile; powerpoint-mcp vendors
this file and adds its own). It imports only the standard library, ``textclean`` (display cleaning) and the OS
file lock from ``checkpoints``.

Layout, one pair of files per document (two servers may share the folder; each file pair has its own OS lock):

    <root>/<stem>-<doc-id8>.jsonl   every event, append-only (proposal, outcome, seen, sync)
    <root>/<stem>-<doc-id8>.md      the net outcomes, newest first, for a person to read
    <root>/<stem>-<doc-id8>.lock    OS lock held while a server reads-and-appends

Precision rule: repeated proposals on one span collapse into one *chain*: the first model proposal, what the
user finally kept, and ``rounds`` (how many proposals). The log keeps every round; the default view never
shows intermediate rounds, lists only outcomes the model has not yet seen (newest first, at most ``limit``,
each item at most ``ITEM_MAX`` characters) and counts pending proposals instead of listing them.

Event schema (``v`` = 1; every event has ``v``, ``event``, ``time``):

    proposal  id, chain, round, prev (the round it rewrites; None for a first round), path, tool, op, undo_token,
              sha_before, sha_after, para, where, anchor_l, anchor_r, start, end, before, after, tracked, author,
              date, revs[{type, author, date, text}]
              ctx_l, ctx_r, nbt (paragraphs near the edited one, cut); decks add slide, shape, paras
    outcome   id (the proposal), chain, outcome (accepted|rejected|reworded|removed|superseded|undone|reopened),
              kept, via (user|model), sha, para. A proposal's last outcome counts. ``superseded``: a later round of
              the model replaced it; ``undone``: undo_document took its write back (the round no longer counts);
              ``reopened``: that undo made it the chain's open round again.
    seen      chains [chain ids the model was shown]
    sync      sha (the document state the outcomes were judged on), ms, checked, warning

Matching uses the stored raw text; text shown to the model goes through ``textclean.clean``.
"""

from __future__ import annotations

import difflib
import json
import os
import re
import uuid
from collections import Counter
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from .checkpoints import _hold_file, _release
from .textclean import clean

VERSION = 1
ITEM_MAX = 300          # characters per item in any view
DEFAULT_LIMIT = 10      # items in the default view
KEPT_MAX = 300          # characters of the user's kept text stored per outcome
ANCHOR_MAX = 40         # characters of unchanged text kept on each side of a proposal
LOCK_TIMEOUT = 1.0       # seconds; feedback writes take milliseconds, and a busy log only costs a warning
OUTCOMES = ("accepted", "rejected", "reworded", "removed")   # what the user decided; the rest are internal
OPEN = (None, "superseded", "reopened")   # last outcomes that leave a chain's last round undecided


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def new_id(prefix: str) -> str:
    return f"{prefix}{uuid.uuid4().hex[:8]}"


def file_stem(stem: str, path_key: str) -> str:
    """``<stem>-<doc-id8>``: the readable name plus 8 hex characters of the normalised path's sha256."""
    import hashlib

    safe = re.sub(r"[^\w.-]", "_", stem)[:40] or "document"
    return f"{safe}-{hashlib.sha256(path_key.encode('utf-8')).hexdigest()[:8]}"


# ---------------------------------------------------------------- one document's log

class DocLog:
    """The event log of one document. Read with ``events()``; write inside ``with log.locked():``."""

    def __init__(self, root: Path, name: str, path: str | None = None) -> None:
        self.root = Path(root)
        self.name = name
        self.path = path
        self.jsonl = self.root / f"{name}.jsonl"
        self.md = self.root / f"{name}.md"
        self.lockfile = self.root / f"{name}.lock"

    def exists(self) -> bool:
        return self.jsonl.exists()

    @contextmanager
    def locked(self):
        self.root.mkdir(parents=True, exist_ok=True)
        fd = _hold_file(self.lockfile, LOCK_TIMEOUT, f"edit feedback for {self.name} is busy")
        try:
            yield self
        finally:
            _release(fd)

    def events(self) -> list[dict]:
        if not self.jsonl.exists():
            return []
        out = []
        with self.jsonl.open("r", encoding="utf-8") as handle:
            for line in handle:
                line = line.strip()
                if not line:
                    continue
                try:
                    out.append(json.loads(line))
                except ValueError:
                    continue  # a torn last line from a crash: skip it, keep the rest
        return out

    def append(self, events: list[dict], *, md: bool = True) -> None:
        """Append ``events`` (caller holds ``locked()``) and rewrite the ``.md`` mirror."""
        if not events:
            return
        self.root.mkdir(parents=True, exist_ok=True)
        torn = False
        if self.jsonl.exists() and self.jsonl.stat().st_size:
            with self.jsonl.open("rb") as raw:  # a crash can leave a last line without its newline
                raw.seek(-1, os.SEEK_END)
                torn = raw.read(1) != b"\n"
        with self.jsonl.open("a", encoding="utf-8", newline="\n") as handle:
            if torn:
                handle.write("\n")  # end the partial line so it cannot swallow the first new event
            for event in events:
                handle.write(json.dumps(event, ensure_ascii=False, separators=(",", ":")) + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        if md and any(e.get("event") in ("proposal", "outcome") for e in events):
            self.write_md()

    def write_md(self) -> None:
        text = render_md(self.events(), self.path or self.name)
        tmp = self.md.with_suffix(f".{uuid.uuid4().hex[:6]}.tmp")
        tmp.write_text(text, encoding="utf-8", newline="\n")
        os.replace(tmp, self.md)


def last_sha(events: list[dict]) -> str | None:
    """The document state the log last accounted for: the latest proposal's sha_after or sync's sha."""
    for event in reversed(events):
        if event.get("event") == "proposal" and event.get("sha_after"):
            return event["sha_after"].lower()
        if event.get("event") == "sync" and event.get("sha"):
            return event["sha"].lower()
    return None


# ---------------------------------------------------------------- proposals and chains

def make_proposal(**fields) -> dict:
    event = {"v": VERSION, "event": "proposal", "id": new_id("p-"), "time": now_iso()}
    event.update(fields)
    return event


def make_outcome(proposal: dict, outcome: str, *, kept: str | None, via: str, sha: str | None,
                 para: int | None = None) -> dict:
    return {"v": VERSION, "event": "outcome", "id": proposal["id"], "chain": proposal.get("chain"),
            "time": now_iso(), "outcome": outcome, "kept": None if kept is None else kept[:KEPT_MAX],
            "via": via, "sha": sha, "para": para}


def last_outcomes(events: list[dict]) -> dict[str, dict]:
    """Each proposal's last outcome event, by proposal id."""
    return {e["id"]: e for e in events if e.get("event") == "outcome"}


def _state(last: dict[str, dict], proposal: dict) -> str | None:
    return (last.get(proposal["id"]) or {}).get("outcome")


def active(events: list[dict]) -> list[dict]:
    """Proposals not taken back by an undo (last outcome ``undone``), in log order."""
    last = last_outcomes(events)
    return [e for e in events if e.get("event") == "proposal" and _state(last, e) != "undone"]


def open_proposals(events: list[dict]) -> list[dict]:
    """Proposals with no outcome yet, or reopened by an undo, oldest first."""
    last = last_outcomes(events)
    return [e for e in events if e.get("event") == "proposal" and _state(last, e) in (None, "reopened")]


def watched_proposals(events: list[dict], limit: int = 20) -> list[dict]:
    """The last proposal of each chain whose last outcome is ``accepted`` (newest ``limit``): text the user kept
    can still be reverted or reworded later, which is recorded as a new outcome of the same proposal."""
    last_prop: dict[str, dict] = {}
    last_out = last_outcomes(events)
    for event in active(events):
        if event.get("chain"):
            last_prop[event["chain"]] = event
    kept = [(last_out[p["id"]]["time"], p) for p in last_prop.values()
            if p["id"] in last_out and last_out[p["id"]]["outcome"] == "accepted"]
    kept.sort(key=lambda pair: pair[0], reverse=True)
    return [p for _, p in kept[:limit]]


def _current_text(chain_events: list[dict]) -> tuple[str | None, dict]:
    """The span's text as it now stands in the document after a chain's last event, and the last proposal (an
    undone round does not count)."""
    props = active(chain_events) or [e for e in chain_events if e.get("event") == "proposal"]
    last_prop = props[-1]
    last_out = last_outcomes(chain_events).get(last_prop["id"])
    if last_out is None or last_out["outcome"] in OPEN:
        return last_prop.get("after", ""), last_prop
    if last_out["outcome"] == "accepted":
        return last_prop.get("after", ""), last_prop
    if last_out["outcome"] == "rejected":
        return last_prop.get("before", ""), last_prop
    if last_out["outcome"] == "reworded":
        return last_out.get("kept") or "", last_prop
    return None, last_prop  # removed: the span is gone


def overlaps(a: str, b: str) -> bool:
    """Two spans are "the same span" when one holds the other, or they share a run of at least half the
    shorter one (and at least 4 characters)."""
    a, b = (a or "").strip(), (b or "").strip()
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    short = min(len(a), len(b))
    if short < 8:
        return False
    match = difflib.SequenceMatcher(None, a, b, autojunk=False).find_longest_match(0, len(a), 0, len(b))
    return match.size >= max(4, short // 2)


def _common_suffix(a: str, b: str) -> int:
    n = 0
    while n < len(a) and n < len(b) and a[-1 - n] == b[-1 - n]:
        n += 1
    return n


def _common_prefix(a: str, b: str) -> int:
    n = 0
    while n < len(a) and n < len(b) and a[n] == b[n]:
        n += 1
    return n


REWRITE_MIN = 0.5       # share of the edited window (context + span) still in a paragraph for it to count as that one
REWRITE_RUN = 4         # characters a shared run needs to count (a parallel sentence shares letters, not runs)
REWRITE_SCAN = 5000     # longer paragraphs are not compared character by character
NEIGHBOUR_MAX = 200     # characters kept of each nearby paragraph (``nbt``), to tell a neighbour from the edited one


def neighbour_key(text: str) -> str:
    """A paragraph's identity, bounded: its text up to NEIGHBOUR_MAX characters, a longer one cut there plus a
    hash of the whole text (two paragraphs that share only an opening are not the same paragraph)."""
    if len(text) <= NEIGHBOUR_MAX:
        return text
    import hashlib

    return f"{text[:NEIGHBOUR_MAX]}#{hashlib.sha1(text.encode('utf-8')).hexdigest()[:12]}"


def neighbour_texts(texts) -> list[str]:
    """What a proposal keeps (``nbt``) of the non-empty paragraphs near the edited one, as the write left them."""
    return [neighbour_key(text) for text in texts if text.strip()]


def _share(window: str, text: str) -> tuple[float, difflib.SequenceMatcher]:
    """How much of ``window`` is in ``text``, in order, counting shared runs of REWRITE_RUN characters or more
    (single shared letters are noise)."""
    matcher = difflib.SequenceMatcher(None, window, text, autojunk=False)
    run = min(REWRITE_RUN, len(window))
    return sum(b.size for b in matcher.get_matching_blocks() if b.size >= run) / len(window), matcher


def rewritten(prop: dict, text: str, others) -> tuple[float, tuple[str, str | None]] | None:
    """(share, verdict) when ``text`` is still recognisably the paragraph the proposal edited, though the user's
    rewrite also changed its anchors (so no exact span fits). The verdict is ``verdict(prop, kept)`` for the
    user's text where the span was, widened to whole words ("" when the user deleted it); ``share`` is how much
    of the edited window (context + proposed or old text) survives in ``text`` (``_share``).

    Identity comes from the paragraphs that were near the edited one (``nbt``), looked for among ``others`` (the
    paragraphs around ``text`` now, ``text`` excluded) by ``neighbour_key``. All found: ``text`` is the edited
    paragraph. One missing and ``text`` is exactly it: ``text`` is that neighbour (the edited paragraph was
    deleted and it slid into place), None. One missing otherwise: ``("unresolved", None)``, since either
    paragraph may have been deleted or reworked into the other's likeness. None also under REWRITE_MIN, and for
    proposals recorded before ``nbt`` existed."""
    if not text or len(text) > REWRITE_SCAN or "nbt" not in prop:
        return None
    ctx_l = prop.get("ctx_l") or prop.get("anchor_l") or ""
    ctx_r = prop.get("ctx_r") or prop.get("anchor_r") or ""
    rows = []  # the window as the model left it first: the user rewrote that text, so the span is placed in it
    for span in (prop.get("after") or "", prop.get("before") or ""):
        window = ctx_l + span + ctx_r
        if window:
            rows.append((*_share(window, text), len(ctx_l), len(ctx_l) + len(span)))
    share = max((row[0] for row in rows), default=0.0)
    if share < REWRITE_MIN:
        return None
    present = Counter(neighbour_key(other) for other in others if other.strip())
    missing = []
    for neighbour in prop["nbt"]:
        if present[neighbour] > 0:
            present[neighbour] -= 1
        else:
            missing.append(neighbour)
    if neighbour_key(text) in missing:
        return None  # it is that neighbour, unchanged, in the deleted paragraph's place
    if missing:
        return share, ("unresolved", None)
    _, matcher, s, e = rows[0]
    lo, hi = None, None
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            if s == e and i1 <= s <= i2:
                a = b = j1 + s - i1  # a deletion's empty span, inside text the user left alone
            else:
                a, b = max(i1, s), min(i2, e)
                if a >= b:
                    continue
                a, b = j1 + a - i1, j1 + b - i1
        elif (i1 < s < i2 or i1 == i2 == s) if s == e else (i1 < e and i2 > s) or (i1 == i2 and s <= i1 <= e):
            a, b = j1, j2  # the user's change inside the span (an empty one: across it, or typed right there)
        else:
            continue
        lo, hi = (a, b) if lo is None else (min(lo, a), max(hi, b))
    if lo is None:
        return None
    if lo == hi:
        return share, verdict(prop, "")  # nothing of the user's where the span was
    while 0 < lo < len(text) and text[lo - 1].isalnum() and text[lo].isalnum():
        lo -= 1  # a cut inside a word: take the whole word
    while 0 < hi < len(text) and text[hi - 1].isalnum() and text[hi].isalnum():
        hi += 1
    return share, verdict(prop, text[lo:hi].strip())


def verdict(prop: dict, kept: str) -> tuple[str, str]:
    """accepted / rejected / reworded for the user's ``kept`` text at the proposal's span. Spans compare with
    their whitespace normalised (an inserted " many" kept reads as "many")."""
    def norm(value) -> str:
        return " ".join((value or "").split())

    if norm(kept) == norm(prop.get("after")):
        return "accepted", kept
    if norm(kept) == norm(prop.get("before")):
        return "rejected", kept
    return "reworded", kept


def _side_ok(agree: int, anchor: str) -> bool:
    return bool(anchor) and agree >= min(len(anchor), 8)


def _locate(text: str, current: str, prop: dict, *, both: bool) -> tuple[int, int] | None:
    """Where a chain's current text sits in ``text``: the occurrence whose surroundings agree best with the
    chain's anchors (at least 8 characters, or the whole anchor, on one side; on both sides when ``both``).
    None when it is absent or two occurrences agree equally well."""
    left, right = prop.get("anchor_l") or "", prop.get("anchor_r") or ""
    n = len(current)
    if n:
        spots, at = [], text.find(current)
        while at >= 0:
            spots.append(at)
            at = text.find(current, at + 1)
    else:
        spots = list(range(len(text) + 1))
    scored = []
    for at in spots:
        lo, ro = _common_suffix(text[:at], left), _common_prefix(text[at + n:], right)
        l_ok = _side_ok(lo, left) or (not left and prop.get("start") and at == 0)
        r_ok = _side_ok(ro, right) or (not right and prop.get("end") and at + n == len(text))
        if (l_ok and r_ok) if both else (l_ok or r_ok):
            scored.append((lo + ro, at))
    if not scored:
        return None
    scored.sort(reverse=True)
    if len(scored) > 1 and scored[1][0] == scored[0][0]:
        return None  # two places fit equally well: not evidence of the same span
    return scored[0][1], scored[0][1] + n


def _anchors_agree(a: dict, b: dict) -> bool:
    """Without the paragraph text: both sides' anchors must agree (8 characters, or the shorter anchor)."""
    al, bl, ar, br = (a.get("anchor_l") or "", b.get("anchor_l") or "", a.get("anchor_r") or "",
                      b.get("anchor_r") or "")
    left = _common_suffix(al, bl) >= min(len(al), len(bl), 8) and bool(al) == bool(bl)
    right = _common_prefix(ar, br) >= min(len(ar), len(br), 8) and bool(ar) == bool(br)
    return left and right


def assign_chain(events: list[dict], proposal: dict, para_text: str | None = None) -> list[dict]:
    """Set ``proposal['chain']`` / ``['round']``: join the chain whose current text this proposal rewrites, else
    start a new chain. Returns ``superseded`` outcomes for a joined chain's still-open proposal (the model
    rewrote its own undecided text).

    The same span means the same place, not the same words: with ``para_text`` (the paragraph as the write
    found it) and the proposal's ``b_start``/``b_end`` offsets in it, the chain's current text is located by
    its anchors and must overlap the proposal's span (one anchor side agreeing in the same paragraph, both in
    another). Without them, the paragraph must be the same, the texts must overlap and both anchors agree. A
    repeated word elsewhere in the paragraph is a separate chain."""
    chains: dict[str, list[dict]] = {}
    for event in events:
        if event.get("chain") and event.get("event") in ("proposal", "outcome"):
            chains.setdefault(event["chain"], []).append(event)
    best = None
    positioned = para_text is not None and proposal.get("b_start") is not None
    for chain_id, chain_events in chains.items():
        current, last_prop = _current_text(chain_events)
        if current is None:
            continue
        same_para = proposal.get("para") is not None and _chain_para(chain_events) == proposal.get("para")
        if positioned:
            spot = _locate(para_text, current, last_prop, both=not same_para)
            if spot is None:
                continue
            b1, b2 = proposal["b_start"], proposal["b_end"]
            hit = (spot[0] < b2 and b1 < spot[1]) or (
                (spot[0] == spot[1] or b1 == b2) and spot[0] <= b2 and b1 <= spot[1])
        else:
            hit = same_para and overlaps(current, proposal.get("before", "")) and _anchors_agree(last_prop, proposal)
        if hit:
            best = (chain_id, chain_events, last_prop)  # newest matching chain wins
    if best is None:
        proposal["chain"] = "c-" + proposal["id"][2:]
        proposal["round"] = 1
        proposal["prev"] = None
        return []
    chain_id, chain_events, last_prop = best
    proposal["chain"] = chain_id
    proposal["round"] = len(active(chain_events)) + 1
    proposal["prev"] = last_prop["id"]  # the round it rewrites (an undo branch it left is not its lineage)
    return [make_outcome(e, "superseded", kept=None, via="model", sha=proposal.get("sha_before"))
            for e in open_proposals(chain_events)]


def _chain_para(chain_events: list[dict]) -> int | None:
    for event in reversed(chain_events):
        if event.get("para") is not None:
            return event["para"]
    return None


def chains(events: list[dict], label: str | None = None) -> list[dict]:
    """Collapse the log into one record per chain, newest first.

    ``outcome`` is the user's last decision (the last round's), ``original`` and ``proposed`` are the first
    round's before/after, ``kept`` what the span reads now; ``pending`` when the last round is undecided. A round
    an undo took back (``undone``) is not counted."""
    seen: set[str] = set()
    by_chain: dict[str, list[dict]] = {}
    order: list[str] = []
    for event in events:
        kind = event.get("event")
        if kind == "seen":
            seen.update(event.get("chains") or [])
        elif kind in ("proposal", "outcome") and event.get("chain"):
            if event["chain"] not in by_chain:
                order.append(event["chain"])
            by_chain.setdefault(event["chain"], []).append(event)
    records = []
    for chain_id in order:
        chain_events = by_chain[chain_id]
        props = active(chain_events) or [e for e in chain_events if e.get("event") == "proposal"]
        if not props:
            continue
        first, last = props[0], props[-1]
        last_out = last_outcomes(chain_events).get(last["id"])
        pending = last_out is None or last_out["outcome"] in OPEN
        current, _ = _current_text(chain_events)
        records.append({
            "id": chain_id,
            "doc": label,
            "para": (last_out or {}).get("para") if (last_out or {}).get("para") is not None else last.get("para"),
            "where": last.get("where"),
            "outcome": None if pending else last_out["outcome"],
            "via": None if pending else last_out.get("via"),
            "rounds": len(props),
            "original": first.get("before", ""),
            "proposed": first.get("after", ""),
            "last_proposed": last.get("after", ""),
            "kept": current,
            "context": first.get("anchor_l", ""),
            "tracked": bool(last.get("tracked")),
            "author": last.get("author"),
            "time": (last_out or last)["time"],
            "pending": pending,
            "seen": chain_id in seen,
        })
    records.reverse()  # outcomes settled in one sync share a time: the later proposal comes first
    records.sort(key=lambda r: r["time"], reverse=True)
    return records


# ---------------------------------------------------------------- text shown to the model

def _show(text: str | None, budget: int) -> str:
    text, _ = clean(text or "")
    text = text.replace("\v", " / ").replace("\n", " / ").replace("\t", " ")
    text = re.sub(r"\s+", " ", text)
    if len(text) <= budget:
        return text
    if budget < 8:
        return text[:budget]
    half = (budget - 1) // 2
    return text[:half] + "…" + text[-(budget - 1 - half):]


def format_item(record: dict, *, with_doc: bool = False) -> str:
    """One record as one line of at most ITEM_MAX characters:
    ``[c-1a2b3c4d] ¶12 reworded (5 rounds, tracked): "…context" "original" → proposed "x" → kept "y"``."""
    head = f"[{record['id']}]"
    if with_doc and record.get("doc"):
        head += f" {_show(record['doc'], 40)}"
    where = record.get("where") or (f"¶{record['para']}" if record.get("para") is not None else "")
    if where:
        head += f" {where}"
    tags = []
    if record["rounds"] > 1:
        tags.append(f"{record['rounds']} rounds")
    if record.get("tracked"):
        tags.append("tracked")
    if record.get("via") == "model":
        tags.append("via model")
    outcome = record["outcome"] or "pending"
    head += f" {outcome}" + (f" ({', '.join(tags)})" if tags else "") + ":"
    original, proposed, kept = record["original"], record["proposed"], record["kept"]
    parts: list[tuple[str, str]] = []
    if record.get("context"):
        parts.append(("…", record["context"]))
    parts.append(("", original))
    parts.append((" → proposed ", proposed))
    if outcome == "reworded" or (record["rounds"] > 1 and outcome != "removed"):
        parts.append((" → kept ", kept or ""))
    budget = ITEM_MAX - len(head) - 1 - sum(len(label) + 2 for label, _ in parts)
    weights = {0: 1, 1: 2, 2: 3, 3: 3} if record.get("context") else {0: 2, 1: 3, 2: 3}
    total_w = sum(weights[i] for i in range(len(parts)))
    lengths = [len(text or "") for _, text in parts]
    shares = [max(6, budget * weights[i] // total_w) for i in range(len(parts))]
    # give unused share of short texts to the long ones
    spare = sum(max(0, s - n) for s, n in zip(shares, lengths))
    for i in sorted(range(len(parts)), key=lambda k: lengths[k] - shares[k], reverse=True):
        if lengths[i] > shares[i] and spare > 0:
            extra = min(spare, lengths[i] - shares[i])
            shares[i] += extra
            spare -= extra
    body = "".join(f'{label}"{_show(text, shares[i])}"' if label != "…" else f'"…{_show(text, shares[i])}" '
                   for i, (label, text) in enumerate(parts))
    line = f"{head} {body}"
    return line if len(line) <= ITEM_MAX else line[:ITEM_MAX - 1] + "…"


def counts(records: list[dict]) -> dict:
    out = {name: 0 for name in OUTCOMES}
    for record in records:
        if record["outcome"] in out:
            out[record["outcome"]] += 1
    return out


def summary_line(new_outcomes: list[dict], pending: int) -> str | None:
    """The one line added to a tool result, or None when the user decided nothing new."""
    tally = {name: 0 for name in OUTCOMES}
    for event in new_outcomes:
        if event.get("outcome") in tally and event.get("via") != "model":
            tally[event["outcome"]] += 1
    if not any(tally.values()):
        return None
    words = ", ".join(f"{n} {name}" for name, n in tally.items() if n)
    tail = f" ({pending} pending)" if pending else ""
    return f"Since your last edit: {words}{tail}. Call edit_feedback for details."


def chain_outcomes(new_outcomes: list[dict], events: list[dict]) -> list[dict]:
    """Collapse outcome events to one per chain (the last), so a chain counts once in the summary."""
    last: dict[str, dict] = {}
    for event in new_outcomes:
        last[event.get("chain") or event["id"]] = event
    return list(last.values())


def view(log: DocLog, *, label: str | None = None, paragraph_index: int | None = None,
         search_text: str | None = None, outcome: str | None = None, limit: int = DEFAULT_LIMIT,
         mark_seen: bool = True) -> dict:
    """The default per-document view: decided chains not yet seen, newest first, capped; with a filter
    (paragraph_index / search_text / outcome) every matching decided chain instead. Pending is a count."""
    with log.locked():
        events = log.events()
        records = chains(events, label)
        decided = [r for r in records if not r["pending"]]
        pending = len(records) - len(decided)
        filtered = bool(paragraph_index is not None or search_text or outcome)
        if filtered:
            picked = [r for r in decided if _matches(r, paragraph_index, search_text, outcome)]
            note = None
        else:
            # Removed paragraphs are counted, never listed, unless asked for (outcome="removed").
            listable = [r for r in decided if r["outcome"] != "removed"]
            picked = [r for r in listable if not r["seen"]]
            note = None
            if not picked and listable:
                picked = listable
                note = "Nothing new since you last looked; these are the most recent outcomes."
        more = max(0, len(picked) - limit)
        picked = picked[:limit]
        warning = None
        if mark_seen and picked:
            fresh = [r["id"] for r in picked if not r["seen"]]
            if fresh:
                try:
                    log.append([{"v": VERSION, "event": "seen", "time": now_iso(), "chains": fresh}], md=False)
                except Exception as exc:  # noqa: BLE001 - the items are still worth returning
                    warning = f"seen marks not saved: {type(exc).__name__}: {exc}"[:300]
    result = {
        "items": [format_item(r) for r in picked],
        "counts": counts(decided),
        "pending": pending,
        "more": more,
    }
    if note:
        result["note"] = note
    if warning:
        result["warning"] = warning
    return result


def _matches(record: dict, paragraph_index, search_text, outcome) -> bool:
    if paragraph_index is not None and record.get("para") != paragraph_index:
        return False
    if outcome and record["outcome"] != outcome:
        return False
    if search_text:
        needle = search_text.lower()
        hay = " ".join(str(record.get(k) or "") for k in ("original", "proposed", "kept", "context",
                                                             "last_proposed")).lower()
        if needle not in hay:
            return False
    return True


def search_all(root: Path, *, query: str | None = None, outcome: str | None = None,
               limit: int = DEFAULT_LIMIT) -> dict:
    """Decided chains across every document in ``root`` matching ``query``, newest first. Marks nothing seen."""
    found: list[dict] = []
    for jsonl in sorted(Path(root).glob("*.jsonl")) if Path(root).exists() else []:
        log = DocLog(root, jsonl.stem)
        events = log.events()
        label = _doc_label(events) or jsonl.stem
        for record in chains(events, label):
            if not record["pending"] and _matches(record, None, query, outcome):
                found.append(record)
    found.sort(key=lambda r: r["time"], reverse=True)
    more = max(0, len(found) - limit)
    return {"items": [format_item(r, with_doc=True) for r in found[:limit]], "counts": counts(found),
            "more": more}


def _doc_label(events: list[dict]) -> str | None:
    for event in events:
        if event.get("path"):
            return Path(str(event["path"]).replace("\\", "/")).name
    return None


def rounds(root: Path, record_id: str, log: DocLog | None = None) -> dict | None:
    """One chain by id: its net ``item`` and every round (``detail="rounds"``), each line at most ITEM_MAX
    characters. Searches ``log`` only when given, else every log in ``root``."""
    logs = [log] if log is not None else [DocLog(root, p.stem) for p in sorted(Path(root).glob("*.jsonl"))]
    for each in logs:
        events = each.events()
        chain_events = [e for e in events if e.get("chain") == record_id and e.get("event") in ("proposal", "outcome")]
        if not chain_events:
            continue
        outcomes = {e["id"]: e for e in chain_events if e.get("event") == "outcome"}
        lines = []
        number = 0
        for event in chain_events:
            if event.get("event") != "proposal":
                continue
            number += 1
            out = outcomes.get(event["id"])
            via_model = out and out.get("via") == "model"
            verdict = (out["outcome"] if out else "pending") + (" via model" if via_model else "")
            budget = (ITEM_MAX - 60) // 3
            line = (f"round {number} {event['time'][:16]} {verdict}: "
                    f'"{_show(event.get("before"), budget)}" → "{_show(event.get("after"), budget)}"')
            if out and out["outcome"] == "reworded":
                line += f' → kept "{_show(out.get("kept"), budget)}"'
            lines.append(line[:ITEM_MAX])
        label = _doc_label(events) or each.name
        record = next(r for r in chains(events, label) if r["id"] == record_id)
        return {"id": record_id, "doc": label, "rounds": lines, "item": format_item(record, with_doc=True)}
    return None


def render_md(events: list[dict], title: str) -> str:
    records = chains(events, None)
    decided = [r for r in records if not r["pending"]]
    tally = counts(decided)
    pending = len(records) - len(decided)
    lines = [f"# Edit feedback: {Path(str(title).replace(chr(92), '/')).name}", "",
             "Net outcome of each model edit (first proposal → what the user kept), newest first. "
             "The .jsonl beside this file keeps every round.", "",
             "Totals: " + ", ".join(f"{n} {k}" for k, n in tally.items()) + f"; {pending} pending.", ""]
    for record in decided:
        lines.append(f"- {record['time'][:16].replace('T', ' ')} UTC {format_item(record)}")
    lines.append("")
    return "\n".join(lines)
