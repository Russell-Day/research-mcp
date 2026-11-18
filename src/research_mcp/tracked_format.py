"""Bold / italic on one phrase as direct OOXML, tracked or plain (plan 864 S6, decision 20).

``style_runs`` sets ``w:b`` / ``w:i`` on the runs of one phrase of one paragraph, matched in the accepted view
(what ``inspect_document`` shows). Word never opens: only the touched runs change, every other byte of
``word/document.xml`` and every other zip entry is copied unchanged (``revisions.rezip``).

Tracked: each run of the phrase is split at the phrase edges and gets the new property plus a ``w:rPrChange``
by ``author`` holding the run's old ``w:rPr`` (as Word writes a formatting revision); ids come from
``revisions.next_revision_id`` and ascend in document order. Untracked: the new property only.

Rules (lead decisions L1-L3):

* L1: a run inside the author's own pending ``w:ins`` (author equal, exact) takes the new formatting in place
  with no ``w:rPrChange`` (Reject of that insertion removes it anyway; Accept gives the formatting).
* L2: a run inside another author's ``w:ins`` is formatted with a ``w:rPrChange`` by the op's author.
* L3: a run that already carries a ``w:rPrChange`` by the same author keeps it (its old rPr is still the state
  before the revision) and only its rPr changes; another author's is refused, naming it. When the phrase covers
  part of such a run, the run is split as Word splits it: every piece keeps the ``w:rPrChange`` with its old rPr
  unchanged, the first piece the original (same id) and each later piece a copy with a fresh id.

Refused, named: text inside a ``w:del`` / move, inside a field result or a phrase crossing a field, a drawing,
footnote reference or other object inside the phrase, a paragraph inside a nested table, and text carrying
another author's tracked formatting change.

Verified before anything is written (L4): bytes outside the paragraph identical; every revision that existed
byte-identical, except a ``w:ins`` holding a styled run (same start tag, same content apart from run
properties and run splits); an L3 copy equal to its original apart from the id; accepted and original text
unchanged; Reject of the new revisions (and of the L1/L3 revisions and copies in both) gives the input paragraph (canonical XML, run splits merged); Accept gives the input's
accepted view with the requested b/i on exactly the selected characters and every other character's
formatting unchanged; new ids unique in the package.
"""

from __future__ import annotations

import copy
import datetime as dt
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import W
from .revisions import (
    DEFAULT_AUTHOR, REV_TAGS, Char, Node, Obj, ParagraphModel, _et_final_chars, _fragment, _head_describe,
    _result, _verify_written, _Writer, _WRAPPERS, attr_value, describe_wrapper, et_resolve, load,
    next_revision_id, parse, revision_stamp, rezip, unescape,
)

OP = "style_text"
FIND_MAX = 5000
_ON = {"1", "true", "on"}
# CT_RPr child order (ECMA-376 EG_RPrBase); w:rPrChange is always last. Anything not listed (w14:..., an
# extension) ranks after the list, so a new w:b / w:i goes before it.
_ORDER = ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
          "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
          "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
          "em", "lang", "eastAsianLayout", "specVanish", "oMath")
_RANK = {"w:" + name: number for number, name in enumerate(_ORDER)}
_LATE = len(_ORDER)
_TEXT_PARTS = ("w:t", "w:tab", "w:br", "w:cr")


def _rank(name: str) -> int:
    if name == "w:rPrChange":
        return _LATE + 2
    if name in ("w:ins", "w:del", "w:moveFrom", "w:moveTo"):
        return -1  # only in a paragraph mark's rPr; never written here
    return _RANK.get(name, _LATE + 1)


rpr_rank = _rank

# CT_PPr child order (ECMA-376 EG_PPrBase), then the mark's w:rPr, w:sectPr and w:pPrChange (always last).
# Anything not listed (an extension) ranks after the base properties and before w:rPr.
_PPR_ORDER = ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
              "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
              "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
              "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle")
_PPR_RANK = {"w:" + name: number for number, name in enumerate(_PPR_ORDER)}
_PPR_TAIL = {"w:rPr": len(_PPR_ORDER) + 1, "w:sectPr": len(_PPR_ORDER) + 2, "w:pPrChange": len(_PPR_ORDER) + 3}


def ppr_rank(name: str) -> int:
    return _PPR_TAIL.get(name, _PPR_RANK.get(name, len(_PPR_ORDER)))


def rpr_change(writer: _Writer, rpr_xml: str, text: str) -> str:
    """A new w:rPrChange by the writer's author holding `rpr_xml`'s complete old properties."""

    return f"<w:rPrChange {writer._attrs('rPrChange', text)}>{_old_rpr(rpr_xml)}</w:rPrChange>"


def old_ppr(ppr_xml: str) -> str:
    """A paragraph's old pPr for its w:pPrChange: every property but the mark's w:rPr, w:sectPr and
    w:pPrChange (as Word writes it; ``<w:pPr/>`` when there is none)."""

    if not ppr_xml:
        return "<w:pPr/>"
    tree = parse(ppr_xml).children[0]
    kept = "".join(ppr_xml[item.start:item.end] for item in tree.children
                   if item.name not in ("w:rPr", "w:sectPr", "w:pPrChange"))
    return f"<w:pPr>{kept}</w:pPr>" if kept else "<w:pPr/>"


def ppr_change(writer: _Writer, ppr_xml: str, text: str = "") -> str:
    """A new w:pPrChange by the writer's author holding `ppr_xml`'s old paragraph properties."""

    return f"<w:pPrChange {writer._attrs('pPrChange', text)}>{old_ppr(ppr_xml)}</w:pPrChange>"


def _toggle(xml: str, node: Node | None) -> bool | None:
    """Explicit value of an on/off property element (w:b, w:i), None when absent."""

    if node is None:
        return None
    value = node.attr(xml, "w:val")
    return value is None or value.strip().lower() in _ON


def _rpr_children(rpr_xml: str) -> tuple[str, list[tuple[str, str]]]:
    """Start tag and (name, raw) children of a run's w:rPr ("" -> no rPr)."""

    if not rpr_xml:
        return "<w:rPr>", []
    tree = parse(rpr_xml).children[0]
    head = re.sub(r"\s*/>\Z", ">", tree.head(rpr_xml)) if tree.closed else tree.head(rpr_xml)
    return head, [(item.name, rpr_xml[item.start:item.end]) for item in tree.children]


def _current(rpr_xml: str, name: str) -> bool | None:
    if not rpr_xml:
        return None
    tree = parse(rpr_xml).children[0]
    return _toggle(rpr_xml, tree.child(name))


def _new_rpr(rpr_xml: str, requested: dict[str, bool], change: str) -> str:
    """`rpr_xml` with each requested toggle (``w:b``/``w:i``) set explicitly, at its schema position, and
    `change` (a new w:rPrChange, or "") appended last."""

    head, children = _rpr_children(rpr_xml)
    for name, value in requested.items():
        children = [item for item in children if item[0] != name]
        element = f"<{name}/>" if value else f'<{name} w:val="0"/>'
        at = next((number for number, item in enumerate(children) if _rank(item[0]) > _rank(name)), len(children))
        children.insert(at, (name, element))
    body = "".join(raw for _, raw in children) + change
    return f"{head}{body}</w:rPr>"


def _old_rpr(rpr_xml: str) -> str:
    """The run's complete old rPr for its w:rPrChange: everything but w:rPrChange and revision markers."""

    _, children = _rpr_children(rpr_xml)
    kept = "".join(raw for name, raw in children
                   if name not in ("w:rPrChange", "w:ins", "w:del", "w:moveFrom", "w:moveTo"))
    return f"<w:rPr>{kept}</w:rPr>" if kept else "<w:rPr/>"


class _Run:
    """One run holding selected characters, and what is done to it."""

    def __init__(self, run: Node) -> None:
        self.run = run
        self.chars: list[Char] = []
        self.kind = "new"  # "new" (rPrChange added), "own_insertion" (L1), "own_format_change" (L3), "plain"
        self.split = False
        self.rpr = ""  # the run's rPr as it stands
        self.new_rpr = ""
        self.change_id: str | None = None
        self.in_place: Node | None = None  # the own w:ins (L1) or own w:rPrChange holder (L3)
        self.copies: list[str] = []  # L3 split: ids of the w:rPrChange copies on the later pieces


class Styler:
    def __init__(self, doc, index: int, *, text: str, bold: bool | None, italic: bool | None, tracked: bool,
                 author: str, allow_whole_paragraph: bool, op: str = OP) -> None:
        from .sections import caps_find

        self.doc, self.index, self.op, self.tracked, self.author = doc, index, op, tracked, author
        self.xml = doc.xml
        self.requested = {name: value for name, value in (("w:b", bold), ("w:i", italic)) if value is not None}
        if not self.requested:
            raise RefusedUnsafe(f"{op}: give bold and/or italic; omitted properties stay unchanged")
        self.node = doc.paragraph(index, op)
        item = doc.layout.para(index)
        if item.nested:
            raise RefusedUnsafe(f"{op}: paragraph {index} is inside a nested table; format it in Word")
        self.model = ParagraphModel(self.xml, self.node)
        layout_chars = [char for char in self.model.chars if char.layout_visible]
        if "".join(char.ch for char in layout_chars) != item.text:
            raise RefusedUnsafe(
                f"{op}: paragraph {index}'s text could not be aligned with the document layout; nothing was changed")
        caps_of = {id(char): flag for char, flag in zip(layout_chars, item.caps)}
        caps = tuple(caps_of[id(char)] for char in self.model.final)
        if not text:
            raise RefusedUnsafe(f"{op}: text must not be empty")
        if len(text) > FIND_MAX:
            raise RefusedUnsafe(f"{op}: text is {len(text)} characters; at most {FIND_MAX} are supported")
        hits = caps_find(self.model.text, caps, text)
        if len(hits) != 1:
            raise RefusedUnsafe(
                f"{op}: text must occur exactly once in paragraph {index}'s accepted text; it occurs {len(hits)} "
                "times" + (" (deleted text does not count)" if self.model.original != self.model.text else ""))
        if text.strip() == self.model.text.strip() and not allow_whole_paragraph:
            raise RefusedUnsafe(f"{op}: whole-paragraph styling is refused; set allow_whole_paragraph=True explicitly")
        self.start, self.end = hits[0], hits[0] + len(text)
        self.selected = self.model.final[self.start:self.end]
        self.selected_text = "".join(char.ch for char in self.selected)
        self.runs: list[_Run] = []
        self._check()
        self._plan()

    # -- refusals ----------------------------------------------------------

    def _refuse(self, what: str) -> RefusedUnsafe:
        return RefusedUnsafe(f"{self.op}: paragraph {self.index}: {what}")

    def _own(self, char: Char) -> Node | None:
        wrappers = [node for node in char.wrappers if node.name in _WRAPPERS]
        if len(wrappers) == 1 and wrappers[0].name == "w:ins" and \
                wrappers[0].attr(self.xml, "w:author") == self.author:
            return wrappers[0]
        return None

    def _check(self) -> None:
        low = self.model.position[id(self.selected[0])]
        high = self.model.position[id(self.selected[-1])]
        for item in self.model.items[low:high + 1]:
            if isinstance(item, Obj):
                if item.kind == "field":
                    raise self._refuse("the text crosses a field (a citation or cross-reference); style the text "
                                       "on either side of it separately")
                if item.kind == "opaque":
                    raise self._refuse(f"the text holds a {item.node.name} element; only text, tabs and line "
                                       "breaks are styled here. Format it in Word")
                continue
            for node in item.wrappers:
                if node.name in ("w:del", "w:moveFrom", "w:moveTo"):
                    what = "deleted" if node.name == "w:del" else "moved"
                    raise self._refuse(f"the text touches {what} text in {describe_wrapper(self.xml, node)}; "
                                       "accept or reject that revision first, or style text outside it")
            if item.field:
                raise self._refuse(f"the text lies inside a field result ({item.ch!r}); fields (citations, "
                                   "cross-references) are never formatted here")
            change = re.search(r"<w:rPrChange\b[^>]*>", item.rpr)
            if change and attr(change.group(0), "w:author") != self.author:
                raise self._refuse(
                    f"the text carries the tracked formatting change {_head_describe(change.group(0))} by another "
                    "author; accept or reject it first")

    # -- planning ------------------------------------------------------------

    def _plan(self) -> None:
        by_run: dict[int, _Run] = {}
        for char in self.selected:
            entry = by_run.get(id(char.run))
            if entry is None:
                entry = by_run[id(char.run)] = _Run(char.run)
                self.runs.append(entry)
            entry.chars.append(char)
        planned = []
        for entry in self.runs:
            run = entry.run
            rpr_node = run.child("w:rPr")
            rpr = self.xml[rpr_node.start:rpr_node.end] if rpr_node is not None else ""
            if all(_current(rpr, name) == value for name, value in self.requested.items()):
                continue  # already formatted as requested: left alone, no revision
            run_chars = [char for char in self.model.chars if char.run is run]
            entry.split = len(run_chars) != len(entry.chars)
            change = re.search(r"<w:rPrChange\b[^>]*>", rpr)
            own = self._own(entry.chars[0])
            if change and not self.tracked:  # plain styling would change what Reject of that revision restores
                raise self._refuse(
                    f"the text carries the tracked formatting change {_head_describe(change.group(0))}; untracked "
                    "styling is refused there (it would add revision ids or be undone by rejecting that change); "
                    "style it tracked, or accept or reject that change first")
            if change:  # the author's own (_check refused another's); split, each piece keeps a copy (L3)
                entry.kind = "own_format_change"
                entry.change_id = attr(change.group(0), "w:id")
            elif not self.tracked:
                entry.kind = "plain"
            elif own is not None:
                entry.kind = "own_insertion"
                entry.in_place = own
            entry.rpr = rpr
            planned.append(entry)
        if not planned:
            raise self._refuse(f"{self.selected_text!r} is already formatted as requested; nothing to change")
        self.runs = planned

    # -- writing -------------------------------------------------------------

    def apply(self, *, first_id: int, stamp: str) -> tuple[str, list[dict]]:
        writer = _Writer(first_id, self.author, stamp, True)
        out, cursor = [], self.node.start
        for entry in self.runs:
            out.append(self.xml[cursor:entry.run.start])
            out.append(self._rebuild(entry, writer))
            cursor = entry.run.end
        out.append(self.xml[cursor:self.node.end])
        return "".join(out), writer.added

    def _rebuild(self, entry: _Run, writer: _Writer) -> str:
        xml, run = self.xml, entry.run
        change = ""
        if entry.kind == "new":
            text = "".join(char.ch for char in entry.chars)
            change = rpr_change(writer, entry.rpr, text)
            entry.change_id = str(writer.next_id - 1)
        entry.new_rpr = _new_rpr(entry.rpr, self.requested, change)
        head = run.head(xml)
        if not entry.split:
            body = "".join(xml[item.start:item.end] for item in run.children if item.name != "w:rPr")
            return head + entry.new_rpr + body + "</w:r>"
        chosen = {(id(char.child), char.offset) for char in entry.chars}
        first = min(self.model.position[id(char)] for char in entry.chars)
        last = max(self.model.position[id(char)] for char in entry.chars)
        pieces: list[list] = [[], [], []]  # before, inside, after: raw strings or (text,) chunks of w:t

        def segment(char: Char | None, position: int) -> int:
            if char is not None:
                return 1 if (id(char.child), char.offset) in chosen else (0 if position < first else 2)
            return 0 if position < first else (2 if position > last else 1)

        def add(where: int, part) -> None:
            bucket = pieces[where]
            if isinstance(part, tuple) and bucket and isinstance(bucket[-1], tuple):
                bucket[-1] = (bucket[-1][0] + part[0],)
            else:
                bucket.append(part)

        chars = {(id(char.child), char.offset): char for char in self.model.chars if char.run is run}
        for child in run.children:
            if child.name == "w:rPr":
                continue
            if child.name == "w:t":
                text = [chars[(id(child), offset)] for offset in range(len(chars_text(xml, child)))]
                where = {segment(char, self.model.position[id(char)]) for char in text}
                if len(where) == 1:
                    pieces[where.pop()].append(xml[child.start:child.end])
                    continue
                for char in text:
                    add(segment(char, self.model.position[id(char)]), (char.ch,))
                continue
            char = chars.get((id(child), 0)) if child.name in _TEXT_PARTS else None
            if char is not None:
                pieces[segment(char, self.model.position[id(char)])].append(xml[child.start:child.end])
                continue
            position = next((number for number, item in enumerate(self.model.items)
                             if isinstance(item, Obj) and item.node is child), None)
            pieces[segment(None, position if position is not None else first)].append(xml[child.start:child.end])
        old_rpr = entry.rpr
        out = []
        for where, bucket in enumerate(pieces):
            if not bucket:
                continue
            body = "".join(part if isinstance(part, str) else
                           f'<w:t xml:space="preserve">{escape(part[0])}</w:t>' for part in bucket)
            rpr = entry.new_rpr if where == 1 else old_rpr
            if entry.kind == "own_format_change" and out:  # a later piece: a copy of the change, fresh id (L3)
                number = writer.next_id
                writer.next_id += 1
                entry.copies.append(str(number))
                rpr = _with_change_id(rpr, number)
            out.append(head + rpr + body + "</w:r>")
        return "".join(out)

    # -- report --------------------------------------------------------------

    def in_place_report(self) -> list[dict]:
        out = []
        for entry in self.runs:
            if entry.kind == "own_insertion":
                node = entry.in_place
                out.append({"kind": "own_insertion", "id": node.attr(self.xml, "w:id"),
                            "author": node.attr(self.xml, "w:author"), "date": node.attr(self.xml, "w:date"),
                            "text": "".join(char.ch for char in entry.chars)})
            elif entry.kind == "own_format_change":
                change = re.search(r"<w:rPrChange\b[^>]*>", entry.rpr).group(0)
                item = {"kind": "own_format_change", "id": attr(change, "w:id"),
                        "author": attr(change, "w:author"), "date": attr(change, "w:date"),
                        "text": "".join(char.ch for char in entry.chars)}
                if entry.copies:
                    item["split_copies"] = list(entry.copies)
                out.append(item)
        return out


def attr(head: str, name: str) -> str | None:
    return attr_value(head, name)


def _with_change_id(rpr: str, number: int) -> str:
    """`rpr` with its w:rPrChange's w:id set to `number` (nothing else changes)."""

    match = re.search(r"<w:rPrChange\b[^>]*>", rpr)
    head = re.sub(r"(\sw:id\s*=\s*)([\"'])[^\"']*\2", lambda m: f"{m.group(1)}{m.group(2)}{number}{m.group(2)}",
                  match.group(0), count=1)
    return rpr[:match.start()] + head + rpr[match.end():]


def chars_text(xml: str, child: Node) -> str:
    return "" if child.closed else unescape(child.inner(xml))


# --------------------------------------------------------------------------
# Verification (L4).
# --------------------------------------------------------------------------

def _strip_run_props(element: ET.Element) -> ET.Element:
    clone = copy.deepcopy(element)
    for run in clone.iter(W + "r"):
        for rpr in run.findall(W + "rPr"):
            run.remove(rpr)
    return clone


def _semantic(key: str, override: dict[str, bool] | None = None) -> tuple:
    """A character's run formatting (`_et_final_chars` key) as (other properties, b, i); `override` sets b/i."""

    other, toggles = [], {"b": None, "i": None}
    if key:
        for child in ET.fromstring(key):
            local = child.tag.rsplit("}", 1)[-1]
            if child.tag.startswith(W) and local in toggles:
                value = child.get(W + "val")
                toggles[local] = value is None or value.lower() in ("1", "true", "on")
            else:
                other.append(ET.tostring(child, encoding="unicode"))
    for name, value in (override or {}).items():
        toggles[name[2:]] = value
    return tuple(sorted(other)), toggles["b"], toggles["i"]


def _verify(styler: Styler, new_xml: str, new_paragraph: str, added: list[dict]) -> None:
    from .sections import canonical

    op, xml, node, author, tracked = styler.op, styler.xml, styler.node, styler.author, styler.tracked
    if new_xml[:node.start] != xml[:node.start] or new_xml[node.start + len(new_paragraph):] != xml[node.end:]:
        raise VerificationFailed(f"{op}: bytes outside paragraph {styler.index} changed")
    old_raw = xml[node.start:node.end]
    old_tree, new_tree = parse(old_raw), parse(new_paragraph)
    new_ids = {item["id"] for item in added}
    if len(added) != sum(entry.kind == "new" for entry in styler.runs) or (added and not tracked):
        raise VerificationFailed(f"{op}: new revision count mismatch")
    holders = {entry.in_place.head(xml) for entry in styler.runs if entry.kind == "own_insertion"}
    holders |= {wrapper.head(xml) for entry in styler.runs for wrapper in entry.chars[0].wrappers
                if wrapper.name == "w:ins"}

    def revisions(tree: Node) -> list[Node]:
        return [item for item in tree.iter() if item.local in REV_TAGS and item.name.startswith("w:")]

    old_revs = revisions(old_tree)
    copy_of = {number: entry.change_id for entry in styler.runs for number in entry.copies}
    originals = {item.attr(old_raw, "w:id"): old_raw[item.start:item.end] for item in old_revs
                 if item.name == "w:rPrChange"}
    fresh, kept, order = [], [], []
    for item in revisions(new_tree):
        number = item.attr(new_paragraph, "w:id")
        if item.name == "w:rPrChange" and number in copy_of:
            raw = new_paragraph[item.start:item.end]
            original = originals.get(copy_of[number], "")
            if not original or _with_change_id(original, int(number)) != raw:
                raise VerificationFailed(f"{op}: the copy {number} of formatting change {copy_of[number]} is not "
                                         "the original with a new id (L3)")
            order.append(int(number))
        elif item.name == "w:rPrChange" and number in new_ids and item.attr(new_paragraph, "w:author") == author:
            fresh.append(item)
            order.append(int(number))
        else:
            kept.append(item)
    if sorted(copy_of) != sorted(str(number) for number in order if str(number) in copy_of):
        raise VerificationFailed(f"{op}: an L3 copy of a formatting change is missing")
    if order != sorted(order) or len(set(order)) != len(order):
        raise VerificationFailed(f"{op}: new revision ids do not ascend left to right")
    if len(fresh) != len(added):
        raise VerificationFailed(f"{op}: new revision count mismatch")
    if len(kept) != len(old_revs):
        raise VerificationFailed(f"{op}: a pre-existing revision in paragraph {styler.index} was added or lost (D4)")
    for old, new in zip(old_revs, kept):
        before, after = old_raw[old.start:old.end], new_paragraph[new.start:new.end]
        if before == after:
            continue
        if old.name != "w:ins" or new.name != "w:ins" or old.head(old_raw) != new.head(new_paragraph) or \
                old.head(old_raw) not in holders:
            raise VerificationFailed(f"{op}: a pre-existing revision in paragraph {styler.index} changed (D4)")
        if canonical(_strip_run_props(_fragment(xml, before))) != canonical(_strip_run_props(_fragment(xml, after))):
            raise VerificationFailed(
                f"{op}: the insertion {_head_describe(old.head(old_raw))} changed beyond its run formatting (D4)")
    previous = None
    for item in fresh:
        number = int(item.attr(new_paragraph, "w:id"))
        if previous is not None and number <= previous:
            raise VerificationFailed(f"{op}: new revision ids do not ascend left to right")
        previous = number
        if not re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$", item.attr(new_paragraph, "w:date") or ""):
            raise VerificationFailed(f"{op}: new revision has the wrong date format")
    new_model = ParagraphModel(new_paragraph, new_tree.children[0])
    if new_model.text != styler.model.text or new_model.original != styler.model.original:
        raise VerificationFailed(f"{op}: the text of paragraph {styler.index} changed")

    old_p, new_p = _fragment(xml, old_raw), _fragment(xml, new_paragraph)
    edited_ins = {entry.in_place.attr(xml, "w:id") for entry in styler.runs if entry.kind == "own_insertion"}
    edited_changes = {entry.change_id for entry in styler.runs if entry.kind == "own_format_change"}
    is_new = lambda el: el.tag == W + "rPrChange" and el.get(W + "id") in new_ids and \
        el.get(W + "author") == author  # noqa: E731
    is_edited = lambda el: el.get(W + "author") == author and (  # noqa: E731
        (el.tag == W + "ins" and el.get(W + "id") in edited_ins)
        or (el.tag == W + "rPrChange" and el.get(W + "id") in edited_changes | set(copy_of)))
    if tracked:
        rejected = copy.deepcopy(new_p)
        et_resolve(rejected, lambda el: is_new(el) or is_edited(el), "reject")
        baseline = copy.deepcopy(old_p)
        et_resolve(baseline, is_edited, "reject")
        if canonical(rejected) != canonical(baseline):
            raise VerificationFailed(
                f"{op}: rejecting the new formatting revisions does not give back paragraph {styler.index} (D5)")
    accepted = copy.deepcopy(new_p)
    et_resolve(accepted, is_new, "accept")
    got, want = _et_final_chars(accepted), _et_final_chars(old_p)
    if "".join(ch for ch, _ in want) != styler.model.text or "".join(ch for ch, _ in got) != styler.model.text:
        raise VerificationFailed(f"{op}: accepting the formatting changes changes the text of paragraph {styler.index}")
    override = dict(styler.requested)
    for number, ((ch, key), (_, old_key)) in enumerate(zip(got, want)):
        if styler.start <= number < styler.end:
            if _semantic(key) != _semantic(old_key, override):
                raise VerificationFailed(f"{op}: character {ch!r} of the phrase is not formatted as requested (D5)")
        elif key != old_key:
            raise VerificationFailed(f"{op}: character {ch!r} outside the phrase changed formatting (D8)")


# --------------------------------------------------------------------------
# Operation.
# --------------------------------------------------------------------------

def style_runs(source, output, *, paragraph_index: int, text: str, bold: bool | None = None,
               italic: bool | None = None, tracked: bool, author: str | None = DEFAULT_AUTHOR,
               allow_whole_paragraph: bool = False, now: dt.datetime | None = None, op: str = OP) -> dict:
    """Bold/italic on one phrase of paragraph `paragraph_index` (accepted view), tracked as w:rPrChange or plain;
    verified, then written to the new file `output`."""

    author = (author or DEFAULT_AUTHOR).strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{op}: author must not be empty")
    parts, doc = load(source)
    styler = Styler(doc, paragraph_index, text=text, bold=bold, italic=italic, tracked=tracked, author=author,
                    allow_whole_paragraph=allow_whole_paragraph, op=op)
    first_id = next_revision_id(parts)
    stamp = revision_stamp(doc.xml, now)
    new_paragraph, added = styler.apply(first_id=first_id, stamp=stamp)
    node = styler.node
    new_xml = doc.xml[:node.start] + new_paragraph + doc.xml[node.end:]
    _verify(styler, new_xml, new_paragraph, added)
    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, op)
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    return _result(
        source, output, op, doc,
        paragraph_index=paragraph_index,
        tracked=tracked,
        author=author if tracked else None,
        selected_text=styler.selected_text,
        requested={{"w:b": "bold", "w:i": "italic"}[name]: value for name, value in styler.requested.items()},
        runs_styled=len(styler.runs),
        revisions_added=added,
        styled_in_place=styler.in_place_report(),
    )



def preview(source, *, paragraph_index: int, text: str, bold: bool | None = None, italic: bool | None = None,
            tracked: bool, author: str | None = DEFAULT_AUTHOR, allow_whole_paragraph: bool = False,
            op: str = OP) -> dict:
    """What style_runs would do, written nowhere: the selection, its context and the runs it would change.
    Every refusal style_runs makes before writing is made here too."""

    author = (author or DEFAULT_AUTHOR).strip()
    _, doc = load(source)
    styler = Styler(doc, paragraph_index, text=text, bold=bold, italic=italic, tracked=tracked, author=author,
                    allow_whole_paragraph=allow_whole_paragraph, op=op)
    accepted = styler.model.text
    kinds = [entry.kind for entry in styler.runs]
    return {
        "selected_text": styler.selected_text,
        "context_before": accepted[max(0, styler.start - 80):styler.start],
        "context_after": accepted[styler.end:styler.end + 80],
        "scope": "whole_paragraph" if styler.selected_text.strip() == accepted.strip() else "exact_text",
        "runs_to_style": len(kinds),
        "revisions_to_add": kinds.count("new"),
        "styled_in_place": [{"kind": entry.kind, "text": "".join(char.ch for char in entry.chars)}
                            for entry in styler.runs if entry.kind in ("own_insertion", "own_format_change")],
    }


def apply_operation(source, output, operation) -> dict:
    """edit_batch entry for the style_text op (tracked/author already resolved by track_setting.route_operation).

    Direct XML when the op is tracked or the document has any revision; a revision-free, untracked op goes
    through Word's own formatting (inline_style), as the style_text tool does."""

    from .revisions import revision_count

    tracked = bool(operation.tracked)
    author = operation.author or DEFAULT_AUTHOR
    if tracked or revision_count(source):
        return style_runs(source, output, paragraph_index=operation.paragraph_index, text=operation.text,
                          bold=operation.bold, italic=operation.italic, tracked=tracked, author=author,
                          allow_whole_paragraph=operation.allow_whole_paragraph)
    from .inline_style import style_exact_text

    result = style_exact_text(source, output, operation.paragraph_index, operation.text, bold=operation.bold,
                              italic=operation.italic, allow_whole_paragraph=operation.allow_whole_paragraph)
    return {**result, "operation": OP, "engine": "word", "tracked": False, "author": None,
            "selected_text": operation.text, "revisions_added": [], "styled_in_place": [], "index_delta": 0}
