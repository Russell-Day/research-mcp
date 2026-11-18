"""Tracked changes as direct OOXML (plan 501 S2).

Word is the gate, not the writer. Every operation here edits
``word/document.xml`` as text: only the edited spans change, every other byte
of the part and every other zip entry is copied unchanged, in the original
entry order. Nothing here starts Word, because any Word save rewrites other
authors' revision records (D4).

Operations (all reached through ``edit_document``):

* ``list_revisions``: id, type, author, date, paragraph and text of each
  revision, plus each listed paragraph's accepted and original text.
* ``tracked_replace_text`` / ``tracked_replace_paragraph``: the plan-328
  minimal diff (partial-word edits widened to whole words when that keeps
  formatting), written as ``w:del`` / ``w:ins`` by ``author`` (default
  "Claude"). ``w:date`` is local wall time with a "Z" suffix and
  ``w16du:dateUtc`` the real UTC time when the document declares w16du,
  Word's own convention. New ``w:id`` values are unique across the package
  and ascend left to right.
* ``accept_revisions`` / ``reject_revisions``: by ids, author, paragraph
  range or section; every selector given must match.
* ``plain_replace``: an untracked replace_text / replace_paragraph on a
  document that has revisions, allowed only when it touches no revised run.

Every write is verified before it is saved (D4/D5): bytes outside the edited
spans are identical; every pre-existing revision not named by the operation
is byte-identical; for tracked edits, rejecting the new revisions gives the
input paragraph back and accepting them gives exactly the requested text with
each new character formatted like its source; for accept/reject, the output
equals an independent ElementTree resolution of the same selection. A failed
check raises and nothing is written. Refusals name the revision (id, type,
author) or the object (field, drawing, move) that blocks the edit.
"""

from __future__ import annotations

import bisect
import copy
import datetime as dt
import io
import re
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import W

DEFAULT_AUTHOR = "Claude"
RUN_REVS = ("ins", "del", "moveFrom", "moveTo")
PROP_REVS = ("rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange",
             "tblGridChange", "tblPrExChange", "numberingChange")
CELL_REVS = ("cellIns", "cellDel", "cellMerge")
REV_TAGS = RUN_REVS + PROP_REVS + CELL_REVS
TABLE_PROP_REVS = ("tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange",
                   "tblPrExChange", "numberingChange")
MOVE_RANGES = ("moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd")
_WRAPPERS = {"w:ins", "w:del", "w:moveFrom", "w:moveTo"}
_DELETED = {"w:del", "w:moveFrom"}
# Revision markers a paragraph mark's rPr may carry; text typed into the paragraph does not inherit them.
_MARK_REVISIONS = {"w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange"}
_INSERTED = {"w:ins", "w:moveTo"}
_PASS_THROUGH = {"w:hyperlink": "h", "w:smartTag": "", "w:customXml": "", "w:dir": "", "w:bdo": "",
                 "w:fldSimple": "f", "w:sdt": "s"}
_MARKS = {"w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd",
          "w:commentReference", "w:proofErr", "w:permStart", "w:permEnd", "w:lastRenderedPageBreak"} | {
    "w:" + name for name in MOVE_RANGES}
_ANCHOR_MARKS = ("w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd",
                 "w:commentReference")
# Zero-width markers a new run may be placed across when it joins the caller's own w:ins at its edge
# (plan 831 S2 round 1); they may move into or out of that w:ins (never across text) to keep Word's anchor order.
_EDGE_MARKS = {"w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd", "w:proofErr",
               "w:permStart", "w:permEnd"}
_DRAWINGS = ("w:drawing", "mc:AlternateContent", "w:pict", "w:object")
_FIELD_PARTS = ("w:fldChar", "w:instrText", "w:delInstrText", "w:fldSimple")
_STAMP = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")


# --------------------------------------------------------------------------
# Raw XML: a light tree over the text, so unchanged bytes are copied, never re-serialized.
# --------------------------------------------------------------------------

_TAG = re.compile(r"<(/?)([A-Za-z_][\w.\-]*(?::[\w.\-]+)?)((?:\s[^>]*?)?)(/?)>|<!--.*?-->|<\?.*?\?>", re.S)


class Node:
    __slots__ = ("name", "start", "head_end", "end", "parent", "children")

    def __init__(self, name: str, start: int, head_end: int, parent: Node | None) -> None:
        self.name, self.start, self.head_end, self.parent = name, start, head_end, parent
        self.end = head_end
        self.children: list[Node] = []

    @property
    def local(self) -> str:
        return self.name.rsplit(":", 1)[-1]

    @property
    def closed(self) -> bool:
        """True for a self-closing element (no content, no end tag)."""
        return self.end == self.head_end

    def close_start(self) -> int:
        return self.end if self.closed else self.end - len(f"</{self.name}>")

    def head(self, xml: str) -> str:
        return xml[self.start:self.head_end]

    def inner(self, xml: str) -> str:
        return xml[self.head_end:self.close_start()]

    def attr(self, xml: str, name: str) -> str | None:
        return attr_value(self.head(xml), name)

    def child(self, name: str) -> Node | None:
        return next((item for item in self.children if item.name == name), None)

    def iter(self):
        yield self
        for item in self.children:
            yield from item.iter()

    def ancestors(self):
        node = self.parent
        while node is not None and node.name != "#root":
            yield node
            node = node.parent


def attr_value(head: str, name: str) -> str | None:
    """Attribute ``name`` of a start tag, in either quote style and with any spacing around "=" (all valid
    XML; a regex on `name="..."` misses `name='31'`: final review terra61/sol61)."""

    found = re.search(r'\s%s\s*=\s*(?:"([^"]*)"|\'([^\']*)\')' % re.escape(name), head)
    if not found:
        return None
    return unescape(found.group(1) if found.group(1) is not None else found.group(2))


def parse(xml: str) -> Node:
    root = Node("#root", 0, 0, None)
    root.end = len(xml)
    stack = [root]
    for match in _TAG.finditer(xml):
        name = match.group(2)
        if name is None:
            continue
        if match.group(1):
            node = stack.pop()
            if node.name != name:
                raise VerificationFailed(f"malformed XML near offset {match.start()}: </{name}> closes <{node.name}>")
            node.end = match.end()
            continue
        node = Node(name, match.start(), match.end(), stack[-1])
        stack[-1].children.append(node)
        if not match.group(4):
            stack.append(node)
    if len(stack) != 1:
        raise VerificationFailed(f"malformed XML: <{stack[-1].name}> is never closed")
    return root


def unescape(text: str) -> str:
    def entity(match: re.Match) -> str:
        name = match.group(1)
        if name.startswith("#x"):
            return chr(int(name[2:], 16))
        if name.startswith("#"):
            return chr(int(name[1:]))
        return {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}[name]
    return re.sub(r"&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);", entity, text)


def _attr_escape(value: str) -> str:
    return escape(value, {'"': "&quot;"})


def strip_rpr_change(rpr: str) -> str:
    return re.sub(r"<w:rPrChange(?=[\s>/])[^>]*?/>|<w:rPrChange(?=[\s>])[^>]*>.*?</w:rPrChange>", "", rpr,
                  flags=re.S)


def revision_stamp(xml: str, now: dt.datetime | None = None) -> str:
    """Date attributes in Word's convention: local wall time + "Z", and true UTC in w16du:dateUtc."""

    now = (now or dt.datetime.now(dt.timezone.utc)).astimezone(dt.timezone.utc)
    attrs = f'w:date="{now.astimezone().strftime("%Y-%m-%dT%H:%M:%SZ")}"'
    head = re.search(r"<w:document\b[^>]*>", xml)
    if head and 'xmlns:w16du="' in head.group(0):
        attrs += f' w16du:dateUtc="{now.strftime("%Y-%m-%dT%H:%M:%SZ")}"'
    return attrs


# --------------------------------------------------------------------------
# Package access.
# --------------------------------------------------------------------------

def read_parts(path: str | Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _package_ids(parts: dict[str, bytes]) -> dict[str, list[int]]:
    """Every numeric WordprocessingML id attribute, per XML part, namespace-aware: any prefix bound to the
    w: namespace, either quote style (a regex on `w:id="..."` misses `w:id='31'` and `x:id="31"`)."""

    key = W + "id"
    found: dict[str, list[int]] = {}
    for name, data in parts.items():
        # no raw-bytes prefilter: a UTF-16 part never contains the UTF-8 namespace bytes (review s2r2 F4)
        if not name.lower().endswith((".xml", ".rels")):
            continue
        try:
            values = [el.get(key) for _, el in ET.iterparse(io.BytesIO(data)) if key in el.attrib]
        except ET.ParseError as exc:
            raise RefusedUnsafe(f"cannot scan {name} for revision ids: {exc}") from exc
        found[name] = [int(v) for v in values if re.fullmatch(r"\s*-?\d+\s*", v)]
    return found


def next_revision_id(parts: dict[str, bytes]) -> int:
    """One past every w:id in every XML part of the package, so a new id is unique everywhere."""

    return max((value for values in _package_ids(parts).values() for value in values), default=0) + 1


def _verify_written(source, output, new_xml: str, first_id: int, op: str) -> None:
    """The written package holds exactly the verified document.xml, and every id the edit added is new to
    the whole package (F4: no collision with an id in a header, footnote or any other part)."""

    after = read_parts(output)
    if after.get("word/document.xml", b"").decode("utf-8") != new_xml:
        raise VerificationFailed(f"{op}: the written word/document.xml is not the verified one")
    before_ids, after_ids = _package_ids(read_parts(source)), _package_ids(after)
    existing = {value for values in before_ids.values() for value in values}
    old_counts = Counter(before_ids.get("word/document.xml", []))
    added = Counter(after_ids.get("word/document.xml", [])) - old_counts
    clashes = sorted(value for value in added if value in existing or value < first_id)
    if clashes:
        raise VerificationFailed(f"{op}: new revision ids {clashes} collide with ids already in the package")


def rezip(source: str | Path, output: str | Path, document_xml: str) -> None:
    """Copy the package entry by entry (same order, same compression); replace word/document.xml only."""

    output = Path(output)
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(output, "x") as target:
        for info in source_zip.infolist():
            data = document_xml.encode("utf-8") if info.filename == "word/document.xml" else source_zip.read(info.filename)
            target.writestr(info, data, compress_type=info.compress_type)


# --------------------------------------------------------------------------
# Document: raw tree, paragraphs in layout numbering, revision records.
# --------------------------------------------------------------------------

@dataclass
class Revision:
    ordinal: int  # position among all revision elements, document order
    node: Node
    type: str
    id: str | None
    author: str | None
    date: str | None
    mark: str | None  # "paragraph" (paragraph-mark revision), "row", or None
    paragraph_index: int | None
    text: str

    def describe(self) -> str:
        what = f"{self.type} id={self.id}" + (f" ({self.mark} mark)" if self.mark else "")
        return f"w:{what} by {self.author!r}" + (f" ({self.date})" if self.date else "")

    def as_dict(self) -> dict:
        return {"id": self.id, "type": self.type, "mark": self.mark, "author": self.author, "date": self.date,
                "paragraph_index": self.paragraph_index, "text": self.text}


class Document:
    def __init__(self, xml: str, layout) -> None:
        self.xml = xml
        self.root = parse(xml)
        body = next((node for node in self.root.iter() if node.name == "w:body"), None)
        if body is None:
            raise RefusedUnsafe("word/document.xml has no w:body")
        self.body = body
        self.layout = layout
        ordinals = [node for node in body.iter() if node.name == "w:p"]
        et_ordinal = {id(node): number for number, node in enumerate(layout.body.iter(W + "p"))} \
            if layout.body is not None else {}
        if len(ordinals) != len(et_ordinal):
            raise VerificationFailed("raw and parsed paragraph counts differ; cannot address paragraphs")
        self.para_nodes: dict[int, Node] = {}
        for item in layout.paragraphs:
            if item.node is not None:
                self.para_nodes[item.index] = ordinals[et_ordinal[id(item.node)]]
        self._starts = sorted((node.start, node.end, index) for index, node in self.para_nodes.items())
        self._start_keys = [start for start, _, _ in self._starts]
        self.revisions: list[Revision] = []
        for node in body.iter():
            if node.local in REV_TAGS and node.name.startswith("w:"):
                self.revisions.append(self._record(len(self.revisions), node))

    def _record(self, ordinal: int, node: Node) -> Revision:
        parent = node.parent
        mark = None
        if node.local in RUN_REVS and parent is not None:
            if parent.name == "w:rPr" and parent.parent is not None and parent.parent.name == "w:pPr":
                mark = "paragraph"
            elif parent.name == "w:trPr":
                mark = "row"
        text = ""
        if node.local in RUN_REVS and not node.closed:
            text = "".join(
                unescape(item.inner(self.xml)) if item.name in ("w:t", "w:delText") else "\t"
                for item in node.iter() if item.name in ("w:t", "w:delText", "w:tab"))
        return Revision(ordinal, node, node.local, node.attr(self.xml, "w:id"), node.attr(self.xml, "w:author"),
                        node.attr(self.xml, "w:date"), mark, self.paragraph_at(node.start), text)

    def paragraph_at(self, offset: int) -> int | None:
        """Layout index of the paragraph holding `offset`, else of the next paragraph (table-level revisions)."""

        if not self._starts:
            return None
        position = bisect.bisect_right(self._start_keys, offset) - 1
        if position >= 0 and offset < self._starts[position][1]:
            return self._starts[position][2]
        if position + 1 < len(self._starts):
            return self._starts[position + 1][2]
        return self._starts[-1][2]

    def paragraph(self, index: int, op: str) -> Node:
        if not 1 <= index <= len(self.layout.paragraphs):
            raise RefusedUnsafe(f"{op}: paragraph_index {index} is outside 1..{len(self.layout.paragraphs)}")
        node = self.para_nodes.get(index)
        if node is None:
            raise RefusedUnsafe(f"{op}: paragraph {index} is a table row-end marker and holds no text")
        return node


def load(path: str | Path):
    from .sections import read_layout

    parts = read_parts(path)
    if "word/document.xml" not in parts:
        raise RefusedUnsafe("package has no word/document.xml")
    xml = parts["word/document.xml"].decode("utf-8")
    return parts, Document(xml, read_layout(path))


# --------------------------------------------------------------------------
# Paragraph text model.
# --------------------------------------------------------------------------

@dataclass(eq=False)
class Char:
    ch: str
    run: Node
    child: Node
    offset: int
    wrappers: tuple
    field: bool
    flags: str
    rpr: str

    @property
    def deleted(self) -> bool:
        return any(node.name in _DELETED for node in self.wrappers)

    @property
    def revision(self) -> Node | None:
        return next((node for node in reversed(self.wrappers) if node.name in _WRAPPERS), None)

    @property
    def layout_visible(self) -> bool:
        # The layout text is the accepted text: no w:delText, and no tab or break inside a tracked deletion.
        return self.child.name != "w:delText" and not (self.deleted and self.child.name != "w:t")

    def fp(self) -> str:
        return self.flags + ("f" if self.field else "") + "|" + re.sub(r"\s+", " ", strip_rpr_change(self.rpr))


@dataclass(eq=False)
class Obj:
    kind: str  # "field", "opaque", "mark"
    node: Node
    run: Node | None
    wrappers: tuple


class ParagraphModel:
    """Characters and objects of one paragraph in document order."""

    def __init__(self, xml: str, paragraph: Node) -> None:
        self.xml = xml
        self.node = paragraph
        self.items: list = []
        self._fields: list[bool] = []
        self._walk(paragraph, (), "")
        self.chars = [item for item in self.items if isinstance(item, Char)]
        self.final = [item for item in self.chars if not item.deleted]
        self.text = "".join(item.ch for item in self.final)
        self.original = "".join(
            item.ch for item in self.chars if not any(node.name in _INSERTED for node in item.wrappers))
        self.position = {id(item): number for number, item in enumerate(self.items)}

    def _walk(self, node: Node, wrappers: tuple, flags: str) -> None:
        for item in node.children:
            name = item.name
            if name == "w:pPr":
                continue
            if name == "w:r":
                self._run(item, wrappers, flags)
            elif name in _WRAPPERS:
                self._walk(item, wrappers + (item,), flags)
            elif name == "w:sdt":
                content = item.child("w:sdtContent")
                if content is not None:
                    self._walk(content, wrappers + (item,), flags + "s")
            elif name in _PASS_THROUGH:
                if name == "w:fldSimple":
                    self.items.append(Obj("field", item, None, wrappers))
                self._walk(item, wrappers + (item,), flags + _PASS_THROUGH[name])
            elif name in _MARKS:
                self.items.append(Obj("mark", item, None, wrappers))
            else:
                self.items.append(Obj("opaque", item, None, wrappers))

    def _run(self, run: Node, wrappers: tuple, flags: str) -> None:
        rpr_node = run.child("w:rPr")
        rpr = self.xml[rpr_node.start:rpr_node.end] if rpr_node is not None else ""
        for item in run.children:
            name = item.name
            if name == "w:rPr":
                continue
            in_field = bool(self._fields) or "f" in flags
            if name == "w:fldChar":
                kind = item.attr(self.xml, "w:fldCharType")
                if kind == "begin":
                    self._fields.append(False)
                elif kind == "separate" and self._fields:
                    self._fields[-1] = True
                elif kind == "end" and self._fields:
                    self._fields.pop()
                self.items.append(Obj("field", item, run, wrappers))
            elif name in ("w:instrText", "w:delInstrText"):
                self.items.append(Obj("field", item, run, wrappers))
            elif name in ("w:t", "w:delText"):
                text = "" if item.closed else unescape(item.inner(self.xml))
                for offset, ch in enumerate(text):
                    self.items.append(Char(ch, run, item, offset, wrappers, in_field, flags, rpr))
            elif name in ("w:tab", "w:br", "w:cr"):
                ch = "\t" if name == "w:tab" else (
                    "\x0c" if name == "w:br" and item.attr(self.xml, "w:type") == "page" else "\v")
                self.items.append(Char(ch, run, item, 0, wrappers, in_field, flags, rpr))
            elif name in _MARKS:
                self.items.append(Obj("mark", item, run, wrappers))
            else:
                self.items.append(Obj("opaque", item, run, wrappers))

    def drawing(self) -> str | None:
        inner = self.xml[self.node.start:self.node.end]
        return next((name for name in _DRAWINGS if f"<{name}" in inner), None)


def describe_wrapper(xml: str, node: Node) -> str:
    return (f"w:{node.local} id={node.attr(xml, 'w:id')} by {node.attr(xml, 'w:author')!r} "
            f"({node.attr(xml, 'w:date')})")


# --------------------------------------------------------------------------
# Text edits (tracked or plain) inside one paragraph.
# --------------------------------------------------------------------------

def _word(ch: str) -> bool:
    return ch.isalnum() or ch == "_"


def whole_words(old: str, fps, caps, edits: list) -> list:
    """Widen edits that cut into a word to the whole word, when the word has one formatting.

    The plan-328 diff works character by character inside a changed span;
    as tracked changes that reads as fragments ("A" -> "Run-to-run a"). A
    reviewer reads whole words, so a change inside a word becomes the
    deletion of the word and the insertion of its new spelling. A word whose
    characters carry different formatting keeps the character-level edits,
    so no character changes formatting.
    """

    from .sections import Edit

    def span(edit) -> tuple[int, int]:
        start, end = edit.start, edit.end
        inside_left = start > 0 and _word(old[start - 1]) and (
            (end > start and _word(old[start])) or (end == start and start < len(old) and _word(old[start])))
        inside_right = end < len(old) and _word(old[end]) and (
            (end > start and _word(old[end - 1])) or (end == start and start > 0 and _word(old[start - 1])))
        if inside_left:
            while start > 0 and _word(old[start - 1]):
                start -= 1
        if inside_right:
            while end < len(old) and _word(old[end]):
                end += 1
        return start, end

    groups: list[list] = []
    for edit in edits:
        start, end = span(edit)
        if groups and start < groups[-1][1]:
            groups[-1][1] = max(groups[-1][1], end)
            groups[-1][2].append(edit)
        else:
            groups.append([start, end, [edit]])
    result = []
    for start, end, members in groups:
        if len(members) == 1 and (start, end) == (members[0].start, members[0].end):
            result.append(members[0])
            continue
        uniform = len({(fps[k], caps[k]) for k in range(start, end)}) == 1 and not any(
            fps[k].split("|", 1)[0] for k in range(start, end))
        if not uniform:
            result.extend(members)
            continue
        text, cursor = [], start
        for edit in members:
            text.append(old[cursor:edit.start] + edit.text)
            cursor = edit.end
        text.append(old[cursor:end])
        new = "".join(text)
        result.append(Edit(start, end, new, "replace" if new else "delete", start if new else None))
    return result


@dataclass
class _Insertion:
    text: str
    rpr: str
    order: int
    # Plan 831 S2: the caller's own w:ins this text joins (written as a plain run inside it, no new w:ins).
    inside: Node | None = None
    # That w:ins's start tag (prefix) or end tag (suffix), moved next to the new run when zero-width markers at
    # its edge must stay on the other side of the new text (`TextEditor._edge_join`); the original tag is skipped.
    prefix: str = ""
    suffix: str = ""


def _render_text(text: str, deleted: bool) -> str:
    tag = "w:delText" if deleted else "w:t"
    parts = []
    # "\t" is a tab and "\v" a manual line break (Word's own text model, as ParagraphModel reads them).
    for chunk in re.split(r"([\t\v])", text):
        if chunk == "\t":
            parts.append("<w:tab/>")
        elif chunk == "\v":
            parts.append("<w:br/>")
        elif chunk:
            parts.append(f'<{tag} xml:space="preserve">{escape(chunk)}</{tag}>')
    return "".join(parts)


class _Writer:
    """Emits new revision elements with ids ascending in document order."""

    def __init__(self, first_id: int, author: str, stamp: str, tracked: bool) -> None:
        self.next_id, self.author, self.stamp, self.tracked = first_id, author, stamp, tracked
        self.added: list[dict] = []

    def _attrs(self, kind: str, text: str) -> str:
        number = self.next_id
        self.next_id += 1
        self.added.append({"id": str(number), "type": kind, "author": self.author, "text": text})
        return f'w:id="{number}" w:author="{_attr_escape(self.author)}" {self.stamp}'

    def insertion(self, item: _Insertion) -> str:
        run = f"<w:r>{item.rpr}{_render_text(item.text, False)}</w:r>"
        if item.inside is not None:
            return run  # typed inside the caller's own insertion: part of that w:ins, as Word types it
        return f"<w:ins {self._attrs('ins', item.text)}>{run}</w:ins>" if self.tracked else run

    def deletion(self, run_xml: str, text: str) -> str:
        return f"<w:del {self._attrs('del', text)}>{run_xml}</w:del>" if self.tracked else ""


def _content(run: Node) -> list[Node]:
    return [item for item in run.children if item.name != "w:rPr"]


def _char_units(xml: str, child: Node) -> int:
    if child.name in ("w:t", "w:delText"):
        return 0 if child.closed else len(unescape(child.inner(xml)))
    return 1 if child.name in ("w:tab", "w:br", "w:cr") else 0


class TextEditor:
    """Apply plan-328 edits (final-view offsets) to one paragraph as tracked or plain changes."""

    def __init__(self, doc: Document, index: int, op: str, *, tracked: bool, own_insertions: bool = False) -> None:
        self.doc, self.index, self.op, self.tracked = doc, index, op, tracked
        # Plan 831 S2: tracked text ops edit the author's own pending insertions in place (sweeps do not).
        self.own_insertions = tracked and own_insertions
        self.author: str | None = None
        self.edited: dict[int, Node] = {}  # own w:ins elements this edit changes, by id(node)
        self.skips: list[tuple[int, int]] = []  # own w:ins tags moved next to a joined run (`_edge_join`)
        # Collapsed bookmarks written before text typed at a point comment (`_reordered`): (target, spans).
        self.reorders: list[tuple[int, list[tuple[int, int]]]] = []
        self.xml = doc.xml
        self.node = doc.paragraph(index, op)
        self.model = ParagraphModel(self.xml, self.node)
        item = doc.layout.para(index)
        drawing = self.model.drawing()
        if drawing:
            raise RefusedUnsafe(
                f"{op}: paragraph {index} holds a drawing ({drawing}); edits there are refused because its "
                "text cannot be matched against Word reliably. Edit it in Word.")
        layout_chars = [char for char in self.model.chars if char.layout_visible]
        if "".join(char.ch for char in layout_chars) != item.text:
            raise RefusedUnsafe(
                f"{op}: paragraph {index}'s text could not be aligned with the document layout; nothing was edited")
        caps_of = {id(char): flag for char, flag in zip(layout_chars, item.caps)}
        self.caps = tuple(caps_of[id(char)] for char in self.model.final)
        self.fps = tuple(char.fp() for char in self.model.final)

    # -- planning --------------------------------------------------------

    def plan(self, new_text: str) -> list:
        from .sections import _check_line, plan_text_diff

        _check_line(new_text, self.op)
        edits = plan_text_diff(self.model.text, new_text, self.caps, self.fps)
        if self.tracked:
            edits = whole_words(self.model.text, self.fps, self.caps, edits)
        return self._slide_onto_own(edits)

    def _slide_onto_own(self, edits: list) -> list:
        """A deletion the diff could place at several equal positions (" were" or "were ") goes where it removes
        the most of the caller's own inserted text (plan 831 S2), so deleting all of an own insertion empties it
        instead of leaving its space and deleting an original one. Ties keep the diff's position."""

        if not self.own_insertions or self.author is None:
            return edits
        from .sections import Edit

        text, final = self.model.text, self.model.final
        out: list = []
        for position, edit in enumerate(edits):
            if edit.mode != "delete" or edit.text:
                out.append(edit)
                continue
            size = edit.end - edit.start
            low = out[-1].end if out else 0
            high = edits[position + 1].start if position + 1 < len(edits) else len(text)
            starts = [edit.start]
            start = edit.start
            while start > low and text[start - 1] == text[start - 1 + size]:
                start -= 1
                starts.append(start)
            start = edit.start
            while start + size < high and text[start] == text[start + size]:
                start += 1
                starts.append(start)
            owned = lambda at: sum(1 for k in range(at, at + size) if self._own(final[k]) is not None)  # noqa: E731
            best = max(starts, key=lambda at: (owned(at), at == edit.start))
            out.append(edit if best == edit.start else Edit(best, best + size, "", "delete", edit.source))
        return out

    def find(self, find: str, replace: str) -> str:
        from .sections import caps_find

        hits = caps_find(self.model.text, self.caps, find)
        if len(hits) != 1:
            raise RefusedUnsafe(
                f"{self.op}: find must occur exactly once in paragraph {self.index}'s accepted text; it occurs "
                f"{len(hits)} times" + (" (deleted text does not count)" if self.model.original != self.model.text
                                       else ""))
        start = hits[0]
        return self.model.text[:start] + replace + self.model.text[start + len(find):]

    # -- refusals ----------------------------------------------------------

    def _blocked(self, what: str, node: Node) -> RefusedUnsafe:
        hint = ("use tracked_replace_text/tracked_replace_paragraph, or accept/reject that revision first"
                if not self.tracked else "accept or reject that revision first, or edit outside it")
        return RefusedUnsafe(f"{self.op}: paragraph {self.index}: {what} {describe_wrapper(self.xml, node)}; {hint}")

    def _own(self, char: Char) -> Node | None:
        """The w:ins holding `char` when it is the caller's own pending insertion (plan 831 S2): the only
        revision wrapper around the character, a w:ins (not a move), whose w:author equals the op's author
        exactly. Such text is edited in place instead of being refused."""

        if not self.own_insertions or self.author is None:
            return None
        wrappers = [node for node in char.wrappers if node.name in _WRAPPERS]
        if len(wrappers) == 1 and wrappers[0].name == "w:ins" and wrappers[0].attr(self.xml, "w:author") == self.author:
            return wrappers[0]
        return None

    def _own_nodes(self) -> list[Node]:
        """Run-level w:ins elements by the caller in this paragraph, not nested in another revision."""

        if not self.own_insertions or self.author is None:
            return []
        return [node for node in self.node.iter() if node.name == "w:ins" and not node.closed
                and node.attr(self.xml, "w:author") == self.author
                and not any(item.name in _WRAPPERS or item.name == "w:rPr" for item in node.ancestors())]

    def _check_char(self, char: Char, action: str) -> None:
        if char.field:
            raise RefusedUnsafe(
                f"{self.op}: paragraph {self.index}: the edit would {action} field text {char.ch!r}; "
                "fields (citations, cross-references) are never edited as text")
        revision = char.revision
        if revision is not None:
            own = self._own(char)
            if own is None:
                if self.own_insertions and revision.name == "w:ins":
                    raise RefusedUnsafe(
                        f"{self.op}: paragraph {self.index}: the edit would {action} text inside "
                        f"{describe_wrapper(self.xml, revision)}, another author's insertion (only insertions by "
                        f"{self.author!r} itself are edited in place); accept or reject that revision first, or "
                        "edit outside it")
                raise self._blocked(f"the edit would {action} text inside", revision)
            change = re.search(r"<w:rPrChange\b[^>]*>", char.rpr)
            if change:
                raise RefusedUnsafe(
                    f"{self.op}: paragraph {self.index}: the edit would {action} text of your own insertion "
                    f"{describe_wrapper(self.xml, own)} that carries the tracked formatting change "
                    f"{_head_describe(change.group(0))}; accept or reject the formatting change first")
            return
        change = re.search(r"<w:rPrChange\b[^>]*>", char.rpr)
        if change and not self.tracked:
            raise RefusedUnsafe(
                f"{self.op}: paragraph {self.index}: the edit would {action} text whose formatting change "
                f"{_head_describe(change.group(0))} is tracked; use the tracked operations or resolve it first")

    # -- application -------------------------------------------------------

    def apply(self, edits: list, *, author: str, first_id: int, stamp: str) -> tuple[str, list[dict]]:
        final = self.model.final
        self.author = author
        self.skips = []
        self.reorders = []
        deleted: set[int] = set()
        cuts: dict[tuple[int, int], list[_Insertion]] = {}
        points: dict[int, list[_Insertion]] = {}
        split_runs: set[int] = set()
        for order, edit in enumerate(edits):
            removed = final[edit.start:edit.end]
            for char in removed:
                self._check_char(char, "delete")
                deleted.add(id(char))
                own = self._own(char)
                if own is not None:
                    self.edited[id(own)] = own
            if removed:
                self._check_between(removed[0], removed[-1])
            if not edit.text:
                continue
            source = final[edit.source] if edit.source is not None else None
            # Text typed into an empty paragraph takes the paragraph mark's formatting, as Word types it.
            rpr = strip_rpr_change(source.rpr) if source is not None else self._mark_rpr()
            if rpr in ("<w:rPr></w:rPr>", "<w:rPr/>"):
                rpr = ""
            insertion = _Insertion(edit.text, rpr, order)
            if source is not None and not self.tracked:
                self._check_char(source, "copy the formatting of")
            if edit.mode == "replace":
                where = ("after", removed[-1])
            elif edit.mode == "after":
                where = ("after", final[edit.start - 1])
            elif edit.mode == "before":
                where = ("before", final[edit.start])
            else:
                where = ("empty", None)
            # A pure insertion (or a rewrite of own inserted text only) at the edge of the caller's own w:ins
            # joins it; replacing original text keeps its own w:del/w:ins pair.
            joinable = all(self._own(char) is not None for char in removed)
            # A pure insertion is typed at a point: Word's typing rules decide which anchors take it
            # (`_typing_point`); replacement text goes right after the text it replaces.
            self._place(where, insertion, cuts, points, split_runs, joinable, typed=edit.mode != "replace")
        # Runs that lose only some characters are split as well.
        for char in final:
            if id(char) in deleted:
                run_chars = [item for item in self.model.chars if item.run is char.run]
                others = [item for item in _content(char.run) if _char_units(self.xml, item) == 0]
                if any(id(item) not in deleted for item in run_chars) or others:
                    split_runs.add(id(char.run))
        for char in self.model.chars:
            if id(char.run) in split_runs:
                change = re.search(r"<w:rPrChange\b[^>]*>", char.rpr)
                if change:
                    raise RefusedUnsafe(
                        f"{self.op}: paragraph {self.index}: the edit would split a run whose formatting change "
                        f"{_head_describe(change.group(0))} is tracked, which would duplicate that revision; "
                        "accept or reject the formatting change first, or edit whole runs only")
        writer = _Writer(first_id, author, stamp, self.tracked)
        if self.node.closed:
            # A self-closing <w:p/> holds nothing: open it, so the new runs land inside the paragraph.
            head = re.sub(r"\s*/>\Z", ">", self.xml[self.node.start:self.node.end])
            body = "".join(writer.insertion(item) for offset in sorted(points) for item in points[offset])
            return f"{head}{body}</{self.node.name}>", writer.added
        touched = {id(char.run): char.run for char in self.model.chars
                   if id(char) in deleted or id(char.run) in split_runs}
        events: list[tuple[int, int, int, object]] = [(run.start, 1, 0, run) for run in touched.values()]
        for offset, items in points.items():
            for item in items:
                events.append((offset, 0, item.order, item))
        events.extend((start, 2, 0, end) for start, end in self.skips)
        events.sort(key=lambda event: (event[0], event[1], event[2]))
        out, cursor = [], self.node.start
        for offset, kind, _, payload in events:
            out.append(self.xml[cursor:offset])
            if kind == 0:
                out.append(payload.prefix + writer.insertion(payload) + payload.suffix)
                cursor = offset
            elif kind == 2:  # an own w:ins tag written next to the joined run instead
                cursor = payload
            else:
                out.append(self._rebuild(payload, deleted, cuts, writer))
                cursor = payload.end
        out.append(self.xml[cursor:self.node.end])
        return self._unwrap_emptied("".join(out)), writer.added

    def _unwrap_emptied(self, paragraph: str) -> str:
        """Remove an own w:ins the edit emptied of all content (its markers, if any, stay where they were)."""

        if not self.edited:
            return paragraph
        heads = {node.head(self.xml) for node in self.edited.values()}
        tree = parse(paragraph)
        empty = []
        for node in tree.iter():
            if node.name != "w:ins" or node.closed or node.head(paragraph) not in heads:
                continue
            if all(item.name in ("w:r", "w:rPr") or item.name in _MARKS or any(a.name == "w:rPr" for a in item.ancestors())
                   for item in node.iter() if item is not node):
                empty.append(node)
        for node in sorted(empty, key=lambda item: item.start, reverse=True):
            paragraph = paragraph[:node.start] + node.inner(paragraph) + paragraph[node.end:]
        return paragraph

    def in_place_report(self, new_paragraph: str) -> list[dict]:
        """Each own w:ins edited in place: id, author, date, its text before and after (after "" and removed
        true when the edit emptied it)."""

        tree = parse(new_paragraph)
        out = []
        for node in sorted(self.edited.values(), key=lambda item: item.start):
            head = node.head(self.xml)
            after = next((item for item in tree.iter() if item.name == "w:ins" and item.head(new_paragraph) == head),
                         None)
            out.append({"id": node.attr(self.xml, "w:id"), "author": node.attr(self.xml, "w:author"),
                        "date": node.attr(self.xml, "w:date"), "before": _inserted_text(self.xml, node),
                        "after": _inserted_text(new_paragraph, after) if after is not None else "",
                        "removed": after is None})
        return out

    def _mark_rpr(self) -> str:
        """The paragraph mark's run formatting (pPr/rPr) without the mark's own revision markers."""

        ppr = self.node.child("w:pPr")
        mark = ppr.child("w:rPr") if ppr is not None else None
        if mark is None or mark.closed:
            return ""
        kept = "".join(self.xml[item.start:item.end] for item in mark.children if item.name not in _MARK_REVISIONS)
        return f"<w:rPr>{kept}</w:rPr>" if kept else ""

    def _check_between(self, first: Char, last: Char) -> None:
        low, high = self.model.position[id(first)], self.model.position[id(last)]
        for item in self.model.items[low:high + 1]:
            if isinstance(item, Obj) and item.kind in ("field", "opaque"):
                label = "a field" if item.kind == "field" else f"a {item.node.name} element"
                raise RefusedUnsafe(
                    f"{self.op}: paragraph {self.index}: the deleted text spans {label}; "
                    "edit the text on either side of it separately")
            if isinstance(item, Char) and item.deleted and any(node.name == "w:moveFrom" for node in item.wrappers):
                raise self._blocked("the deleted text spans moved text", item.revision)

    def _place(self, where, insertion: _Insertion, cuts, points, split_runs, joinable: bool = False,
               typed: bool = False) -> None:
        side, char = where
        if side == "empty":
            # After everything the paragraph holds (pPr, markers, empty or deleted runs): a collapsed bookmark
            # there stays before the typed text, as in Word; a point comment, or a comment starting there,
            # stays after it (`_typing_point`). A self-closing <w:p/> is opened in apply.
            points.setdefault(self._typing_point(self.node.children, None, "empty", insertion), []).append(insertion)
            return
        if char.field:
            raise RefusedUnsafe(
                f"{self.op}: paragraph {self.index}: the new text {insertion.text!r} would land inside a field "
                "result; add it outside the citation or field")
        content = _content(char.run)
        position = content.index(char.child)
        units = _char_units(self.xml, char.child)
        if side == "after":
            at_edge = (char.offset == units - 1) and position == len(content) - 1
        else:
            at_edge = char.offset == 0 and position == 0
        own = self._own(char)  # anchored on the caller's own inserted text: the new text goes inside it
        if own is not None:
            insertion.inside = own
            self.edited[id(own)] = own
        if at_edge:
            if own is not None:  # stays inside the own insertion, next to its run
                edge = self._own_edge(char.run, own, side)
                if edge and self._edge_join(own, edge, self._edge_items(own, edge == "end"), insertion, points):
                    return
                points.setdefault(char.run.end if side == "after" else char.run.start, []).append(insertion)
                return
            node = self._hoisted(char.run, side, insertion)
            offset = node.end if side == "after" else node.start
            if joinable and self._join(offset, side, insertion, points):
                return
            if typed and node.parent is not None:
                offset = self._typing_point(node.parent.children, node, side, insertion)
            points.setdefault(offset, []).append(insertion)
            return
        revision = char.revision
        if revision is not None and own is None:
            raise self._blocked(f"the new text {insertion.text!r} would land strictly inside", revision)
        split_runs.add(id(char.run))
        boundary = char.offset + 1 if side == "after" else char.offset
        cuts.setdefault((id(char.run), (id(char.child), boundary)), []).append(insertion)

    def _join(self, offset: int, side: str, insertion: _Insertion, points) -> bool:
        """An insertion right at the edge of the caller's own w:ins goes inside it (Word extends your own
        insertion as you type next to it): after a character, an own w:ins starting at `offset`; before one, an
        own w:ins ending there. Zero-width markers (comment and bookmark ranges, proofErr, permissions,
        comment-reference runs) between `offset` and that edge do not stop the join (`_edge_join`). False
        leaves the text for a new w:ins at `offset`."""

        for node in self._own_nodes():
            if side == "before":  # the new text precedes a character after `node`: node's end edge
                outside = self._edge_items(node, True)
                reach = outside[-1].end if outside else node.end
                if node.end <= offset and reach <= offset and not self.xml[reach:offset].strip():
                    return self._edge_join(node, "end", outside, insertion, points)
            else:  # the new text follows a character before `node`: node's start edge
                outside = self._edge_items(node, False)
                reach = outside[0].start if outside else node.start
                if offset <= reach and offset <= node.start and not self.xml[offset:reach].strip():
                    return self._edge_join(node, "start", outside, insertion, points)
        return False

    @staticmethod
    def _zero_width(node: Node) -> bool:
        """A marker element of `_EDGE_MARKS`, or a run holding only a comment reference."""

        if node.name in _EDGE_MARKS:
            return True
        if node.name == "w:r" and not node.closed:
            names = [item.name for item in node.children]
            return "w:commentReference" in names and all(name in ("w:rPr", "w:commentReference") for name in names)
        return False

    def _edge_items(self, node: Node, forward: bool) -> list[Node]:
        """The zero-width siblings right after `node` (forward) or right before it, in document order."""

        if node.parent is None:
            return []
        siblings = node.parent.children
        at = next(number for number, item in enumerate(siblings) if item is node)
        return self._zero_width_run(siblings[at + 1:] if forward else siblings[:at][::-1], node, forward)

    def _zero_width_run(self, items: list[Node], start: Node, forward: bool) -> list[Node]:
        out, previous = [], start
        for item in items:
            gap = self.xml[previous.end:item.start] if forward else self.xml[item.end:previous.start]
            if not self._zero_width(item) or gap.strip():
                break
            out.append(item)
            previous = item
        return out if forward else out[::-1]

    def _own_edge(self, run: Node, own: Node, side: str) -> str | None:
        """"end" when `run` is the last run with content of `own` and the new text follows it, "start" when it
        is the first and the new text precedes it; else None."""

        if run.parent is not own:
            return None
        children = own.children
        at = next(number for number, item in enumerate(children) if item is run)
        if side == "after" and len(self._zero_width_run(children[at + 1:], run, True)) == len(children) - at - 1:
            return "end"
        if side == "before" and len(self._zero_width_run(children[:at][::-1], run, False)) == at:
            return "start"
        return None

    def _point_comment(self, cid: str | None, units: list[Node]) -> bool:
        """True when comment `cid`'s range start and end are both among `units`, start first: a zero-length
        (point) comment at the join point (`units` hold no accepted text)."""

        at = {item.name: number for number, item in enumerate(units)
              if item.name in ("w:commentRangeStart", "w:commentRangeEnd") and item.attr(self.xml, "w:id") == cid}
        return len(at) == 2 and at["w:commentRangeStart"] < at["w:commentRangeEnd"]

    def _edge_role(self, item: Node, units: list[Node], position: int) -> str:
        """Where a zero-width item at the join point must be relative to the new text, as Word types it and as
        `sections.map_anchor_point` checks it: "before" (every anchor end, a bookmark start, both ends of a
        collapsed bookmark, a comment reference with no range), "after" (a comment start, and the start, end and
        reference of a point comment: Word types before it, plan 924) or "any"."""

        name = item.name
        if name in ("w:bookmarkEnd", "w:bookmarkStart"):
            return "before"
        if name in ("w:commentRangeStart", "w:commentRangeEnd"):
            point = self._point_comment(item.attr(self.xml, "w:id"), units)
            if name == "w:commentRangeEnd":
                return "after" if point else "before"
            return "after"
        if name == "w:r":
            cid = item.child("w:commentReference").attr(self.xml, "w:id")
            if self._point_comment(cid, units):
                return "after"
            ranged = re.search(r"<w:commentRange(?:Start|End)\b[^>]*\bw:id=([\"'])" + re.escape(cid or "") + r"\1",
                               self.xml)
            return "any" if ranged else "before"
        return "any"

    def _edge_join(self, own: Node, edge: str, outside: list[Node], insertion: _Insertion, points) -> bool:
        """Put the new run inside `own` at its `edge` ("start" or "end"), keeping every zero-width item there
        on the side of the new text Word's anchor rules put it (`_edge_role`): items that must precede it stay
        before it, items that must follow it stay after it, and the others stay where they are. When that needs
        a marker on the other side of own's start or end tag, the tag is written next to the new run instead
        (`skips`, `prefix`/`suffix`), so the marker crosses only the tag, never text, and the w:ins keeps its one
        start tag. Only `_EDGE_MARKS` elements cross a tag (a comment-reference run never enters or leaves the
        w:ins). False (nothing changed) when no order fits; the caller then keeps today's placement."""

        inner_items = own.children
        if edge == "end":
            last = next((item for item in reversed(inner_items) if not self._zero_width(item)), None)
            inner = self._zero_width_run(inner_items[inner_items.index(last) + 1:], last, True) if last else []
            if last is None or len(inner) != len(inner_items) - inner_items.index(last) - 1:
                return False
            units, tag_at = inner + outside, len(inner)
        else:
            first = next((item for item in inner_items if not self._zero_width(item)), None)
            at = inner_items.index(first) if first is not None else -1
            inner = self._zero_width_run(inner_items[:at][::-1], first, False) if first is not None else []
            if first is None or len(inner) != at:
                return False
            units, tag_at = outside + inner, len(outside)
        roles = [self._edge_role(item, units, number) for number, item in enumerate(units)]
        low = max((number + 1 for number, role in enumerate(roles) if role == "before"), default=0)
        high = min((number for number, role in enumerate(roles) if role == "after"), default=len(units))
        if low > high:
            return False
        cut = min(max(tag_at, low), high)
        moved = units[cut:tag_at] if cut < tag_at else units[tag_at:cut]
        if any(item.name not in _EDGE_MARKS for item in moved):
            return False
        if cut == tag_at:
            point = own.close_start() if edge == "end" else own.head_end
        else:
            tag = (own.close_start(), own.end) if edge == "end" else (own.start, own.head_end)
            if tag in self.skips:
                return False
            self.skips.append(tag)
            if edge == "end":
                insertion.suffix = self.xml[tag[0]:tag[1]]
            else:
                insertion.prefix = self.xml[tag[0]:tag[1]]
            point = units[cut].start if cut < len(units) else units[-1].end
        insertion.inside = own
        self.edited[id(own)] = own
        points.setdefault(point, []).append(insertion)
        return True

    def _typing_point(self, siblings: list[Node], node: Node | None, side: str,
                      insertion: _Insertion | None = None) -> int:
        """Where text typed next to `node` goes among the zero-width items beside it, by Word's typing rules
        (AGENTS §4; `sections.map_anchor_point` checks them): after every anchor end, a bookmark start and both
        ends of a collapsed bookmark, before a comment start and before a point comment (its start, end and
        reference). `side` "after" scans the items following `node`, "before" those preceding it, "empty" the
        items at the end of an empty paragraph (`node` None). Deleted text (no anchor marks inside) and empty
        runs may sit among them; they are not in the accepted text. Within the allowed range the text stays as
        close to `node` as it can (after a comment reference that follows its range end). When no order fits
        (a point comment, then a collapsed bookmark), the markers that must precede the text are written right
        before it (`insertion.prefix`, the originals skipped), as Word reorders them; they cross only other
        zero-width items, never text. Otherwise the near edge is kept and the anchor check refuses."""

        def passable(item: Node) -> bool:
            if self._zero_width(item):
                return True
            if item.name == "w:r" and not item.closed:
                return all(child.name == "w:rPr" for child in item.children)
            if item.name in _DELETED and not item.closed:
                return not any(inner.name in _ANCHOR_MARKS for inner in item.iter())
            return False

        if side == "after":
            at = next(number for number, item in enumerate(siblings) if item is node)
            ordered, edge, previous = siblings[at + 1:], node.end, node
        elif side == "before":
            at = next(number for number, item in enumerate(siblings) if item is node)
            ordered, edge, previous = siblings[:at][::-1], node.start, node
        else:
            ordered, edge, previous = siblings[::-1], self.node.close_start(), None
        units: list[Node] = []
        for item in ordered:
            if side == "after":
                gap = self.xml[previous.end:item.start]
            else:
                gap = self.xml[item.end:previous.start] if previous is not None else self.xml[item.end:edge]
            if not passable(item) or gap.strip():
                break
            units.append(item)
            previous = item
        if side != "after":
            units.reverse()
        if not units:
            return edge
        roles = ["any" if not self._zero_width(item) else self._edge_role(item, units, number)
                 for number, item in enumerate(units)]
        low = max((number + 1 for number, role in enumerate(roles) if role == "before"), default=0)
        high = min((number for number, role in enumerate(roles) if role == "after"), default=len(units))
        if low > high:
            return self._reordered(units, roles, high, low, insertion) or edge
        if side == "after":
            cut = low
            while cut < high and units[cut].name == "w:r" and self._zero_width(units[cut]):
                cut += 1  # a comment reference stays with the range end before it
        else:
            cut = high
        if cut == len(units):
            return units[-1].end if side == "after" else edge
        return units[cut].start

    def _reordered(self, units: list[Node], roles: list[str], high: int, low: int,
                   insertion: _Insertion | None) -> int | None:
        """The typing point when an item that must follow the new text (`roles` "after", first at `high`)
        precedes one that must precede it (last before `low`): the text goes before `units[high]`, and every
        "before" item from there on is written right before the text instead (Word's own order: a collapsed
        bookmark after a point comment ends up before the typed text, the comment after it). A moved comment
        range end takes its reference run along. Only `_EDGE_MARKS` move, across zero-width items only; None
        (nothing changed) when anything else would have to move or the items were already moved."""

        if insertion is None or insertion.prefix:
            return None
        moved: list[Node] = []
        for number in range(high, low):
            item = units[number]
            if roles[number] != "before" or any(item is done for done in moved):
                continue
            if item.name not in _EDGE_MARKS:
                return None
            moved.append(item)
            following = units[number + 1] if number + 1 < len(units) else None
            if (item.name == "w:commentRangeEnd" and following is not None and following.name == "w:r"
                    and self._zero_width(following) and roles[number + 1] == "any"
                    and following.child("w:commentReference").attr(self.xml, "w:id") == item.attr(self.xml, "w:id")):
                moved.append(following)
        spans = [(item.start, item.end) for item in moved]
        if any(span in self.skips for span in spans):
            return None
        self.skips.extend(spans)
        self.reorders.append((units[high].start, spans))
        insertion.prefix = "".join(self.xml[start:end] for start, end in spans)
        return units[high].start

    def reordered_input(self) -> str:
        """The input paragraph with the `_reordered` markers moved as the edit moves them (no text changed):
        what Reject of the new revisions must give back."""

        moves = sorted([(target, 0, spans) for target, spans in self.reorders]
                       + [(start, 1, end) for _, spans in self.reorders for start, end in spans],
                       key=lambda item: (item[0], item[1]))
        out, cursor = [], self.node.start
        for offset, kind, payload in moves:
            out.append(self.xml[cursor:offset])
            if kind == 0:
                out.append("".join(self.xml[start:end] for start, end in payload))
                cursor = offset
            else:
                cursor = payload
        out.append(self.xml[cursor:self.node.end])
        return "".join(out)

    def _hoisted(self, run: Node, side: str, insertion: _Insertion) -> Node:
        """`run`, or the revision wrapper it is the `side` edge of (stepping out of nested ones)."""

        node = run
        while node.parent is not None and node.parent.name in _WRAPPERS:
            parent = node.parent
            siblings = parent.children
            edge = siblings[-1] is node if side == "after" else siblings[0] is node
            gap = (self.xml[node.end:parent.close_start()] if side == "after"
                   else self.xml[parent.head_end:node.start])
            if not edge or gap.strip():
                raise self._blocked(f"the new text {insertion.text!r} would land strictly inside", parent)
            node = parent
        return node

    def _hoist(self, run: Node, side: str, insertion: _Insertion) -> int:
        """Offset just outside `run`, stepping out of revision wrappers whose edge it is."""

        node = self._hoisted(run, side, insertion)
        return node.end if side == "after" else node.start

    def _rebuild(self, run: Node, deleted: set[int], cuts, writer: _Writer) -> str:
        xml = self.xml
        head = run.head(xml) if not run.closed else "<w:r>"
        if run.closed:
            return xml[run.start:run.end]
        rpr_node = run.child("w:rPr")
        rpr = xml[rpr_node.start:rpr_node.end] if rpr_node is not None else ""
        chars = {(id(char.child), char.offset): char for char in self.model.chars if char.run is run}
        # Characters of the caller's own insertion are removed outright (no w:del), as Word does (plan 831 S2).
        own = any(self._own(char) is not None for char in chars.values())
        run_cuts = {key[1]: value for key, value in cuts.items() if key[0] == id(run)}
        pieces: list[tuple[str, list]] = []
        state = {"kind": None, "parts": []}

        def flush() -> None:
            if state["parts"]:
                pieces.append((state["kind"], state["parts"]))
            state["parts"] = []

        def add(kind: str, part) -> None:
            if kind != state["kind"]:
                flush()
                state["kind"] = kind
            state["parts"].append(part)

        def cut(key) -> None:
            if key in run_cuts:
                flush()
                state["kind"] = None
                for item in run_cuts[key]:
                    pieces.append(("ins", item))

        for child in run.children:
            if child is rpr_node:
                continue
            units = _char_units(xml, child)
            if child.name in ("w:t", "w:delText") or units:
                for offset in range(units):
                    cut((id(child), offset))
                    char = chars[(id(child), offset)]
                    kind = ("drop" if own else "del") if id(char) in deleted else "keep"
                    add(kind, ("char", char) if child.name in ("w:t", "w:delText") else ("raw", xml[child.start:child.end], char))
                cut((id(child), units))
            else:
                add("keep", ("raw", xml[child.start:child.end], None))
        flush()
        out = []
        for kind, parts in pieces:
            if kind == "ins":
                out.append(writer.insertion(parts))
                continue
            if kind == "drop":
                continue
            body, text = [], []
            for part in parts:
                if part[0] == "char":
                    if body and isinstance(body[-1], list):
                        body[-1].append(part[1].ch)
                    else:
                        body.append([part[1].ch])
                    text.append(part[1].ch)
                else:
                    body.append(part[1])
                    if part[2] is not None:
                        text.append(part[2].ch)
            rendered = "".join(_render_text("".join(item), kind == "del") if isinstance(item, list) else item
                               for item in body)
            run_xml = head + rpr + rendered + "</w:r>"
            out.append(writer.deletion(run_xml, "".join(text)) if kind == "del" else run_xml)
        return "".join(out)


def _inserted_text(xml: str, node: Node) -> str:
    """Accepted text of a w:ins element: w:t text, tabs and breaks, not text inside a nested deletion."""

    out = []

    def walk(item: Node) -> None:
        for child in item.children:
            if child.name in _DELETED or child.name == "w:rPr":
                continue
            if child.name == "w:t":
                out.append("" if child.closed else unescape(child.inner(xml)))
            elif child.name == "w:tab":
                out.append("\t")
            elif child.name in ("w:br", "w:cr"):
                out.append("\x0c" if child.name == "w:br" and child.attr(xml, "w:type") == "page" else "\v")
            else:
                walk(child)

    walk(node)
    return "".join(out)


def _head_describe(head: str) -> str:
    return (f"w:rPrChange id={attr_value(head, 'w:id')} by {attr_value(head, 'w:author') or ''!r} "
            f"({attr_value(head, 'w:date')})")


# --------------------------------------------------------------------------
# ElementTree side: an independent resolution used only to verify writes.
# --------------------------------------------------------------------------

def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


# Markers that may sit between two paragraphs (body or cell level) and so join the merged paragraph, and the
# markers a removed wrapper leaves behind (a comment or bookmark is never lost with the text around it).
_BETWEEN = frozenset({"w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd", "w:proofErr",
                      "w:permStart", "w:permEnd"} | {"w:" + name for name in MOVE_RANGES})
_HOIST = frozenset({"w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd", "w:permStart",
                    "w:permEnd"} | {"w:" + name for name in MOVE_RANGES})
_ET_BETWEEN = frozenset(W + name[2:] for name in _BETWEEN)
_ET_HOIST = frozenset(W + name[2:] for name in _HOIST)


def mark_removed(name: str, mode: str) -> bool:
    """True when resolving a paragraph-mark revision `name` removes the mark (the paragraph merges with the
    next one): accepting a deletion (w:del, w:moveFrom) or rejecting an insertion (w:ins, w:moveTo)."""

    return (name in ("del", "moveFrom")) == (mode == "accept")


def wrapper_removed(name: str, mode: str) -> bool:
    """True when resolving run-level `name` removes its content (accepted deletion, rejected insertion)."""

    return (name in ("ins", "moveTo")) != (mode == "accept")


def _et_retag(node: ET.Element) -> None:
    """Deleted text back to text, outside nested deletions (which keep their own state)."""

    for child in node:
        if child.tag in (W + "del", W + "moveFrom"):
            continue
        if child.tag == W + "delText":
            child.tag = W + "t"
        elif child.tag == W + "delInstrText":
            child.tag = W + "instrText"
        _et_retag(child)


def _et_hoisted(node: ET.Element) -> list[ET.Element]:
    out: list[ET.Element] = []
    for child in node:
        if child.tag in _ET_HOIST:
            out.append(child)
        elif child.tag == W + "r":
            if child.find(W + "commentReference") is not None:
                out.append(child)
        else:
            out.extend(_et_hoisted(child))
    return out


_ET_COMMENT_MARKS = (W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference")


def _et_free_comments(root: ET.Element, parents: dict, chosen, mode: str) -> dict:
    """Comments a whole-table rejection takes with it, as Word deletes a comment whose every marker lies in a
    table it removes: for each top-level table every row of which has a chosen w:ins row mark, the ids of the
    comments all of whose range and reference markers lie inside it (plan 916 S5)."""

    if mode != "reject":
        return {}
    everywhere = Counter(el.get(W + "id") for el in root.iter() if el.tag in _ET_COMMENT_MARKS)
    out = {}
    for table in root.iter(W + "tbl"):
        body = parents.get(table)
        rows = table.findall(W + "tr")
        if body is None or body.tag != W + "body" or not rows or not all(
                any(chosen(mark) for mark in row.findall(f"{W}trPr/{W}ins")) for row in rows):
            continue
        inside = Counter(el.get(W + "id") for el in table.iter() if el.tag in _ET_COMMENT_MARKS)
        out[table] = {cid for cid, count in inside.items() if count == everywhere[cid]}
    return out


def _et_comment_marker(child: ET.Element, free: set) -> bool:
    """A comment range marker, or a run holding only comment references, of a comment in ``free``."""

    if child.tag in (W + "commentRangeStart", W + "commentRangeEnd"):
        return child.get(W + "id") in free
    if child.tag != W + "r":
        return False
    references = child.findall(W + "commentReference")
    return bool(references) and all(item.tag in (W + "rPr", W + "commentReference") for item in child) \
        and all(item.get(W + "id") in free for item in references)


def _et_remove_row(row: ET.Element, parents: dict, merging: list, mode: str, free: set = frozenset()) -> None:
    """Remove a table row whose row mark goes (rejected insertion, accepted deletion), as Word does: only
    when everything in it went with the same selection (every run removed, every paragraph mark removed).
    Rejecting the insertion of every row of a top-level table removes the whole table, as Word does, when the
    table holds nothing but its properties, grid and those rows and no other revision (plan 916 S5); the
    markers of the comments in ``free`` (wholly inside that table) go with it."""

    table = parents.get(row)
    if table is None or table.tag != W + "tbl":
        raise RefusedUnsafe("a row-mark resolution on a row outside a table")
    for cell in row:
        if cell.tag in (W + "trPr", W + "tblPrEx"):
            continue
        if cell.tag != W + "tc":
            raise RefusedUnsafe(f"a row-mark resolution would remove a row holding w:{_local(cell.tag)}")
        for item in cell:
            if item.tag == W + "tcPr":
                continue
            if item.tag != W + "p":
                raise RefusedUnsafe(f"a row-mark resolution would remove a row holding w:{_local(item.tag)}")
            left = [child for child in item if child.tag not in (W + "pPr", W + "proofErr")
                    and not _et_comment_marker(child, free)]
            if left or item not in merging:
                raise RefusedUnsafe("a row-mark resolution would remove a row whose content is not all "
                                    "resolved away by the same selection")
    for cell in row:
        for item in cell:
            if item in merging:
                merging.remove(item)
    table.remove(row)
    if table.find(W + "tr") is None:
        body = parents.get(table)
        if mode != "reject" or body is None or body.tag != W + "body" \
                or any(child.tag not in (W + "tblPr", W + "tblGrid") for child in table) \
                or any(el.tag.startswith(W) and _local(el.tag) in REV_TAGS for el in table.iter()):
            raise RefusedUnsafe("a row-mark resolution would leave a table with no rows")
        body.remove(table)


def et_resolve(root: ET.Element, chosen, mode: str, drop=None) -> None:
    """Accept (mode "accept") or reject, in place, every revision element with chosen(el) true.

    Processed in reverse document order, so a revision nested in another is resolved before its container
    moves it. `drop(el)` names extra elements to remove first (the range markers of resolved moves). A
    removed wrapper leaves its bookmark/comment markers (and comment-reference runs) where it stood. A
    removed paragraph mark merges the paragraph into the next one, which keeps its own pPr (Word's rule).
    A removed row mark (rejected w:ins, accepted w:del in w:trPr) removes its row, as Word does, when the
    selection resolved away everything in it; an emptied trPr is dropped (Word drops it).
    Rejecting the insertion of every row of a top-level table removes the table (Word's rule), with the
    comments wholly inside it.
    Raises on what the string engine refuses as well (cell revisions, partial rows, impossible merges).
    """

    parents = {child: parent for parent in root.iter() for child in parent}
    if drop is not None:
        for el in [el for el in root.iter() if drop(el)]:
            parents[el].remove(el)
    order = {el: number for number, el in enumerate(root.iter())}
    targets = [el for el in root.iter() if el.tag.startswith(W) and _local(el.tag) in REV_TAGS and chosen(el)]
    free_comments = _et_free_comments(root, parents, chosen, mode)
    merging: list[ET.Element] = []
    touched_rows: list[ET.Element] = []  # trPr elements a resolution emptied go, as Word drops them
    for el in reversed(targets):
        name, parent = _local(el.tag), parents[el]
        if parent.tag == W + "trPr":
            touched_rows.append(parent)
        if name in PROP_REVS:
            if mode == "accept":
                parent.remove(el)
                continue
            old = el.find(W + _local(parent.tag))
            old_children = list(old) if old is not None else []
            if _local(parent.tag) == "rPr":
                keep = [k for k in parent if _local(k.tag) in RUN_REVS]
                parent[:] = keep + old_children
            elif "w:" + _local(parent.tag) in _PROP_KEEP or _local(parent.tag) == "tblPr":
                kept = _PROP_KEEP.get("w:" + _local(parent.tag), ())
                parent[:] = old_children + [k for k in parent if "w:" + _local(k.tag) in kept]
            else:
                raise RefusedUnsafe(f"reject of w:{name} is not supported")
            continue
        if name in CELL_REVS:
            raise RefusedUnsafe(f"w:{name} cannot be resolved here")
        grand = parents.get(parent)
        if _local(parent.tag) == "rPr" and grand is not None and grand.tag == W + "pPr":
            parent.remove(el)
            if mark_removed(name, mode):
                paragraph = parents[grand]
                if paragraph not in merging:
                    merging.append(paragraph)
            continue
        if _local(parent.tag) == "trPr" and name in ("ins", "del"):
            # cell content comes later in document order, so it is resolved already (reverse order)
            parent.remove(el)
            if mark_removed(name, mode):
                row = parents[parent]
                _et_remove_row(row, parents, merging, mode, free_comments.get(parents.get(row), set()))
            continue
        if _local(parent.tag) in ("rPr", "trPr"):
            if (name == "ins") == (mode == "accept"):
                parent.remove(el)
                continue
            raise RefusedUnsafe(f"{mode} of a {_local(parent.tag)} w:{name} mark would merge rows")
        index = list(parent).index(el)
        parent.remove(el)
        if wrapper_removed(name, mode):
            moved = _et_hoisted(el)
        else:
            if name in ("del", "moveFrom"):
                _et_retag(el)
            moved = list(el)
        for offset, child in enumerate(moved):
            parent.insert(index + offset, child)
            parents[child] = parent
    for trpr in touched_rows:
        row = parents[trpr]
        if not len(trpr) and trpr in list(row):
            row.remove(trpr)
    for paragraph in sorted(merging, key=lambda el: order[el], reverse=True):
        parent = parents[paragraph]
        siblings = list(parent)
        at = siblings.index(paragraph) + 1
        between = []
        while at < len(siblings) and siblings[at].tag in _ET_BETWEEN:
            between.append(siblings[at])
            at += 1
        if at == len(siblings) or siblings[at].tag != W + "p":
            raise RefusedUnsafe("a paragraph-mark resolution would merge a paragraph with a non-paragraph")
        following = siblings[at]
        ppr = paragraph.find(W + "pPr")
        if ppr is not None and (ppr.find(W + "sectPr") is not None or any(
                _local(item.tag) in REV_TAGS for item in ppr.iter() if item is not ppr)):
            raise RefusedUnsafe("a paragraph-mark resolution would drop a section break or a revision")
        content = [child for child in paragraph if child.tag != W + "pPr"] + between
        start = 1 if len(following) and following[0].tag == W + "pPr" else 0
        for offset, child in enumerate(content):
            following.insert(start + offset, child)
            parents[child] = following
        parent.remove(paragraph)
        for item in between:
            parent.remove(item)


_OMML = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def _et_final_chars(paragraph: ET.Element) -> list[tuple[str, str]]:
    from .sections import canonical

    out: list[tuple[str, str]] = []

    def key(run: ET.Element) -> str:
        rpr = run.find(W + "rPr")
        if rpr is None:
            return ""
        clone = copy.deepcopy(rpr)
        for change in clone.findall(W + "rPrChange"):
            clone.remove(change)
        return canonical(clone) if len(clone) else ""

    def walk(node: ET.Element, deleted: bool) -> None:
        for child in node:
            name = _local(child.tag)
            if name == "pPr":
                continue
            if name == "r":
                if deleted:
                    continue
                run_key = key(child)
                for part in child:
                    part_name = _local(part.tag)
                    if part_name == "t":
                        out.extend((ch, run_key) for ch in part.text or "")
                    elif part_name == "tab":
                        out.append(("\t", run_key))
                    elif part_name in ("br", "cr"):
                        out.append(("\x0c" if part.get(W + "type") == "page" and part_name == "br" else "\v", run_key))
                continue
            if name in ("sdtPr", "sdtEndPr") or child.tag.startswith(_OMML):
                continue  # OMML math is opaque to ParagraphModel too (its m:r/m:t are not text runs)
            walk(child, deleted or name in ("del", "moveFrom"))

    walk(paragraph, False)
    return out


def _fragment(doc_xml: str, fragment: str) -> ET.Element:
    """Parse a paragraph (or other) fragment with the document's namespace declarations."""

    head = re.search(r"<w:document\b[^>]*>", doc_xml).group(0)
    return ET.fromstring(head + fragment + "</w:document>")[0]


def _revision_raw(xml: str, node: Node) -> str:
    return xml[node.start:node.end]


def _check_stamps(xml: str, nodes: list[Node], author: str, op: str) -> None:
    previous = None
    for node in nodes:
        number = int(node.attr(xml, "w:id"))
        if previous is not None and number <= previous:
            raise VerificationFailed(f"{op}: new revision ids do not ascend left to right")
        previous = number
        if node.attr(xml, "w:author") != author or not _STAMP.match(node.attr(xml, "w:date") or ""):
            raise VerificationFailed(f"{op}: new revision has the wrong author or date format")


# --------------------------------------------------------------------------
# Operations.
# --------------------------------------------------------------------------

def _detached_comments(before: str, after: str) -> list[str]:
    """Comment ids whose range held text before the resolution and holds none after it.

    The comment and its anchors are kept (Word would delete such a comment); the caller is told.
    """
    def spans(xml):
        # parsed markers (any quote style or spacing), the text between a start and its end in document order
        out: dict[str, str] = {}
        open_ids: dict[str, list[str]] = {}
        for node in parse(xml).iter():
            if node.name in ("w:commentRangeStart", "w:commentRangeEnd"):
                cid = node.attr(xml, "w:id")
                if cid is None:
                    continue
                if node.name == "w:commentRangeStart":
                    open_ids.setdefault(cid, [])
                elif cid in open_ids:
                    out[cid] = "".join(open_ids.pop(cid))
            elif node.name in ("w:t", "w:delText") and open_ids and not node.closed:
                for pieces in open_ids.values():
                    pieces.append(node.inner(xml))
        return out
    old, new = spans(before), spans(after)
    return [cid for cid, text in new.items() if not text and old.get(cid)]


def _result(source, output, op: str, doc: Document, **extra) -> dict:
    from .audit import audit_docx, file_sha256

    after = audit_docx(output)
    return {
        "source": str(Path(source).resolve()),
        "output": str(Path(output).resolve()),
        "operation": op,
        "engine": "direct_xml",
        "output_sha256": file_sha256(output),
        "paragraph_count": after.paragraph_count,
        "index_delta": 0,
        "indices_after_shifted": False,
        **extra,
    }


def text_edit(source, output, *, op: str, paragraph_index: int, tracked: bool, author: str = DEFAULT_AUTHOR,
              find: str | None = None, replace: str | None = None, text: str | None = None,
              now: dt.datetime | None = None) -> dict:
    """Tracked (or plain) replace in one paragraph; verified, then written to the new file `output`."""

    from .sections import read_layout

    author = (author or "").strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{op}: author must not be empty")
    parts, doc = load(source)
    editor = TextEditor(doc, paragraph_index, op, tracked=tracked, own_insertions=True)
    editor.author = author  # own insertions are known before planning (plan 831 S2)
    if find is not None:
        if not find:
            raise RefusedUnsafe(f"{op}: find must not be empty")
        requested = editor.find(find, replace or "")
    else:
        requested = text if text is not None else ""
    edits = editor.plan(requested)
    if not edits:
        raise RefusedUnsafe(f"{op}: paragraph {paragraph_index} already reads as requested; nothing to change")
    first_id = next_revision_id(parts)
    stamp = revision_stamp(doc.xml, now)
    new_paragraph, added = editor.apply(edits, author=author, first_id=first_id, stamp=stamp)
    node = editor.node
    new_xml = doc.xml[:node.start] + new_paragraph + doc.xml[node.end:]
    expected_text = _expected_text(editor.model.text, edits)
    if not _same_display(expected_text, requested, _planned_caps(editor.caps, edits)):
        # the planner's edits must spell out what the caller asked for, not merely what the planner expects (F3)
        raise VerificationFailed(f"{op}: the planned edits do not give the requested text (D5)",
                                 expected=requested, actual=expected_text)
    _verify_text_edit(doc, editor, new_xml, new_paragraph, edits, expected_text, added, author, tracked, op,
                      requested=requested)
    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, op)
        after_layout = read_layout(output)
        _verify_anchors(doc, editor, after_layout, edits, op)
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    return _result(
        source, output, op, doc,
        paragraph_index=paragraph_index,
        tracked=tracked,
        author=author if tracked else None,
        text_before=editor.model.text,
        text_after=expected_text,
        requested_text=requested,
        diff=[edit.describe(editor.model.text) for edit in edits],
        revisions_added=added if tracked else [],
        **({"revised_in_place": editor.in_place_report(new_paragraph)} if editor.edited else {}),
    )


def _expected_text(old: str, edits: list) -> str:
    out, cursor = [], 0
    for edit in edits:
        out.append(old[cursor:edit.start] + edit.text)
        cursor = edit.end
    out.append(old[cursor:])
    return "".join(out)


def _planned_caps(caps, edits: list) -> list[bool]:
    """Per character of the edited text: does it display in capitals? Kept characters keep their input flag;
    new characters take their source character's flag (no source: not caps)."""

    out: list[bool] = []
    cursor = 0
    for edit in edits:
        out.extend(caps[cursor:edit.start])
        out.extend([bool(caps[edit.source]) if edit.source is not None else False] * len(edit.text))
        cursor = edit.end
    out.extend(caps[cursor:])
    return out


def _same_display(text: str, requested: str, caps: list[bool]) -> bool:
    """`text` reads as `requested`: equal, except that a caps character may be stored in either case."""

    return len(text) == len(requested) == len(caps) and all(
        a == b or (flag and a.upper() == b.upper()) for a, b, flag in zip(text, requested, caps))


def _verify_text_edit(doc: Document, editor: TextEditor, new_xml: str, new_paragraph: str, edits: list,
                      expected_text: str, added: list[dict], author: str, tracked: bool, op: str, *,
                      requested: str) -> None:
    node = editor.node
    xml = doc.xml
    if new_xml[:node.start] != xml[:node.start] or new_xml[node.start + len(new_paragraph):] != xml[node.end:]:
        raise VerificationFailed(f"{op}: bytes outside paragraph {editor.index} changed")
    old_tree, new_tree = parse(xml[node.start:node.end]), parse(new_paragraph)
    shift_old, shift_new = xml[node.start:node.end], new_paragraph
    new_ids = {item["id"] for item in added}
    # Plan 831 S2: the caller's own insertions edited in place. They are left out of the byte-identical check
    # (their content changed by design) but keep their start tag byte for byte: same id, author and date.
    edited = {item.attr(xml, "w:id"): item.head(xml) for item in editor.edited.values()} if tracked else {}
    if any(item.attr(xml, "w:author") != author for item in editor.edited.values()):
        raise VerificationFailed(f"{op}: an insertion by another author was edited in place")

    def in_place(tree_xml: str, item: Node) -> bool:
        return item.name == "w:ins" and item.attr(tree_xml, "w:id") in edited and \
            item.attr(tree_xml, "w:author") == author

    old_revs = [_revision_raw(shift_old, item) for item in old_tree.iter() if item.local in REV_TAGS
                and item.name.startswith("w:") and not in_place(shift_old, item)]
    kept, fresh, seen = [], [], Counter()
    for item in new_tree.iter():
        if item.local in REV_TAGS and item.name.startswith("w:"):
            if item.attr(shift_new, "w:id") in new_ids and item.attr(shift_new, "w:author") == author and \
                    item.local in ("ins", "del") and tracked:
                fresh.append(item)
            elif in_place(shift_new, item):
                rid = item.attr(shift_new, "w:id")
                seen[rid] += 1
                if item.head(shift_new) != edited[rid]:
                    raise VerificationFailed(
                        f"{op}: the insertion w:ins id={rid} edited in place lost its id, author or date")
            else:
                kept.append(_revision_raw(shift_new, item))
    if kept != old_revs:
        raise VerificationFailed(f"{op}: a pre-existing revision in paragraph {editor.index} changed (D4)")
    if any(count > 1 for count in seen.values()):
        raise VerificationFailed(f"{op}: an insertion edited in place was split in two")
    if tracked:
        if len(fresh) != len(added):
            raise VerificationFailed(f"{op}: new revision count mismatch")
        if any(in_place(shift_new, outer) for item in fresh for outer in item.ancestors()):
            raise VerificationFailed(f"{op}: a new revision was written inside an insertion edited in place")
        _check_stamps(shift_new, fresh, author, op)
    from .sections import canonical

    old_p, new_p = _fragment(xml, xml[node.start:node.end]), _fragment(xml, new_paragraph)
    if editor.reorders:  # collapsed bookmarks moved before text typed at a point comment, as Word moves them
        old_p = _fragment(xml, editor.reordered_input())
    is_new = lambda el: tracked and el.get(W + "id") in new_ids and el.get(W + "author") == author  # noqa: E731
    is_edited = lambda el: el.tag == W + "ins" and el.get(W + "id") in edited and \
        el.get(W + "author") == author  # noqa: E731
    rejected = copy.deepcopy(new_p)
    et_resolve(rejected, lambda el: is_new(el) or is_edited(el), "reject")
    baseline = old_p
    if edited:
        # Reject All of the result equals Reject All of the input for the own insertions edited in place.
        baseline = copy.deepcopy(old_p)
        et_resolve(baseline, is_edited, "reject")
    if tracked and canonical(rejected) != canonical(baseline):
        raise VerificationFailed(f"{op}: rejecting the new revisions does not give back paragraph {editor.index} (D5)")
    accepted = copy.deepcopy(new_p)
    et_resolve(accepted, is_new, "accept")
    got = _et_final_chars(accepted)
    old_chars = _et_final_chars(old_p)
    if "".join(ch for ch, _ in old_chars) != editor.model.text:
        raise VerificationFailed(f"{op}: paragraph {editor.index} text model disagrees with its XML")
    want: list[tuple[str, str | None]] = []
    cursor = 0
    for edit in edits:
        want.extend(old_chars[cursor:edit.start])
        source_key = old_chars[edit.source][1] if edit.source is not None else None
        want.extend((ch, source_key) for ch in edit.text)
        cursor = edit.end
    want.extend(old_chars[cursor:])
    got_text = "".join(ch for ch, _ in got)
    # D5 twice: the written XML equals the plan, and reads as the caller's own request (F3: not just the plan).
    if got_text != expected_text or not _same_display(got_text, requested, _planned_caps(editor.caps, edits)):
        raise VerificationFailed(
            f"{op}: accepting the new revisions does not give the requested text (D5)",
            expected=requested, actual=got_text)
    for (ch, key), (_, wanted) in zip(got, want):
        if wanted is not None and key != wanted:
            raise VerificationFailed(f"{op}: character {ch!r} did not keep the formatting of its source (D8)")


def _verify_anchors(doc: Document, editor: TextEditor, after_layout, edits: list, op: str) -> None:
    from .sections import StreamEdit, anchor_errors_for_edits, _paragraph_starts

    before = doc.layout
    layout_chars = [char for char in editor.model.chars if char.layout_visible]
    position = {id(char): number for number, char in enumerate(layout_chars)}
    final = editor.model.final
    base = _paragraph_starts([len(item.text) for item in before.paragraphs])[editor.index - 1]
    stream: list[StreamEdit] = []
    delta = 0
    for edit in edits:
        removed = final[edit.start:edit.end]
        if removed:
            at = position[id(removed[0])]
            stream.append(StreamEdit(base + at + delta, len(removed), len(edit.text)))
            delta += len(edit.text) - len(removed)
            continue
        if edit.mode == "before":
            at = position[id(final[edit.start])]
        elif edit.mode == "after":
            at = position[id(final[edit.start - 1])] + 1
        else:
            at = 0
        stream.append(StreamEdit(base + at + delta, 0, len(edit.text)))
        delta += len(edit.text)
    lengths = [len(item.text) for item in before.paragraphs]
    lengths[editor.index - 1] += delta
    problems = anchor_errors_for_edits(before, after_layout, stream, lengths)
    if problems:
        raise VerificationFailed(f"{op}: bookmark/comment anchors moved: " + "; ".join(problems[:5]))


# -- selection --------------------------------------------------------------

def select(doc: Document, *, op: str, ids=None, author: str | None = None, paragraph_index: int | None = None,
           through_paragraph_index: int | None = None, heading: str | None = None,
           until_heading: str | None = None) -> tuple[list[Revision], tuple[int, int] | None]:
    from .sections import locate_section

    span = None
    if heading:
        section = locate_section(doc.layout, heading, until_heading=until_heading)
        members = [section.heading.index] + [item.index for item in section.paragraphs]
        for row in section.rows:
            members.extend(index for cell in row.cells for index in cell)
            members.append(row.end_index)
        span = (min(members), max(members))
    if paragraph_index is not None or through_paragraph_index is not None:
        first = paragraph_index if paragraph_index is not None else through_paragraph_index
        last = through_paragraph_index if through_paragraph_index is not None else first
        if last < first:
            raise RefusedUnsafe(f"{op}: through_paragraph_index {last} is before paragraph_index {first}")
        for index in (first, last):
            if not 1 <= index <= len(doc.layout.paragraphs):
                raise RefusedUnsafe(f"{op}: paragraph {index} is outside 1..{len(doc.layout.paragraphs)}")
        span = (first, last) if span is None else (max(span[0], first), min(span[1], last))
        if span[0] > span[1]:
            raise RefusedUnsafe(f"{op}: the paragraph range and the section do not overlap")
    wanted = {str(item) for item in ids} if ids else None
    chosen = []
    for revision in doc.revisions:
        if wanted is not None and revision.id not in wanted:
            continue
        if author is not None and revision.author != author:
            continue
        if span is not None and (revision.paragraph_index is None or not span[0] <= revision.paragraph_index <= span[1]
                                 or revision.type in TABLE_PROP_REVS):
            continue
        chosen.append(revision)
    if wanted is not None:
        missing = sorted(wanted - {revision.id for revision in chosen}, key=lambda value: int(value) if value.lstrip("-").isdigit() else 0)
        if missing:
            raise RefusedUnsafe(f"{op}: no revision with id {', '.join(missing)} matches the other selectors")
    return chosen, span


def list_revisions(source, *, ids=None, author: str | None = None, paragraph_index: int | None = None,
                   through_paragraph_index: int | None = None, heading: str | None = None,
                   until_heading: str | None = None, limit: int = 200) -> dict:
    _, doc = load(source)
    chosen, span = select(doc, op="list_revisions", ids=ids, author=author, paragraph_index=paragraph_index,
                          through_paragraph_index=through_paragraph_index, heading=heading,
                          until_heading=until_heading)
    listed = [revision.as_dict() for revision in chosen[:limit]]
    indices = sorted({item["paragraph_index"] for item in listed if item["paragraph_index"] is not None})
    if span is not None:
        indices = [index for index in range(span[0], span[1] + 1) if index in doc.para_nodes]
    paragraphs = []
    for index in indices[:limit]:
        node = doc.para_nodes.get(index)
        if node is None:
            continue
        model = ParagraphModel(doc.xml, node)
        paragraphs.append({"paragraph_index": index, "accepted_text": model.text, "original_text": model.original})
    authors: dict[str, int] = {}
    for revision in doc.revisions:
        authors[revision.author or ""] = authors.get(revision.author or "", 0) + 1
    return {
        "operation": "list_revisions",
        "total_in_document": len(doc.revisions),
        "authors": authors,
        "matched": len(chosen),
        "truncated": len(chosen) > limit,
        "scope": list(span) if span else None,
        "revisions": listed,
        "paragraphs": paragraphs,
    }


# -- accept / reject ---------------------------------------------------------

def _contains(xml: str, node: Node, names) -> str | None:
    return next((item.name for item in node.iter() if item is not node and item.name in names), None)


def _balanced_fields(xml: str, node: Node) -> bool:
    """True when every field part inside `node` belongs to a field that begins and ends inside it."""

    depth = 0
    for item in node.iter():
        if item.name == "w:fldChar":
            kind = item.attr(xml, "w:fldCharType")
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth -= 1
                if depth < 0:
                    return False
            elif depth == 0:
                return False
        elif item.name in ("w:instrText", "w:delInstrText") and depth == 0:
            return False
    return depth == 0


# Table-property changes a reject may undo, with the children of the property element that are not
# properties and so stay (structural revision marks); and the properties that define the table grid.
_REJECTABLE_TABLE_PROPS = ("tblPrChange", "trPrChange", "tcPrChange")
_PROP_KEEP = {"w:pPr": ("w:rPr", "w:sectPr"), "w:trPr": ("w:ins", "w:del"),
              "w:tcPr": ("w:cellIns", "w:cellDel", "w:cellMerge")}
_GRID_PROPS = {"w:tcPr": ("w:gridSpan", "w:vMerge", "w:hMerge"), "w:trPr": ("w:gridBefore", "w:gridAfter")}


def _table_structure_change(xml: str, node: Node) -> str | None:
    """The grid property a reject of the table-property change `node` would alter, else None."""

    parent = node.parent
    old = node.child(parent.name)

    def grid(element):
        if element is None:
            return {}
        return {child.name: re.sub(r"\s+", " ", xml[child.start:child.end]) for child in element.children
                if child.name in _GRID_PROPS.get(parent.name, ())}

    now, before = grid(parent), grid(old)
    return next((name for name in _GRID_PROPS.get(parent.name, ()) if now.get(name) != before.get(name)), None)


def _refuse_resolution(doc: Document, revision: Revision, mode: str, op: str) -> None:
    xml, node, name = doc.xml, revision.node, revision.type
    label = revision.describe()
    if name in CELL_REVS:
        raise RefusedUnsafe(f"{op}: {label} is a table-cell revision; resolve it in Word")
    if name in TABLE_PROP_REVS and mode == "reject":
        if name == "tblGridChange":
            # Word does not write the old grid back: it lays the old grid out again against the cells, so
            # its result is a layout (old 3000/5000 came back as 3858/4462 in a fixed 8320-dxa table).
            raise RefusedUnsafe(f"{op}: rejecting {label} (table grid) is not supported: Word recomputes the "
                                "grid from its layout; accept it, or reject it in Word")
        if name not in _REJECTABLE_TABLE_PROPS:
            raise RefusedUnsafe(f"{op}: rejecting {label} (section/numbering formatting) is not supported; accept "
                                "it, or reject it in Word")
        changed = _table_structure_change(xml, node)
        if changed:
            raise RefusedUnsafe(f"{op}: rejecting {label} would change the table's structure ({changed}); "
                                "resolve it in Word")
    if revision.mark in ("paragraph", "row") or name not in RUN_REVS:
        return  # paragraph marks: checked by _plan_merges
    removes = wrapper_removed(name, mode)
    fields = _contains(xml, node, _FIELD_PARTS)
    # A whole field (a Zotero citation) goes or comes back with its wrapper, as in Word; part of one would
    # leave a broken field behind, so that is refused (accepting an insertion only unwraps it: always fine).
    if fields and (removes or name in ("del", "moveFrom")) and not _balanced_fields(xml, node):
        raise RefusedUnsafe(f"{op}: {mode}ing {label} would {'remove' if removes else 'restore'} part of a field "
                            f"({fields}); resolve it in Word")
    if not removes:
        return
    blocker = _contains(xml, node, _DRAWINGS)
    if blocker:
        raise RefusedUnsafe(f"{op}: {mode}ing {label} would delete a {blocker} inside it; resolve it in Word")
    for run in (item for item in node.iter() if item.name == "w:r" and item.child("w:commentReference")):
        rpr = run.child("w:rPr")
        if any(child.name not in ("w:rPr", "w:commentReference") for child in run.children) or (
                rpr is not None and _contains(xml, rpr, {"w:" + tag for tag in REV_TAGS})):
            raise RefusedUnsafe(f"{op}: {mode}ing {label} would delete text that shares a run with a "
                                "w:commentReference; resolve it in Word")


def _nested_in_dropped_moves(doc: Document, selected: list[Revision], mode: str,
                             op: str) -> list[tuple[Revision, str]]:
    """A move side that goes (moveTo rejected, moveFrom accepted) holding a revision that was not selected:
    Word keeps such a nested revision (another reviewer's pending deletion survives the rejected move; plan
    562 brief, final reviews terra61/sol61), and this engine would delete it with the side, so it is refused
    before anything is written. (A rejected insertion holding a deletion still removes both: that text only
    ever existed as an insertion. OOXML-derived, not re-checked in Word in the final fix round.)"""

    chosen = {id(revision.node) for revision in selected}
    by_node = {id(revision.node): revision for revision in doc.revisions}
    out = []
    for revision in selected:
        if revision.mark or revision.type not in ("moveFrom", "moveTo") or not wrapper_removed(revision.type, mode):
            continue
        nested = [by_node[id(item)] for item in revision.node.iter()
                  if item is not revision.node and id(item) in by_node and id(item) not in chosen]
        if not nested:
            continue
        shown = ", ".join(item.describe() for item in nested[:5]) + (
            f" and {len(nested) - 5} more" if len(nested) > 5 else "")
        ids = ", ".join(item.id or "?" for item in nested)
        out.append((revision, f"{op}: {mode}ing {revision.describe()} would delete {len(nested)} revision(s) "
                              f"inside it that were not selected ({shown}); Word keeps them. Select id "
                              f"{ids} too, or resolve it in Word"))
    return out


def _nested_moves(doc: Document, selected: list[Revision], mode: str, op: str) -> list[tuple[Revision, str]]:
    """A selected move revision that sits inside another move: inside another w:moveFrom/w:moveTo wrapper,
    or inside move ranges of two different names. Resolving either pair would resolve one move and leave
    half of the other pending (a later decision then duplicates or loses its text: recheck terra61/sol61),
    and Word's handling of nested moves is not verified, so the selection is refused before anything is
    written. One move written in several pieces inside one range is not nested."""

    kind_of = {"moveFrom": "from", "moveTo": "to"}
    ranges = _move_ranges(doc)
    by_node = {id(revision.node): revision for revision in doc.revisions}
    out = []
    for revision in selected:
        if revision.type not in kind_of:
            continue
        where = None
        parent = revision.node.parent
        while parent is not None:
            if parent.name in ("w:moveFrom", "w:moveTo") and id(parent) in by_node:
                where = by_node[id(parent)].describe()
                break
            parent = parent.parent
        if where is None:
            names = sorted({name for kind, name, start, end in ranges
                            if kind == kind_of[revision.type] and start.start <= _logical(revision) <= end.start})
            if len(names) > 1:
                where = "move ranges " + " and ".join(repr(name) for name in names)
        if where is not None:
            out.append((revision, f"{op}: {mode}ing {revision.describe()} would resolve a move nested inside "
                                  f"{where}; nested moves are not supported here (Word's handling is not "
                                  "verified). Resolve them in Word"))
    return out


_REFUSAL_IDS_SHOWN = 20
_REFUSAL_REASONS_SHOWN = 10


def _raise_refusals(op: str, mode: str, selected: int, refusals: list[tuple[Revision, str]]) -> None:
    """One error for every refused revision: the original message when there is one; otherwise the count,
    then each reason once with its ids (capped, "and N more")."""

    if not refusals:
        return
    if len(refusals) == 1:
        raise RefusedUnsafe(refusals[0][1])
    groups: dict[str, list[str]] = {}
    for revision, message in refusals:
        reason = message.removeprefix(f"{op}: ").replace(revision.describe(), f"w:{revision.type}")
        ids = groups.setdefault(reason, [])
        if (revision.id or "?") not in ids:
            ids.append(revision.id or "?")
    count = len({id(revision.node) for revision, _ in refusals})
    lines = [f"{op}: {count} of the {selected} selected revisions cannot be {mode}ed here; nothing was "
             "changed:"]
    for reason, ids in list(groups.items())[:_REFUSAL_REASONS_SHOWN]:
        shown = ", ".join(ids[:_REFUSAL_IDS_SHOWN])
        more = f" and {len(ids) - _REFUSAL_IDS_SHOWN} more" if len(ids) > _REFUSAL_IDS_SHOWN else ""
        lines.append(f"- ({len(ids)}) {reason}: ids {shown}{more}")
    if len(groups) > _REFUSAL_REASONS_SHOWN:
        rest = list(groups.values())[_REFUSAL_REASONS_SHOWN:]
        lines.append(f"- and {len(rest)} more reasons ({sum(len(ids) for ids in rest)} revisions)")
    lines.append("Deselect these ids to resolve the rest, or resolve them in Word.")
    raise RefusedUnsafe("\n".join(lines))


# -- moves: a pair is resolved together, found by the w:name of its range starts ----------------------------

def _mark_paragraph(revision: Revision) -> Node:
    return revision.node.parent.parent.parent


def _logical(revision: Revision) -> int:
    """Document offset a revision stands for: a paragraph mark sits at the end of its paragraph."""

    if revision.mark == "paragraph":
        return _mark_paragraph(revision).close_start()
    return revision.node.start


def _move_ranges(doc: Document) -> list[tuple[str, str, Node, Node]]:
    """(kind "from"/"to", name, start, end) for every complete move range in the body."""

    starts: dict[tuple[str, str | None], Node] = {}
    ends: dict[tuple[str, str | None], Node] = {}
    for node in doc.body.iter():
        if node.name.startswith("w:") and node.local in MOVE_RANGES:
            kind = "from" if node.local.startswith("moveFrom") else "to"
            key = (kind, node.attr(doc.xml, "w:id"))
            (starts if node.local.endswith("Start") else ends)[key] = node
    found = []
    for key, start in starts.items():
        end = ends.get(key)
        if end is not None:
            found.append((key[0], start.attr(doc.xml, "w:name") or f"#{key[0]}{key[1]}", start, end))
    return found


def _expand_moves(doc: Document, chosen: list[Revision]):
    """Selected revisions plus every member of each selected move pair, the range markers to remove, the
    partners pulled in, and one record per move."""

    ranges = _move_ranges(doc)
    kind_of = {"moveFrom": "from", "moveTo": "to"}

    def inside(revision: Revision, kind: str, start: Node, end: Node) -> bool:
        return kind_of.get(revision.type) == kind and start.start <= _logical(revision) <= end.start

    selected = {id(revision.node): revision for revision in chosen}
    pulled: list[Revision] = []
    drop: list[Node] = []
    moves: list[dict] = []
    seen: set[str] = set()
    queue = [revision for revision in chosen if revision.type in kind_of]
    while queue:
        revision = queue.pop(0)
        home = next((item for item in ranges if inside(revision, item[0], item[2], item[3])), None)
        if home is None:
            moves.append({"name": None, "members": [revision]})
            continue
        if home[1] in seen:
            continue
        seen.add(home[1])
        group = [item for item in ranges if item[1] == home[1]]
        members = [item for item in doc.revisions if any(inside(item, k, s, e) for k, _, s, e in group)]
        for _, _, start, end in group:
            drop.extend((start, end))
        for member in members:
            if id(member.node) not in selected:
                selected[id(member.node)] = member
                pulled.append(member)
        moves.append({"name": home[1], "members": members})
    ordered = sorted(selected.values(), key=lambda item: item.ordinal)
    return ordered, sorted(pulled, key=lambda item: item.ordinal), drop, moves


def affected_paragraphs(doc: Document, chosen: list[Revision], mode: str) -> list[int]:
    """Layout indices a resolution of `chosen` touches: each revision's paragraph, the move partners it
    pulls in, and the paragraph a removed mark joins (for batch groups; selection errors are ignored)."""

    selected = _expand_moves(doc, chosen)[0]
    found = {revision.paragraph_index for revision in selected if revision.paragraph_index is not None}
    index_of = {id(node): index for index, node in doc.para_nodes.items()}
    for revision in selected:
        if revision.mark == "paragraph" and mark_removed(revision.type, mode):
            paragraph = _mark_paragraph(revision)
            siblings = paragraph.parent.children
            at = siblings.index(paragraph) + 1
            while at < len(siblings) and siblings[at].name in _BETWEEN:
                at += 1
            if at < len(siblings) and id(siblings[at]) in index_of:
                found.add(index_of[id(siblings[at])])
    return sorted(found)


# -- row marks: a removed row mark removes the whole row (plan 864 S3b) ---------------------------------------

def _row_label(doc: Document, row: Node, index_of: dict[int, int]) -> str:
    table = row.parent
    tables = [node for node in doc.body.children if node.name == "w:tbl"]
    number = next((n for n, node in enumerate(tables, 1) if node is table), None)
    rows = [node for node in table.children if node.name == "w:tr"]
    indices = [index_of[id(node)] for node in row.iter() if node.name == "w:p" and id(node) in index_of]
    where = f" (paragraphs {indices[0]}-{indices[-1] + 1})" if indices else ""
    return f"table {number} row {rows.index(row) + 1}{where}"


def _plan_rows(doc: Document, selected: list[Revision], mode: str, op: str, index_of: dict[int, int],
               refusals: list[tuple[Revision, str]]):
    """Rows whose row mark goes (rejected w:ins, accepted w:del in w:trPr): Word removes the whole w:tr.
    Allowed only when everything in the row goes with the same selection: every run inside a selected
    wrapper that is removed, every paragraph mark selected and removed, no other revision, no anchor or
    object left behind, and the table keeps a row. Rejecting the insertion of every row of a top-level table
    removes the whole w:tbl instead, as Word does, when the table holds nothing but its properties, grid and
    rows and no revision outside the rows (plan 916 S5); a comment whose range and reference markers all lie
    inside that table goes with it, as Word deletes it. Returns the rows (or whole tables) to remove
    (document order), their removed layout indices (paragraphs plus the row-end marker), the ids of every
    revision node inside a row with a selected removing row mark (they are planned with the row, not as
    merges), and the legacy ids of the comments removed with whole tables."""

    chosen = {id(revision.node): revision for revision in selected}
    by_node = {id(revision.node): revision for revision in doc.revisions}
    rows: list[tuple[Node, Revision]] = []
    for revision in selected:
        if revision.mark == "row" and mark_removed(revision.type, mode):
            rows.append((revision.node.parent.parent, revision))
    inside: set[int] = set()
    planned: list[Node] = []
    removed_indices: list[int] = []
    for row, lead in rows:
        inside.update(id(item) for item in row.iter() if id(item) in by_node)
    going: dict[int, int] = {}
    whole: dict[int, Node] = {}  # tables every row of which goes (rejected insertions): the w:tbl goes
    free = _whole_table_comments(doc, rows, mode, by_node)  # table id -> comments wholly inside it
    for row, lead in rows:
        label = _row_label(doc, row, index_of) if row.name == "w:tr" else "a row"
        loose = free.get(id(row.parent), set()) if row.parent is not None else set()

        def goes_with_table(node: Node) -> bool:
            if node.name in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
                return node.attr(doc.xml, "w:id") in loose
            return node.name == "w:r" and _comment_reference_only(node) and all(
                item.attr(doc.xml, "w:id") in loose for item in node.children if item.name == "w:commentReference")
        what = f"{op}: {mode}ing {lead.describe()} would remove {label}"

        def refuse(message: str) -> None:
            refusals.append((lead, f"{what}, but {message}"))

        if row.name != "w:tr" or row.parent is None or row.parent.name != "w:tbl" \
                or row.parent.parent is None or row.parent.parent.name != "w:body":
            refuse("it is not a row of a top-level table; resolve it in Word")
            continue
        other = next((by_node[id(item)] for item in row.iter()
                      if id(item) in by_node and id(item) not in chosen), None)
        if other is not None:
            refuse(f"the row also holds {other.describe()}, which is not part of the selection; select it "
                   "too, or resolve it in Word")
            continue
        problem = None
        for cell in row.children:
            if cell.name in ("w:trPr", "w:tblPrEx"):
                continue
            if cell.name != "w:tc":
                problem = f"it holds {cell.name}; resolve it in Word"
                break
            for item in cell.children:
                if item.name == "w:tcPr":
                    continue
                if item.name != "w:p":
                    kind = {"w:tbl": "a nested table"}.get(item.name, item.name)
                    problem = f"a cell holds {kind}; resolve it in Word"
                    break
                ppr = item.child("w:pPr")
                rpr = ppr.child("w:rPr") if ppr is not None else None
                marks = [chosen.get(id(child)) for child in (rpr.children if rpr is not None else ())]
                if not any(mark is not None and mark.mark == "paragraph" and mark_removed(mark.type, mode)
                           for mark in marks):
                    problem = (f"paragraph {index_of.get(id(item))} keeps its paragraph mark (it is not part of "
                               "the selection); select it too, or resolve it in Word")
                    break
                if ppr is not None and ppr.child("w:sectPr") is not None:
                    problem = f"paragraph {index_of.get(id(item))} ends a section; resolve it in Word"
                    break
                for child in item.children:
                    if child.name in ("w:pPr", "w:proofErr") or goes_with_table(child):
                        continue
                    if child.name in _HOIST:
                        problem = (f"paragraph {index_of.get(id(item))} holds {child.name}, which Word would "
                                   "delete with the row; remove it first, or resolve it in Word")
                        break
                    revision = chosen.get(id(child))
                    if revision is None or revision.mark or revision.type not in RUN_REVS \
                            or not wrapper_removed(revision.type, mode):
                        text = "".join(doc.xml[t.head_end:t.close_start()] for t in child.iter()
                                       if t.name in ("w:t", "w:delText") and not t.closed)
                        thing = f"text {text[:40]!r}" if text else child.name
                        problem = (f"paragraph {index_of.get(id(item))} holds {thing}, which is not part of the "
                                   "selection; resolve it in Word")
                        break
                    anchor = next((node for node in child.iter() if (node.name in _HOIST or (
                        node.name == "w:commentReference")) and not goes_with_table(node)), None)
                    if anchor is not None:
                        problem = (f"paragraph {index_of.get(id(item))} holds {anchor.name}, which Word would "
                                   "delete with the row; remove it first, or resolve it in Word")
                        break
                if problem:
                    break
            if problem:
                break
        if problem:
            refuse(problem)
            continue
        table = row.parent
        going[id(table)] = going.get(id(table), 0) + 1
        if going[id(table)] == sum(1 for node in table.children if node.name == "w:tr"):
            blocker = _whole_table_problem(table, mode, by_node)
            if blocker:
                refuse(blocker)
                continue
            whole[id(table)] = table
        planned.append(row)
        indices = [index_of[id(node)] for node in row.iter() if node.name == "w:p" and id(node) in index_of]
        removed_indices.extend(indices + [indices[-1] + 1])
    planned = [node for node in planned if id(node.parent) not in whole] + list(whole.values())
    planned.sort(key=lambda node: node.start)
    comments = sorted({cid for key in whole for cid in free.get(key, ())}, key=lambda cid: (len(cid), cid))
    return planned, sorted(removed_indices), inside, comments


def _drop_comment_entries(source, legacy_ids: list[str], op: str) -> tuple[dict[str, bytes], list[dict]]:
    """The comment parts without the entries of ``legacy_ids`` (comments.xml, commentsExtended.xml,
    commentsIds.xml, commentsExtensible.xml, as delete_comment removes them) and the removed comments
    ({comment_id, legacy_id, author}). A thread is removed whole: a comment whose replies (or whose thread's
    first comment) stay is refused, named."""

    if not legacy_ids:
        return {}, []
    from .comment_ops import _Comments, _Package

    package = _Package(source, op)
    comments = _Comments(package)
    wanted = set(legacy_ids)
    victims = [record for record in comments.records if record.legacy_id in wanted]
    missing = wanted - {record.legacy_id for record in victims}
    if missing:
        raise RefusedUnsafe(f"{op}: comment markers {sorted(missing)} inside the table have no entry in "
                            "word/comments.xml; resolve it in Word")
    for record in victims:
        root = comments.root_of(record) or record
        thread = [root] + comments.replies(root)
        kept = [item for item in thread if item.legacy_id not in wanted]
        if kept:
            raise RefusedUnsafe(f"{op}: comment {record.comment_id} lies inside the table but its thread has "
                                f"comments anchored outside it ({', '.join(item.comment_id for item in kept)}); "
                                "delete the comment first, or resolve it in Word")
    part = package.part("comments")
    for record in victims:
        part.splice(record.node.start, record.node.end, "")
        for kind, node in (("extended", record.extended), ("ids", record.ids), ("extensible", record.extensible)):
            if node is not None:
                package.part(kind).splice(node.start, node.end, "")
    removed = [{"comment_id": record.comment_id, "legacy_id": record.legacy_id,
                "author": record.node.attr(part.xml, "w:author")} for record in victims]
    return package.render(), removed


def _write_without_comments(source, output, new_xml: str, comment_parts: dict[str, bytes], legacy_ids: list[str],
                            removed: list[dict], op: str) -> None:
    """Write the resolved document with the removed comments' entries gone from every comment part, then
    check that no marker of theirs is left in any story and every other thread is unchanged."""

    from .comment_ops import _check_threads, _story_markers

    output = Path(output)
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    changed = dict(comment_parts)
    changed["word/document.xml"] = new_xml.encode("utf-8")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(output, "x") as target:
        for info in source_zip.infolist():
            data = changed.get(info.filename)
            target.writestr(info, source_zip.read(info.filename) if data is None else data,
                            compress_type=info.compress_type)
    try:
        with zipfile.ZipFile(output) as archive:
            written = {name: archive.read(name) for name in archive.namelist()}
        leftover = sorted(_story_markers(written, set(legacy_ids)))
        if leftover:
            raise VerificationFailed(f"{op}: markers of the removed comments remain in {leftover}")
    except Exception:
        output.unlink(missing_ok=True)
        raise
    _check_threads(Path(source), output, set(), op, removed=set(legacy_ids))


def _whole_table_comments(doc: Document, rows: list, mode: str, by_node: dict) -> dict[int, set[str]]:
    """For each top-level table every row of which is a selected rejected insertion (and that may go whole):
    the legacy ids of the comments whose range and reference markers all lie inside it. Word deletes such a
    comment with the table; a comment reaching outside the table is not listed and keeps refusing."""

    if mode != "reject":
        return {}
    count: dict[int, int] = {}
    tables: dict[int, Node] = {}
    for row, _ in rows:
        table = row.parent
        if row.name == "w:tr" and table is not None and table.name == "w:tbl" and table.parent is not None \
                and table.parent.name == "w:body":
            count[id(table)] = count.get(id(table), 0) + 1
            tables[id(table)] = table
    marks = ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
    everywhere = Counter(node.attr(doc.xml, "w:id") for node in doc.root.iter() if node.name in marks)
    out: dict[int, set[str]] = {}
    for key, table in tables.items():
        if count[key] != sum(1 for node in table.children if node.name == "w:tr") \
                or _whole_table_problem(table, mode, by_node):
            continue
        inside = Counter(node.attr(doc.xml, "w:id") for node in table.iter() if node.name in marks)
        out[key] = {cid for cid, number in inside.items() if cid is not None and number == everywhere[cid]}
    return out


def _whole_table_problem(table: Node, mode: str, by_node: dict) -> str | None:
    """Why a table whose every row goes cannot go as a whole (None when it can: rejected row insertions of a
    table holding only its properties, grid and rows, with no revision in its properties or grid)."""

    if mode != "reject":
        return "that would leave the table with no rows; resolve it in Word"
    for child in table.children:
        if child.name == "w:tr":
            continue
        if child.name not in ("w:tblPr", "w:tblGrid"):
            return f"the table also holds {child.name} outside its rows; resolve it in Word"
        revision = next((by_node[id(item)] for item in child.iter() if id(item) in by_node), None)
        if revision is not None:
            return f"the table's properties hold {revision.describe()}; resolve it in Word"
    return None


# -- paragraph marks: a removed mark merges its paragraph into the next one ---------------------------------

def _plan_merges(doc: Document, selected: list[Revision], mode: str, op: str, index_of: dict[int, int],
                 refusals: list[tuple[Revision, str]] | None = None, gone_tables: frozenset[int] = frozenset()):
    """Paragraphs whose mark goes (each with the paragraph it joins and the markers between), and the mark
    revisions that only lose their marker. Refuses merges Word's result cannot be matched for: raises at the
    first, or, given ``refusals``, appends (revision, message) for each and plans the rest. A table removed as
    a whole (``gone_tables``, ids) is passed over: the paragraph joins the one after it."""

    def refuse(revision: Revision, message: str) -> None:
        if refusals is None:
            raise RefusedUnsafe(message)
        refusals.append((revision, message))

    by_node = {id(revision.node): revision for revision in doc.revisions}
    chosen_ids = {id(revision.node) for revision in selected}
    paragraphs: dict[int, tuple[Node, list[Revision]]] = {}
    for revision in selected:
        if revision.mark == "paragraph":
            paragraph = _mark_paragraph(revision)
            paragraphs.setdefault(id(paragraph), (paragraph, []))[1].append(revision)
    merges: list[tuple[Node, Node, list[Node]]] = []
    marker_only: list[Revision] = []
    for paragraph, revs in sorted(paragraphs.values(), key=lambda item: item[0].start):
        if not any(mark_removed(revision.type, mode) for revision in revs):
            marker_only.extend(revs)
            continue
        lead = next(revision for revision in revs if mark_removed(revision.type, mode))
        label = lead.describe()
        index = index_of.get(id(paragraph))
        what = f"{mode}ing {label} would merge paragraph {index}"
        ppr = paragraph.child("w:pPr")
        if ppr is not None and ppr.child("w:sectPr") is not None:
            refuse(lead, f"{op}: {what} with the next one, but paragraph {index} ends a section; "
                         "resolve it in Word")
            continue
        siblings = paragraph.parent.children
        at = siblings.index(paragraph) + 1
        between = []
        while at < len(siblings) and (siblings[at].name in _BETWEEN or id(siblings[at]) in gone_tables):
            if id(siblings[at]) not in gone_tables:
                between.append(siblings[at])
            at += 1
        following = siblings[at] if at < len(siblings) else None
        if following is None or following.name == "w:sectPr":
            where = {"w:tc": "its table cell", "w:body": "the document"}.get(
                paragraph.parent.name, f"its {paragraph.parent.name}")
            refuse(lead, f"{op}: {what} with the next one, but it is the last paragraph of {where}; "
                         "resolve it in Word")
            continue
        if following.name != "w:p":
            kind = {"w:tbl": "a table", "w:sdt": "a content control"}.get(following.name, following.name)
            refuse(lead, f"{op}: {what} with {kind}; resolve it in Word")
            continue
        dropped = next((by_node[id(item)] for item in (ppr.iter() if ppr is not None else ())
                        if item is not ppr and id(item) in by_node and id(item) not in chosen_ids), None)
        if dropped is not None:
            refuse(lead, f"{op}: {what} with the next one, which would drop {dropped.describe()} from its "
                         "properties; select it too, or resolve it in Word")
            continue
        merges.append((paragraph, following, between))
    return merges, marker_only


def _chains(merges, index_of: dict[int, int]) -> list[list[int]]:
    chains: list[list[int]] = []
    previous = None
    for paragraph, following, _ in merges:
        if previous is not None and previous is paragraph:
            chains[-1][-1:] = [index_of[id(paragraph)], index_of[id(following)]]
        else:
            chains.append([index_of[id(paragraph)], index_of[id(following)]])
        previous = following
    return chains


# -- rendering ----------------------------------------------------------------------------------------------

def _hoisted(xml: str, node: Node, actions: dict[int, str]) -> str:
    """The bookmark/comment/permission/move-range markers (and comment-reference runs) inside a removed
    wrapper, in order, so the anchors stay where the removed text was."""

    out = []
    for child in node.children:
        if child.name in _HOIST:
            if actions.get(id(child)) != "remove":
                out.append(xml[child.start:child.end])
        elif child.name == "w:r":
            if child.child("w:commentReference") is not None:
                out.append(xml[child.start:child.end])
        else:
            out.append(_hoisted(xml, child, actions))
    return "".join(out)


def _paragraph_content(xml: str, node: Node, actions, dirty) -> str:
    """A merged-away paragraph's content without its pPr (its mark and properties go)."""

    pieces, cursor = [], node.head_end
    for child in node.children:
        if child.name != "w:pPr":
            pieces.append(xml[cursor:child.start])
            pieces.append(_render(xml, child, actions, dirty))
        cursor = child.end
    if not node.closed:
        pieces.append(xml[cursor:node.close_start()])
    return "".join(pieces)


def _render_joined(xml: str, node: Node, actions, dirty, prefix: str) -> str:
    """The paragraph that survives a merge: its own head and pPr, then the joined content, then its own."""

    if node.closed:
        return node.head(xml)[:-2].rstrip() + ">" + prefix + "</w:p>"
    pieces, cursor = [node.head(xml)], node.head_end
    placed = not node.children or node.children[0].name != "w:pPr"
    if placed:
        pieces.append(prefix)
    for child in node.children:
        pieces.append(xml[cursor:child.start])
        pieces.append(_render(xml, child, actions, dirty))
        cursor = child.end
        if not placed:
            pieces.append(prefix)
            placed = True
    pieces.append(xml[cursor:node.end])
    return "".join(pieces)


def _render_children(xml: str, node: Node, actions, dirty) -> str:
    pieces, cursor, carry = [], node.head_end, None
    for child in node.children:
        gap = xml[cursor:child.start]
        cursor = child.end
        if carry is not None:
            carry.append(gap)
            if actions.get(id(child)) == "merge":
                carry.append(_paragraph_content(xml, child, actions, dirty))
            elif child.name == "w:p":
                pieces.append(_render_joined(xml, child, actions, dirty, "".join(carry)))
                carry = None
            else:
                carry.append(_render(xml, child, actions, dirty))
            continue
        pieces.append(gap)
        if actions.get(id(child)) == "merge":
            carry = [_paragraph_content(xml, child, actions, dirty)]
        else:
            pieces.append(_render(xml, child, actions, dirty))
    if carry is not None:
        raise VerificationFailed("a merged paragraph has no following paragraph")
    pieces.append(xml[cursor:node.close_start()])
    return "".join(pieces)


def _render(xml: str, node: Node, actions: dict[int, str], dirty: set[int]) -> str:
    """The node's text with `actions` applied to the selected descendants; untouched spans are copied."""

    action = actions.get(id(node))
    if id(node) not in dirty and action is None:
        return xml[node.start:node.end]
    if action == "remove":
        return ""
    if action == "drop":
        return _hoisted(xml, node, actions)
    if action == "retag":
        new = "w:t" if node.name == "w:delText" else "w:instrText"
        head = "<" + new + node.head(xml)[len(node.name) + 1:]
        return head if node.closed else head + node.inner(xml) + f"</{new}>"
    if node.closed:
        return xml[node.start:node.end]
    reject_prop = next((child for child in node.children if actions.get(id(child)) == "reject_prop"), None)
    if reject_prop is not None:
        old = reject_prop.child(node.name)
        old_inner = old.inner(xml) if old is not None and not old.closed else ""
        if node.name == "w:rPr":
            keep = "".join(_render(xml, child, actions, dirty) for child in node.children
                           if child.local in RUN_REVS)
            inner = keep + old_inner
        else:  # pPr, tblPr, trPr, tcPr: the old properties, then the kept structural children
            keep = "".join(_render(xml, child, actions, dirty) for child in node.children
                           if child.name in _PROP_KEEP.get(node.name, ()))
            inner = old_inner + keep
    else:
        inner = _render_children(xml, node, actions, dirty)
    if action == "unwrap":
        return inner
    if node.name in ("w:rPr", "w:pPr", "w:trPr") and node.children and not inner.strip():
        return ""  # every child was a resolved marker; an empty rPr/pPr trips Word's save check (Word drops
        # an emptied trPr too)
    return node.head(xml) + inner + xml[node.close_start():node.end]


def _retag_targets(node: Node) -> list[Node]:
    """delText / delInstrText inside a restored deletion, outside nested deletions."""

    out = []
    for child in node.children:
        if child.name in _DELETED:
            continue
        if child.name in ("w:delText", "w:delInstrText"):
            out.append(child)
        else:
            out.extend(_retag_targets(child))
    return out


def _paired(before: list[Node], after: list[Node], xml: str, new_xml: str) -> list[tuple[Node, Node]]:
    """Match kept revisions to the output by (element, w:id), in order: a merge moves the following
    paragraph's pPr revisions ahead of the joined content, so plain document order may differ."""

    def key(node: Node, text: str):
        return node.local, node.attr(text, "w:id")

    pending: dict = {}
    for node in after:
        pending.setdefault(key(node, new_xml), []).append(node)
    pairs = []
    for node in before:
        bucket = pending.get(key(node, xml))
        if not bucket:
            raise VerificationFailed(f"revision {node.local} id={node.attr(xml, 'w:id')} is missing from the result")
        pairs.append((node, bucket.pop(0)))
    return pairs


def _report_row(doc: Document, row: Node, index_of: dict[int, int], out: list[dict]) -> None:
    indices = [index_of[id(node)] for node in row.iter() if node.name == "w:p" and id(node) in index_of]
    out.append({"row": _row_label(doc, row, index_of), "paragraphs": indices,
                "row_end_index": indices[-1] + 1,
                "text": ["".join(doc.xml[t.head_end:t.close_start()] for t in node.iter()
                                 if t.name in ("w:t", "w:delText") and not t.closed)
                         for node in row.iter() if node.name == "w:p"]})


def resolve_revisions(source, output, *, mode: str, ids=None, author: str | None = None,
                      paragraph_index: int | None = None, through_paragraph_index: int | None = None,
                      heading: str | None = None, until_heading: str | None = None) -> dict:
    """Accept or reject the selected existing revisions; every other byte is kept.

    A move is resolved as a pair (the partner and both range markers come with either side); a removed
    paragraph mark merges its paragraph into the next one, which keeps its own pPr, as Word does."""

    from .sections import canonical

    op = f"{mode}_revisions"
    if not (ids or author or paragraph_index or through_paragraph_index or heading):
        raise RefusedUnsafe(f"{op}: give ids, author, a paragraph range or a section heading")
    parts, doc = load(source)
    chosen, span = select(doc, op=op, ids=ids, author=author, paragraph_index=paragraph_index,
                          through_paragraph_index=through_paragraph_index, heading=heading,
                          until_heading=until_heading)
    if not chosen:
        raise RefusedUnsafe(f"{op}: no revision matches the selection; nothing to {mode}")
    selected, pulled, range_nodes, moves = _expand_moves(doc, chosen)
    # every refusal is collected first, so one error names all of them (final review: no iterative discovery)
    refusals: list[tuple[Revision, str]] = []
    for revision in selected:
        try:
            _refuse_resolution(doc, revision, mode, op)
        except RefusedUnsafe as error:
            refusals.append((revision, str(error)))
    refusals.extend(_nested_in_dropped_moves(doc, selected, mode, op))
    refusals.extend(_nested_moves(doc, selected, mode, op))
    index_of = {id(node): index for index, node in doc.para_nodes.items()}
    rows, row_indices, in_rows, comments_gone = _plan_rows(doc, selected, mode, op, index_of, refusals)
    gone_tables = frozenset(id(node) for node in rows if node.name == "w:tbl")
    merges, _ = _plan_merges(doc, [revision for revision in selected if id(revision.node) not in in_rows],
                             mode, op, index_of, refusals, gone_tables)
    _raise_refusals(op, mode, len(selected), refusals)
    comment_parts, comments_removed = _drop_comment_entries(source, comments_gone, op)

    actions: dict[int, str] = {}
    acted: list[Node] = []

    def act(node: Node, action: str) -> None:
        actions[id(node)] = action
        acted.append(node)

    removed: list[Node] = []  # subtrees that disappear (their revisions go with them)
    for revision in selected:
        name, node = revision.type, revision.node
        if name in PROP_REVS:
            act(node, "remove" if mode == "accept" else "reject_prop")
        elif revision.mark:
            act(node, "remove")
        elif wrapper_removed(name, mode):
            act(node, "drop")
            removed.append(node)
        else:
            act(node, "unwrap")
            if name in ("del", "moveFrom"):
                for item in _retag_targets(node):
                    act(item, "retag")
    for node in range_nodes:
        act(node, "remove")
    for paragraph, _, _ in merges:
        act(paragraph, "merge")
        ppr = paragraph.child("w:pPr")
        if ppr is not None:
            removed.append(ppr)
    for row in rows:
        act(row, "remove")
        removed.append(row)
    dirty: set[int] = set()
    for node in acted:
        dirty.update(id(ancestor) for ancestor in node.ancestors())
    new_body = _render(doc.xml, doc.body, actions, dirty)
    new_xml = doc.xml[:doc.body.start] + new_body + doc.xml[doc.body.end:]

    # D4: every revision not selected is byte-identical (those containing a selected one keep their start
    # tag), except those inside removed text, which go with it and are reported.
    chosen_ids = {id(revision.node) for revision in selected}
    gone = {id(item) for node in removed for item in node.iter() if item is not node}
    removed_with = [revision for revision in doc.revisions
                    if id(revision.node) in gone and id(revision.node) not in chosen_ids]
    containing = {id(ancestor) for revision in selected for ancestor in revision.node.ancestors()}
    after_root = parse(new_xml)
    after_revs = [item for item in after_root.iter() if item.local in REV_TAGS and item.name.startswith("w:")
                  and any(a.name == "w:body" for a in item.ancestors())]
    before_keep = [revision.node for revision in doc.revisions
                   if id(revision.node) not in chosen_ids and id(revision.node) not in gone]
    if len(after_revs) != len(before_keep):
        raise VerificationFailed(f"{op}: {len(before_keep)} revisions should remain, found {len(after_revs)}")
    for old, new in _paired(before_keep, after_revs, doc.xml, new_xml):
        if id(old) in containing:
            same = old.head(doc.xml) == new.head(new_xml)
        else:
            same = doc.xml[old.start:old.end] == new_xml[new.start:new.end]
        if not same:
            raise VerificationFailed(f"{op}: revision {old.local} id={old.attr(doc.xml, 'w:id')} changed although "
                                     "it was not selected (D4)")
    # D5: equal to an independent ElementTree resolution of the same selection.
    selected_ordinals = {revision.ordinal for revision in selected}
    range_ordinals = {number for number, node in enumerate(
        item for item in doc.body.iter() if item.name.startswith("w:") and item.local in MOVE_RANGES)
        if id(node) in {id(item) for item in range_nodes}}
    old_root = ET.fromstring(doc.xml.encode("utf-8"))
    old_body = old_root.find(W + "body")
    ordinal_of: dict[int, int] = {}
    for number, el in enumerate(item for item in old_body.iter()
                                if item.tag.startswith(W) and _local(item.tag) in REV_TAGS):
        ordinal_of[id(el)] = number
    range_of: dict[int, int] = {}
    for number, el in enumerate(item for item in old_body.iter()
                                if item.tag.startswith(W) and _local(item.tag) in MOVE_RANGES):
        range_of[id(el)] = number
    et_resolve(old_body, lambda el: ordinal_of.get(id(el)) in selected_ordinals, mode,
               drop=lambda el: range_of.get(id(el)) in range_ordinals)
    new_root = ET.fromstring(new_xml.encode("utf-8"))
    if canonical(old_body) != canonical(new_root.find(W + "body")):
        raise VerificationFailed(f"{op}: the result differs from an independent resolution of the same "
                                 "revisions; nothing was written")
    anchor_pattern = r"<w:(?:%s)\b[^>]*>" % "|".join(name[2:] for name in _ANCHOR_MARKS)
    gone_marks = {f"<w:{name}" for name in ("commentRangeStart", "commentRangeEnd", "commentReference")}

    def removed_comment(tag: str) -> bool:  # a marker of a comment removed with its table
        found = re.search(r"\bw:id=([\"'])(.*?)\1", tag)
        return tag.split()[0].rstrip("/>") in gone_marks and found is not None and found.group(2) in comments_gone

    marks_before = [m.group(0) for m in re.finditer(anchor_pattern, doc.xml) if not removed_comment(m.group(0))]
    marks_after = [m.group(0) for m in re.finditer(anchor_pattern, new_xml)]
    if marks_before != marks_after:
        raise VerificationFailed(f"{op}: a bookmark or comment anchor was lost; nothing was written")
    if comments_gone:
        _write_without_comments(source, output, new_xml, comment_parts, comments_gone, comments_removed, op)
    else:
        rezip(source, output, new_xml)

    chains = _chains(merges, index_of)
    merged_away = sorted(index_of[id(paragraph)] for paragraph, _, _ in merges)
    gone_indices = sorted(merged_away + row_indices)
    merged = []
    if chains:
        _, after_doc = load(output)
        for chain in chains:
            new_index = chain[0] - sum(1 for index in gone_indices if index < chain[0])
            node = after_doc.para_nodes.get(new_index)
            merged.append({"paragraphs": chain, "kept_properties_of": chain[-1],
                           "text": ParagraphModel(after_doc.xml, node).text if node is not None else None})
    moved = []
    for move in moves:
        members = move["members"]
        moved.append({
            "name": move["name"],
            "ids": [member.id for member in members],
            "kept": "to" if mode == "accept" else "from",
            "from_paragraphs": sorted({m.paragraph_index for m in members if m.type == "moveFrom"
                                       and m.paragraph_index is not None}),
            "to_paragraphs": sorted({m.paragraph_index for m in members if m.type == "moveTo"
                                     and m.paragraph_index is not None}),
        })
    rows_removed = []
    tables_removed = []
    for node in rows:
        if node.name == "w:tbl":
            rows_inside = [row for row in node.children if row.name == "w:tr"]
            indices = [index_of[id(p)] for p in node.iter() if p.name == "w:p" and id(p) in index_of]
            tables_removed.append({"table": _row_label(doc, rows_inside[0], index_of).split(" row ")[0],
                                   "rows": len(rows_inside), "paragraphs": [indices[0], indices[-1] + 1]})
        else:
            rows_inside = [node]
        for row in rows_inside:
            _report_row(doc, row, index_of, rows_removed)
    # a removed row takes its paragraphs and its row-end marker (layout indices) with it
    shifts = {"index_delta": -len(gone_indices), "indices_after_shifted": bool(gone_indices),
              "index_shifts": [[index + 1, -1] for index in gone_indices]}
    if gone_indices:
        shifts["shift_base"] = gone_indices[0] + 1
    return _result(
        source, output, op, doc,
        mode=mode,
        scope=list(span) if span else None,
        resolved=[revision.as_dict() for revision in selected],
        resolved_count=len(selected),
        pulled_in=[revision.as_dict() for revision in pulled],
        removed_with=[revision.as_dict() for revision in removed_with],
        merged=merged,
        rows_removed=rows_removed,
        **({"tables_removed": tables_removed} if tables_removed else {}),
        **({"comments_removed": comments_removed} if comments_removed else {}),
        moved=moved,
        detached_comments=_detached_comments(doc.xml, new_xml),
        **shifts,
    )


# -- character sweep (plan 551) ------------------------------------------------------
#
# The table (textclean / char_map.json) decides what changes; this decides where it may change. Only plain
# body text is swept: characters of w:t elements outside fields, outside every existing revision, in
# paragraphs without drawings. OMML math (m:oMath / m:oMathPara), field codes and field results (Zotero
# citations included) and images are never touched; their would-be changes are counted as skipped.

SWEEP_OP = "sweep_characters"
_SEPARATORS = {" ", " "}
_SKIP_KINDS = ("math", "field", "revision", "drawing", "unaligned")
# Word's own manual line break (U+000B, written as <w:br/>): never cleaned in the Word integration, whatever
# char_map.json says (it strips U+000B for MCPs whose native break is something else).
_WORD_NATIVE = ("\v",)


def _word_keep(keep) -> tuple[str, ...]:
    return (*(keep or ()), *_WORD_NATIVE)


def _positional_clean(text: str, keep, line_break: str = "\v") -> tuple[list[str | None], list[dict]]:
    """Per character of `text`: its replacement (None = unchanged), using textclean's own decisions.

    textclean maps U+2028 / U+2029 to "\\n"; `line_break` replaces that (default "\\v", Word's manual line
    break, since a sweep or a one-paragraph field cannot split a paragraph). Pass "\\n" to keep textclean's."""

    from . import textclean as tc

    keep = _word_keep(keep)
    cleaned, changes = tc.clean(text, keep)
    if not changes:
        return [None] * len(text), []
    table = tc._table()[0]
    kept = tc._keep_set(keep)
    cps = [ord(ch) for ch in text]
    flag_tags = tc._flag_tags(cps) if any(0xE0000 <= cp <= 0xE007F for cp in cps) else set()
    out: list[str | None] = []
    for i, cp in enumerate(cps):
        entry = table.get(cp)
        if entry is None or cp in kept or i in flag_tags or (cp == tc._ZWJ and tc._zwj_in_emoji(cps, i)):
            out.append(None)
            continue
        out.append(tc._replacement(cp, entry))
    rebuilt = "".join(ch if repl is None else repl for ch, repl in zip(text, out))
    if rebuilt != cleaned:
        raise VerificationFailed(f"{SWEEP_OP}: positional cleaning disagrees with textclean.clean",
                                 expected=cleaned, actual=rebuilt)
    if line_break != "\n":
        for number, ch in enumerate(text):
            if ch in _SEPARATORS and out[number] is not None:
                out[number] = line_break
        for record in changes:
            if record["char"] in _SEPARATORS:
                record["to"] = line_break
    return out, changes


def clean_new_text(text: str, keep=(), *, old: str | None = None, single_paragraph: bool = True
                   ) -> tuple[str, list[dict]]:
    """Clean text an agent is about to write (plan 551).

    `old`: the text being replaced; characters the new text keeps from it (difflib's equal blocks) are left as
    they are, so text echoed from the document (field results, deliberate NBSPs) is not rewritten. A one-
    paragraph field gets "\\v" for U+2028/U+2029; a multi-paragraph field ("\\n" native) gets "\\n"."""

    from . import textclean as tc

    if not text:
        return text, []
    repl, _ = _positional_clean(text, keep, "\v" if single_paragraph else "\n")
    if old:
        import difflib

        for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, old, text, autojunk=False).get_opcodes():
            if tag == "equal":
                repl[j1:j2] = [None] * (j2 - j1)
    changed = "".join(ch for ch, item in zip(text, repl) if item is not None)
    if not changed:
        return text, []
    _, changes = _positional_clean(changed, (), "\v" if single_paragraph else "\n")
    return "".join(ch if item is None else item for ch, item in zip(text, repl)), changes


def _scan_count(text: str, keep) -> list[dict]:
    from . import textclean as tc

    return tc.scan(text, _word_keep(keep)) if text else []


def _math_texts(xml: str, node: Node, out: list[str], raws: list[str] | None = None) -> None:
    """Text (m:t) of each outermost OMML math element under `node`."""

    for child in node.children:
        if child.local in ("oMath", "oMathPara") and not child.name.startswith("w:"):
            if raws is not None:
                raws.append(xml[child.start:child.end])
            out.append("".join(unescape(item.inner(xml)) for item in child.iter()
                               if item.local == "t" and not item.name.startswith("w:") and not item.closed))
        else:
            _math_texts(xml, child, out, raws)


def _all_math(xml: str, root: Node) -> list[str]:
    raws: list[str] = []
    _math_texts(xml, root, [], raws)
    return raws


@dataclass
class _SweepParagraph:
    index: int
    model: ParagraphModel
    edits: list
    changes: list[dict]
    skipped: dict[str, list[dict]]
    refused: str | None = None


_DELETED_WRAPPERS = frozenset({"w:del", "w:moveFrom"})
_INSERTED_WRAPPERS = frozenset({"w:ins", "w:moveTo"})


@dataclass
class _FieldViews:
    """Field depth across the whole document, in three views: as stored (every w:fldChar), with all
    revisions accepted (fldChars inside w:del/w:moveFrom gone) and with all rejected (fldChars inside
    w:ins/w:moveTo gone). Text is field text when it lies inside a field in ANY view.

    ParagraphModel follows fields inside one paragraph only and counts deleted field markers as live, so
    neither a field that spans paragraphs (a Zotero bibliography, a TOC) nor a field whose begin/separate/end
    markers sit inside tracked revisions can be read from it alone."""

    views: list[list[tuple[int, int]]]
    revised: bool

    def protected(self, offset: int) -> bool:
        return any(_depth_at(marks, offset) > 0 for marks in self.views)


def _field_views(doc: Document) -> _FieldViews:
    events: list[tuple[int, int, str]] = []
    for node in doc.root.iter():
        if node.name != "w:fldChar":
            continue
        kind = node.attr(doc.xml, "w:fldCharType")
        if kind not in ("begin", "end"):
            continue
        names = {item.name for item in node.ancestors()}
        where = "del" if names & _DELETED_WRAPPERS else ("ins" if names & _INSERTED_WRAPPERS else "")
        events.append((node.start, 1 if kind == "begin" else -1, where))
    views = []
    for hidden in ("", "del", "ins"):
        depth, marks = 0, []
        for offset, step, where in events:
            if hidden and where == hidden:
                continue
            depth = max(0, depth + step)
            marks.append((offset, depth))
        views.append(marks)
    return _FieldViews(views, any(where for _, _, where in events))


def _depth_at(marks: list[tuple[int, int]], offset: int) -> int:
    position = bisect.bisect_left(marks, (offset, -1)) - 1
    return marks[position][1] if position >= 0 else 0


def _plan_paragraph(doc: Document, index: int, keep, fields: _FieldViews | None = None) -> _SweepParagraph:
    from .sections import Edit

    node = doc.paragraph(index, SWEEP_OP)
    xml = doc.xml
    model = ParagraphModel(xml, node)
    fields = fields if fields is not None else _field_views(doc)
    cache: dict[int, bool] = {}

    def in_field(char) -> bool:
        if char.field:
            return True
        key = id(char.child)
        if key not in cache:
            cache[key] = fields.protected(char.child.start)
        return cache[key]

    skipped: dict[str, list[dict]] = {kind: [] for kind in _SKIP_KINDS}
    maths: list[str] = []
    _math_texts(xml, node, maths)
    skipped["math"] += [_scan_count(text, keep) for text in maths]
    field_text = "".join(char.ch for char in model.chars if char.child.name == "w:t" and in_field(char))
    instr = "".join(unescape(item.node.inner(xml)) for item in model.items
                    if isinstance(item, Obj) and item.node.name in ("w:instrText", "w:delInstrText")
                    and not item.node.closed)
    simple = "".join(unescape(item.node.attr(xml, "w:instr") or "") for item in model.items
                     if isinstance(item, Obj) and item.node.name == "w:fldSimple")
    skipped["field"] += [_scan_count(field_text, keep), _scan_count(instr, keep), _scan_count(simple, keep)]
    revised = "".join(char.ch for char in model.chars
                      if char.revision is not None and char.child.name in ("w:t", "w:delText") and not in_field(char))
    skipped["revision"].append(_scan_count(revised, keep))
    final = model.final
    eligible = [char.revision is None and char.child.name == "w:t" and not in_field(char) for char in final]
    plan = _SweepParagraph(index, model, [], [], skipped)
    drawing = model.drawing()
    if drawing:
        plan.refused = f"holds a drawing ({drawing})"
        skipped["drawing"].append(_scan_count("".join(c.ch for c, ok in zip(final, eligible) if ok), keep))
        return plan
    replacements: list[str | None] = [None] * len(final)
    changes: list[list[dict]] = []
    position = 0
    while position < len(final):
        if not eligible[position]:
            position += 1
            continue
        end = position
        while end < len(final) and eligible[end]:
            end += 1
        segment = "".join(char.ch for char in final[position:end])
        repl, found = _positional_clean(segment, keep)
        replacements[position:end] = repl
        changes.append(found)
        position = end
    from . import textclean as tc

    plan.changes = tc.merge_changes(*changes)
    position = 0
    while position < len(final):
        if replacements[position] is None:
            position += 1
            continue
        end = position
        while end < len(final) and replacements[end] is not None and final[end].run is final[position].run:
            end += 1
        text = "".join(replacements[position:end])
        plan.edits.append(Edit(position, end, text, "replace" if text else "delete", position if text else None))
        position = end
    if plan.edits:
        layout_text = "".join(char.ch for char in model.chars if char.layout_visible)
        if layout_text != doc.layout.para(index).text:
            plan.refused = "could not be aligned with the document layout"
            skipped["unaligned"].append(plan.changes)
            plan.edits, plan.changes = [], []
    return plan


def _sweep_range(doc: Document, first: int | None, last: int | None) -> range:
    count = len(doc.layout.paragraphs)
    first = 1 if first is None else first
    last = count if last is None else last
    if not 1 <= first <= count or not first <= last <= count:
        raise RefusedUnsafe(f"{SWEEP_OP}: paragraph range {first}..{last} is outside 1..{count}")
    return range(first, last + 1)


def _sweep_plans(doc: Document, first, last, keep) -> list[_SweepParagraph]:
    fields = _field_views(doc)
    return [_plan_paragraph(doc, index, keep, fields)
            for index in _sweep_range(doc, first, last) if doc.para_nodes.get(index) is not None]


def _summary(plans: list[_SweepParagraph]) -> dict:
    from . import textclean as tc

    cleaned = tc.merge_changes(*(plan.changes for plan in plans))
    skipped = {kind: tc.total(tc.merge_changes(*(found for plan in plans for found in plan.skipped[kind])))
               for kind in _SKIP_KINDS}
    skipped_by = tc.merge_changes(*(found for plan in plans for kind in _SKIP_KINDS for found in plan.skipped[kind]))
    return {
        "cleaned": cleaned,
        "total": tc.total(cleaned),
        "by_codepoint": {rec["codepoint"]: {"name": rec["name"], "action": rec["action"], "count": rec["count"]}
                         for rec in cleaned},
        "skipped": {kind: count for kind, count in skipped.items() if count},
        "skipped_total": sum(skipped.values()),
        "skipped_by_codepoint": {rec["codepoint"]: rec["count"] for rec in skipped_by},
        "paragraphs_skipped": [{"paragraph_index": plan.index, "reason": plan.refused}
                               for plan in plans if plan.refused][:50],
    }


def scan_characters(path, *, first: int | None = None, last: int | None = None, keep=()) -> dict:
    """What sweep_characters would change (same skip rules), without writing: for inspect_document."""

    _, doc = load(path)
    summary = _summary(_sweep_plans(doc, first, last, keep))
    return {"total": summary["total"], "by_codepoint": summary["by_codepoint"],
            "skipped_total": summary["skipped_total"], "skipped": summary["skipped"]}


def _verify_anchors_multi(doc: Document, plans: list[tuple[TextEditor, list]], after_layout, op: str) -> None:
    from .sections import StreamEdit, anchor_errors_for_edits, _paragraph_starts

    before = doc.layout
    starts = _paragraph_starts([len(item.text) for item in before.paragraphs])
    lengths = [len(item.text) for item in before.paragraphs]
    stream: list[StreamEdit] = []
    delta = 0
    for editor, edits in plans:
        layout_chars = [char for char in editor.model.chars if char.layout_visible]
        position = {id(char): number for number, char in enumerate(layout_chars)}
        final = editor.model.final
        base = starts[editor.index - 1]
        own = 0
        for edit in edits:
            if edit.end > edit.start or edit.mode == "before":
                at = position[id(final[edit.start])]
            elif edit.mode == "after":
                at = position[id(final[edit.start - 1])] + 1
            else:
                at = 0
            stream.append(StreamEdit(base + at + delta + own, edit.end - edit.start, len(edit.text)))
            own += len(edit.text) - (edit.end - edit.start)
        lengths[editor.index - 1] += own
        delta += own
    problems = anchor_errors_for_edits(before, after_layout, stream, lengths)
    if problems:
        raise VerificationFailed(f"{op}: bookmark/comment anchors moved: " + "; ".join(problems[:5]))


def _instr_raws(xml: str) -> list[str]:
    return re.findall(r"<w:instrText\b.*?</w:instrText>|<w:fldSimple\b[^>]*>", xml, re.S)


def sweep_characters(source, output, *, paragraph_index: int | None = None,
                     through_paragraph_index: int | None = None, keep=(), author: str = DEFAULT_AUTHOR,
                     preview_only: bool = False, now: dt.datetime | None = None) -> dict:
    """Clean unusual characters in existing text as tracked changes (w:del/w:ins by `author`).

    Verified before saving, per paragraph as for tracked_replace_paragraph (rejecting the new revisions gives
    the input back; accepting gives the cleaned text; other revisions byte-identical), plus: every OMML math
    element and every field instruction is byte-identical, and bytes outside the swept paragraphs are too."""

    from .audit import audit_docx, file_sha256
    from .sections import read_layout

    author = (author or "").strip()
    if not author:
        raise RefusedUnsafe(f"{SWEEP_OP}: author must not be empty")
    last = through_paragraph_index if through_paragraph_index is not None else (
        paragraph_index if paragraph_index is not None else None)
    parts, doc = load(source)
    plans = _sweep_plans(doc, paragraph_index, last, keep)
    summary = _summary(plans)
    changing = [plan for plan in plans if plan.edits]
    base = {"tracked": True, "author": author, "keep": list(keep),
            "paragraph_range": [plans[0].index, plans[-1].index] if plans else None, **summary}
    if preview_only or not changing:
        if preview_only:
            # nothing is written, so nothing was cleaned: the would-be changes go under `would_clean`
            base["would_clean"], base["cleaned"] = base["cleaned"], []
        if Path(output).exists():
            raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
        Path(output).write_bytes(Path(source).read_bytes())
        return {"source": str(Path(source).resolve()), "output": str(Path(output).resolve()),
                "operation": SWEEP_OP, "engine": "direct_xml", "read_only": True, "preview_only": preview_only,
                "output_sha256": file_sha256(output), "paragraph_count": audit_docx(output).paragraph_count,
                "index_delta": 0, "paragraphs_changed": [plan.index for plan in changing],
                "revisions_added": 0, **base,
                "message": ("Preview: nothing was written." if preview_only
                            else "Nothing to clean outside math, fields and revisions; nothing was written.")}
    first_id = next_revision_id(parts)
    next_id = first_id
    stamp = revision_stamp(doc.xml, now)
    xml = doc.xml
    edited: list[tuple[TextEditor, list, str, str, list[dict]]] = []
    for plan in changing:
        editor = TextEditor(doc, plan.index, SWEEP_OP, tracked=True)
        if editor.model.text != plan.model.text:
            raise VerificationFailed(f"{SWEEP_OP}: paragraph {plan.index} text model changed between plan and edit")
        new_paragraph, added = editor.apply(plan.edits, author=author, first_id=next_id, stamp=stamp)
        next_id += len(added)
        expected = _expected_text(editor.model.text, plan.edits)
        edited.append((editor, plan.edits, new_paragraph, expected, added))
    pieces, cursor = [], 0
    for editor, _, new_paragraph, _, _ in sorted(edited, key=lambda item: item[0].node.start):
        pieces.append(xml[cursor:editor.node.start])
        pieces.append(new_paragraph)
        cursor = editor.node.end
    pieces.append(xml[cursor:])
    new_xml = "".join(pieces)
    for editor, edits, new_paragraph, expected, added in edited:
        node = editor.node
        alone = xml[:node.start] + new_paragraph + xml[node.end:]
        _verify_text_edit(doc, editor, alone, new_paragraph, edits, expected, added, author, True, SWEEP_OP,
                          requested=expected)
    if _all_math(new_xml, parse(new_xml)) != _all_math(xml, doc.root):
        raise VerificationFailed(f"{SWEEP_OP}: an equation (m:oMath) changed")
    if _instr_raws(new_xml) != _instr_raws(xml):
        raise VerificationFailed(f"{SWEEP_OP}: a field instruction changed")
    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, SWEEP_OP)
        _verify_anchors_multi(doc, [(editor, edits) for editor, edits, *_ in edited], read_layout(output), SWEEP_OP)
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    added_all = [item for *_, added in edited for item in added]
    return _result(
        source, output, SWEEP_OP, doc, **base,
        paragraphs_changed=[editor.index for editor, *_ in edited],
        revisions_added=len(added_all),
        revisions_added_sample=added_all[:20],
    )


# -- entry point from edit_batch -------------------------------------------------

def apply_operation(source, output, operation) -> dict:
    kind = operation.type
    if kind == "tracked_replace_text":
        return text_edit(source, output, op=kind, paragraph_index=operation.paragraph_index, tracked=True,
                         author=operation.author, find=operation.find, replace=operation.replace)
    if kind == "tracked_replace_paragraph":
        return text_edit(source, output, op=kind, paragraph_index=operation.paragraph_index, tracked=True,
                         author=operation.author, text=operation.text)
    if kind == SWEEP_OP:
        return sweep_characters(source, output, paragraph_index=operation.paragraph_index,
                                through_paragraph_index=operation.through_paragraph_index, keep=operation.keep,
                                author=operation.author, preview_only=operation.preview_only)
    selectors = dict(ids=operation.ids, author=operation.author, paragraph_index=operation.paragraph_index,
                     through_paragraph_index=operation.through_paragraph_index, heading=operation.heading,
                     until_heading=operation.until_heading)
    if kind in ("accept_revisions", "reject_revisions"):
        return resolve_revisions(source, output, mode=kind.split("_", 1)[0], **selectors)
    if kind == "list_revisions":
        listing = list_revisions(source, limit=operation.limit, **selectors)
        if Path(output).exists():
            raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
        Path(output).write_bytes(Path(source).read_bytes())
        from .audit import audit_docx, file_sha256

        return {**listing, "read_only": True, "output_sha256": file_sha256(output),
                "paragraph_count": audit_docx(output).paragraph_count, "index_delta": 0}
    raise RefusedUnsafe(f"unsupported revision operation {kind!r}")


def revision_count(path: str | Path) -> int:
    """Every revision element (ins/del/moves, every *PrChange, cell revisions) in every word/*.xml part.

    Namespace-aware (any prefix bound to the WordprocessingML namespace), unlike the audit's counter, which
    sees only w:ins/w:del/w:moveFrom/w:moveTo in the body. Every refusal guard and the direct-XML routing use
    this, so a document whose only revisions are property changes is never handed to Word."""

    wanted = {W + name for name in REV_TAGS}
    total = 0
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                try:
                    with archive.open(name) as stream:
                        total += sum(1 for _, el in ET.iterparse(stream) if el.tag in wanted)
                except ET.ParseError as exc:
                    raise RefusedUnsafe(f"cannot scan {name} for tracked revisions: {exc}") from exc
    return total


def has_revisions(path: str | Path) -> bool:
    return revision_count(path) > 0


def plain_replace(source, output, *, op: str, paragraph_index: int, find: str | None = None,
                  replace: str | None = None, text: str | None = None) -> dict:
    """Untracked replace on a document with revisions: allowed only where no revised run is touched."""

    return text_edit(source, output, op=op, paragraph_index=paragraph_index, tracked=False, find=find,
                     replace=replace, text=text)


# --------------------------------------------------------------------------
# Plan 864 S2: insert_paragraph as direct XML (decisions 9, 10, 12).
# --------------------------------------------------------------------------

_INSERT_OP = "insert_paragraph"


def _pstyle(style_id: str) -> str:
    return f'<w:pStyle w:val="{_attr_escape(style_id)}"/>'


def _mark_rpr_kept(xml: str, ppr: Node | None) -> str:
    """The paragraph mark's run properties (pPr/rPr children) without the mark's own revision markers."""

    rpr = ppr.child("w:rPr") if ppr is not None else None
    if rpr is None or rpr.closed:
        return ""
    return "".join(xml[item.start:item.end] for item in rpr.children if item.name not in _MARK_REVISIONS)


def _copied_ppr(xml: str, ppr: Node | None, *, style_id: str | None, mark: str, mark_rpr: str) -> str:
    """Decision 9: the anchor's pPr minus w:sectPr, w:pPrChange and the mark's revision markers; `style_id`
    replaces (or adds) w:pStyle; the mark keeps `mark_rpr` and takes `mark` (a new w:ins, tracked) first."""

    out: list[str] = []
    styled = False
    for child in (ppr.children if ppr is not None else []):
        if child.name in ("w:sectPr", "w:pPrChange", "w:rPr"):
            continue
        if child.name == "w:pStyle" and style_id is not None:
            out.append(_pstyle(style_id))
            styled = True
            continue
        out.append(xml[child.start:child.end])
    if style_id is not None and not styled:
        out.insert(0, _pstyle(style_id))
    if mark or mark_rpr:
        out.append(f"<w:rPr>{mark}{mark_rpr}</w:rPr>")  # rPr follows the other pPr children (CT_PPr)
    return f"<w:pPr>{''.join(out)}</w:pPr>" if out else ""


def _restyled_ppr(xml: str, ppr: Node | None, style_id: str, change: str | None) -> str:
    """The anchor's original pPr with w:pStyle set to `style_id`; tracked, a w:pPrChange (attributes `change`)
    records the original paragraph properties so Reject restores them."""

    children = ppr.children if ppr is not None else []
    out: list[str] = []
    styled = False
    for child in children:
        if child.name == "w:pStyle":
            out.append(_pstyle(style_id))
            styled = True
        else:
            out.append(xml[child.start:child.end])
    if not styled:
        out.insert(0, _pstyle(style_id))
    if change is not None:
        old = "".join(xml[child.start:child.end] for child in children
                      if child.name not in ("w:rPr", "w:sectPr", "w:pPrChange"))
        out.append(f"<w:pPrChange {change}>" + (f"<w:pPr>{old}</w:pPr>" if old else "<w:pPr/>") + "</w:pPrChange>")
    return f"<w:pPr>{''.join(out)}</w:pPr>"


def _with_ppr(xml: str, node: Node, ppr_xml: str) -> str:
    """Paragraph `node` with its pPr replaced by `ppr_xml` (added when it had none); content byte-identical."""

    if node.closed:
        head = node.head(xml)
        return head[:-2].rstrip() + ">" + ppr_xml + "</w:p>"
    old = node.child("w:pPr")
    if old is not None:
        return xml[node.start:old.start] + ppr_xml + xml[old.end:node.end]
    return xml[node.start:node.head_end] + ppr_xml + xml[node.head_end:node.end]


class _InsertPlan:
    """The new paragraphs (and the anchor, Word's way) as text replacing xml[start:end]."""

    def __init__(self, start: int, end: int, replacement: str, added: list[dict], word_way: bool) -> None:
        self.start, self.end, self.replacement, self.added, self.word_way = start, end, replacement, added, word_way


def _plan_insert(doc: Document, node: Node, pieces: list[str], *, style_id: str | None, default_style: bool,
                 tracked: bool, author: str, first_id: int, stamp: str, word_way: bool,
                 following: Node | None, before: bool = False, text_rpr: str | None = None) -> _InsertPlan:
    xml = doc.xml
    ppr = node.child("w:pPr")
    writer = _Writer(first_id, author, stamp, tracked)
    mark_rpr = "" if style_id is not None else _mark_rpr_kept(xml, ppr)
    run_rpr = f"<w:rPr>{mark_rpr}</w:rPr>" if mark_rpr else ""
    if text_rpr is not None:  # S3a: typed like the anchor's first (before) or last (after) character
        run_rpr = text_rpr

    def mark() -> str:
        return f"<w:ins {writer._attrs('ins', '¶')}/>" if tracked else ""

    def body(piece: str) -> str:
        run = f"<w:r>{run_rpr}{_render_text(piece, False)}</w:r>"
        return f"<w:ins {writer._attrs('ins', piece)}>{run}</w:ins>" if tracked else run

    def copied(piece: str) -> str:
        ppr_xml = _copied_ppr(xml, ppr, style_id=style_id, mark=mark(), mark_rpr=mark_rpr)
        return f"<w:p>{ppr_xml}{body(piece)}</w:p>"

    if before:
        return _InsertPlan(node.start, node.start, "".join(copied(piece) for piece in pieces), writer.added, False)
    if not word_way:
        at = following.start if following is not None else node.end
        replacement = "".join(copied(piece) for piece in pieces)
        return _InsertPlan(at, at, replacement, writer.added, False)
    # Decision 10, Word's way: the anchor's mark becomes the inserted one and the last new paragraph takes the
    # anchor's original pPr (section break, other authors' mark revisions and property changes included), so
    # Reject merges the anchor back into it unchanged.
    anchor_ppr = _copied_ppr(xml, ppr, style_id=None, mark=mark(), mark_rpr=_mark_rpr_kept(xml, ppr))
    parts = [_with_ppr(xml, node, anchor_ppr)]
    parts.extend(copied(piece) for piece in pieces[:-1])
    original = xml[ppr.start:ppr.end] if ppr is not None else ""
    last_ppr = original
    current = ppr.child("w:pStyle") if ppr is not None else None
    current_id = current.attr(xml, "w:val") if current is not None else None
    if style_id is not None and current_id != style_id and not (current_id is None and default_style):
        if ppr is not None and ppr.child("w:pPrChange") is not None:
            change = ppr.child("w:pPrChange")
            raise RefusedUnsafe(
                f"{_INSERT_OP}: the new last paragraph takes paragraph properties that already hold a tracked "
                f"formatting change ({describe_wrapper(xml, change)}), so its style cannot be changed as well; "
                "accept or reject that change first, or insert without `style`")
        change = f"{writer._attrs('pPrChange', '')}" if tracked else None
        last_ppr = _restyled_ppr(xml, ppr, style_id, change)
    parts.append(f"<w:p>{last_ppr}{body(pieces[-1])}</w:p>")
    return _InsertPlan(node.start, node.end, "".join(parts), writer.added, True)


def _props_key(element: ET.Element | None) -> str:
    """Canonical pPr/rPr with an empty rPr or pPr read as absent (Accept of a mark-only w:ins leaves one)."""

    from .sections import canonical

    if element is None:
        return ""
    clone = copy.deepcopy(element)
    for rpr in clone.findall(W + "rPr"):
        if not len(rpr) and not rpr.attrib:
            clone.remove(rpr)
    return canonical(clone) if len(clone) or clone.attrib else ""


def _set_pstyle(ppr: ET.Element, style_id: str) -> None:
    current = ppr.find(W + "pStyle")
    if current is None:
        current = ET.Element(W + "pStyle")
        ppr.insert(0, current)
    current.attrib.clear()
    current.set(W + "val", style_id)


def _verify_insert_accepted(op: str, old: ET.Element, accepted: ET.Element, *, anchor: int, at: int,
                            pieces: list[str], style_id: str | None, default_style: bool, word_way: bool,
                            run_key: str | None = None) -> None:
    """F1 (S2 review r1): the Accept view of the result, checked from the input container alone (never from
    the insert's own plan). `old` is the input container, `accepted` the result's with the new revisions
    accepted (the result itself when untracked); `anchor` is the anchor's child index in `old` and `at` the
    child index where the new paragraphs must start (derived from the input: right after the anchor, or right
    before the following paragraph). Checks: the n new paragraphs sit at at..at+n-1 and every other child
    equals the input's, in order (Word's way: the anchor equal but for its pPr); each new paragraph's pPr is the
    input anchor's pPr minus w:sectPr, w:pPrChange and the mark's revision markers (pStyle = `style_id`, mark
    formatting dropped, when a style is given), and Word's way the anchor takes that pPr (no style) and the last
    new paragraph the anchor's original one (restyled); each new paragraph holds only runs, every character
    formatted like the anchor's mark (no formatting when a style is given), reading exactly as requested."""

    from .sections import canonical

    n = len(pieces)
    old_children, new_children = list(old), list(accepted)
    if len(new_children) != len(old_children) + n:
        raise VerificationFailed(f"{op}: the result does not hold exactly {n} new paragraph(s) after Accept")
    added = new_children[at:at + n]
    others = new_children[:at] + new_children[at + n:]
    for number, (before, after) in enumerate(zip(old_children, others)):
        if word_way and number == anchor:
            strip = lambda el: [canonical(child) for child in el if child.tag != W + "pPr"]  # noqa: E731
            if after.tag != W + "p" or strip(after) != strip(before):
                raise VerificationFailed(f"{op}: the anchor paragraph's content changed")
            continue
        if canonical(after) != canonical(before):
            raise VerificationFailed(f"{op}: the new paragraphs are not right after the anchor, or another "
                                     "paragraph changed (Accept view)")
    source_ppr = old_children[anchor].find(W + "pPr")
    base = copy.deepcopy(source_ppr) if source_ppr is not None else ET.Element(W + "pPr")
    for child in list(base):
        if child.tag in (W + "sectPr", W + "pPrChange"):
            base.remove(child)
    mark = base.find(W + "rPr")
    if mark is not None:
        for child in list(mark):
            if "w:" + _local(child.tag) in _MARK_REVISIONS:
                mark.remove(child)
    text_like = run_key is not None
    if run_key is None:
        run_key = canonical(mark) if mark is not None and len(mark) and style_id is None else ""
    copied = copy.deepcopy(base)
    if style_id is not None:
        if mark is not None:
            copied.remove(copied.find(W + "rPr"))
        _set_pstyle(copied, style_id)
    expected = [_props_key(copied)] * n
    if word_way:
        if _props_key(others[anchor].find(W + "pPr")) != _props_key(base):
            raise VerificationFailed(f"{op}: the anchor paragraph's properties are not its own minus the section "
                                     "break, property change and mark revisions (Word's way)")
        last = copy.deepcopy(source_ppr) if source_ppr is not None else ET.Element(W + "pPr")
        current = last.find(W + "pStyle")
        current_id = current.get(W + "val") if current is not None else None
        if style_id is not None and current_id != style_id and not (current_id is None and default_style):
            _set_pstyle(last, style_id)
        expected[-1] = _props_key(last)
    for number, (paragraph, piece, want) in enumerate(zip(added, pieces, expected)):
        label = f"new paragraph {number + 1}"
        if paragraph.tag != W + "p":
            raise VerificationFailed(f"{op}: {label} is not a paragraph")
        if _props_key(paragraph.find(W + "pPr")) != want:
            raise VerificationFailed(f"{op}: {label}'s paragraph properties are not the anchor's (decision 9/10)")
        if any(child.tag not in (W + "pPr", W + "r") for child in paragraph):
            raise VerificationFailed(f"{op}: {label} holds something other than runs")
        chars = _et_final_chars(paragraph)
        if "".join(ch for ch, _ in chars) != piece:
            raise VerificationFailed(f"{op}: {label} does not read as requested")
        if any(key != run_key for _, key in chars):
            raise VerificationFailed(f"{op}: {label}'s text is not formatted like the anchor's "
                                     + ("text" if text_like else "paragraph mark"))


def _text_rpr(op: str, doc: Document, node: Node, index: int, *, first: bool) -> tuple[str, str] | None:
    """The run formatting new text typed at the anchor's start (`first`) or end takes, Word's way: that of the
    anchor's first or last visible character, as raw rPr XML and its comparison key. None for an anchor with no
    visible text (the paragraph mark's formatting applies then). Refused when that character sits in a field,
    hyperlink or content control, or carries a tracked formatting change."""

    from .sections import canonical

    model = ParagraphModel(doc.xml, node)
    if not model.final:
        return None
    char = model.final[0 if first else -1]
    where = "first" if first else "last"
    if char.field or char.flags:
        raise RefusedUnsafe(f"{op}: the {where} character of paragraph {index} is inside a field, hyperlink or "
                            "content control, so new lines cannot take its formatting; edit that line in Word")
    if "<w:rPrChange" in char.rpr:
        raise RefusedUnsafe(f"{op}: the {where} character of paragraph {index} carries a tracked formatting change, "
                            "so new lines cannot take its formatting; accept or reject that change first")
    if not char.rpr:
        return "", ""
    element = _fragment(doc.xml, char.rpr)
    return char.rpr, (canonical(element) if len(element) else "")


def insert_paragraphs(source, output, *, after_paragraph_index: int, text: str, style_id: str | None = None,
                      default_style: bool = False, style_name: str | None = None, tracked: bool,
                      author: str = DEFAULT_AUTHOR, now: dt.datetime | None = None, before: bool = False,
                      run_rpr: bool = False) -> dict:
    """New paragraph(s) after `after_paragraph_index`, one per "\\n"-separated line, as direct XML.

    Tracked: each new paragraph mark and run is a w:ins by `author` (ids ascending in document order); the
    new paragraphs go right before the following paragraph, or, when there is none in the anchor's container
    (end of body, last paragraph of a cell, before a table or a block content control) or the anchor holds a
    section break, Word's way (decision 10). Untracked: the same paragraphs without revision markers, Word's
    way only for an anchor that holds a section break. Verified before anything is published (decision 12).

    Plan 864 S3a (row and section edits): `before` puts the new paragraphs right before the anchor
    (`after_paragraph_index` then names the paragraph they precede), each with the anchor's copied pPr, never
    Word's way; `run_rpr` gives the new text the run formatting of the anchor's first (before) or last (after)
    visible character, as Word types it, instead of the paragraph mark's (no style may be given with it).
    """

    from .audit import file_sha256
    from .sections import StreamEdit, _paragraph_starts, anchor_errors_for_edits, canonical, read_layout

    op = _INSERT_OP
    author = (author or "").strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{op}: author must not be empty")
    if not text:
        raise RefusedUnsafe(f"{op}: text must not be empty")
    if "\r" in text:
        raise RefusedUnsafe(f"{op}: text holds a carriage return; separate paragraphs with \"\\n\"")
    pieces = text.split("\n")
    if any(not piece for piece in pieces):
        raise RefusedUnsafe(f"{op}: text has an empty line; every \"\\n\"-separated line becomes a paragraph and "
                            "must hold text")
    parts, doc = load(source)
    index = after_paragraph_index
    node = doc.paragraph(index, op)
    para = doc.layout.para(index)
    if para.nested:
        raise RefusedUnsafe(f"{op}: paragraph {index} is inside a nested table; insert there in Word")
    xml = doc.xml
    ppr = node.child("w:pPr")
    mark_rpr = ppr.child("w:rPr") if ppr is not None else None
    for item in (mark_rpr.children if mark_rpr is not None and not before else []):
        if item.name in ("w:del", "w:moveFrom"):
            raise RefusedUnsafe(
                f"{op}: paragraph {index}'s mark is a tracked deletion ({describe_wrapper(xml, item)}), so it "
                "merges with the next paragraph; accept or reject that revision first, or insert after "
                "another paragraph")
    container = node.parent
    siblings = container.children
    if before:
        back = siblings.index(node) - 1
        while back >= 0 and siblings[back].name in _BETWEEN:
            back -= 1
        prior = siblings[back] if back >= 0 and siblings[back].name == "w:p" else None
        prior_ppr = prior.child("w:pPr") if prior is not None else None
        prior_mark = prior_ppr.child("w:rPr") if prior_ppr is not None else None
        for item in (prior_mark.children if prior_mark is not None else []):
            if item.name in ("w:del", "w:moveFrom"):
                raise RefusedUnsafe(
                    f"{op}: the paragraph before paragraph {index} has a tracked deletion of its mark "
                    f"({describe_wrapper(xml, item)}), so Accept would merge it into the new paragraph; accept "
                    "or reject that revision first")
    text_rpr = run_key = None
    if run_rpr:
        if style_id is not None:
            raise RefusedUnsafe(f"{op}: a style and the anchor's text formatting cannot both be given")
        typed = _text_rpr(op, doc, node, index, first=before)
        if typed is not None:
            text_rpr, run_key = typed
    at = siblings.index(node) + 1
    while at < len(siblings) and siblings[at].name in _BETWEEN:
        at += 1
    following = siblings[at] if at < len(siblings) and siblings[at].name == "w:p" else None
    has_section = ppr is not None and ppr.child("w:sectPr") is not None
    word_way = not before and (has_section or (tracked and following is None))
    first_id = next_revision_id(parts)
    stamp = revision_stamp(xml, now)
    plan = _plan_insert(doc, node, pieces, tracked=tracked, style_id=style_id, default_style=default_style,
                        author=author, first_id=first_id, stamp=stamp, word_way=word_way, following=following,
                        before=before, text_rpr=text_rpr)
    new_xml = xml[:plan.start] + plan.replacement + xml[plan.end:]

    # -- decision 12: verified in memory first ---------------------------------------------------------------
    # Every check below reads the input and the result only, never `plan` (S2 review r1, F1).
    delta = len(new_xml) - len(xml)
    if new_xml[:container.start] != xml[:container.start] or new_xml[container.end + delta:] != xml[container.end:]:
        raise VerificationFailed(f"{op}: bytes outside the anchor's container changed")
    old_box, new_box = xml[container.start:container.end], new_xml[container.start:container.end + delta]
    old_tree, new_tree = parse(old_box), parse(new_box)
    # New revisions: by `author`, numbered from the input's next free id (not the plan's list).
    new_ids = {item.attr(new_box, "w:id") for item in new_tree.iter()
               if tracked and item.local in REV_TAGS and item.name.startswith("w:")
               and item.attr(new_box, "w:author") == author and (item.attr(new_box, "w:id") or "").isdigit()
               and int(item.attr(new_box, "w:id")) >= first_id}

    def is_rev(tree_xml: str, item: Node) -> bool:
        return item.local in REV_TAGS and item.name.startswith("w:")

    def fresh(tree_xml: str, item: Node) -> bool:
        return tracked and item.attr(tree_xml, "w:id") in new_ids and item.attr(tree_xml, "w:author") == author

    old_revs = Counter(_revision_raw(old_box, item) for item in old_tree.iter() if is_rev(old_box, item))
    kept = Counter(_revision_raw(new_box, item) for item in new_tree.iter()
                   if is_rev(new_box, item) and not fresh(new_box, item))
    if kept != old_revs:
        raise VerificationFailed(f"{op}: a pre-existing revision changed")
    added_nodes = [item for item in new_tree.iter() if is_rev(new_box, item) and fresh(new_box, item)]
    if len(added_nodes) != len(new_ids) or any(item.name not in ("w:ins", "w:pPrChange") for item in added_nodes):
        raise VerificationFailed(f"{op}: the new revisions are not one insertion per paragraph mark and run")
    _check_stamps(new_box, added_nodes, author, op)
    old_el, new_el = _fragment(xml, old_box), _fragment(new_xml, new_box)
    is_new = lambda el: tracked and el.get(W + "id") in new_ids and el.get(W + "author") == author  # noqa: E731
    accepted = new_el
    if tracked:
        rejected = copy.deepcopy(new_el)
        et_resolve(rejected, is_new, "reject")
        if canonical(rejected) != canonical(old_el):
            raise VerificationFailed(f"{op}: rejecting the new paragraphs does not give back the input (D5)")
        accepted = copy.deepcopy(new_el)
        et_resolve(accepted, is_new, "accept")
    anchor_at = siblings.index(node)
    if before:
        new_at = anchor_at
    else:
        new_at = anchor_at + 1 if word_way or following is None else siblings.index(following)
    _verify_insert_accepted(op, old_el, accepted, anchor=anchor_at, at=new_at, pieces=pieces, style_id=style_id,
                            default_style=default_style, word_way=word_way, run_key=run_key)
    before_lengths = [len(item.text) for item in doc.layout.paragraphs]
    split = index - 1 if before else index  # layout list position the new paragraphs take
    lengths = before_lengths[:split] + [len(piece) for piece in pieces] + before_lengths[split:]
    position = (_paragraph_starts(before_lengths)[split] if split < len(before_lengths)
                else sum(length + 1 for length in before_lengths))
    stream = [StreamEdit(position, 0, sum(len(piece) + 1 for piece in pieces), either_side=True)]

    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, op)
        after_layout = read_layout(output)
        texts = [item.text for item in after_layout.paragraphs]
        expected = [item.text for item in doc.layout.paragraphs]
        expected[split:split] = pieces
        if texts != expected:
            raise VerificationFailed(f"{op}: the document does not read as requested after the insert",
                                     expected=expected[split:split + len(pieces)],
                                     actual=texts[split:split + len(pieces)])
        if style_id is not None:
            got = [after_layout.para(number).style for number in range(split + 1, split + len(pieces) + 1)]
            if any(value != style_id and not (value is None and default_style) for value in got):
                raise VerificationFailed(f"{op}: the new paragraphs do not carry style {style_id!r}", actual=got)
        problems = anchor_errors_for_edits(doc.layout, after_layout, stream, lengths)
        if problems:
            raise VerificationFailed(f"{op}: bookmark/comment anchors moved: " + "; ".join(problems[:5]))
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    new_index = split + 1
    return _result(
        source, output, op, doc,
        source_sha256=file_sha256(source),
        **({"before_paragraph_index": index} if before else {"after_paragraph_index": index}),
        new_paragraph_index=new_index,
        new_paragraph_indices=list(range(new_index, new_index + len(pieces))),
        paragraph_texts=pieces,
        index_delta=len(pieces),
        indices_after_shifted=True,
        stale_from_index=new_index,
        style=style_name,
        style_id=style_id,
        tracked=tracked,
        author=author if tracked else None,
        revisions_added=plan.added if tracked else [],
        anchor_mark_moved=word_way,
    )


# --------------------------------------------------------------------------
# Plan 864 S3a: paragraph deletion as direct XML (decision 11), for row and section edits.
# --------------------------------------------------------------------------

_DELETE_OP = "delete_paragraphs"
_DELETABLE_RUN_PARTS = frozenset({"w:rPr", "w:t", "w:tab", "w:br", "w:cr", "w:lastRenderedPageBreak",
                                  "w:noBreakHyphen", "w:softHyphen", "w:sym"})
_OBJECT_NAMES = {
    "w:bookmarkStart": "a bookmark", "w:bookmarkEnd": "a bookmark", "w:commentRangeStart": "a comment",
    "w:commentRangeEnd": "a comment", "w:commentReference": "a comment", "w:fldChar": "a field",
    "w:instrText": "a field", "w:fldSimple": "a field", "w:hyperlink": "a hyperlink", "w:sdt": "a content control",
    "w:drawing": "a drawing", "mc:AlternateContent": "a drawing", "w:pict": "a drawing", "w:object": "a drawing",
    "w:footnoteReference": "a footnote", "w:endnoteReference": "an endnote",
}


def _mark_revision(ppr: Node | None) -> Node | None:
    """A paragraph's own revision: a w:pPrChange, or an ins/del/move/rPrChange on its mark."""

    if ppr is None:
        return None
    change = ppr.child("w:pPrChange")
    if change is not None:
        return change
    rpr = ppr.child("w:rPr")
    return next((item for item in (rpr.children if rpr is not None else []) if item.name in _MARK_REVISIONS), None)


_KEPT_DELETION_RUN_PARTS = (_DELETABLE_RUN_PARTS - {"w:t"}) | {"w:delText"}


def _kept_deletion_problem(xml: str, wrapper: Node) -> str | None:
    """Plan 864 S3c: why an existing w:del cannot stay as it is inside a deleted paragraph (None: it holds only
    plain deleted text runs and proofing marks)."""

    for child in wrapper.children:
        if child.name == "w:proofErr":
            continue
        if child.name != "w:r":
            found = next((sub for sub in child.iter() if sub.name in _OBJECT_NAMES), child)
            return _OBJECT_NAMES.get(found.name, f"<{found.name}>")
        rpr = child.child("w:rPr")
        change = next((item for item in rpr.iter() if item.name == "w:rPrChange"), None) if rpr is not None else None
        if change is not None:
            return f"a tracked formatting change ({describe_wrapper(xml, change)})"
        for item in child.children:
            if item.name not in _KEPT_DELETION_RUN_PARTS:
                found = next((sub for sub in item.iter() if sub.name in _OBJECT_NAMES), item)
                return _OBJECT_NAMES.get(found.name, f"<{found.name}>")
    return None


def _refuse_paragraph(op: str, xml: str, node: Node, index: int, role: str, *, content: bool,
                      keep_deletions: bool = False) -> None:
    """Refused (named) unless the paragraph can be deleted (`content`) or lose its mark: no section break, no
    revision of its own, and (`content`) only plain runs of text. With `keep_deletions` (plan 864 S3c) an
    existing w:del of plain text runs (any author) is allowed: it stays as it is."""

    label = f"paragraph {index} ({role})"
    ppr = node.child("w:pPr")
    if ppr is not None and ppr.child("w:sectPr") is not None:
        raise RefusedUnsafe(f"{op}: {label} holds a section break, which deleting it would remove; delete it in Word")
    revision = _mark_revision(ppr)
    if revision is not None:
        raise RefusedUnsafe(f"{op}: {label} holds a tracked revision ({describe_wrapper(xml, revision)}); accept or "
                            "reject it first, then edit again")
    if not content:
        return
    for child in node.children:
        if child.name in ("w:pPr", "w:proofErr"):
            continue
        if keep_deletions and child.name == "w:del":
            problem = _kept_deletion_problem(xml, child)
            if problem is not None:
                raise RefusedUnsafe(f"{op}: {label} holds {problem} inside tracked deletion "
                                    f"({describe_wrapper(xml, child)}), which deleting the paragraph would remove; "
                                    "accept or reject that deletion first, or delete the paragraph in Word")
            continue
        parts = [child] if child.name != "w:r" else [item for item in child.children
                                                     if item.name not in _DELETABLE_RUN_PARTS]
        if child.name == "w:r":
            rpr = child.child("w:rPr")
            change = next((item for item in rpr.iter() if item.name == "w:rPrChange"), None) \
                if rpr is not None else None
            if change is not None:
                parts = [change]
        for item in parts:
            if item.local in REV_TAGS and item.name.startswith("w:"):
                raise RefusedUnsafe(f"{op}: {label} holds a tracked revision ({describe_wrapper(xml, item)}); "
                                    "accept or reject it first, then edit again")
            found = next((sub for sub in item.iter() if sub.name in _OBJECT_NAMES), item)
            what = _OBJECT_NAMES.get(found.name, f"<{found.name}>")
            raise RefusedUnsafe(f"{op}: {label} holds {what}, which deleting the paragraph would remove; delete "
                                "it in Word, or keep that line")


def _ppr_with_mark(ppr_xml: str, mark: str) -> str:
    """`ppr_xml` (a whole w:pPr, or "") with `mark` first in its paragraph-mark rPr (created when missing)."""

    if not ppr_xml:
        return f"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>"
    ppr = parse(ppr_xml).children[0]
    if ppr.closed:
        return f"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>"
    rpr = ppr.child("w:rPr")
    if rpr is not None:
        if rpr.closed:
            return ppr_xml[:rpr.start] + f"<w:rPr>{mark}</w:rPr>" + ppr_xml[rpr.end:]
        return ppr_xml[:rpr.head_end] + mark + ppr_xml[rpr.head_end:]
    after = next((item for item in ppr.children if item.name in ("w:sectPr", "w:pPrChange")), None)
    at = after.start if after is not None else ppr.close_start()
    return ppr_xml[:at] + f"<w:rPr>{mark}</w:rPr>" + ppr_xml[at:]


def _deleted_content(xml: str, node: Node, writer: _Writer, text: str, *, keep_deletions: bool = False) -> str:
    """Everything after the paragraph's pPr in one w:del, text retagged as w:delText; "" for an empty one.

    `keep_deletions` (plan 864 S3c): existing w:del children stay byte-identical and each stretch of live
    children between them that holds a run goes into its own new w:del (ids ascending left to right)."""

    if node.closed:
        return ""
    ppr = node.child("w:pPr")
    body = xml[ppr.end if ppr is not None else node.head_end:node.close_start()]
    if not any(item.name == "w:r" for item in node.children):
        return body  # nothing to delete but the mark (proofing marks stay as they are)
    if keep_deletions and any(item.name == "w:del" for item in node.children):
        out, cursor = "", ppr.end if ppr is not None else node.head_end
        for stretch in _stretches(node):
            if isinstance(stretch, list):
                out += xml[cursor:stretch[0].start] + _deleted_stretch(xml, stretch, writer)
                cursor = stretch[-1].end
        return out + xml[cursor:node.close_start()]
    body = re.sub(r"<w:t(?=[\s>/])", "<w:delText", body).replace("</w:t>", "</w:delText>")
    return writer.deletion(body, text)


def _stretches(node: Node) -> list:
    """A paragraph's children after its pPr: each existing w:del alone, the live children between them as lists."""

    out: list = []
    for child in node.children:
        if child.name == "w:pPr":
            continue
        if child.name == "w:del":
            out.append(child)
        elif out and isinstance(out[-1], list):
            out[-1].append(child)
        else:
            out.append([child])
    return out


def _deleted_stretch(xml: str, stretch: list[Node], writer: _Writer) -> str:
    body = xml[stretch[0].start:stretch[-1].end]
    if not any(item.name == "w:r" for item in stretch):
        return body  # proofing marks only: nothing to delete
    text = "".join(unescape(sub.inner(xml)) if sub.name == "w:t" and not sub.closed else
                   "\t" if sub.name == "w:tab" else "\n" if sub.name in ("w:br", "w:cr") else ""
                   for item in stretch if item.name == "w:r" for sub in item.children)
    body = re.sub(r"<w:t(?=[\s>/])", "<w:delText", body).replace("</w:t>", "</w:delText>")
    return writer.deletion(body, text)


def _paragraph(xml: str, node: Node, ppr_xml: str, content: str) -> str:
    head = node.head(xml)[:-2].rstrip() + ">" if node.closed else xml[node.start:node.head_end]
    return head + ppr_xml + content + "</w:p>"


def _ppr_parts(xml: str, ppr: Node | None, skip=("w:rPr", "w:sectPr", "w:pPrChange")) -> str:
    return "".join(xml[item.start:item.end] for item in (ppr.children if ppr is not None else [])
                   if item.name not in skip)


def _normalized(element: ET.Element) -> str:
    """Canonical XML with an empty mark rPr and an empty pPr read as absent (a rejected mark deletion leaves
    the rPr/pPr that were created to hold it)."""

    from .sections import canonical

    clone = copy.deepcopy(element)
    for paragraph in clone.iter(W + "p"):
        ppr = paragraph.find(W + "pPr")
        if ppr is None:
            continue
        rpr = ppr.find(W + "rPr")
        if rpr is not None and not len(rpr) and not rpr.attrib:
            ppr.remove(rpr)
        if not len(ppr) and not ppr.attrib:
            paragraph.remove(ppr)
    return canonical(clone)


def delete_paragraphs(source, output, *, first: int, last: int, trailing: bool = False,
                      previous: int | None = None, fix_format: bool = False, tracked: bool,
                      author: str = DEFAULT_AUTHOR, now: dt.datetime | None = None,
                      keep_deletions: bool = False) -> dict:
    """Delete layout paragraphs first..last (one container, consecutive) as direct XML (decision 11).

    Tracked: each paragraph's content in one w:del by `author` (w:t -> w:delText) and its mark deleted
    (`<w:del/>` first in the mark's rPr); Accept merges the emptied paragraphs into the next one. `trailing`
    (the last lines of a table cell, whose last mark cannot go): Word's way, the mark of `previous` (the line
    right before) is deleted instead of the last doomed line's, whose content is deleted and whose mark stays,
    so Accept leaves `previous`'s text on that mark; with `fix_format` and paragraph properties that differ
    (the mark's run formatting aside), that mark takes `previous`'s properties with a w:pPrChange holding its
    own, as Word's restore of the previous paragraph's format reads. A non-trailing range with no paragraph
    after it in its container is deleted the trailing way (`fallback_trailing`). Untracked: the paragraphs are
    removed (trailing: `previous` keeps its text on the surviving mark).

    Refused, naming the paragraph: a section break, a revision of its own (mark or properties) on any
    paragraph whose mark changes, and in a deleted paragraph anything but plain text runs (another author's
    revision, a field, a drawing, a hyperlink or content control, a bookmark or comment marker), or markers
    between the paragraphs. Verified before anything is published: bytes outside the container identical,
    every pre-existing revision byte-identical, new ids ascending by `author`, Reject of the new revisions =
    the input, Accept (untracked: the result) = the input with the paragraphs deleted as described.

    `keep_deletions` (plan 864 S3c, the delete_paragraph op): a deleted paragraph may hold existing w:del
    wrappers (any author) of plain text runs; they stay byte-identical, the live stretches around them get
    their own new w:del, and Accept of the new revisions carries the old ones into the paragraph the deleted
    ones merge into (Word's merge). Untracked, such a paragraph is refused (it would drop that deletion).
    """

    from .audit import file_sha256
    from .sections import read_layout

    op = _DELETE_OP if not keep_deletions else "delete_paragraph"
    author = (author or "").strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{op}: author must not be empty")
    if last < first:
        raise RefusedUnsafe(f"{op}: last ({last}) is before first ({first})")
    parts, doc = load(source)
    xml = doc.xml
    doomed = [doc.paragraph(index, op) for index in range(first, last + 1)]
    for index in range(first, last + 1):
        if doc.layout.para(index).nested:
            raise RefusedUnsafe(f"{op}: paragraph {index} is inside a nested table; delete it in Word")
    container = doomed[0].parent
    siblings = container.children
    start_at = siblings.index(doomed[0])
    if any(node.parent is not container for node in doomed) or \
            siblings[start_at:start_at + len(doomed)] != doomed:
        raise RefusedUnsafe(f"{op}: paragraphs {first}..{last} are not consecutive paragraphs of one container "
                            "(something lies between them); delete them one by one")
    after_at = start_at + len(doomed)
    while after_at < len(siblings) and siblings[after_at].name in _BETWEEN:
        after_at += 1
    fallback = False
    if not trailing and not (after_at < len(siblings) and siblings[after_at].name == "w:p"):
        trailing, fix_format, fallback = True, True, True
        previous = first - 1
    prior: Node | None = None
    if trailing:
        if previous is None or previous != first - 1:
            raise RefusedUnsafe(f"{op}: a trailing delete merges into the line right before paragraph {first}; "
                                f"previous is {previous}")
        prior = doc.paragraph(previous, op) if previous >= 1 else None
        if prior is None or start_at == 0 or siblings[start_at - 1] is not prior:
            raise RefusedUnsafe(f"{op}: paragraph {first} has no paragraph right before it in its container to keep "
                                "the last mark; a cell or container must keep at least one paragraph")
        _refuse_paragraph(op, xml, prior, previous, "the line before, whose mark is deleted", content=False)
    for index, node in zip(range(first, last + 1), doomed):
        _refuse_paragraph(op, xml, node, index, "deleted", content=True, keep_deletions=keep_deletions)
        held = next((child for child in node.children if child.name == "w:del"), None)
        if held is not None and not tracked:
            raise RefusedUnsafe(f"{op}: paragraph {index} holds tracked deletion ({describe_wrapper(xml, held)}), "
                                "which an untracked delete would drop; delete it tracked, or accept or reject that "
                                "deletion first")

    first_id = next_revision_id(parts)
    stamp = revision_stamp(xml, now)
    writer = _Writer(first_id, author, stamp, tracked)
    texts_before = [ParagraphModel(xml, node).text for node in doomed]
    last_node = doomed[-1]
    last_ppr = last_node.child("w:pPr")
    fixed_parts = None
    if trailing and fix_format:
        prior_ppr = prior.child("w:pPr")
        own, theirs = _ppr_parts(xml, last_ppr), _ppr_parts(xml, prior_ppr)
        if _props_key(_fragment(xml, f"<w:pPr>{own}</w:pPr>")) != _props_key(_fragment(xml, f"<w:pPr>{theirs}</w:pPr>")):
            mark_rpr = last_ppr.child("w:rPr") if last_ppr is not None else None
            kept = xml[mark_rpr.start:mark_rpr.end] if mark_rpr is not None else ""
            fixed_parts = (theirs, kept, own)
    if tracked:
        mark = lambda text: f"<w:del {writer._attrs('del', text)}/>"  # noqa: E731
        span_start = prior.start if trailing else doomed[0].start
        pieces: list[str] = []
        if trailing:
            prior_ppr = prior.child("w:pPr")
            ppr_xml = _ppr_with_mark(xml[prior_ppr.start:prior_ppr.end] if prior_ppr is not None else "", mark("¶"))
            content = "" if prior.closed else xml[prior_ppr.end if prior_ppr is not None else prior.head_end:
                                                  prior.close_start()]
            pieces.append(_paragraph(xml, prior, ppr_xml, content))
        for number, (node, text) in enumerate(zip(doomed, texts_before)):
            ppr = node.child("w:pPr")
            raw = xml[ppr.start:ppr.end] if ppr is not None else ""
            if trailing and number == len(doomed) - 1:
                if fixed_parts is not None:
                    theirs, kept, own = fixed_parts
                    change = f"<w:pPrChange {writer._attrs('pPrChange', '')}>" + (
                        f"<w:pPr>{own}</w:pPr>" if own else "<w:pPr/>") + "</w:pPrChange>"
                    raw = f"<w:pPr>{theirs}{kept}{change}</w:pPr>"
                pieces.append(_paragraph(xml, node, raw, _deleted_content(xml, node, writer, text,
                                                                          keep_deletions=keep_deletions)))
                continue
            ppr_xml = _ppr_with_mark(raw, mark("¶"))
            pieces.append(_paragraph(xml, node, ppr_xml, _deleted_content(xml, node, writer, text,
                                                                          keep_deletions=keep_deletions)))
        replacement = xml[span_start:span_start]
        cursor = span_start
        for node, piece in zip(([prior] if trailing else []) + doomed, pieces):
            replacement += xml[cursor:node.start] + piece
            cursor = node.end
        new_xml = xml[:span_start] + replacement + xml[cursor:]
    elif trailing:
        if fixed_parts is not None:
            theirs, kept, _ = fixed_parts
            ppr_xml = f"<w:pPr>{theirs}{kept}</w:pPr>"
        else:
            ppr_xml = xml[last_ppr.start:last_ppr.end] if last_ppr is not None else ""
        new_xml = xml[:prior.start] + _with_ppr(xml, prior, ppr_xml) + xml[last_node.end:]
    else:
        new_xml = xml[:doomed[0].start] + xml[last_node.end:]

    # -- decision 12, read from the input and the result only ---------------------------------------------------
    delta = len(new_xml) - len(xml)
    if new_xml[:container.start] != xml[:container.start] or new_xml[container.end + delta:] != xml[container.end:]:
        raise VerificationFailed(f"{op}: bytes outside the paragraphs' container changed")
    old_box, new_box = xml[container.start:container.end], new_xml[container.start:container.end + delta]
    old_tree, new_tree = parse(old_box), parse(new_box)

    def is_rev(item: Node) -> bool:
        return item.local in REV_TAGS and item.name.startswith("w:")

    def fresh(item: Node) -> bool:
        number = item.attr(new_box, "w:id") or ""
        return tracked and number.isdigit() and int(number) >= first_id and item.attr(new_box, "w:author") == author

    old_revs = Counter(_revision_raw(old_box, item) for item in old_tree.iter() if is_rev(item))
    kept_revs = Counter(_revision_raw(new_box, item) for item in new_tree.iter() if is_rev(item) and not fresh(item))
    if kept_revs != old_revs:
        raise VerificationFailed(f"{op}: a pre-existing revision changed")
    added_nodes = [item for item in new_tree.iter() if is_rev(item) and fresh(item)]
    if len(added_nodes) != len(writer.added) or any(item.name not in ("w:del", "w:pPrChange") for item in added_nodes):
        raise VerificationFailed(f"{op}: the new revisions are not the paragraphs' deletions")
    _check_stamps(new_box, added_nodes, author, op)
    new_ids = {item.attr(new_box, "w:id") for item in added_nodes}
    old_el, new_el = _fragment(xml, old_box), _fragment(new_xml, new_box)
    is_new = lambda el: tracked and el.get(W + "id") in new_ids and el.get(W + "author") == author  # noqa: E731
    accepted = new_el
    if tracked:
        rejected = copy.deepcopy(new_el)
        et_resolve(rejected, is_new, "reject")
        if _normalized(rejected) != _normalized(old_el):
            raise VerificationFailed(f"{op}: rejecting the deletion does not give back the input (D5)")
        accepted = copy.deepcopy(new_el)
        et_resolve(accepted, is_new, "accept")
    expected = copy.deepcopy(old_el)
    children = list(expected)
    kill = children[start_at:start_at + len(doomed)]
    if trailing:
        keep, mark_of = children[start_at - 1], children[start_at + len(doomed) - 1]
        old_ppr = keep.find(W + "pPr")
        if old_ppr is not None:
            keep.remove(old_ppr)
        surviving = mark_of.find(W + "pPr")
        if fixed_parts is not None:
            surviving = copy.deepcopy(old_ppr) if old_ppr is not None else ET.Element(W + "pPr")
            for item in list(surviving):
                if item.tag in (W + "rPr", W + "sectPr", W + "pPrChange"):
                    surviving.remove(item)
            own_ppr = mark_of.find(W + "pPr")
            mark_rpr = own_ppr.find(W + "rPr") if own_ppr is not None else None
            if mark_rpr is not None:
                surviving.append(copy.deepcopy(mark_rpr))
        if surviving is not None:
            keep.insert(0, copy.deepcopy(surviving))
    if keep_deletions:
        # Accept of the new revisions merges each emptied paragraph into the next one, carrying what stays in
        # it: its existing w:del wrappers and proofing marks with no run beside them, in document order.
        leftovers: list[ET.Element] = []
        for paragraph in kill:
            stretch: list[ET.Element] = []
            for child in [*paragraph, None]:
                if child is None or child.tag == W + "del":
                    if stretch and not any(item.tag == W + "r" for item in stretch):
                        leftovers.extend(stretch)
                    stretch = []
                    if child is not None:
                        leftovers.append(child)
                elif child.tag != W + "pPr":
                    stretch.append(child)
        if leftovers:
            if trailing:
                keep.extend(copy.deepcopy(item) for item in leftovers)
            else:
                rest = children[start_at + len(doomed):]
                target = next(item for item in rest if item.tag == W + "p")
                at = 1 if len(target) and target[0].tag == W + "pPr" else 0
                for offset, item in enumerate(leftovers):
                    target.insert(at + offset, copy.deepcopy(item))
    for item in kill:
        expected.remove(item)
    if _normalized(accepted) != _normalized(expected):
        raise VerificationFailed(f"{op}: the result does not read as the input with paragraphs {first}..{last} "
                                 "deleted" + (" (Accept view)" if tracked else ""))

    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, op)
        got = [item.text for item in read_layout(output).paragraphs]
        want = [item.text for item in doc.layout.paragraphs]
        if tracked:
            want[first - 1:last] = [""] * len(doomed)
        else:
            del want[first - 1:last]
        if got != want:
            raise VerificationFailed(f"{op}: the document does not read as expected after the deletion",
                                     expected=want[max(first - 2, 0):last + 1], actual=got[max(first - 2, 0):last + 1])
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    delta_paragraphs = 0 if tracked else -len(doomed)
    return _result(
        source, output, op, doc,
        source_sha256=file_sha256(source),
        paragraph_indices=list(range(first, last + 1)),
        deleted_texts=texts_before,
        trailing=trailing,
        fallback_trailing=fallback,
        previous_paragraph_index=previous if trailing else None,
        format_restored=fixed_parts is not None,
        index_delta=delta_paragraphs,
        indices_after_shifted=delta_paragraphs != 0,
        tracked=tracked,
        author=author if tracked else None,
        revisions_added=writer.added if tracked else [],
    )


# --------------------------------------------------------------------------
# Plan 864 S3b: table rows added or removed as direct XML (decision 11), tracked as Word tracks them.
# --------------------------------------------------------------------------

_DELETE_ROWS_OP = "delete_rows"
_INSERT_ROWS_OP = "insert_rows"
_ROW_MERGES = ("w:vMerge", "w:gridSpan", "w:hMerge")
# Anchors a copied row leaves behind (a copy never duplicates a bookmark name or a comment's range).
_COPY_STRIP = frozenset({"w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd",
                         "w:permStart", "w:permEnd"})
_ET_COPY_STRIP = frozenset(W + name[2:] for name in _COPY_STRIP)
_PARA_ID = re.compile(r"""\s+[\w.\-]+:(?:paraId|textId)\s*=\s*(?:"[^"]*"|'[^']*')""")


def _table_row(doc: Document, index: int, op: str, role: str):
    """The layout Row holding paragraph `index` (a row-end marker counts) and its raw w:tr."""

    item = doc.layout.para(index)
    row = doc.layout.row_of(index)
    if row is None:
        where = "inside a nested table" if item.nested else "not in a table"
        raise RefusedUnsafe(f"{op}: {role} paragraph {index} is {where}; give a paragraph of a row of a top-level "
                            "table")
    tables = [node for node in doc.body.children if node.name == "w:tbl"]
    node = [item for item in tables[row.table].children if item.name == "w:tr"][row.row]
    return row, node


def _row_name(row) -> str:
    return f"table {row.table + 1} row {row.row + 1} (paragraphs {row.first_index}-{row.end_index})"


def _comment_reference_only(run: Node) -> bool:
    return run.child("w:commentReference") is not None and all(
        item.name in ("w:rPr", "w:commentReference") for item in run.children)


def _refuse_row(op: str, xml: str, node: Node, label: str, *, copy_of: bool) -> None:
    """Refused, naming the row, the object and the way forward, unless the row is plain table text: no
    revision anywhere in it, no nested table, no merged cells, and only plain text runs in its paragraphs.
    A row to copy (`copy_of`) may hold bookmark, comment and permission anchors: the copy leaves them out."""

    way = ("add the row in Word, or use a row without it as the template" if copy_of
           else "delete the row in Word, or keep it")
    verb = "copying" if copy_of else "deleting"
    for item in node.iter():
        if item.name.startswith("w:") and (item.local in REV_TAGS or item.local in MOVE_RANGES):
            what = describe_wrapper(xml, item)
            raise RefusedUnsafe(f"{op}: {label} holds a tracked revision ({what}); accept or reject it first, "
                                "then edit again")
        if item.name in _ROW_MERGES:
            raise RefusedUnsafe(f"{op}: {label} has merged cells (<{item.name}>), which {verb} the row would "
                                f"break; {way}")
        if item.name == "w:tbl" and item is not node:
            raise RefusedUnsafe(f"{op}: {label} holds a nested table; {way}")

    def refuse(found: Node) -> None:
        hit = next((sub for sub in found.iter() if sub.name in _OBJECT_NAMES), found)
        what = _OBJECT_NAMES.get(hit.name, f"<{hit.name}>")
        raise RefusedUnsafe(f"{op}: {label} holds {what}, which {verb} the row would "
                            f"{'duplicate' if copy_of else 'remove'}; {way}")

    for child in node.children:
        if child.name in ("w:tblPrEx", "w:trPr") or (copy_of and child.name in _COPY_STRIP):
            continue
        if child.name != "w:tc":
            refuse(child)
        for item in child.children:
            if item.name == "w:tcPr" or (copy_of and item.name in _COPY_STRIP):
                continue
            if item.name != "w:p":
                refuse(item)
            for part in item.children:
                if part.name in ("w:pPr", "w:proofErr") or (copy_of and part.name in _COPY_STRIP):
                    continue
                if part.name != "w:r":
                    refuse(part)
                if copy_of and _comment_reference_only(part):
                    continue
                bad = next((sub for sub in part.children if sub.name not in _DELETABLE_RUN_PARTS
                            and not (copy_of and sub.name == "w:commentReference")), None)
                if bad is not None:
                    refuse(bad)


def _clean_row_copy(xml: str, node: Node) -> tuple[str, list[str]]:
    """The row's bytes without its anchors (bookmarks, comment ranges and references, permissions) and without
    w14:paraId / w14:textId on its rows and paragraphs (Word gives a copy new ones)."""

    edits: list[tuple[int, int, str]] = []
    stripped: list[str] = []
    skip: set[int] = set()
    for item in node.iter():
        if id(item) in skip:
            continue
        if item.name in ("w:tr", "w:p"):
            head = item.head(xml)
            cleaned = _PARA_ID.sub("", head)
            if cleaned != head:
                edits.append((item.start, item.head_end, cleaned))
        elif item.name in _COPY_STRIP:
            edits.append((item.start, item.end, ""))
            name = item.attr(xml, "w:name")
            stripped.append(f"<{item.name} w:id={item.attr(xml, 'w:id')}" + (f" w:name={name}>" if name else ">"))
        elif item.name == "w:r" and item.child("w:commentReference") is not None:
            targets = [item] if _comment_reference_only(item) else [
                sub for sub in item.children if sub.name == "w:commentReference"]
            for target in targets:
                edits.append((target.start, target.end, ""))
                skip.update(id(sub) for sub in target.iter())
            reference = item.child("w:commentReference")
            stripped.append(f"<w:commentReference w:id={reference.attr(xml, 'w:id')}>")
    out, cursor = [], node.start
    for start, end, text in sorted(edits):
        if start < cursor:
            raise VerificationFailed(f"{_INSERT_ROWS_OP}: overlapping edits while cleaning the row copy")
        out.append(xml[cursor:start] + text)
        cursor = end
    out.append(xml[cursor:node.end])
    return "".join(out), stripped


def _plain_text(xml: str, node: Node) -> str:
    out: list[str] = []
    for item in node.iter():
        if item.name in ("w:t", "w:delText") and not item.closed:
            out.append(unescape(item.inner(xml)))
        elif item.name == "w:tab" and item.parent is not None and item.parent.name == "w:r":
            out.append("\t")
    return "".join(out)


def _track_row(row_xml: str, kind: str, writer: _Writer) -> str:
    """One row (its own bytes) as Word tracks a row deletion (`kind` "del") or insertion ("ins"): the mark in
    w:trPr (created after w:tblPrEx when missing; last before a w:trPrChange), every paragraph mark marked,
    and each paragraph's content in one wrapper (deleted text as w:delText). Ids ascend in document order:
    the row mark, then per paragraph its mark and its content. An empty paragraph gets only its mark."""

    tree = parse(row_xml)
    row = tree.children[0]
    edits: list[tuple[int, int, str]] = []
    mark = f"<w:{kind} {writer._attrs(kind, '')}/>"
    trpr = row.child("w:trPr")
    if trpr is None:
        prex = row.child("w:tblPrEx")
        at = prex.end if prex is not None else row.head_end
        edits.append((at, at, f"<w:trPr>{mark}</w:trPr>"))
    elif trpr.closed:
        edits.append((trpr.start, trpr.end, f"<w:trPr>{mark}</w:trPr>"))
    else:
        change = trpr.child("w:trPrChange")
        at = change.start if change is not None else trpr.close_start()
        edits.append((at, at, mark))
    for cell in row.children:
        if cell.name != "w:tc":
            continue
        for paragraph in cell.children:
            if paragraph.name != "w:p":
                continue
            text = _plain_text(row_xml, paragraph)
            ppr = paragraph.child("w:pPr")
            ppr_xml = _ppr_with_mark(row_xml[ppr.start:ppr.end] if ppr is not None else "",
                                     f"<w:{kind} {writer._attrs(kind, '¶')}/>")
            if kind == "del":
                content = _deleted_content(row_xml, paragraph, writer, text)
            else:
                body = "" if paragraph.closed else row_xml[ppr.end if ppr is not None else paragraph.head_end:
                                                           paragraph.close_start()]
                content = f"<w:ins {writer._attrs('ins', text)}>{body}</w:ins>" if any(
                    item.name == "w:r" for item in paragraph.children) else body
            edits.append((paragraph.start, paragraph.end, _paragraph(row_xml, paragraph, ppr_xml, content)))
    out, cursor = [], 0
    for start, end, text in edits:
        out.append(row_xml[cursor:start] + text)
        cursor = end
    out.append(row_xml[cursor:])
    return "".join(out)


def _verify_row_box(op: str, xml: str, new_xml: str, lo: int, hi: int, *, first_id: int, author: str,
                    writer: _Writer, tracked: bool, kind: str):
    """Decision 12 for a row edit, read from the input and the result only: bytes outside [lo, hi) identical,
    every pre-existing revision byte-identical, the new revisions exactly the writer's (all `kind`, ids
    ascending, author and date), and Reject of them = the input. Returns (old body, new body, is_new)."""

    delta = len(new_xml) - len(xml)
    if new_xml[:lo] != xml[:lo] or new_xml[hi + delta:] != xml[hi:]:
        raise VerificationFailed(f"{op}: bytes outside the table changed")
    old_box, new_box = xml[lo:hi], new_xml[lo:hi + delta]
    old_tree, new_tree = parse(old_box), parse(new_box)

    def is_rev(item: Node) -> bool:
        return item.local in REV_TAGS and item.name.startswith("w:")

    def fresh(item: Node) -> bool:
        number = item.attr(new_box, "w:id") or ""
        return tracked and number.isdigit() and int(number) >= first_id and item.attr(new_box, "w:author") == author

    old_revs = Counter(_revision_raw(old_box, item) for item in old_tree.iter() if is_rev(item))
    kept_revs = Counter(_revision_raw(new_box, item) for item in new_tree.iter() if is_rev(item) and not fresh(item))
    if kept_revs != old_revs:
        raise VerificationFailed(f"{op}: a pre-existing revision changed")
    added_nodes = [item for item in new_tree.iter() if is_rev(item) and fresh(item)]
    if len(added_nodes) != len(writer.added) or any(item.name != f"w:{kind}" for item in added_nodes):
        raise VerificationFailed(f"{op}: the new revisions are not the rows' {'deletions' if kind == 'del' else 'insertions'}")
    _check_stamps(new_box, added_nodes, author, op)
    new_ids = {item.attr(new_box, "w:id") for item in added_nodes}
    old_el = _fragment(xml, f"<w:body>{old_box}</w:body>")
    new_el = _fragment(new_xml, f"<w:body>{new_box}</w:body>")
    is_new = lambda el: tracked and el.get(W + "id") in new_ids and el.get(W + "author") == author  # noqa: E731
    if tracked:
        rejected = copy.deepcopy(new_el)
        et_resolve(rejected, is_new, "reject")
        if _row_normalized(rejected) != _row_normalized(old_el):
            raise VerificationFailed(f"{op}: rejecting the new revisions does not give back the input")
    return old_el, new_el, is_new


def _row_normalized(element: ET.Element) -> str:
    """`_normalized`, with an empty w:trPr read as absent too: rejecting a row deletion drops the trPr that
    held only its mark (as Word does), so an input `<w:trPr/>` comes back absent."""

    clone = copy.deepcopy(element)
    for row in clone.iter(W + "tr"):
        trpr = row.find(W + "trPr")
        if trpr is not None and not len(trpr) and not trpr.attrib:
            row.remove(trpr)
    return _normalized(clone)


def _accepted_view(op: str, new_el: ET.Element, is_new, tracked: bool, expected: ET.Element, what: str) -> None:
    accepted = new_el
    if tracked:
        accepted = copy.deepcopy(new_el)
        et_resolve(accepted, is_new, "accept")
    if _row_normalized(accepted) != _row_normalized(expected):
        raise VerificationFailed(f"{op}: the result does not read as the input with {what}"
                                 + (" (Accept view)" if tracked else ""))


def _check_layout(op: str, output, want: list[str], at: int) -> None:
    from .sections import read_layout

    got = [item.text for item in read_layout(output).paragraphs]
    if got != want:
        raise VerificationFailed(f"{op}: the document does not read as expected after the row edit",
                                 expected=want[max(at - 3, 0):at + 6], actual=got[max(at - 3, 0):at + 6])


def delete_rows(source, output, *, first_indices, tracked: bool, author: str = DEFAULT_AUTHOR,
                now: dt.datetime | None = None) -> dict:
    """Delete the top-level table rows holding the paragraphs `first_indices` (any paragraph of a row, its
    row-end marker included) as direct XML (decision 11).

    Tracked, as Word records a row deleted with Track Changes on: `<w:del/>` in the row's w:trPr, every
    paragraph mark of every cell deleted (the last of each cell included) and each paragraph's runs in one
    w:del by `author` (w:t -> w:delText); ids ascend in document order (row mark, then per paragraph its mark
    and its runs). Accept removes the row; Reject gives it back. Untracked: the w:tr is removed.

    Refused, naming the row, the object and the way forward: any revision in the row (another author's, or
    the row already tracked), a nested table, merged cells (vMerge, gridSpan, hMerge), anything but plain text
    runs (a field, bookmark, comment, hyperlink, content control, drawing, footnote), and deleting every row
    of a table. Verified before anything is published: bytes outside the tables identical, every pre-existing
    revision byte-identical, new ids fresh and ascending by `author`, Reject of the new revisions = the input,
    Accept (untracked: the result) = the input without the rows, and the written layout reads so.
    """

    from .audit import file_sha256

    op = _DELETE_ROWS_OP
    author = (author or "").strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{op}: author must not be empty")
    parts, doc = load(source)
    xml = doc.xml
    found: dict[tuple[int, int], tuple] = {}
    for index in first_indices:
        row, node = _table_row(doc, index, op, "row")
        found.setdefault((row.table, row.row), (row, node))
    if not found:
        raise RefusedUnsafe(f"{op}: no rows given")
    rows = [found[key] for key in sorted(found)]
    for number in sorted({row.table for row, _ in rows}):
        doomed = sum(1 for row, _ in rows if row.table == number)
        if doomed == len(doc.layout.tables[number]):
            what = "its only row" if doomed == 1 else f"all {doomed} of its rows"
            raise RefusedUnsafe(f"{op}: deleting {what} would leave table {number + 1} with no rows; a table keeps "
                                "at least one row: delete the table in Word")
    for row, node in rows:
        _refuse_row(op, xml, node, _row_name(row), copy_of=False)

    first_id = next_revision_id(parts)
    writer = _Writer(first_id, author, revision_stamp(xml, now), tracked)
    texts = [[doc.layout.para(index).text for index in range(row.first_index, row.end_index)] for row, _ in rows]
    pieces, cursor = [], 0
    for row, node in rows:
        replacement = _track_row(xml[node.start:node.end], "del", writer) if tracked else ""
        pieces.append(xml[cursor:node.start] + replacement)
        cursor = node.end
    pieces.append(xml[cursor:])
    new_xml = "".join(pieces)

    tables = [node for node in doc.body.children if node.name == "w:tbl"]
    involved = sorted({row.table for row, _ in rows})
    lo, hi = tables[involved[0]].start, tables[involved[-1]].end
    old_el, new_el, is_new = _verify_row_box(op, xml, new_xml, lo, hi, first_id=first_id, author=author,
                                             writer=writer, tracked=tracked, kind="del")
    expected = copy.deepcopy(old_el)
    et_tables = [item for item in expected if item.tag == W + "tbl"]
    in_box = [number for number, node in enumerate(tables) if lo <= node.start and node.end <= hi]
    if len(et_tables) != len(in_box):
        raise VerificationFailed(f"{op}: the tables could not be matched to the input")
    by_number = dict(zip(in_box, et_tables))
    for table_number in involved:
        element = by_number[table_number]
        trs = element.findall(W + "tr")
        for row, _ in rows:
            if row.table == table_number:
                element.remove(trs[row.row])
    _accepted_view(op, new_el, is_new, tracked, expected, "the rows deleted")

    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, op)
        want = [item.text for item in doc.layout.paragraphs]
        for row, _ in reversed(rows):
            if tracked:
                want[row.first_index - 1:row.end_index] = [""] * (row.end_index - row.first_index + 1)
            else:
                del want[row.first_index - 1:row.end_index]
        _check_layout(op, output, want, rows[0][0].first_index)
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    size = sum(row.end_index - row.first_index + 1 for row, _ in rows)
    delta = 0 if tracked else -size
    return _result(
        source, output, op, doc,
        source_sha256=file_sha256(source),
        rows_deleted=[{"row": _row_name(row), "paragraphs": list(range(row.first_index, row.end_index)),
                       "row_end_index": row.end_index, "text": text} for (row, _), text in zip(rows, texts)],
        index_delta=delta,
        indices_after_shifted=delta != 0,
        tracked=tracked,
        author=author if tracked else None,
        revisions_added=writer.added if tracked else [],
    )


def insert_rows(source, output, *, template: int, before: int | None = None, append_after: int | None = None,
                tracked: bool, author: str = DEFAULT_AUTHOR, now: dt.datetime | None = None) -> dict:
    """Add a copy of the top-level table row holding paragraph `template`, right before the row holding
    `before`, or after the table's last row (`append_after`, a paragraph of that row), as direct XML.

    The copy keeps the template's row, cell and paragraph properties and its text; it leaves out bookmarks,
    comment ranges and references and permission ranges (listed in `stripped`) and w14:paraId / w14:textId.
    Tracked, as Word records a row added with Track Changes on: `<w:ins/>` in the copy's w:trPr, every
    paragraph mark inserted and each paragraph's runs in one w:ins by `author` (an empty paragraph has only
    its mark); a later tracked text edit by the same author inside the copy is made in place (plan 831), so
    the copy stays one insertion. Untracked: a plain copy.

    Refused, naming the template row and the way forward: a revision in it, a nested table, merged cells,
    and any object but text runs and the anchors above (a field, hyperlink, content control, drawing,
    footnote). Verified as delete_rows: bytes outside the table identical, pre-existing revisions identical,
    new ids fresh and ascending, Reject = the input, Accept (untracked: the result) = the input with the
    copy in place, and the written layout reads so.
    """

    from .audit import file_sha256

    op = _INSERT_ROWS_OP
    author = (author or "").strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{op}: author must not be empty")
    if (before is None) == (append_after is None):
        raise RefusedUnsafe(f"{op}: give exactly one of before or append_after")
    parts, doc = load(source)
    xml = doc.xml
    source_row, source_node = _table_row(doc, template, op, "template")
    if before is not None:
        target, target_node = _table_row(doc, before, op, "before")
        at, new_first, position = target_node.start, target.first_index, target.row
    else:
        target, target_node = _table_row(doc, append_after, op, "append_after")
        if target.row != len(doc.layout.tables[target.table]) - 1:
            raise RefusedUnsafe(f"{op}: append_after names {_row_name(target)}, not the table's last row; give "
                                "before= for a row inside the table")
        at, new_first, position = target_node.end, target.end_index + 1, target.row + 1
    if target.table != source_row.table:
        raise RefusedUnsafe(f"{op}: the template is in table {source_row.table + 1} but the copy would go in table "
                            f"{target.table + 1}; a row is copied within its own table")
    label = _row_name(source_row)
    _refuse_row(op, xml, source_node, f"template {label}", copy_of=True)

    first_id = next_revision_id(parts)
    writer = _Writer(first_id, author, revision_stamp(xml, now), tracked)
    cleaned, stripped = _clean_row_copy(xml, source_node)
    piece = _track_row(cleaned, "ins", writer) if tracked else cleaned
    if _PARA_ID.search(piece):
        raise VerificationFailed(f"{op}: the copy still carries a paraId or textId")
    new_xml = xml[:at] + piece + xml[at:]

    table_node = [node for node in doc.body.children if node.name == "w:tbl"][target.table]
    old_el, new_el, is_new = _verify_row_box(op, xml, new_xml, table_node.start, table_node.end, first_id=first_id,
                                             author=author, writer=writer, tracked=tracked, kind="ins")
    expected = copy.deepcopy(old_el)
    element = next(item for item in expected if item.tag == W + "tbl")
    trs = element.findall(W + "tr")
    duplicate = copy.deepcopy(trs[source_row.row])
    for parent in list(duplicate.iter()):
        for child in list(parent):
            if child.tag in _ET_COPY_STRIP:
                parent.remove(child)
            elif child.tag == W + "r" and child.find(W + "commentReference") is not None:
                for reference in child.findall(W + "commentReference"):
                    child.remove(reference)
                if all(item.tag == W + "rPr" for item in child):
                    parent.remove(child)
    slot = list(element).index(trs[position]) if position < len(trs) else len(element)
    element.insert(slot, duplicate)
    _accepted_view(op, new_el, is_new, tracked, expected, f"a copy of {label} added")

    size = source_row.end_index - source_row.first_index + 1
    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, op)
        want = [item.text for item in doc.layout.paragraphs]
        copied = [doc.layout.para(index).text for index in range(source_row.first_index, source_row.end_index + 1)]
        want[new_first - 1:new_first - 1] = copied
        _check_layout(op, output, want, new_first)
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    return _result(
        source, output, op, doc,
        source_sha256=file_sha256(source),
        template=label,
        new_row={"table": target.table + 1, "row": position + 1, "first_index": new_first,
                 "end_index": new_first + size - 1, "text": copied[:-1]},
        stripped=stripped,
        index_delta=size,
        indices_after_shifted=True,
        tracked=tracked,
        author=author if tracked else None,
        revisions_added=writer.added if tracked else [],
    )
