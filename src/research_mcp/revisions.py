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
_INSERTED = {"w:ins", "w:moveTo"}
_PASS_THROUGH = {"w:hyperlink": "h", "w:smartTag": "", "w:customXml": "", "w:dir": "", "w:bdo": "",
                 "w:fldSimple": "f", "w:sdt": "s"}
_MARKS = {"w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd",
          "w:commentReference", "w:proofErr", "w:permStart", "w:permEnd", "w:lastRenderedPageBreak"} | {
    "w:" + name for name in MOVE_RANGES}
_ANCHOR_MARKS = ("w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd",
                 "w:commentReference")
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
        return self.child.name != "w:delText"

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

    def __init__(self, doc: Document, index: int, op: str, *, tracked: bool) -> None:
        self.doc, self.index, self.op, self.tracked = doc, index, op, tracked
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
        return edits

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

    def _check_char(self, char: Char, action: str) -> None:
        if char.field:
            raise RefusedUnsafe(
                f"{self.op}: paragraph {self.index}: the edit would {action} field text {char.ch!r}; "
                "fields (citations, cross-references) are never edited as text")
        revision = char.revision
        if revision is not None:
            raise self._blocked(f"the edit would {action} text inside", revision)
        change = re.search(r"<w:rPrChange\b[^>]*>", char.rpr)
        if change and not self.tracked:
            raise RefusedUnsafe(
                f"{self.op}: paragraph {self.index}: the edit would {action} text whose formatting change "
                f"{_head_describe(change.group(0))} is tracked; use the tracked operations or resolve it first")

    # -- application -------------------------------------------------------

    def apply(self, edits: list, *, author: str, first_id: int, stamp: str) -> tuple[str, list[dict]]:
        final = self.model.final
        deleted: set[int] = set()
        cuts: dict[tuple[int, int], list[_Insertion]] = {}
        points: dict[int, list[_Insertion]] = {}
        split_runs: set[int] = set()
        for order, edit in enumerate(edits):
            removed = final[edit.start:edit.end]
            for char in removed:
                self._check_char(char, "delete")
                deleted.add(id(char))
            if removed:
                self._check_between(removed[0], removed[-1])
            if not edit.text:
                continue
            source = final[edit.source] if edit.source is not None else None
            rpr = strip_rpr_change(source.rpr) if source is not None else ""
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
            self._place(where, insertion, cuts, points, split_runs)
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
        touched = {id(char.run): char.run for char in self.model.chars
                   if id(char) in deleted or id(char.run) in split_runs}
        events: list[tuple[int, int, int, object]] = [(run.start, 1, 0, run) for run in touched.values()]
        for offset, items in points.items():
            for item in items:
                events.append((offset, 0, item.order, item))
        events.sort(key=lambda event: (event[0], event[1], event[2]))
        out, cursor = [], self.node.start
        for offset, kind, _, payload in events:
            out.append(self.xml[cursor:offset])
            if kind == 0:
                out.append(writer.insertion(payload))
                cursor = offset
            else:
                out.append(self._rebuild(payload, deleted, cuts, writer))
                cursor = payload.end
        out.append(self.xml[cursor:self.node.end])
        return "".join(out), writer.added

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

    def _place(self, where, insertion: _Insertion, cuts, points, split_runs) -> None:
        side, char = where
        if side == "empty":
            ppr = self.node.child("w:pPr")
            points.setdefault(ppr.end if ppr is not None else self.node.head_end, []).append(insertion)
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
        if at_edge:
            points.setdefault(self._hoist(char.run, side, insertion), []).append(insertion)
            return
        revision = char.revision
        if revision is not None:
            raise self._blocked(f"the new text {insertion.text!r} would land strictly inside", revision)
        split_runs.add(id(char.run))
        boundary = char.offset + 1 if side == "after" else char.offset
        cuts.setdefault((id(char.run), (id(char.child), boundary)), []).append(insertion)

    def _hoist(self, run: Node, side: str, insertion: _Insertion) -> int:
        """Offset just outside `run`, stepping out of revision wrappers whose edge it is."""

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
        return node.end if side == "after" else node.start

    def _rebuild(self, run: Node, deleted: set[int], cuts, writer: _Writer) -> str:
        xml = self.xml
        head = run.head(xml) if not run.closed else "<w:r>"
        if run.closed:
            return xml[run.start:run.end]
        rpr_node = run.child("w:rPr")
        rpr = xml[rpr_node.start:rpr_node.end] if rpr_node is not None else ""
        chars = {(id(char.child), char.offset): char for char in self.model.chars if char.run is run}
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
                    kind = "del" if id(char) in deleted else "keep"
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


def et_resolve(root: ET.Element, chosen, mode: str, drop=None) -> None:
    """Accept (mode "accept") or reject, in place, every revision element with chosen(el) true.

    Processed in reverse document order, so a revision nested in another is resolved before its container
    moves it. `drop(el)` names extra elements to remove first (the range markers of resolved moves). A
    removed wrapper leaves its bookmark/comment markers (and comment-reference runs) where it stood. A
    removed paragraph mark merges the paragraph into the next one, which keeps its own pPr (Word's rule).
    Raises on what the string engine refuses as well (cell revisions, row marks, impossible merges).
    """

    parents = {child: parent for parent in root.iter() for child in parent}
    if drop is not None:
        for el in [el for el in root.iter() if drop(el)]:
            parents[el].remove(el)
    order = {el: number for number, el in enumerate(root.iter())}
    targets = [el for el in root.iter() if el.tag.startswith(W) and _local(el.tag) in REV_TAGS and chosen(el)]
    merging: list[ET.Element] = []
    for el in reversed(targets):
        name, parent = _local(el.tag), parents[el]
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
    editor = TextEditor(doc, paragraph_index, op, tracked=tracked)
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
    old_revs = [_revision_raw(shift_old, item) for item in old_tree.iter() if item.local in REV_TAGS
                and item.name.startswith("w:")]
    kept, fresh = [], []
    for item in new_tree.iter():
        if item.local in REV_TAGS and item.name.startswith("w:"):
            if item.attr(shift_new, "w:id") in new_ids and item.attr(shift_new, "w:author") == author and \
                    item.local in ("ins", "del") and tracked:
                fresh.append(item)
            else:
                kept.append(_revision_raw(shift_new, item))
    if kept != old_revs:
        raise VerificationFailed(f"{op}: a pre-existing revision in paragraph {editor.index} changed (D4)")
    if tracked:
        if len(fresh) != len(added):
            raise VerificationFailed(f"{op}: new revision count mismatch")
        _check_stamps(shift_new, fresh, author, op)
    from .sections import canonical

    old_p, new_p = _fragment(xml, xml[node.start:node.end]), _fragment(xml, new_paragraph)
    is_new = lambda el: tracked and el.get(W + "id") in new_ids and el.get(W + "author") == author  # noqa: E731
    rejected = copy.deepcopy(new_p)
    et_resolve(rejected, is_new, "reject")
    if tracked and canonical(rejected) != canonical(old_p):
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
        if edit.end > edit.start or edit.mode == "before":
            at = position[id(final[edit.start])]
        elif edit.mode == "after":
            at = position[id(final[edit.start - 1])] + 1
        else:
            at = 0
        stream.append(StreamEdit(base + at + delta, edit.end - edit.start, len(edit.text)))
        delta += len(edit.text) - (edit.end - edit.start)
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
    if revision.mark == "row":
        raise RefusedUnsafe(f"{op}: {label} inserts or deletes a table row; resolve it in Word")
    if revision.mark == "paragraph" or name not in RUN_REVS:
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


# -- paragraph marks: a removed mark merges its paragraph into the next one ---------------------------------

def _plan_merges(doc: Document, selected: list[Revision], mode: str, op: str, index_of: dict[int, int],
                 refusals: list[tuple[Revision, str]] | None = None):
    """Paragraphs whose mark goes (each with the paragraph it joins and the markers between), and the mark
    revisions that only lose their marker. Refuses merges Word's result cannot be matched for: raises at the
    first, or, given ``refusals``, appends (revision, message) for each and plans the rest."""

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
        while at < len(siblings) and siblings[at].name in _BETWEEN:
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
    if node.name in ("w:rPr", "w:pPr") and node.children and not inner.strip():
        return ""  # every child was a resolved marker; an empty rPr/pPr trips Word's save check
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
    merges, _ = _plan_merges(doc, selected, mode, op, index_of, refusals)
    _raise_refusals(op, mode, len(selected), refusals)

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
    marks_before = [m.group(0) for m in re.finditer(r"<w:(?:%s)\b[^>]*>" % "|".join(
        name[2:] for name in _ANCHOR_MARKS), doc.xml)]
    marks_after = [m.group(0) for m in re.finditer(r"<w:(?:%s)\b[^>]*>" % "|".join(
        name[2:] for name in _ANCHOR_MARKS), new_xml)]
    if marks_before != marks_after:
        raise VerificationFailed(f"{op}: a bookmark or comment anchor was lost; nothing was written")
    rezip(source, output, new_xml)

    chains = _chains(merges, index_of)
    merged_away = sorted(index_of[id(paragraph)] for paragraph, _, _ in merges)
    merged = []
    if chains:
        _, after_doc = load(output)
        for chain in chains:
            new_index = chain[0] - sum(1 for index in merged_away if index < chain[0])
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
    shifts = {"index_delta": -len(merged_away), "indices_after_shifted": bool(merged_away),
              "index_shifts": [[index + 1, -1] for index in merged_away]}
    if merged_away:
        shifts["shift_base"] = merged_away[0] + 1
    return _result(
        source, output, op, doc,
        mode=mode,
        scope=list(span) if span else None,
        resolved=[revision.as_dict() for revision in selected],
        resolved_count=len(selected),
        pulled_in=[revision.as_dict() for revision in pulled],
        removed_with=[revision.as_dict() for revision in removed_with],
        merged=merged,
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
