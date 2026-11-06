"""Word comments as direct OOXML (plan 562 S1).

Add, reply, resolve, reopen and delete comment threads without Word, on documents with or without tracked
changes. Like ``revisions.py`` every part is edited as text: only the spans that belong to the comment being
changed are written; every other byte of every part, and every other zip entry, is copied unchanged in the
original entry order. New parts (``commentsExtended``, ``commentsIds``, ``commentsExtensible``, or
``comments`` itself) are appended with their content-type override and document relationship.

Parts, as Word writes them:

* ``word/comments.xml``: ``w:comment`` (w:id, author, local-time ``w:date`` with "Z", initials) holding
  paragraphs, each with a unique ``w14:paraId``. The LAST paragraph's paraId keys the comment in the
  other parts.
* ``word/commentsExtended.xml``: ``w15:commentEx`` paraId, ``paraIdParent`` (reply -> root) and ``done``.
* ``word/commentsIds.xml``: ``w16cid:commentId`` paraId -> ``durableId`` (the id ``inspect_document``
  reports).
* ``word/commentsExtensible.xml``: ``w16cex:commentExtensible`` durableId -> real UTC ``dateUtc``.
* ``word/document.xml``: ``w:commentRangeStart``/``w:commentRangeEnd`` around the anchor and a run holding
  ``w:commentReference``. A reply gets its own markers next to its thread's (Word's layout).

Anchors on revised documents: the anchor is matched in the paragraph's accepted view (deleted text is not
visible) and must be unique there. The markers go immediately around the anchor; an edge that falls at the
boundary of a ``w:ins``/``w:del`` is lifted outside it, otherwise the marker sits inside the insertion (as
Word itself does) and the revision keeps its id, author, date, text and formatting. The reference run is
always outside every revision wrapper, so rejecting an insertion never deletes a comment. An anchor that
lies only in deleted text, or cuts a field, is refused.

Every write is verified before it is saved: other threads' entries in every comment part are
byte-identical; the revision count is unchanged and every revision is byte-identical, with one exception:
an inserted revision (``w:ins`` or ``w:moveTo``) that holds an anchor edge differs only by the comment markers and the split of the run at that
edge (same id, author and date, same text and run formatting); ``word/document.xml`` with the changed
comments' markers removed has the same characters and run formatting everywhere; the thread read back
through ``comments.inspect_comments`` has the requested text, author, anchor and done flag.
"""

from __future__ import annotations

import datetime as dt
import io
import random
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape

from . import config
from .errors import RefusedUnsafe, VerificationFailed

DISPOSITION_LABELS = {
    "addressed": "Addressed",
    "rejected": "Rejected",
    "no_change": "No change",
    "needs_information": "Needs information",
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX_NS = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
_CT_PREFIX = "application/vnd.openxmlformats-officedocument.wordprocessingml."
_DOC_RELS = "word/_rels/document.xml.rels"
_CONTENT_TYPES = "[Content_Types].xml"
_TEXT_ID = "77777777"


@dataclass(frozen=True)
class _PartSpec:
    kind: str
    name: str
    rel_type: str
    content_type: str
    prefix: str
    uri: str
    root: str


_PARTS = {
    "comments": _PartSpec("comments", "word/comments.xml",
                          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                          _CT_PREFIX + "comments+xml", "w", W_NS, "comments"),
    "extended": _PartSpec("extended", "word/commentsExtended.xml",
                          "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                          _CT_PREFIX + "commentsExtended+xml", "w15", W15_NS, "commentsEx"),
    "ids": _PartSpec("ids", "word/commentsIds.xml",
                     "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
                     _CT_PREFIX + "commentsIds+xml", "w16cid", W16CID_NS, "commentsIds"),
    "extensible": _PartSpec("extensible", "word/commentsExtensible.xml",
                            "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
                            _CT_PREFIX + "commentsExtensible+xml", "w16cex", W16CEX_NS, "commentsExtensible"),
}
_NEW_ROOT = {
    "comments": f'<w:comments xmlns:mc="{MC_NS}" xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" mc:Ignorable="w14">',
    "extended": f'<w15:commentsEx xmlns:mc="{MC_NS}" xmlns:w14="{W14_NS}" xmlns:w15="{W15_NS}" '
                f'mc:Ignorable="w14 w15">',
    "ids": f'<w16cid:commentsIds xmlns:mc="{MC_NS}" xmlns:w16cid="{W16CID_NS}" mc:Ignorable="w16cid">',
    "extensible": f'<w16cex:commentsExtensible xmlns:mc="{MC_NS}" xmlns:w16cex="{W16CEX_NS}" '
                  f'mc:Ignorable="w16cex">',
}
_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
_HEX8 = re.compile(r"^[0-9A-Fa-f]{8}$")


# --------------------------------------------------------------------------
# Small text helpers.
# --------------------------------------------------------------------------

# Attribute values in either quote style and with any spacing around "=" (valid XML that a regex on
# `name="..."` misses: final reviews terra61/sol61).
_VALUE = r"""\s*=\s*(?:"(?P<d>[^"]*)"|'(?P<s>[^']*)')"""


def _attr_span(head: str, name: str) -> tuple[int, int, str] | None:
    """(start, end, raw value) of attribute ``name`` in a start tag, or None."""

    found = re.search(r"\s%s%s" % (re.escape(name), _VALUE), head)
    if not found:
        return None
    group = "d" if found.group("d") is not None else "s"
    return found.start(group), found.end(group), found.group(group)


def _attr(head: str, name: str) -> str | None:
    from .revisions import attr_value

    return attr_value(head, name)


def _prefixed_values(raw: str, local: str) -> list[str]:
    """Values of every ``<any prefix>:local`` attribute in ``raw`` (e.g. w14:paraId), either quote style."""

    # finditer, not findall: findall gives "" (not None) for the alternative that did not match, which
    # turned every single-quoted value into "" (recheck sol61 P2)
    found = re.finditer(r"\s[\w.\-]+:%s%s" % (re.escape(local), _VALUE), raw)
    return [match.group("d") if match.group("d") is not None else match.group("s") for match in found]


def _prefix_for(head: str, uri: str) -> str | None:
    for found in re.finditer(r"\sxmlns:([\w.\-]+)%s" % _VALUE, head):
        if (found.group("d") if found.group("d") is not None else found.group("s")) == uri:
            return found.group(1)
    return None


def _ensure_namespace(head: str, prefix: str, uri: str) -> str:
    """Return the root start tag with ``prefix`` bound to ``uri`` and listed in mc:Ignorable."""

    bound = _attr(head, f"xmlns:{prefix}")
    if bound is not None and bound != uri:
        raise RefusedUnsafe(f"the part binds prefix {prefix!r} to {bound!r}; cannot add {uri}")
    closing = "/>" if head.endswith("/>") else ">"
    body = head[: -len(closing)]
    if bound is None:
        body += f' xmlns:{prefix}="{uri}"'
    mc = _prefix_for(body, MC_NS)
    if mc is None:
        mc = "mc" if _attr(body, "xmlns:mc") is None else "mce"
        body += f' xmlns:{mc}="{MC_NS}"'
    ignorable = _attr_span(body, f"{mc}:Ignorable")
    if ignorable is None:
        body += f' {mc}:Ignorable="{prefix}"'
    elif prefix not in ignorable[2].split():
        value = (ignorable[2] + " " + prefix).strip()
        body = body[: ignorable[0]] + value + body[ignorable[1]:]
    return body + closing


def _attr_escape(value: str) -> str:
    return escape(value, {'"': "&quot;"})


def _raw_offset(inner: str, chars: int) -> int:
    """Raw offset in escaped element text that comes after ``chars`` characters."""

    position = count = 0
    while count < chars:
        if inner[position] == "&":
            position = inner.index(";", position) + 1
        else:
            position += 1
        count += 1
    return position


def _now() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)


def _local_stamp(now: dt.datetime) -> str:
    return now.astimezone().strftime("%Y-%m-%dT%H:%M:%SZ")


def _utc_stamp(now: dt.datetime) -> str:
    return now.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# --------------------------------------------------------------------------
# Package: parts, splices, new entries.
# --------------------------------------------------------------------------

class _Part:
    """One XML part edited by splices on its original text."""

    def __init__(self, name: str, xml: str, *, new: bool = False) -> None:
        from .revisions import parse

        self.name, self.xml, self.new = name, xml, new
        self.tree = parse(xml)
        self.root = next(node for node in self.tree.children)
        self.splices: list[tuple[int, int, str]] = []
        self.appended: list[str] = []
        self.new_head: str | None = None

    def head(self) -> str:
        return self.root.head(self.xml)

    def entries(self, local: str):
        return [node for node in self.root.children if node.local == local]

    def splice(self, start: int, end: int, text: str) -> None:
        self.splices.append((start, end, text))

    def append(self, text: str) -> None:
        self.appended.append(text)

    @property
    def dirty(self) -> bool:
        return bool(self.splices or self.appended or self.new_head or self.new)

    def render(self) -> str:
        xml = self.xml
        edits = list(self.splices)
        if self.appended:
            if self.root.closed:
                head = self.new_head or self.root.head(xml)
                body = head[:-2] + ">" + "".join(self.appended) + f"</{self.root.name}>"
                edits.append((self.root.start, self.root.end, body))
            else:
                edits.append((self.root.close_start(), self.root.close_start(), "".join(self.appended)))
        if self.new_head and not (self.appended and self.root.closed):
            edits.append((self.root.start, self.root.head_end, self.new_head))
        edits.sort(key=lambda item: (item[0], item[1]))
        for (start, end, _), (next_start, _, _) in zip(edits, edits[1:]):
            if end > next_start:
                raise VerificationFailed(f"{self.name}: overlapping edits")
        for start, end, text in reversed(edits):
            xml = xml[:start] + text + xml[end:]
        return xml


class _Package:
    def __init__(self, source: str | Path, op: str) -> None:
        from .revisions import read_parts

        self.source, self.op = Path(source), op
        self.parts = read_parts(source)
        if "word/document.xml" not in self.parts:
            raise RefusedUnsafe(f"{op}: package has no word/document.xml")
        self.text: dict[str, _Part] = {}
        self.added: list[str] = []
        self._locate()

    def _locate(self) -> None:
        rels = self.parts.get(_DOC_RELS, b"").decode("utf-8")
        tags = [match.group(0) for match in re.finditer(r"<Relationship\b[^>]*>", rels)]
        self.rel_ids = [value for value in (_attr(tag, "Id") for tag in tags) if value is not None]
        by_type = {spec.rel_type: spec for spec in _PARTS.values()}
        self.present: dict[str, bool] = {}
        seen: dict[str, str] = {}
        for tag in tags:
            kind = _attr(tag, "Type")
            spec = by_type.get(kind) if kind else None
            if spec is None:
                continue
            target = _attr(tag, "Target") or ""
            resolved = target[1:] if target.startswith("/") else "word/" + target.removeprefix("./")
            if resolved != spec.name or spec.kind in seen:
                raise RefusedUnsafe(f"{self.op}: {spec.root} part at non-standard location {target!r}; "
                                    "edit this document's comments in Word")
            seen[spec.kind] = resolved
        for kind, spec in _PARTS.items():
            in_zip = spec.name in self.parts
            if in_zip != (kind in seen):
                state = "has no relationship" if in_zip else "is referenced but missing"
                raise RefusedUnsafe(f"{self.op}: {spec.name} {state}; edit this document's comments in Word")
            self.present[kind] = in_zip

    def part(self, kind_or_name: str, *, create: bool = False) -> _Part | None:
        spec = _PARTS.get(kind_or_name)
        name = spec.name if spec else kind_or_name
        if name in self.text:
            return self.text[name]
        if name in self.parts:
            raw = self.parts[name]
            try:
                xml = raw.decode("utf-8")
            except UnicodeDecodeError as exc:
                raise RefusedUnsafe(f"{self.op}: {name} is not UTF-8") from exc
            self.text[name] = _Part(name, xml)
            return self.text[name]
        if not create or spec is None:
            return None
        root_head = _NEW_ROOT[spec.kind]
        xml = _XML_DECL + root_head + f"</{spec.prefix}:{spec.root}>"
        self.text[name] = _Part(name, xml, new=True)
        self.added.append(spec.kind)
        return self.text[name]

    def render(self) -> dict[str, bytes]:
        changed = {name: part.render().encode("utf-8") for name, part in self.text.items() if part.dirty}
        if self.added:
            types = self.parts[_CONTENT_TYPES].decode("utf-8")
            rels = self.parts[_DOC_RELS].decode("utf-8")
            used = set(self.rel_ids)
            number = max([int(m) for m in re.findall(r"rId(\d+)", " ".join(used))] + [0])
            type_lines, rel_lines = [], []
            for kind in self.added:
                spec = _PARTS[kind]
                part_name = "/" + spec.name
                if f'PartName="{part_name}"' not in types:
                    type_lines.append(f'<Override PartName="{part_name}" ContentType="{spec.content_type}"/>')
                number += 1
                while f"rId{number}" in used:
                    number += 1
                used.add(f"rId{number}")
                target = spec.name.split("/", 1)[1]
                rel_lines.append(f'<Relationship Id="rId{number}" Type="{spec.rel_type}" Target="{target}"/>')
            close = types.rindex("</Types>")
            changed[_CONTENT_TYPES] = (types[:close] + "".join(type_lines) + types[close:]).encode("utf-8")
            close = rels.rindex("</Relationships>")
            changed[_DOC_RELS] = (rels[:close] + "".join(rel_lines) + rels[close:]).encode("utf-8")
        return changed

    def write(self, output: str | Path, changed: dict[str, bytes]) -> None:
        output = Path(output)
        if output.exists():
            raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(self.source) as source_zip, zipfile.ZipFile(output, "x") as target:
            names = set()
            for info in source_zip.infolist():
                names.add(info.filename)
                data = changed.get(info.filename)
                if data is None:
                    data = source_zip.read(info.filename)
                target.writestr(info, data, compress_type=info.compress_type)
            for name, data in changed.items():
                if name not in names:
                    target.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)


# --------------------------------------------------------------------------
# Comment records.
# --------------------------------------------------------------------------

@dataclass
class _Record:
    legacy_id: str
    node: object  # revisions.Node in comments.xml
    para_id: str | None
    parent: str | None
    done: bool
    durable: str | None
    extended: object | None  # w15:commentEx node
    ids: object | None  # w16cid:commentId node
    extensible: object | None  # w16cex:commentExtensible node

    @property
    def comment_id(self) -> str:
        return self.durable or f"legacy:{self.legacy_id}"


class _Comments:
    """Records read from the raw parts, keyed the way ``comments.inspect_comments`` keys them."""

    def __init__(self, package: _Package) -> None:
        self.package = package
        self.comments = package.part("comments")
        self.records: list[_Record] = []
        if self.comments is None:
            return
        part = self.comments
        head = part.head()
        if _prefix_for(head, W_NS) != "w":
            raise RefusedUnsafe(f"{package.op}: word/comments.xml does not use the w: prefix")
        self.p14 = _prefix_for(head, W14_NS) or "w14"
        extended, ids, extensible = package.part("extended"), package.part("ids"), package.part("extensible")
        ext_by_para = {}
        if extended is not None:
            p15 = _prefix_for(extended.head(), W15_NS) or "w15"
            for node in extended.entries("commentEx"):
                para = node.attr(extended.xml, f"{p15}:paraId")
                if para:
                    ext_by_para[para] = (node, node.attr(extended.xml, f"{p15}:paraIdParent"),
                                         node.attr(extended.xml, f"{p15}:done") == "1")
        ids_by_para = {}
        if ids is not None:
            pcid = _prefix_for(ids.head(), W16CID_NS) or "w16cid"
            for node in ids.entries("commentId"):
                para, durable = node.attr(ids.xml, f"{pcid}:paraId"), node.attr(ids.xml, f"{pcid}:durableId")
                if para and durable:
                    ids_by_para[para] = (node, durable.upper())
        cex_by_durable = {}
        if extensible is not None:
            pcex = _prefix_for(extensible.head(), W16CEX_NS) or "w16cex"
            for node in extensible.entries("commentExtensible"):
                durable = node.attr(extensible.xml, f"{pcex}:durableId")
                if durable:
                    cex_by_durable[durable.upper()] = node
        for node in part.entries("comment"):
            paragraphs = [item for item in node.iter() if item.name == "w:p"]
            para_id = next((item.attr(part.xml, f"{self.p14}:paraId") for item in reversed(paragraphs)
                            if item.attr(part.xml, f"{self.p14}:paraId")), None)
            ext = ext_by_para.get(para_id or "")
            id_entry = ids_by_para.get(para_id or "")
            durable = id_entry[1] if id_entry else None
            self.records.append(_Record(
                legacy_id=node.attr(part.xml, "w:id") or "", node=node, para_id=para_id,
                parent=ext[1] if ext else None, done=bool(ext and ext[2]), durable=durable,
                extended=ext[0] if ext else None, ids=id_entry[0] if id_entry else None,
                extensible=cex_by_durable.get(durable) if durable else None))

    def find(self, comment_id: str) -> _Record:
        wanted = comment_id.strip()
        upper = wanted.upper()
        matches = [record for record in self.records if record.durable and record.durable == upper]
        if not matches:
            legacy = wanted[len("legacy:"):] if wanted.lower().startswith("legacy:") else wanted
            matches = [record for record in self.records if record.legacy_id == legacy]
        if len(matches) != 1:
            known = ", ".join(record.comment_id for record in self.roots()) or "none"
            what = "matched no comment" if not matches else f"matched {len(matches)} comments"
            raise RefusedUnsafe(f"{self.package.op}: comment_id {comment_id!r} {what}; "
                                f"inspect_document lists the thread ids ({known})")
        return matches[0]

    def roots(self) -> list[_Record]:
        return [record for record in self.records if not record.parent]

    def root_of(self, record: _Record) -> _Record | None:
        if not record.parent:
            return record
        return next((item for item in self.records if item.para_id == record.parent and not item.parent), None)

    def replies(self, root: _Record) -> list[_Record]:
        if not root.para_id:
            return []
        return [record for record in self.records if record.parent == root.para_id]

    def thread_root(self, comment_id: str) -> _Record:
        record = self.find(comment_id)
        if record.parent:
            root = self.root_of(record)
            name = root.comment_id if root else record.parent
            raise RefusedUnsafe(f"{self.package.op}: {comment_id!r} is a reply in thread {name!r}; "
                                f"use the thread id {name!r}")
        return record


# --------------------------------------------------------------------------
# Identity: paraIds, durable ids, comment ids.
# --------------------------------------------------------------------------

class _Ids:
    def __init__(self, package: _Package) -> None:
        self.para: set[str] = set()
        self.durable: set[str] = set()
        for name, data in package.parts.items():
            if not name.endswith(".xml"):
                continue
            text = data.decode("utf-8", "ignore")
            for local, into in (("paraId", self.para), ("paraIdParent", self.para), ("textId", self.para),
                                ("durableId", self.durable)):
                into.update(value.upper() for value in _prefixed_values(text, local) if _HEX8.match(value))
        from .revisions import next_revision_id

        self.next_comment = next_revision_id(package.parts)
        self._random = random.SystemRandom()

    def _fresh(self, used: set[str]) -> str:
        while True:
            value = f"{self._random.randrange(1, 0x7FFFFFFF):08X}"
            if value not in used and value != _TEXT_ID:
                used.add(value)
                return value

    def para_id(self) -> str:
        return self._fresh(self.para)

    def durable_id(self) -> str:
        return self._fresh(self.durable)

    def comment_id(self) -> str:
        value = self.next_comment
        self.next_comment += 1
        return str(value)


# --------------------------------------------------------------------------
# Building comment XML.
# --------------------------------------------------------------------------

def _styles(package: _Package) -> tuple[bool, bool]:
    styles = package.parts.get("word/styles.xml", b"").decode("utf-8", "ignore")
    return 'w:styleId="CommentText"' in styles, 'w:styleId="CommentReference"' in styles


def _paragraph_runs(line: str) -> str:
    out: list[str] = []
    for piece in re.split(r"([\t\v])", line):
        if piece == "\t":
            out.append("<w:r><w:tab/></w:r>")
        elif piece == "\v":
            out.append("<w:r><w:br/></w:r>")
        elif piece:
            out.append(f'<w:r><w:t xml:space="preserve">{escape(piece)}</w:t></w:r>')
    return "".join(out)


def _comment_xml(package: _Package, ids: _Ids, *, legacy_id: str, author: str, initials: str, text: str,
                 now: dt.datetime, p14: str) -> tuple[str, str]:
    """``(xml, key paraId)`` for a new ``w:comment``; the key is the last paragraph's paraId."""

    text_style, reference_style = _styles(package)
    ppr = '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' if text_style else ""
    rpr = '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' if reference_style else ""
    paragraphs, para_id = [], ""
    for number, line in enumerate(text.split("\n")):
        para_id = ids.para_id()
        mark = f"<w:r>{rpr}<w:annotationRef/></w:r>" if number == 0 else ""
        paragraphs.append(f'<w:p {p14}:paraId="{para_id}" {p14}:textId="{_TEXT_ID}">{ppr}{mark}'
                          f"{_paragraph_runs(line)}</w:p>")
    initials_attr = f' w:initials="{_attr_escape(initials)}"' if initials else ""
    xml = (f'<w:comment w:id="{legacy_id}" w:author="{_attr_escape(author)}" w:date="{_local_stamp(now)}"'
           f"{initials_attr}>{''.join(paragraphs)}</w:comment>")
    return xml, para_id


def _reference_run(package: _Package, legacy_id: str) -> str:
    _, reference_style = _styles(package)
    rpr = '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' if reference_style else ""
    return f'<w:r>{rpr}<w:commentReference w:id="{legacy_id}"/></w:r>'


class _Writer:
    """Adds entries to the three side parts with the prefixes each part already binds."""

    def __init__(self, package: _Package, comments: _Comments) -> None:
        self.package, self.comments = package, comments

    def _part(self, kind: str) -> tuple[_Part, str]:
        spec = _PARTS[kind]
        part = self.package.part(kind, create=True)
        prefix = _prefix_for(part.new_head or part.head(), spec.uri)
        if prefix is None:
            raise RefusedUnsafe(f"{self.package.op}: {spec.name} does not declare {spec.uri}")
        return part, prefix

    def extended(self, para_id: str, parent: str | None, done: bool) -> None:
        part, prefix = self._part("extended")
        parent_attr = f' {prefix}:paraIdParent="{parent}"' if parent else ""
        part.append(f'<{prefix}:commentEx {prefix}:paraId="{para_id}"{parent_attr} '
                    f'{prefix}:done="{1 if done else 0}"/>')

    def ids(self, para_id: str, durable: str) -> None:
        part, prefix = self._part("ids")
        part.append(f'<{prefix}:commentId {prefix}:paraId="{para_id}" {prefix}:durableId="{durable}"/>')

    def extensible(self, durable: str, now: dt.datetime) -> None:
        part, prefix = self._part("extensible")
        part.append(f'<{prefix}:commentExtensible {prefix}:durableId="{durable}" '
                    f'{prefix}:dateUtc="{_utc_stamp(now)}"/>')


def _comments_part_for_write(package: _Package) -> tuple[_Part, str]:
    part = package.part("comments", create=True)
    head = part.new_head or part.head()
    p14 = _prefix_for(head, W14_NS)
    if p14 is None:
        part.new_head = _ensure_namespace(head, "w14", W14_NS)
        p14 = "w14"
    return part, p14


def _set_attr(xml: str, node, name: str, value: str) -> tuple[int, int, str]:
    """Splice that sets attribute ``name`` on ``node``'s start tag."""

    head = node.head(xml)
    found = _attr_span(head, name)
    if found:
        new = head[: found[0]] + value + head[found[1]:]
    else:
        closing = "/>" if head.endswith("/>") else ">"
        new = head[: -len(closing)] + f' {name}="{value}"' + closing
    return node.start, node.head_end, new


def _ensure_identity(package: _Package, comments: _Comments, ids: _Ids, writer: _Writer,
                     record: _Record) -> str:
    """Give an older comment (no paraId) a paraId and durable id so it can be keyed; return its paraId."""

    if record.para_id:
        if not record.durable:
            durable = ids.durable_id()
            writer.ids(record.para_id, durable)
            record.durable = durable
        return record.para_id
    part, p14 = _comments_part_for_write(package)
    paragraphs = [item for item in record.node.iter() if item.name == "w:p"]
    if not paragraphs:
        raise RefusedUnsafe(f"{package.op}: comment {record.legacy_id} has no paragraph")
    para_id = ids.para_id()
    part.splice(*_set_attr(part.xml, paragraphs[-1], f"{p14}:paraId", para_id))
    record.para_id = para_id
    durable = ids.durable_id()
    writer.ids(para_id, durable)
    record.durable = durable
    return para_id


# --------------------------------------------------------------------------
# Document markers.
# --------------------------------------------------------------------------

_REVISION_WRAPPERS = {"w:ins", "w:del", "w:moveFrom", "w:moveTo"}


def _element_children(node) -> list:
    return [item for item in node.children if item.name not in ("w:rPr", "w:pPr")]


def _climb(node, side: str):
    """Lift a run-level boundary out of revision wrappers whose matching edge it is."""

    while node.parent is not None and node.parent.name in _REVISION_WRAPPERS:
        siblings = _element_children(node.parent)
        if not siblings or (siblings[0] is not node if side == "start" else siblings[-1] is not node):
            break
        node = node.parent
    return node


def _outermost_wrapper(node):
    found = None
    for item in node.ancestors():
        if item.name in _REVISION_WRAPPERS:
            found = item
        if item.name == "w:p":
            break
    return found


@dataclass
class _Edge:
    offset: int
    pre: str  # text before the marker (closes a split run)
    post: str  # text after the marker (reopens a split run)
    container: object  # node whose children the marker joins
    head_fix: tuple[int, int, str] | None = None


def _run_parts(xml: str, run) -> tuple[str, str]:
    rpr = run.child("w:rPr")
    return run.head(xml), (xml[rpr.start:rpr.end] if rpr is not None else "")


def _split_guard(xml: str, run, op: str) -> None:
    _, rpr = _run_parts(xml, run)
    if "<w:rPrChange" in rpr:
        tag = re.search(r"<w:rPrChange\b[^>]*>", rpr)
        change = _attr(tag.group(0), "w:id") if tag else None
        author = _attr(tag.group(0), "w:author") if tag else None
        raise RefusedUnsafe(
            f"{op}: the anchor edge falls inside a run with a tracked formatting change "
            f"(w:rPrChange id={change or '?'} by {author or '?'!r}); "
            "widen or narrow the anchor to that run's whole text")


def _tracked_run(xml: str, run) -> bool:
    return run is not None and "<w:rPrChange" in _run_parts(xml, run)[1]


def _safe_anchors(model, anchor: str) -> list[str]:
    """Anchors next to `anchor` whose edges fall on the boundaries of the tracked-formatting runs it cuts:
    widened to the whole run(s) first, then narrowed to exclude them. Only unique visible text is offered."""

    final, text = model.final, model.text
    start = text.find(anchor)
    if start < 0:
        return []
    stop = start + len(anchor)

    def run_span(index: int) -> tuple[int, int]:
        run, low, high = final[index].run, index, index
        while low > 0 and final[low - 1].run is run:
            low -= 1
        while high + 1 < len(final) and final[high + 1].run is run:
            high += 1
        return low, high + 1

    wide_start, narrow_start, wide_stop, narrow_stop = start, start, stop, stop
    if 0 < start and final[start - 1].run is final[start].run and _tracked_run(model.xml, final[start].run):
        wide_start, narrow_start = run_span(start)
    if stop < len(final) and final[stop].run is final[stop - 1].run and _tracked_run(model.xml, final[stop].run):
        narrow_stop, wide_stop = run_span(stop - 1)
    found: list[str] = []
    for low, high in ((wide_start, wide_stop), (narrow_start, narrow_stop)):
        candidate = text[low:high]
        if candidate.strip() == candidate and candidate and candidate != anchor and text.count(candidate) == 1 \
                and candidate not in found:
            found.append(candidate)
    return found


def _start_edge(xml: str, item, op: str) -> _Edge:
    from .revisions import Char

    run = item.run
    if run is None:  # fldSimple or another run-level object
        node = _climb(item.node, "start")
        return _Edge(node.start, "", "", node.parent)
    child = item.child if isinstance(item, Char) else item.node
    offset = item.offset if isinstance(item, Char) else 0
    content = _element_children(run)
    if offset == 0 and content and content[0] is child:
        node = _climb(run, "start")
        return _Edge(node.start, "", "", node.parent)
    _split_guard(xml, run, op)
    head, rpr = _run_parts(xml, run)
    if offset == 0:
        return _Edge(child.start, "</w:r>", head + rpr, run.parent)
    raw = _raw_offset(child.inner(xml), offset)
    fix = _preserve_fix(xml, child, raw, before=True)
    return _Edge(child.head_end + raw, "</w:t></w:r>", head + rpr + '<w:t xml:space="preserve">', run.parent, fix)


def _end_edge(xml: str, item, op: str) -> _Edge:
    from .revisions import Char

    run = item.run
    if run is None:
        node = _climb(item.node, "end")
        return _Edge(node.end, "", "", node.parent)
    child = item.child if isinstance(item, Char) else item.node
    is_text = isinstance(item, Char) and child.name == "w:t"
    length = len(_unescaped(xml, child)) if is_text else 1
    at_end = (item.offset + 1 == length) if isinstance(item, Char) else True
    content = _element_children(run)
    if at_end and content and content[-1] is child:
        node = _climb(run, "end")
        return _Edge(node.end, "", "", node.parent)
    _split_guard(xml, run, op)
    head, rpr = _run_parts(xml, run)
    if at_end:
        return _Edge(child.end, "</w:r>", head + rpr, run.parent)
    raw = _raw_offset(child.inner(xml), item.offset + 1)
    fix = _preserve_fix(xml, child, raw, before=False)
    return _Edge(child.head_end + raw, "</w:t></w:r>", head + rpr + '<w:t xml:space="preserve">', run.parent, fix)


def _unescaped(xml: str, child) -> str:
    from .revisions import unescape

    return "" if child.closed else unescape(child.inner(xml))


def _preserve_fix(xml: str, child, raw: int, *, before: bool) -> tuple[int, int, str] | None:
    """A split w:t whose first half ends (or starts) with a space needs xml:space="preserve"."""

    head = child.head(xml)
    if "xml:space=" in head:
        return None
    first = child.inner(xml)[:raw]
    if first and first[-1].isspace() or first[:1].isspace():
        return child.start, child.head_end, head[:-1] + ' xml:space="preserve">'
    return None


def _field_extents(model) -> list[tuple[int, int, object]]:
    """(first item, last item, fldSimple node or None) of every field that begins and ends inside the paragraph.

    A complex field (fldChar begin .. end) ends on its own end item. A w:fldSimple has no end item: its result
    runs are children of the container, so the container node is returned and its closing tag is the edge."""

    extents, stack = [], []
    for index, item in enumerate(model.items):
        if item.__class__.__name__ != "Obj" or item.kind != "field":
            continue
        if item.node.name == "w:fldSimple":
            last = index
            for later in range(index + 1, len(model.items)):
                if item.node in model.items[later].wrappers:
                    last = later
            extents.append((index, last, item.node))
        elif item.node.name == "w:fldChar":
            kind = item.node.attr(model.xml, "w:fldCharType")
            if kind == "begin":
                stack.append(index)
            elif kind == "end" and stack:
                extents.append((stack.pop(), index, None))
    return extents


def final_positions(model) -> set[int]:
    cached = getattr(model, "_final_positions", None)
    if cached is None:
        cached = {model.position[id(char)] for char in model.final}
        model._final_positions = cached
    return cached


def _anchor_span(doc, model, paragraph_index: int, anchor: str, op: str) -> tuple[int, int, object]:
    """Item indices (first, last) of the anchor in the accepted view, widened to whole fields, plus the
    w:fldSimple whose closing tag is the end edge when the anchor ends on a whole fldSimple result (else None)."""

    text = model.text
    starts = [match.start() for match in re.finditer("(?=%s)" % re.escape(anchor), text)]
    if not starts:
        if anchor in model.original or anchor in "".join(char.ch for char in model.chars):
            raise RefusedUnsafe(
                f"{op}: anchor {anchor!r} in paragraph {paragraph_index} lies in deleted (tracked) text, which "
                "is not visible once changes are accepted; anchor on visible text")
        raise RefusedUnsafe(f"{op}: anchor {anchor!r} was not found in paragraph {paragraph_index} "
                            f"(visible text: {text[:120]!r})")
    if len(starts) > 1:
        raise RefusedUnsafe(f"{op}: anchor {anchor!r} is not unique in paragraph {paragraph_index} "
                            f"({len(starts)} matches in the visible text); use a longer anchor")
    first_char = model.final[starts[0]]
    last_char = model.final[starts[0] + len(anchor) - 1]
    first, last = model.position[id(first_char)], model.position[id(last_char)]
    # A field the anchor touches must be covered whole (its visible result, at least): then the markers go
    # around the whole field (begin .. end). Covering only part of a result would cut the field.
    simple_end = None
    for begin, end, simple in _field_extents(model):
        visible = [index for index in range(begin, end + 1) if index in final_positions(model)]
        touched = [index for index in visible if first <= index <= last]
        if not touched:
            continue
        if len(touched) != len(visible):
            raise RefusedUnsafe(f"{op}: anchor {anchor!r} cuts a field (citation, cross-reference or similar); "
                                "anchor on text outside the field or on the whole field result")
        first = min(first, begin)
        if end >= last:
            last, simple_end = end, simple
    edges = [model.items[first]] + ([model.items[last]] if simple_end is None else [])
    for edge in edges:
        if getattr(edge, "field", False):
            raise RefusedUnsafe(f"{op}: anchor {anchor!r} cuts a field (citation, cross-reference or similar); "
                                "anchor on text outside the field or on the whole field result")
    return first, last, simple_end


def _anchor_edits(doc, paragraph_index: int, anchor: str, legacy_id: str, reference: str,
                  op: str) -> list[tuple[int, int, str]]:
    from .revisions import ParagraphModel, _field_views

    node = doc.paragraph(paragraph_index, op)
    model = ParagraphModel(doc.xml, node)
    first, last, simple_end = _anchor_span(doc, model, paragraph_index, anchor, op)
    try:
        start = _start_edge(doc.xml, model.items[first], op)
        if simple_end is not None:  # whole fldSimple result: the end marker goes after </w:fldSimple>
            node = _climb(simple_end, "end")
            end = _Edge(node.end, "", "", node.parent)
        else:
            end = _end_edge(doc.xml, model.items[last], op)
    except RefusedUnsafe as exc:
        if "tracked formatting change" not in str(exc):
            raise
        safe = _safe_anchors(model, anchor)
        if not safe:
            raise
        raise RefusedUnsafe(f"{exc}; anchors that start and end on that run's boundaries work, e.g. "
                            + " or ".join(repr(item) for item in safe)) from None
    fields = _field_views(doc)
    if fields.protected(start.offset) or fields.protected(end.offset):
        raise RefusedUnsafe(f"{op}: anchor {anchor!r} lies inside a field result that spans paragraphs "
                            "(a bibliography, table of contents or similar); comment outside the field")
    if start.offset > end.offset:
        raise VerificationFailed(f"{op}: anchor edges are out of order")
    edits: list[tuple[int, int, str]] = []
    start_marker = f'<w:commentRangeStart w:id="{legacy_id}"/>'
    end_marker = f'<w:commentRangeEnd w:id="{legacy_id}"/>'
    wrapper = _outermost_wrapper(end.container) if end.container.name in _REVISION_WRAPPERS or any(
        item.name in _REVISION_WRAPPERS for item in end.container.ancestors()) else None
    if end.container.name in _REVISION_WRAPPERS:
        outer = _outermost_wrapper(end.container)
        wrapper = outer if outer is not None else end.container
    if wrapper is None:
        end_text = end.pre + end_marker + reference + end.post
        edits.append((end.offset, end.offset, end_text))
    else:
        edits.append((end.offset, end.offset, end.pre + end_marker + end.post))
        edits.append((wrapper.end, wrapper.end, reference))
    edits.append((start.offset, start.offset, start.pre + start_marker + start.post))
    for fix in (start.head_fix, end.head_fix):
        if fix is not None and fix not in edits:
            edits.append(fix)
    return edits


def _apply_splices(xml: str, edits: list[tuple[int, int, str]]) -> str:
    # Insertions at one offset keep the order they were added in (end marker before reference, start after
    # an end at the same point would cross, so starts sort before ends at the same offset).
    ordered = sorted(enumerate(edits), key=lambda item: (item[1][0], item[1][1], item[0]))
    for (_, (_, end, _)), (_, (start, _, _)) in zip(ordered, ordered[1:]):
        if end > start:
            raise VerificationFailed("overlapping document edits")
    for _, (start, end, text) in reversed(ordered):
        xml = xml[:start] + text + xml[end:]
    return xml


def _markers(xml_root, legacy_ids: set[str], xml: str):
    starts, ends, references = [], [], []
    for node in xml_root.iter():
        if node.name in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
            if node.attr(xml, "w:id") in legacy_ids:
                {"w:commentRangeStart": starts, "w:commentRangeEnd": ends,
                 "w:commentReference": references}[node.name].append(node)
    return starts, ends, references


# --------------------------------------------------------------------------
# Verification.
# --------------------------------------------------------------------------

def _owner_map(package_parts: dict[str, bytes]) -> dict[str, dict[str, list[tuple[str, str]]]]:
    """Per comment part: owner key -> raw entries. Owners are legacy ids (via paraId/durableId)."""

    from .revisions import parse

    out: dict[str, dict[str, list[tuple[str, str]]]] = {}

    def entries(name: str):
        if name not in package_parts:
            return None, []
        xml = package_parts[name].decode("utf-8")
        root = next(iter(parse(xml).children))
        return xml, [(node, xml[node.start:node.end]) for node in root.children]

    xml, comments = entries("word/comments.xml")
    para_owner, durable_owner = {}, {}
    out["word/comments.xml"] = {}
    for node, raw in comments:
        legacy = node.attr(xml, "w:id") or "?"
        out["word/comments.xml"].setdefault(legacy, []).append(("comment", raw))
        for para in _prefixed_values(raw, "paraId"):
            if _HEX8.match(para):
                para_owner[para.upper()] = legacy
    xml, items = entries("word/commentsIds.xml")
    out["word/commentsIds.xml"] = {}
    for node, raw in items:
        para = next(iter(_prefixed_values(raw, "paraId")), None)
        durable = next(iter(_prefixed_values(raw, "durableId")), None)
        owner = para_owner.get(para.upper(), "orphan:" + raw) if para is not None else "orphan:" + raw
        if durable is not None:
            durable_owner[durable.upper()] = owner
        out["word/commentsIds.xml"].setdefault(owner, []).append(("id", raw))
    xml, items = entries("word/commentsExtended.xml")
    out["word/commentsExtended.xml"] = {}
    for node, raw in items:
        para = next(iter(_prefixed_values(raw, "paraId")), None)
        owner = para_owner.get(para.upper(), "orphan:" + raw) if para is not None else "orphan:" + raw
        out["word/commentsExtended.xml"].setdefault(owner, []).append(("ex", raw))
    xml, items = entries("word/commentsExtensible.xml")
    out["word/commentsExtensible.xml"] = {}
    for node, raw in items:
        durable = next(iter(_prefixed_values(raw, "durableId")), None)
        owner = durable_owner.get(durable.upper(), "orphan:" + raw) if durable is not None else "orphan:" + raw
        out["word/commentsExtensible.xml"].setdefault(owner, []).append(("cex", raw))
    return out


def _run_chars(fragment: str) -> list[tuple[str, str]]:
    """(character, run properties) of a revision fragment, so a split run reads the same as the whole."""

    from .revisions import parse, unescape

    tree = parse(fragment)
    out = []
    for run in (node for node in tree.iter() if node.name == "w:r"):
        rpr = run.child("w:rPr")
        props = fragment[rpr.start:rpr.end] if rpr is not None else ""
        for child in run.children:
            if child.name in ("w:t", "w:delText", "w:instrText", "w:delInstrText"):
                text = "" if child.closed else unescape(child.inner(fragment))
                out.extend((ch, child.name + props) for ch in text)
            elif child.name != "w:rPr" and child.name != "w:commentReference":
                out.append((fragment[child.start:child.end], props))
    return out


def _strip_markers(fragment: str, legacy_ids: set[str]) -> str:
    """``fragment`` without the parsed markers of ``legacy_ids`` (any quote style or spacing): range starts
    and ends, and each reference with its run when the run holds nothing else (as delete_comment removes)."""

    from .revisions import parse

    if not legacy_ids:
        return fragment
    cuts = []
    for node in parse(fragment).iter():
        if node.name not in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
            continue
        if node.attr(fragment, "w:id") not in legacy_ids:
            continue
        run = node.parent
        whole_run = node.name == "w:commentReference" and run.name == "w:r" and _element_children(run) == [node]
        cuts.append((run.start, run.end) if whole_run else (node.start, node.end))
    for start, end in sorted(set(cuts), reverse=True):
        fragment = fragment[:start] + fragment[end:]
    return fragment


_COMMENT_MARKERS = {f"{{{W_NS}}}{local}" for local in ("commentRangeStart", "commentRangeEnd",
                                                         "commentReference")}


def _story_markers(parts: dict[str, bytes], legacy_ids: set[str]) -> dict[str, int]:
    """Per word/*.xml part (document, footnotes, endnotes, headers, footers, any other story): how many
    comment range starts/ends/references carry one of ``legacy_ids``. Namespace-aware (any prefix bound to
    the WordprocessingML namespace) and independent of quote style, spacing or attribute order: sol61's
    footnote markers `<w:commentReference w:id='0' />` were invisible to the regex that preceded this."""

    key = f"{{{W_NS}}}id"
    found: dict[str, int] = {}
    for name, data in parts.items():
        if not (name.startswith("word/") and name.lower().endswith(".xml")):
            continue
        try:
            count = sum(1 for _, el in ET.iterparse(io.BytesIO(data))
                        if el.tag in _COMMENT_MARKERS and (el.get(key) or "").strip() in legacy_ids)
        except ET.ParseError as exc:
            raise RefusedUnsafe(f"cannot scan {name} for comment markers: {exc}") from exc
        if count:
            found[name] = count
    return found


def _verify_document(before: str, after: str, changed: set[str], op: str) -> None:
    """Revision count unchanged; every revision byte-identical except one holding an anchor edge, which may
    differ only by the changed comments' markers and the run split at that edge (same record head: id,
    author, date; same characters and run formatting). The document is identical once those markers are
    removed and split runs are read as characters."""

    from .revisions import REV_TAGS, parse

    def revisions(xml: str):
        root = parse(xml)
        return [node for node in root.iter() if node.name.startswith("w:") and node.local in REV_TAGS]

    old, new = revisions(before), revisions(after)
    if len(old) != len(new):
        raise VerificationFailed(f"{op}: revision count changed ({len(old)} -> {len(new)})")
    for old_node, new_node in zip(old, new):
        old_raw, new_raw = before[old_node.start:old_node.end], after[new_node.start:new_node.end]
        if old_raw == new_raw:
            continue
        if old_node.head(before) != new_node.head(after):
            raise VerificationFailed(f"{op}: revision {old_node.head(before)[:80]} changed its record")
        if _run_chars(_strip_markers(old_raw, changed)) != _run_chars(_strip_markers(new_raw, changed)):
            raise VerificationFailed(f"{op}: revision {old_node.head(before)[:80]} changed its text or format")
    stripped_old, stripped_new = _strip_markers(before, changed), _strip_markers(after, changed)
    if stripped_old != stripped_new and _run_chars(stripped_old) != _run_chars(stripped_new):
        raise VerificationFailed(f"{op}: document text or formatting changed outside the comment markers")
    if parse(stripped_new) is None:  # well-formed (parse raises otherwise)
        raise VerificationFailed(f"{op}: document.xml is malformed")


def _verify_package(package: _Package, changed_parts: dict[str, bytes], touched: set[str], op: str,
                    *, allowed_parts: set[str]) -> None:
    after_parts = dict(package.parts)
    after_parts.update(changed_parts)
    for name, data in changed_parts.items():
        if name not in allowed_parts:
            raise VerificationFailed(f"{op}: unexpected change to {name}")
        if name.endswith((".xml", ".rels")):
            try:
                ET.fromstring(data)
            except ET.ParseError as exc:
                raise VerificationFailed(f"{op}: {name} is not well-formed: {exc}") from exc
    before_owners, after_owners = _owner_map(package.parts), _owner_map(after_parts)
    for part in before_owners:
        old, new = before_owners[part], after_owners.get(part, {})
        for owner in set(old) | set(new):
            if owner in touched:
                continue
            if old.get(owner, []) != new.get(owner, []):
                raise VerificationFailed(f"{op}: {part} entry for comment {owner} changed; only the comment "
                                         "being edited may change")
        old_order = [owner for owner in old if owner not in touched]
        new_order = [owner for owner in new if owner not in touched]
        if old_order != new_order:
            raise VerificationFailed(f"{op}: {part} reordered other comments")
    for name in (_CONTENT_TYPES, _DOC_RELS):
        if name in changed_parts:
            old, new = package.parts[name].decode("utf-8"), changed_parts[name].decode("utf-8")
            close = "</Types>" if name == _CONTENT_TYPES else "</Relationships>"
            if not new.startswith(old[: old.rindex(close)]) or not new.endswith(old[old.rindex(close):]):
                raise VerificationFailed(f"{op}: {name} changed beyond added entries")


def _thread_dicts(path: Path) -> dict[str, dict]:
    from .comments import inspect_comments
    from .paragraphs import read_snapshot

    return {thread.legacy_id: thread.to_dict() for thread in inspect_comments(read_snapshot(path))}


# --------------------------------------------------------------------------
# Operations.
# --------------------------------------------------------------------------

def _finish(package: _Package, output: str | Path, doc_xml: str | None, touched: set[str], op: str,
            *, extra_allowed: set[str] = frozenset()) -> dict[str, bytes]:
    from .revisions import revision_count

    changed = package.render()
    before_doc = package.parts["word/document.xml"].decode("utf-8")
    if doc_xml is not None and doc_xml != before_doc:
        changed["word/document.xml"] = doc_xml.encode("utf-8")
        _verify_document(before_doc, doc_xml, touched, op)
    allowed = {spec.name for spec in _PARTS.values()} | {_CONTENT_TYPES, _DOC_RELS, "word/document.xml"}
    _verify_package(package, changed, touched, op, allowed_parts=allowed | set(extra_allowed))
    before_revisions = revision_count(package.source)
    package.write(output, changed)
    try:
        after_revisions = revision_count(output)
        if after_revisions != before_revisions:
            raise VerificationFailed(f"{op}: tracked revision count changed ({before_revisions} -> "
                                     f"{after_revisions})")
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    return changed


def _result(source, output, op: str, **extra) -> dict:
    from .audit import audit_docx, file_sha256
    from .revisions import revision_count

    return {
        "source": str(Path(source).resolve()), "output": str(Path(output).resolve()), "operation": op,
        "engine": "direct_xml", "index_delta": 0, "source_sha256": file_sha256(source), "output_sha256": file_sha256(output),
        "paragraph_count": audit_docx(output).paragraph_count, "indices_after_shifted": False,
        "protected_structure_unchanged": True, "revision_count": revision_count(output), **extra,
    }


def _check_threads(source: Path, output: Path, touched: set[str], op: str, removed: set[str] = frozenset(),
                   added: str | None = None) -> dict[str, dict]:
    before, after = _thread_dicts(source), _thread_dicts(output)
    for legacy, thread in before.items():
        if legacy in touched or legacy in removed:
            continue
        if after.get(legacy) != thread:
            Path(output).unlink(missing_ok=True)
            raise VerificationFailed(f"{op}: comment thread {thread['comment_id']} changed; only the thread "
                                     "being edited may change")
    expected = (set(before) - set(removed)) | ({added} if added else set())
    if set(after) != expected:
        Path(output).unlink(missing_ok=True)
        raise VerificationFailed(f"{op}: threads after the edit are {sorted(after)}, expected {sorted(expected)}")
    return after


def _clean_author(author: str | None, initials: str | None, op: str) -> tuple[str, str]:
    """A given author is used as is (initials derived from it when not given). Only a missing author asks
    ``config.comment_author()`` (RESEARCH_MCP_AUTHOR, else Word's user name, else "Author")."""

    if author is None:
        author, default_initials = config.comment_author()
        initials = default_initials if initials is None else initials
    author = author.strip()
    if not author:
        raise RefusedUnsafe(f"{op}: comment author must not be empty")
    if initials is None:
        initials = config.initials_for(author)
    return author, initials.strip()


def add_comment(source, output, *, paragraph_index: int, anchor: str, text: str, author: str | None = None,
                initials: str | None = None, now: dt.datetime | None = None) -> dict:
    """New open thread on the unique visible ``anchor`` in ``paragraph_index``."""

    op = "add_comment"
    if not 1 <= len(anchor) <= 255:
        raise RefusedUnsafe("anchor must contain 1..255 characters")
    if not text.strip():
        raise RefusedUnsafe("comment text must not be empty")
    text = text.strip().replace("\r\n", "\n").replace("\r", "\n")
    author, initials = _clean_author(author, initials, op)
    now = now or _now()
    from .revisions import Document
    from .sections import read_layout

    source, output = Path(source), Path(output)
    package = _Package(source, op)
    comments = _Comments(package)
    ids = _Ids(package)
    doc = Document(package.parts["word/document.xml"].decode("utf-8"), read_layout(source))
    legacy = ids.comment_id()
    edits = _anchor_edits(doc, paragraph_index, anchor, legacy, _reference_run(package, legacy), op)
    new_xml = _apply_splices(doc.xml, edits)
    part, p14 = _comments_part_for_write(package)
    comment_xml, para_id = _comment_xml(package, ids, legacy_id=legacy, author=author, initials=initials,
                                        text=text, now=now, p14=p14)
    part.append(comment_xml)
    writer = _Writer(package, comments)
    durable = ids.durable_id()
    writer.extended(para_id, None, False)
    writer.ids(para_id, durable)
    writer.extensible(durable, now)
    _finish(package, output, new_xml, {legacy}, op)
    after = _check_threads(source, output, set(), op, added=legacy)
    thread = after[legacy]
    problems = []
    if thread["comment_id"] != durable:
        problems.append(f"id {thread['comment_id']} != {durable}")
    if thread["text"] != text or thread["author"] != author or thread["done"]:
        problems.append("text/author/done")
    if thread["anchor_paragraph_index"] != paragraph_index or thread["anchor_text"] != anchor:
        problems.append(f"anchor {thread['anchor_paragraph_index']}:{thread['anchor_text']!r}")
    if problems:
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{op}: the new thread did not read back as written ({'; '.join(problems)})")
    return _result(source, output, op, comment_id=durable, legacy_id=legacy, author=author, initials=initials,
                   anchor_paragraph_index=paragraph_index, anchor_text=anchor, thread_open=True,
                   parts_created=[_PARTS[kind].name for kind in package.added])


def _thread_markers(doc_xml: str, legacy_ids: list[str], op: str, root_id: str):
    from .revisions import parse

    root = parse(doc_xml)
    starts, _, references = _markers(root, set(legacy_ids), doc_xml)
    if not references:
        raise RefusedUnsafe(f"{op}: comment {root_id!r} has no reference in word/document.xml (it may be anchored "
                            "in a header, footnote or text box); edit it in Word")
    return starts, references


def reply_comment(source, output, *, comment_id: str, disposition: str, reply_text: str,
                  resolve: bool = False, author: str | None = None, initials: str | None = None,
                  now: dt.datetime | None = None) -> dict:
    op = "reply_comment"
    if disposition not in DISPOSITION_LABELS:
        raise RefusedUnsafe(f"unsupported comment disposition: {disposition!r}; use one of "
                            + ", ".join(DISPOSITION_LABELS))
    if not reply_text.strip():
        raise RefusedUnsafe("reply_text must not be empty")
    author, initials = _clean_author(author, initials, op)
    now = now or _now()
    rendered = f"{DISPOSITION_LABELS[disposition]}: {reply_text.strip()}"
    source, output = Path(source), Path(output)
    package = _Package(source, op)
    comments = _Comments(package)
    root = comments.thread_root(comment_id)
    if root.done and not resolve:
        raise RefusedUnsafe(f"{op}: thread {root.comment_id!r} is resolved; reopen_comment it first "
                            "(or pass resolve=true to reply and keep it resolved)")
    ids = _Ids(package)
    writer = _Writer(package, comments)
    thread = [root] + comments.replies(root)
    root_para = _ensure_identity(package, comments, ids, writer, root)
    doc_xml = package.parts["word/document.xml"].decode("utf-8")
    starts, references = _thread_markers(doc_xml, [item.legacy_id for item in thread], op, root.comment_id)
    legacy = ids.comment_id()
    edits = []
    if starts:
        last_start = max(starts, key=lambda node: node.start)
        edits.append((last_start.end, last_start.end, f'<w:commentRangeStart w:id="{legacy}"/>'))
    last_reference = max(references, key=lambda node: node.start)
    run = last_reference.parent if last_reference.parent.name == "w:r" else last_reference
    wrapper = _outermost_wrapper(run)
    point = (wrapper or run).end
    end_marker = f'<w:commentRangeEnd w:id="{legacy}"/>' if starts else ""
    edits.append((point, point, end_marker + _reference_run(package, legacy)))
    new_xml = _apply_splices(doc_xml, edits)
    part, p14 = _comments_part_for_write(package)
    comment_xml, para_id = _comment_xml(package, ids, legacy_id=legacy, author=author, initials=initials,
                                        text=rendered, now=now, p14=p14)
    part.append(comment_xml)
    durable = ids.durable_id()
    writer.extended(para_id, root_para, bool(resolve or root.done))
    writer.ids(para_id, durable)
    writer.extensible(durable, now)
    touched = {item.legacy_id for item in thread} | {legacy}
    if resolve:
        _set_done(package, comments, writer, thread, root_para, True)
    _finish(package, output, new_xml, touched, op)
    after = _check_threads(source, output, {root.legacy_id}, op)
    view = after[root.legacy_id]
    reply = next((item for item in view["replies"] if item["legacy_id"] == legacy), None)
    if reply is None or reply["text"] != rendered or reply["author"] != author or view["done"] != bool(resolve):
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{op}: the reply did not read back as written")
    return _result(source, output, op, comment_id=view["comment_id"], legacy_id=root.legacy_id,
                   reply_id=durable, reply_legacy_id=legacy, disposition=disposition, reply_text=rendered,
                   author=author, thread_open=not view["done"], resolved=view["done"],
                   anchor_paragraph_index=view["anchor_paragraph_index"], anchor_text=view["anchor_text"])


def _set_done(package: _Package, comments: _Comments, writer: _Writer, thread: list[_Record], root_para: str,
              done: bool) -> None:
    value = "1" if done else "0"
    extended = package.part("extended")
    p15 = _prefix_for(extended.head(), W15_NS) if extended is not None and not extended.new else None
    for record in thread:
        if record.extended is not None and p15 is not None:
            extended.splice(*_set_attr(extended.xml, record.extended, f"{p15}:done", value))
        else:
            if not record.para_id:
                _ensure_identity(package, comments, _Ids(package), writer, record)
            writer.extended(record.para_id, None if record is thread[0] else root_para, done)


def set_comment_done(source, output, *, comment_id: str, done: bool) -> dict:
    op = "resolve_comment" if done else "reopen_comment"
    source, output = Path(source), Path(output)
    package = _Package(source, op)
    comments = _Comments(package)
    root = comments.thread_root(comment_id)
    if root.done == done:
        raise RefusedUnsafe(f"{op}: thread {root.comment_id!r} is already {'resolved' if done else 'open'}; "
                            "nothing to change")
    ids = _Ids(package)
    writer = _Writer(package, comments)
    thread = [root] + comments.replies(root)
    root_para = _ensure_identity(package, comments, ids, writer, root)
    _set_done(package, comments, writer, thread, root_para, done)
    _finish(package, output, None, {item.legacy_id for item in thread}, op)
    after = _check_threads(source, output, {root.legacy_id}, op)
    view = after[root.legacy_id]
    if view["done"] != done or any(reply["done"] != done for reply in view["replies"]):
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{op}: the thread did not read back as {'resolved' if done else 'open'}")
    return _result(source, output, op, comment_id=view["comment_id"], legacy_id=root.legacy_id,
                   thread_open=not done, resolved=done, reply_count=len(view["replies"]),
                   anchor_paragraph_index=view["anchor_paragraph_index"], anchor_text=view["anchor_text"])


def delete_comment(source, output, *, comment_id: str, include_replies: bool = True) -> dict:
    op = "delete_comment"
    source, output = Path(source), Path(output)
    package = _Package(source, op)
    comments = _Comments(package)
    record = comments.find(comment_id)
    if record.parent:
        if include_replies:
            root = comments.root_of(record)
            name = root.comment_id if root else record.parent
            raise RefusedUnsafe(f"{op}: {comment_id!r} is a reply in thread {name!r}; to delete only this reply "
                                f"pass include_replies=false, to delete the thread use {name!r}")
        victims = [record]
    else:
        replies = comments.replies(record)
        if replies and not include_replies:
            raise RefusedUnsafe(f"{op}: thread {record.comment_id!r} has {len(replies)} replies; deleting only "
                                "its first comment would orphan them. Use include_replies=true")
        victims = [record] + replies
    legacy_ids = {item.legacy_id for item in victims}
    doc_xml = package.parts["word/document.xml"].decode("utf-8")
    from .revisions import parse

    tree = parse(doc_xml)
    starts, ends, references = _markers(tree, legacy_ids, doc_xml)
    edits = []
    for node in starts + ends:
        edits.append((node.start, node.end, ""))
    for node in references:
        run = node.parent
        content = _element_children(run) if run.name == "w:r" else []
        target = run if run.name == "w:r" and content == [node] else node
        edits.append((target.start, target.end, ""))
    # every story, parsed (not a lexical pattern): a marker left in a footnote, endnote, header or footer
    # would reference the deleted comment, so such a comment is refused before anything is written
    stories = _story_markers(package.parts, legacy_ids)
    other_parts = sorted(name for name in stories if name != "word/document.xml")
    if other_parts:
        raise RefusedUnsafe(f"{op}: comment {record.comment_id!r} is anchored in {', '.join(other_parts)}; "
                            "delete it in Word")
    if stories.get("word/document.xml", 0) != len(starts) + len(ends) + len(references):
        raise RefusedUnsafe(f"{op}: comment {record.comment_id!r} has markers in word/document.xml written with "
                            "a namespace prefix other than w:; delete it in Word")
    new_xml = _apply_splices(doc_xml, edits)
    part = package.part("comments")
    for item in victims:
        part.splice(item.node.start, item.node.end, "")
        for kind, node in (("extended", item.extended), ("ids", item.ids), ("extensible", item.extensible)):
            if node is not None:
                package.part(kind).splice(node.start, node.end, "")
    _finish(package, output, new_xml, legacy_ids, op)
    with zipfile.ZipFile(output) as archive:
        written = {name: archive.read(name) for name in archive.namelist()}
    try:
        leftover = sorted(_story_markers(written, legacy_ids))
    except RefusedUnsafe as exc:
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{op}: {exc}") from exc
    if leftover:
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{op}: comment markers remain in {leftover}")
    if record.parent:
        root = comments.root_of(record)
        _check_threads(source, output, {root.legacy_id} if root else set(), op)
    else:
        _check_threads(source, output, set(), op, removed={record.legacy_id})
    return _result(source, output, op, comment_id=record.comment_id, legacy_id=record.legacy_id,
                   deleted=sorted(item.comment_id for item in victims),
                   deleted_legacy_ids=sorted(legacy_ids, key=int))


def comment_paragraphs(path, comment_id: str) -> list[int] | None:
    """Layout paragraph indices a comment's whole thread is anchored on (first marker .. last marker).

    Used by edit_document's grouping so resolve/reopen/delete/reply (which carry no paragraph_index) join the
    table-row or section group of the text they comment on. None when the id is unknown or the thread has no
    markers in the body; the operation itself then reports its own error."""

    from .revisions import load

    try:
        package = _Package(path, "comment grouping")
        comments = _Comments(package)
        record = comments.find(comment_id)
        root = comments.root_of(record) or record
        thread = [root, record] + comments.replies(root)
        _, doc = load(path)
    except (RefusedUnsafe, VerificationFailed):
        return None
    starts, ends, references = _markers(doc.root, {item.legacy_id for item in thread}, doc.xml)
    indices = [doc.paragraph_at(node.start) for node in starts + ends + references]
    indices = [index for index in indices if index is not None]
    if not indices:
        return None
    return list(range(min(indices), max(indices) + 1))


def apply_operation(source, output, operation) -> dict:
    kind = operation.type
    if kind == "add_comment":
        return add_comment(source, output, paragraph_index=operation.paragraph_index, anchor=operation.anchor,
                           text=operation.text, author=operation.author, initials=operation.initials)
    if kind == "reply_comment":
        return reply_comment(source, output, comment_id=operation.comment_id, disposition=operation.disposition,
                             reply_text=operation.reply_text, resolve=operation.resolve,
                             author=operation.author, initials=operation.initials)
    if kind in ("resolve_comment", "reopen_comment"):
        return set_comment_done(source, output, comment_id=operation.comment_id, done=kind == "resolve_comment")
    if kind == "delete_comment":
        return delete_comment(source, output, comment_id=operation.comment_id,
                              include_replies=operation.include_replies)
    raise RefusedUnsafe(f"unsupported comment operation {kind!r}")


COMMENT_OPERATION_TYPES = ("add_comment", "reply_comment", "resolve_comment", "reopen_comment", "delete_comment")
