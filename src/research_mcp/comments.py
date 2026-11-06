"""Offline, durable Word comment-thread inspection."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from io import BytesIO
import xml.etree.ElementTree as ET
import zipfile

from .ooxml import W
from .paragraphs import (
    DocumentSnapshot,
    _paragraph_nodes,
    _paragraph_text,
)


W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
W16CID = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"


@dataclass(frozen=True)
class CommentReplyInfo:
    comment_id: str
    legacy_id: str
    author: str
    date: str | None
    text: str
    done: bool

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass(frozen=True)
class CommentThreadInfo:
    comment_id: str
    legacy_id: str
    author: str
    initials: str
    date: str | None
    text: str
    done: bool
    anchor_paragraph_index: int | None
    anchor_text: str
    anchor_text_exact: bool
    replies: tuple[CommentReplyInfo, ...]

    def to_dict(self) -> dict:
        return asdict(self)


def _read_parts(snapshot: DocumentSnapshot) -> dict[str, bytes]:
    wanted = {
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
    }
    with zipfile.ZipFile(BytesIO(snapshot.package)) as archive:
        names = set(archive.namelist())
        return {name: archive.read(name) for name in wanted if name in names}


def _comment_body(element: ET.Element) -> str:
    return "\n".join(_paragraph_text(paragraph) for paragraph in element.findall(f".//{W}p"))


def _anchor_data(
    snapshot: DocumentSnapshot,
) -> dict[str, tuple[int | None, str, bool]]:
    """Legacy comment ID -> (first COM paragraph, visible scope, exact)."""

    active: set[str] = set()
    starts: dict[str, int] = {}
    references: dict[str, int] = {}
    text: dict[str, list[str]] = {}
    exact: dict[str, bool] = {}

    def emit(paragraph: ET.Element, paragraph_index: int) -> None:
        def walk(node: ET.Element) -> None:
            for child in node:
                if child.tag == W + "pPr":
                    continue
                if child.tag == W + "commentRangeStart":
                    legacy_id = child.get(W + "id", "")
                    if legacy_id:
                        active.add(legacy_id)
                        starts.setdefault(legacy_id, paragraph_index)
                        text.setdefault(legacy_id, [])
                        exact.setdefault(legacy_id, True)
                    continue
                if child.tag == W + "commentRangeEnd":
                    legacy_id = child.get(W + "id", "")
                    active.discard(legacy_id)
                    continue
                if child.tag == W + "commentReference":
                    legacy_id = child.get(W + "id", "")
                    if legacy_id:
                        references.setdefault(legacy_id, paragraph_index)
                    continue
                if child.tag == W + "t":
                    value = child.text or ""
                    for legacy_id in active:
                        text.setdefault(legacy_id, []).append(value)
                    continue
                if child.tag == W + "tab":
                    for legacy_id in active:
                        text.setdefault(legacy_id, []).append("\t")
                    continue
                if child.tag in (W + "br", W + "cr"):
                    value = "\x0c" if child.get(W + "type") == "page" else "\v"
                    for legacy_id in active:
                        text.setdefault(legacy_id, []).append(value)
                    continue
                if child.tag in (W + "drawing", W + "pict", W + "object"):
                    for legacy_id in active:
                        exact[legacy_id] = False
                walk(child)

        walk(paragraph)
        # A comment that remains active after this paragraph spans its paragraph
        # mark. Word exposes that separator in Scope.Text.
        for legacy_id in active:
            text.setdefault(legacy_id, []).append("\r")

    for index, paragraph, _ in _paragraph_nodes(snapshot):
        if paragraph is not None:
            emit(paragraph, index)

    ids = set(starts) | set(references) | set(text)
    return {
        legacy_id: (
            starts.get(legacy_id, references.get(legacy_id)),
            "".join(text.get(legacy_id, ())).rstrip("\r"),
            exact.get(legacy_id, True),
        )
        for legacy_id in ids
    }


def inspect_comments(snapshot: DocumentSnapshot) -> tuple[CommentThreadInfo, ...]:
    parts = _read_parts(snapshot)
    raw_comments = parts.get("word/comments.xml")
    if raw_comments is None:
        return ()

    comments_root = ET.fromstring(raw_comments)
    extended: dict[str, tuple[str | None, bool]] = {}
    if raw := parts.get("word/commentsExtended.xml"):
        for item in ET.fromstring(raw):
            para_id = item.get(W15 + "paraId")
            if para_id:
                extended[para_id] = (
                    item.get(W15 + "paraIdParent"),
                    item.get(W15 + "done", "0") == "1",
                )

    durable: dict[str, str] = {}
    if raw := parts.get("word/commentsIds.xml"):
        for item in ET.fromstring(raw):
            para_id = item.get(W16CID + "paraId")
            durable_id = item.get(W16CID + "durableId")
            if para_id and durable_id:
                durable[para_id] = durable_id.upper()

    records: list[dict] = []
    for item in comments_root.findall(f"{W}comment"):
        paragraphs = item.findall(f".//{W}p")
        para_id = next(
            (paragraph.get(W14 + "paraId") for paragraph in reversed(paragraphs)
             if paragraph.get(W14 + "paraId")),
            None,
        )
        parent, done = extended.get(para_id or "", (None, False))
        legacy_id = item.get(W + "id", "")
        records.append(
            {
                "para_id": para_id,
                "parent": parent,
                "done": done,
                "legacy_id": legacy_id,
                "comment_id": durable.get(para_id or "", f"legacy:{legacy_id}"),
                "author": item.get(W + "author", ""),
                "initials": item.get(W + "initials", ""),
                "date": item.get(W + "date"),
                "text": _comment_body(item),
            }
        )

    replies_by_parent: dict[str, list[CommentReplyInfo]] = {}
    for record in records:
        if not record["parent"]:
            continue
        replies_by_parent.setdefault(record["parent"], []).append(
            CommentReplyInfo(
                comment_id=record["comment_id"],
                legacy_id=record["legacy_id"],
                author=record["author"],
                date=record["date"],
                text=record["text"],
                done=record["done"],
            )
        )

    anchors = _anchor_data(snapshot)
    threads: list[CommentThreadInfo] = []
    for record in records:
        if record["parent"]:
            continue
        paragraph_index, anchor_text, anchor_exact = anchors.get(
            record["legacy_id"], (None, "", True)
        )
        threads.append(
            CommentThreadInfo(
                comment_id=record["comment_id"],
                legacy_id=record["legacy_id"],
                author=record["author"],
                initials=record["initials"],
                date=record["date"],
                text=record["text"],
                done=record["done"],
                anchor_paragraph_index=paragraph_index,
                anchor_text=anchor_text,
                anchor_text_exact=anchor_exact,
                replies=tuple(replies_by_parent.get(record["para_id"] or "", ())),
            )
        )
    return tuple(threads)
