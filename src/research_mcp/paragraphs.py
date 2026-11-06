"""Offline paragraph addressing that agrees with Word's COM numbering.

`insert_zotero_citation` and the citation restore path both address text as
`document.Paragraphs(paragraph_index)` -- a *COM* coordinate. Any tool that
hands an agent a paragraph index has to speak that same coordinate system, or
the agent reads an index from one tool and silently edits a different
paragraph with the next.

The two numberings are not the same. Measured on a long manuscript with many tables:

    OOXML  .//w:p          502
    COM    Paragraphs.Count 552

Word's collection contains one extra paragraph per table row -- the row-end
mark -- placed *after* that row's cell paragraphs. Reconstructing the sequence
under that rule and diffing element-wise against a live COM capture of all 552
paragraphs gives exact positional agreement: same length, same order, and
`Range.Information(wdWithInTable)` matching this module's `in_table` for every
paragraph. A flat `.//w:p` walk instead diverges from index 133 onward and is
wrong by 50 by the end of the document.

That result is what makes this module offline. It reads the package, takes no
Word instance and no integration lease, and still emits indices that
`insert_zotero_citation` will resolve to the paragraph the agent meant.

One honest limit: paragraph *text* is recovered from `w:t` runs, which do not
carry field results. In the same comparison 35 of 552 paragraphs differed in
text -- 28 by whitespace inside Zotero bibliography fields, 7 where Word
renders a field or drawing character that the package does not store as text.
Those paragraphs are flagged `has_fields`, because an `anchor` matched
against the text this module returns will not match what Word sees.
"""

from __future__ import annotations

from io import BytesIO
from dataclasses import asdict, dataclass
import hashlib
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

from .audit import is_file_locked
from .ooxml import NS, W


# A whole-document dump is the failure mode this project just spent a payload
# audit removing: a tool result is permanent context, carried by every later
# turn. Callers page through instead.
DEFAULT_WINDOW = 40
MAX_WINDOW = 200


@dataclass(frozen=True)
class DocumentSnapshot:
    """One immutable package read used for both identity and inspection."""

    path: Path
    package: bytes
    sha256: str


def read_snapshot(path: str | Path) -> DocumentSnapshot:
    resolved = Path(path).resolve(strict=True)
    package = resolved.read_bytes()
    return DocumentSnapshot(
        path=resolved,
        package=package,
        sha256=hashlib.sha256(package).hexdigest().upper(),
    )


@dataclass(frozen=True)
class ParagraphInfo:
    """One paragraph, addressed the way Word addresses it."""

    index: int
    style: str | None
    in_table: bool
    has_fields: bool
    text_exact: bool
    text: str

    def to_dict(self) -> dict:
        return asdict(self)


# Word's Range.Text renders an inline drawing as a placeholder character. Its
# identity is a Word implementation detail -- the recorded capture shows "/"
# for every picture in v49 -- so this module refuses to guess it and marks the
# paragraph inexact instead. An `anchor` matched against such a paragraph's
# text would not be what Word actually sees.
_OPAQUE_TAGS = ("drawing", "pict", "object")


def _paragraph_text(paragraph: ET.Element) -> str:
    """Text in document order, including tabs.

    A `w:t`-only join silently drops `w:tab`, which is how every numbered
    bibliography entry ("1.\tKrippendorff K...") came out one character short
    against Word. Offsets are the whole point of this data, so separators are
    reproduced rather than collapsed.
    """

    parts: list[str] = []

    def emit(node: ET.Element) -> None:
        for child in node:
            # `w:pPr/w:tabs/w:tab` are tab *stop definitions*, not tab
            # characters. Walking the whole subtree picks them up and prepends
            # a phantom tab to every paragraph that declares a stop.
            if child.tag == W + "pPr":
                continue
            if child.tag == W + "t":
                parts.append(child.text or "")
            elif child.tag == W + "tab":
                parts.append("\t")
            elif child.tag in (W + "br", W + "cr"):
                # Both verified against live Word captures: a page break reads
                # \x0c and a plain line break reads \v (tests/test_paragraphs.py).
                parts.append("\x0c" if child.get(W + "type") == "page" else "\v")
            else:
                emit(child)

    emit(paragraph)
    return "".join(parts)


def _text_is_exact(paragraph: ET.Element) -> bool:
    return not any(
        paragraph.find(f".//w:{tag}", NS) is not None for tag in _OPAQUE_TAGS
    )


def _paragraph_style(paragraph: ET.Element) -> str | None:
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return style.get(W + "val") if style is not None else None


def _has_fields(paragraph: ET.Element) -> bool:
    return any(
        paragraph.find(f".//w:{tag}", NS) is not None
        for tag in ("fldChar", "instrText", "fldSimple")
    )


def _row_end_marker(index: int) -> ParagraphInfo:
    """The paragraph Word puts at the end of every table row.

    It holds no text and cannot be addressed usefully, but it occupies an
    index, so omitting it shifts every later paragraph.
    """

    return ParagraphInfo(
        index=index, style=None, in_table=True, has_fields=False, text_exact=True, text=""
    )


def _paragraph_nodes(
    snapshot: DocumentSnapshot,
) -> tuple[tuple[int, ET.Element | None, bool], ...]:
    """Actual paragraph elements plus Word's synthetic table-row paragraphs."""

    with zipfile.ZipFile(BytesIO(snapshot.package)) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        return ()

    found: list[tuple[int, ET.Element | None, bool]] = []

    def add(paragraph: ET.Element | None, in_table: bool) -> None:
        found.append((len(found) + 1, paragraph, in_table))

    def walk_table(table: ET.Element) -> None:
        for row in table.findall("w:tr", NS):
            for cell in row.findall("w:tc", NS):
                for nested in cell:
                    if nested.tag == W + "p":
                        add(nested, True)
                    elif nested.tag == W + "tbl":
                        walk_table(nested)
            add(None, True)

    for child in body:
        if child.tag == W + "p":
            add(child, False)
        elif child.tag == W + "tbl":
            walk_table(child)
    return tuple(found)


def _read_paragraphs(snapshot: DocumentSnapshot) -> tuple[ParagraphInfo, ...]:
    """Every paragraph in COM order, offline.

    Body children are walked in document order. A `w:tbl` contributes, per
    row, that row's cell paragraphs followed by one row-end paragraph.
    """

    found: list[ParagraphInfo] = []
    for index, paragraph, in_table in _paragraph_nodes(snapshot):
        if paragraph is None:
            found.append(_row_end_marker(index))
        else:
            found.append(
                ParagraphInfo(
                    index=index,
                    style=_paragraph_style(paragraph),
                    in_table=in_table,
                    has_fields=_has_fields(paragraph),
                    text_exact=_text_is_exact(paragraph),
                    text=_paragraph_text(paragraph),
                )
            )

    return tuple(found)


def read_paragraphs(path: str | Path) -> tuple[ParagraphInfo, ...]:
    """Every paragraph in COM order from one immutable package read."""

    return _read_paragraphs(read_snapshot(path))


def _outline_styles(snapshot: DocumentSnapshot) -> set[str]:
    """Style ids that Word treats as headings, resolved through `w:basedOn`.

    Matching style *names* against "heading" only works for documents that use
    the built-in styles. A journal template with `SectionHead` would return an
    empty outline, and an empty outline is worse than no tool: the agent falls
    back to windowing a 552-paragraph document 200 at a time.

    `w:outlineLvl` is what actually makes a style a heading, and it lives in
    `styles.xml` rather than on the paragraph -- this manuscript has zero
    paragraph-level `outlineLvl` elements and six styles carrying one.
    """

    try:
        with zipfile.ZipFile(BytesIO(snapshot.package)) as archive:
            styles = ET.fromstring(archive.read("word/styles.xml"))
    except (KeyError, ET.ParseError):
        return set()

    levels: dict[str, str | None] = {}
    based_on: dict[str, str] = {}
    for style in styles.findall("w:style", NS):
        style_id = style.get(W + "styleId")
        if not style_id:
            continue
        level = style.find("w:pPr/w:outlineLvl", NS)
        levels[style_id] = level.get(W + "val") if level is not None else None
        parent = style.find("w:basedOn", NS)
        if parent is not None and parent.get(W + "val"):
            based_on[style_id] = parent.get(W + "val")

    def has_level(style_id: str, depth: int = 0) -> bool:
        if depth > 10 or style_id not in levels:
            return False
        if levels[style_id] is not None:
            return True
        parent = based_on.get(style_id)
        return has_level(parent, depth + 1) if parent else False

    return {style_id for style_id in levels if has_level(style_id)}


def _outline(snapshot: DocumentSnapshot, paragraphs: tuple[ParagraphInfo, ...]) -> tuple[ParagraphInfo, ...]:
    """Headings only, so an agent can find a section without paging.

    A 552-paragraph manuscript cannot be located by windowing alone; this is
    the cheap index that makes a targeted `read_paragraphs` window possible.

    `Title` carries no outline level, so it is matched by name as well.
    """

    heading_styles = _outline_styles(snapshot)
    return tuple(
        item
        for item in paragraphs
        if item.style
        and (
            item.style in heading_styles
            or item.style.lower().startswith(("heading", "title"))
        )
    )


def outline(path: str | Path) -> tuple[ParagraphInfo, ...]:
    snapshot = read_snapshot(path)
    return _outline(snapshot, _read_paragraphs(snapshot))


def _unusual_characters(path) -> dict:
    """Plan 551: what edit_document's sweep_characters would clean in the body (same skip rules)."""

    from .revisions import scan_characters

    try:
        return scan_characters(path)
    except Exception as exc:  # the count is advisory; inspection itself must not fail on it
        return {"total": None, "by_codepoint": {}, "error": f"{type(exc).__name__}: {exc}"[:300]}


def inspect_document(
    path: str | Path,
    *,
    first_paragraph: int | None = None,
    last_paragraph: int | None = None,
    include_comments: bool = False,
    search_text: str | None = None,
) -> dict:
    """Outline by default; paragraph search or a bounded paragraph window when requested.

    Calling with no range returns the document's shape and headings.
    Supplying `search_text` finds matching paragraphs and returns their indices and text.
    Supplying a range returns the paragraphs themselves, clamped to `MAX_WINDOW`.
    """

    snapshot = read_snapshot(path)
    all_paragraphs = _read_paragraphs(snapshot)
    total = len(all_paragraphs)
    locked, lock_reason = is_file_locked(snapshot.path)

    summary: dict = {
        "path": str(snapshot.path),
        "source_sha256": snapshot.sha256,
        "is_locked": locked,
        "paragraph_count": total,
        "paragraph_count_basis": (
            "Word-compatible 1-based Paragraphs collection, including table-row end markers"
        ),
        "table_paragraph_count": sum(1 for item in all_paragraphs if item.in_table),
        "field_paragraph_count": sum(1 for item in all_paragraphs if item.has_fields),
        "inexact_text_paragraph_count": sum(1 for item in all_paragraphs if not item.text_exact),
        "paragraph_index_matches_word": True,
    }
    if lock_reason:
        summary["lock_reason"] = lock_reason
    summary["unusual_characters"] = _unusual_characters(snapshot.path)

    if include_comments:
        # Local import avoids a module cycle: comment anchor reconstruction
        # deliberately reuses this module's COM-equivalent paragraph walk.
        from .comments import inspect_comments

        threads = inspect_comments(snapshot)
        summary["comments"] = [thread.to_dict() for thread in threads]
        summary["comment_thread_count"] = len(threads)

    if search_text:
        query_lower = search_text.lower()
        matches = [
            item.to_dict()
            for item in all_paragraphs
            if query_lower in item.text.lower()
        ]
        summary["search_query"] = search_text
        summary["match_count"] = len(matches)
        summary["matches"] = matches[:50]
        summary["search_truncated"] = len(matches) > 50
        # Preserve the legacy field while making its search meaning explicit.
        summary["truncated"] = summary["search_truncated"]
        return summary

    if first_paragraph is None and last_paragraph is None:
        headings = _outline(snapshot, all_paragraphs)
        summary["outline"] = [
            {"index": item.index, "style": item.style, "text": item.text}
            for item in headings
        ]
        summary["outline_empty"] = not headings
        return summary

    start = max(1, int(first_paragraph or 1))
    if start > total:
        raise ValueError(
            f"first_paragraph {start} is past the end of the document ({total} paragraphs)"
        )
    end = int(last_paragraph) if last_paragraph is not None else start + DEFAULT_WINDOW - 1
    end = min(end, total)
    if end < start:
        raise ValueError(f"last_paragraph {end} precedes first_paragraph {start}")

    window_truncated = (end - start + 1) > MAX_WINDOW
    if window_truncated:
        end = start + MAX_WINDOW - 1

    summary["first_paragraph"] = start
    summary["last_paragraph"] = end
    summary["window_truncated"] = window_truncated
    # Preserve the legacy field for clients that already consume it.
    summary["truncated"] = window_truncated
    summary["returned_paragraph_count"] = end - start + 1
    summary["paragraphs"] = [item.to_dict() for item in all_paragraphs[start - 1 : end]]
    return summary
