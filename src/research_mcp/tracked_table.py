"""Table insertion as direct XML, tracked or plain (plan 916 S5).

``insert_table`` writes the journal-neutral table the Word path (``editing.insert_table_to_version``) makes:
single 0.5 pt borders on the header row's top and bottom and the last row's bottom and no others; the header row
and the descriptive columns left-aligned, the other cells right-aligned; an optional caption paragraph before the
table; and vertical merges of repeated labels (``MergeGroup``, the first cell keeps the label and the others are
emptied). Word never opens, so it works on documents that already have tracked changes.

Tracked, it is written the way Word records a table typed with Track Changes on (lead decision 8, checked
against Word's own output): the anchor paragraph's mark becomes an inserted one (``w:ins`` in ``pPr/rPr``), then
come the caption paragraph (mark and run inserted), the table (``w:ins`` in every row's ``w:trPr``, on every cell
paragraph's mark and around every run) and an empty paragraph that keeps the anchor's original paragraph
properties. Reject removes the table and the caption and merges the anchor back into that empty paragraph, which
gives the input; Accept leaves anchor, caption, table and the empty paragraph after it, as the Word path does.
The table's formatting is part of the inserted table, so it carries no property-change revisions. Merges cannot
be recorded as tracked changes: they are written plainly (``w:vMerge``) and each merged block gets one comment
by the author on its first cell, listed in the result under ``untracked``.

Plain (``tracked=False``), the same table goes in after the anchor with no revision markers; an anchor holding a
section break is split Word's way so the break stays after the table.

Verified before anything is published: bytes outside the anchor paragraph unchanged; every existing revision
byte-identical; the new revisions are ``w:ins`` by the author with ids unique in the package and ascending in
document order; Reject of the new revisions gives the input body (canonical XML); the Accept view, checked
against the request alone, holds the anchor, the caption, the table (cell text, merges, borders, alignment)
and the empty paragraph, with every other body element unchanged; the written file reads as expected and its
bookmarks and comments keep their extents.
"""

from __future__ import annotations

import copy
import datetime as dt
import shutil
import tempfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import W
from .revisions import (
    DEFAULT_AUTHOR,
    REV_TAGS,
    Document,
    Node,
    _check_stamps,
    _copied_ppr,
    _et_final_chars,
    _fragment,
    _mark_rpr_kept,
    _render_text,
    _revision_raw,
    _verify_written,
    _with_ppr,
    _Writer,
    describe_wrapper,
    et_resolve,
    load,
    next_revision_id,
    parse,
    read_parts,
    revision_stamp,
    resolve_revisions,
    rezip,
)
from .tables import MergeGroup

OP = "insert_table"
DEFAULT_TEXT_WIDTH = 9360  # twips: a Letter page with 1" margins, Word's text width when no section says otherwise
_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")
_TBL_PR = ('<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>'
           '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
           'w:noHBand="0" w:noVBand="0"/></w:tblPr>')
MERGE_COMMENT = "Cells in rows {start}-{end} of column {column} were merged (not tracked)"


def _border(side: str) -> str:
    return f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'


def _normalize(text: str) -> str:
    return " ".join(text.split())


# --------------------------------------------------------------------------
# The request: grid, merges and what each cell holds.
# --------------------------------------------------------------------------

def _validate(headers: Sequence[str], rows: Sequence[Sequence[str]], descriptive_columns: Sequence[int],
              merges: Sequence[MergeGroup], caption: str | None) -> list[list[str]]:
    """The table's text grid (header row first); refuses what the Word path refuses, merges included."""

    if not headers or not rows:
        raise RefusedUnsafe("a table requires at least one header and one body row")
    columns = len(headers)
    if any(len(row) != columns for row in rows):
        raise RefusedUnsafe("every body row must have the same width as headers")
    if not descriptive_columns or any(column < 1 or column > columns for column in descriptive_columns):
        raise RefusedUnsafe("descriptive columns must be inside the table width")
    if caption is not None:
        if not caption.strip():
            raise RefusedUnsafe("caption must not be empty; omit it for a table without a caption")
        if "\n" in caption or "\r" in caption:
            raise RefusedUnsafe("caption must be one line")
    grid = [list(headers), *[list(row) for row in rows]]
    if any("\r" in value for row in grid for value in row):
        raise RefusedUnsafe("a cell holds a carriage return; separate lines in a cell with \"\\n\"")
    taken: set[tuple[int, int]] = set()
    for group in merges:
        if group.column < 1 or group.column > columns:
            raise RefusedUnsafe(f"merge column {group.column} out of bounds")
        if group.start_row < 2 or group.end_row > len(grid) or group.start_row >= group.end_row:
            raise RefusedUnsafe(
                f"merge rows {group.start_row}..{group.end_row} invalid for table with {len(grid)} rows")
        texts = [grid[row - 1][group.column - 1] for row in range(group.start_row, group.end_row + 1)]
        for row, text in zip(range(group.start_row, group.end_row + 1), texts):
            if not text.strip():
                raise RefusedUnsafe(f"refusing merge over empty cell at row {row}, column {group.column}")
        if len({_normalize(text) for text in texts}) != 1 or _normalize(texts[0]) != _normalize(group.label):
            raise RefusedUnsafe(
                f"refusing merge: cells in column {group.column} rows {group.start_row}..{group.end_row} "
                f"do not uniformly contain label {group.label!r} (found: {texts})")
        cells = {(row, group.column) for row in range(group.start_row, group.end_row + 1)}
        if cells & taken:
            raise RefusedUnsafe(f"merges overlap in column {group.column}")
        taken |= cells
    return grid


def _cell_lines(grid: list[list[str]], merges: Sequence[MergeGroup]) -> list[list[list[str]]]:
    """Each cell's paragraphs, as the table reads once inserted: a merged block's later cells are empty."""

    emptied = {(row, group.column) for group in merges for row in range(group.start_row + 1, group.end_row + 1)}
    return [[[""] if (r, c) in emptied else value.split("\n") for c, value in enumerate(values, start=1)]
            for r, values in enumerate(grid, start=1)]


def expected_layout(grid: list[list[str]], merges: Sequence[MergeGroup], caption: str | None) -> list[str]:
    """The layout paragraphs the insertion adds after the anchor, in order: the caption, every cell paragraph
    and each row's end marker (empty), then the empty paragraph after the table."""

    out = [caption] if caption is not None else []
    for row in _cell_lines(grid, merges):
        for lines in row:
            out.extend(lines)
        out.append("")
    out.append("")
    return out


def _text_width(doc: Document, node: Node) -> int:
    """The text width of the anchor's section (page width minus left, right and gutter margins)."""

    xml = doc.xml
    section = next((item for item in doc.body.iter() if item.name == "w:sectPr" and item.start >= node.start), None)
    if section is None:
        return DEFAULT_TEXT_WIDTH
    size, margins = section.child("w:pgSz"), section.child("w:pgMar")
    try:
        width = int(size.attr(xml, "w:w")) - int(margins.attr(xml, "w:left")) - int(margins.attr(xml, "w:right"))
        width -= int(margins.attr(xml, "w:gutter") or 0)
    except (AttributeError, TypeError, ValueError):
        return DEFAULT_TEXT_WIDTH
    return width if width > 0 else DEFAULT_TEXT_WIDTH


# --------------------------------------------------------------------------
# Writing.
# --------------------------------------------------------------------------

@dataclass
class _Plan:
    start: int
    end: int
    replacement: str
    word_way: bool
    merge_offsets: dict[int, int]  # merge number -> position of its first cell's paragraph in the new layout block


def _table_xml(lines: list[list[list[str]]], descriptive: set[int], merges: Sequence[MergeGroup], width: int,
               writer: _Writer, tracked: bool, offset: int) -> tuple[str, dict[int, int]]:
    n_rows, columns = len(lines), len(lines[0])
    column_width = width // columns
    restart = {(group.start_row, group.column): number for number, group in enumerate(merges)}
    later = {(row, group.column) for group in merges for row in range(group.start_row + 1, group.end_row + 1)}
    merge_offsets: dict[int, int] = {}
    position = offset

    def mark() -> str:
        return f"<w:rPr><w:ins {writer._attrs('ins', '¶')}/></w:rPr>" if tracked else ""

    out = ["<w:tbl>", _TBL_PR, "<w:tblGrid>", f'<w:gridCol w:w="{column_width}"/>' * columns, "</w:tblGrid>"]
    for r, row in enumerate(lines, start=1):
        out.append("<w:tr>")
        if tracked:
            out.append(f"<w:trPr><w:ins {writer._attrs('ins', '')}/></w:trPr>")
        for c, cell in enumerate(row, start=1):
            tcpr = [f'<w:tcW w:w="{column_width}" w:type="dxa"/>']
            if (r, c) in restart:
                tcpr.append('<w:vMerge w:val="restart"/>')
                merge_offsets[restart[(r, c)]] = position
            elif (r, c) in later:
                tcpr.append("<w:vMerge/>")
            sides = (["top", "bottom"] if r == 1 else []) + (["bottom"] if r == n_rows else [])
            if sides:
                tcpr.append("<w:tcBorders>" + "".join(_border(side) for side in dict.fromkeys(sides)) + "</w:tcBorders>")
            jc = "left" if r == 1 or c in descriptive else "right"
            out.append(f"<w:tc><w:tcPr>{''.join(tcpr)}</w:tcPr>")
            for line in cell:
                out.append(f'<w:p><w:pPr><w:jc w:val="{jc}"/>{mark()}</w:pPr>')
                if line:
                    run = f"<w:r>{_render_text(line, False)}</w:r>"
                    out.append(f"<w:ins {writer._attrs('ins', line)}>{run}</w:ins>" if tracked else run)
                out.append("</w:p>")
                position += 1
            out.append("</w:tc>")
        out.append("</w:tr>")
        position += 1  # the row-end marker
    out.append("</w:tbl>")
    return "".join(out), merge_offsets


def _plan(doc: Document, node: Node, lines, descriptive: set[int], merges: Sequence[MergeGroup],
          caption: str | None, caption_style_id: str | None, writer: _Writer, tracked: bool) -> _Plan:
    xml = doc.xml
    ppr = node.child("w:pPr")
    has_section = ppr is not None and ppr.child("w:sectPr") is not None
    word_way = tracked or has_section

    def mark() -> str:
        return f"<w:ins {writer._attrs('ins', '¶')}/>" if tracked else ""

    pieces: list[str] = []
    if word_way:
        # Word's way (as Word records a tracked table insert): the anchor's mark becomes the inserted one and the
        # empty paragraph after the table keeps the anchor's original pPr, so Reject merges the anchor back.
        anchor_ppr = _copied_ppr(xml, ppr, style_id=None, mark=mark(), mark_rpr=_mark_rpr_kept(xml, ppr))
        pieces.append(_with_ppr(xml, node, anchor_ppr))
        start, end = node.start, node.end
    else:
        start = end = node.end
    offset = 0
    if caption is not None:
        # The caption is a new paragraph made as insert_paragraph makes one with a style (plan 864 S2's helpers):
        # the anchor's pPr with that style and no mark formatting, its mark and run inserted when tracked.
        caption_ppr = _copied_ppr(xml, ppr, style_id=caption_style_id, mark=mark(), mark_rpr="")
        run = f"<w:r>{_render_text(caption, False)}</w:r>"
        body = f"<w:ins {writer._attrs('ins', caption)}>{run}</w:ins>" if tracked else run
        pieces.append(f"<w:p>{caption_ppr}{body}</w:p>")
        offset = 1
    width = _text_width(doc, node)
    table, merge_offsets = _table_xml(lines, descriptive, merges, width, writer, tracked, offset)
    pieces.append(table)
    if word_way:
        trailing = xml[ppr.start:ppr.end] if ppr is not None else ""
    else:
        trailing = _copied_ppr(xml, ppr, style_id=None, mark="", mark_rpr=_mark_rpr_kept(xml, ppr))
    pieces.append(f"<w:p>{trailing}</w:p>" if trailing else "<w:p/>")
    return _Plan(start, end, "".join(pieces), word_way, merge_offsets)


# --------------------------------------------------------------------------
# Verification (from the input and the request only, never from the plan).
# --------------------------------------------------------------------------

def _is_rev(item: Node) -> bool:
    return item.local in REV_TAGS and item.name.startswith("w:")


def _body_element(xml: str) -> tuple[ET.Element, Node]:
    tree = parse(xml)
    body = next((item for item in tree.iter() if item.name == "w:body"), None)
    if body is None:
        raise VerificationFailed(f"{OP}: the result has no w:body")
    return _fragment(xml, xml[body.start:body.end]), body


def _table_is_inserted(table: ET.Element, is_new) -> bool:
    """Every row of `table` carries a new row insertion, every cell paragraph's mark is a new insertion and
    every run sits in a new w:ins (comment markers aside: a merge comment's range and reference go with it)."""

    rows = table.findall(W + "tr")
    if not rows:
        return False
    for row in rows:
        trpr = row.find(W + "trPr")
        marks = [] if trpr is None else [item for item in trpr if item.tag == W + "ins"]
        if len(marks) != 1 or not is_new(marks[0]):
            return False
        for paragraph in row.iter(W + "p"):
            ppr = paragraph.find(W + "pPr")
            rpr = ppr.find(W + "rPr") if ppr is not None else None
            mark = rpr.find(W + "ins") if rpr is not None else None
            if mark is None or not is_new(mark):
                return False
            for child in paragraph:
                if child.tag in (W + "pPr", W + "commentRangeStart", W + "commentRangeEnd"):
                    continue
                if child.tag == W + "ins" and is_new(child):
                    continue
                if child.tag == W + "r" and [part.tag for part in child if part.tag != W + "rPr"] == [
                        W + "commentReference"]:
                    continue
                return False
    return True


def _pstyle(paragraph: ET.Element) -> str | None:
    ppr = paragraph.find(W + "pPr")
    style = ppr.find(W + "pStyle") if ppr is not None else None
    return style.get(W + "val") if style is not None else None


def _props(element: ET.Element | None) -> str:
    from .sections import canonical

    if element is None:
        return ""
    clone = copy.deepcopy(element)
    for rpr in clone.findall(W + "rPr"):
        if not len(rpr) and not rpr.attrib:
            clone.remove(rpr)
    return canonical(clone) if len(clone) or clone.attrib else ""


def _base_ppr(source_ppr: ET.Element | None) -> ET.Element:
    """The anchor's pPr minus w:sectPr, w:pPrChange and the mark's revision markers (decision 9)."""

    base = copy.deepcopy(source_ppr) if source_ppr is not None else ET.Element(W + "pPr")
    for child in list(base):
        if child.tag in (W + "sectPr", W + "pPrChange"):
            base.remove(child)
    mark = base.find(W + "rPr")
    if mark is not None:
        for child in list(mark):
            if child.tag in (W + "ins", W + "del", W + "moveFrom", W + "moveTo", W + "rPrChange"):
                mark.remove(child)
    return base


def _check_accepted_table(table: ET.Element, grid: list[list[str]], descriptive: set[int],
                          merges: Sequence[MergeGroup]) -> None:
    lines = _cell_lines(grid, merges)
    rows = table.findall(W + "tr")
    if len(rows) != len(grid):
        raise VerificationFailed(f"{OP}: the table has {len(rows)} rows after Accept, expected {len(grid)}")
    if any(el.tag.startswith(W) and el.tag[len(W):] in REV_TAGS for el in table.iter()):
        raise VerificationFailed(f"{OP}: the table still holds a revision after Accept")
    tblpr = table.find(W + "tblPr")
    borders = tblpr.find(W + "tblBorders") if tblpr is not None else None
    if borders is not None and any(borders.find(W + side) is not None
                                   and borders.find(W + side).get(W + "val") not in ("nil", "none")
                                   for side in ("left", "right", "insideV")):
        raise VerificationFailed(f"{OP}: the table has vertical borders")
    restart = {(group.start_row, group.column) for group in merges}
    later = {(row, group.column) for group in merges for row in range(group.start_row + 1, group.end_row + 1)}
    for r, (row, want_row) in enumerate(zip(rows, lines), start=1):
        cells = row.findall(W + "tc")
        if len(cells) != len(want_row):
            raise VerificationFailed(f"{OP}: row {r} has {len(cells)} cells, expected {len(want_row)}")
        for c, (cell, want) in enumerate(zip(cells, want_row), start=1):
            where = f"row {r}, column {c}"
            paragraphs = cell.findall(W + "p")
            got = ["".join(ch for ch, _ in _et_final_chars(paragraph)) for paragraph in paragraphs]
            if got != want:
                raise VerificationFailed(f"{OP}: cell {where} does not read as requested")
            tcpr = cell.find(W + "tcPr")
            vmerge = tcpr.find(W + "vMerge") if tcpr is not None else None
            state = None if vmerge is None else (vmerge.get(W + "val") or "continue")
            want_state = "restart" if (r, c) in restart else "continue" if (r, c) in later else None
            if state != want_state:
                raise VerificationFailed(f"{OP}: cell {where} merge is {state}, expected {want_state}")
            tcb = tcpr.find(W + "tcBorders") if tcpr is not None else None
            have = {_local(side.tag) for side in tcb} if tcb is not None else set()
            sides = ({"top", "bottom"} if r == 1 else set()) | ({"bottom"} if r == len(grid) else set())
            if have != sides or any(side.get(W + "val") != "single" for side in (tcb if tcb is not None else [])):
                raise VerificationFailed(f"{OP}: cell {where} borders are {sorted(have)}, expected {sorted(sides)}")
            jc = "left" if r == 1 or c in descriptive else "right"
            for paragraph in paragraphs:
                ppr = paragraph.find(W + "pPr")
                value = ppr.find(W + "jc") if ppr is not None else None
                if value is None or value.get(W + "val") != jc:
                    raise VerificationFailed(f"{OP}: cell {where} is not aligned {jc}")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _verify(xml: str, new_xml: str, doc: Document, node: Node, plan_start: int, plan_end: int, *, first_id: int,
            author: str, tracked: bool, word_way: bool, grid, descriptive: set[int], merges, caption: str | None,
            caption_style_id: str | None, final: bool = False) -> list[Node]:
    """Decision-12 checks of `new_xml` against the input `xml`; returns the new revisions (document order).
    `final` is the check of the written file after the merge comments were added (markers in the table)."""

    from .sections import canonical

    if not final:
        tail = len(xml) - plan_end
        if new_xml[:plan_start] != xml[:plan_start] or new_xml[len(new_xml) - tail:] != xml[plan_end:]:
            raise VerificationFailed(f"{OP}: bytes outside the anchor paragraph changed")
    new_tree = parse(new_xml)
    new_body = next(item for item in new_tree.iter() if item.name == "w:body")
    old_revs = Counter(_revision_raw(xml, item) for item in doc.body.iter() if _is_rev(item))

    def fresh(item: Node) -> bool:
        number = item.attr(new_xml, "w:id") or ""
        return tracked and number.isdigit() and int(number) >= first_id and item.attr(new_xml, "w:author") == author

    added = [item for item in new_body.iter() if _is_rev(item) and fresh(item)]
    kept = Counter(_revision_raw(new_xml, item) for item in new_body.iter() if _is_rev(item) and not fresh(item))
    if not final and kept != old_revs:
        raise VerificationFailed(f"{OP}: a pre-existing revision changed")
    if final and kept - old_revs:
        raise VerificationFailed(f"{OP}: a pre-existing revision changed when the merge comments were added")
    if any(item.name != "w:ins" for item in added):
        raise VerificationFailed(f"{OP}: a new revision is not an insertion")
    if [int(item.attr(new_xml, "w:id")) for item in added] != list(range(first_id, first_id + len(added))):
        raise VerificationFailed(f"{OP}: the new revision ids are not unique and ascending in document order")
    _check_stamps(new_xml, added, author, OP)
    if tracked and not added:
        raise VerificationFailed(f"{OP}: no revision was added")

    old_el = _fragment(xml, xml[doc.body.start:doc.body.end])
    new_el = _fragment(new_xml, new_xml[new_body.start:new_body.end])
    new_ids = {item.attr(new_xml, "w:id") for item in added}
    is_new = lambda el: tracked and el.get(W + "id") in new_ids and el.get(W + "author") == author  # noqa: E731
    anchor_at = doc.body.children.index(node)
    table_at = anchor_at + 1 + (caption is not None)
    children = list(new_el)
    if table_at >= len(children) or children[table_at].tag != W + "tbl":
        raise VerificationFailed(f"{OP}: the new table is not where it was asked for")
    accepted = new_el
    if tracked:
        rejected = copy.deepcopy(new_el)
        table = list(rejected)[table_at]
        if not _table_is_inserted(table, is_new):
            raise VerificationFailed(f"{OP}: the table is not one tracked insertion (rows, marks and runs)")
        # Reject removes the whole table (every row's insertion rejected: Word's rule, as resolve_revisions
        # does), and with it the merge comments wholly inside it (final check), as Word deletes them.
        et_resolve(rejected, is_new, "reject")
        if canonical(rejected) != canonical(old_el):
            raise VerificationFailed(f"{OP}: rejecting the new table does not give back the input")
        accepted = copy.deepcopy(new_el)
        et_resolve(accepted, is_new, "accept")
    if final:
        return added  # the comment markers in the table were checked by comment_ops; the rest is unchanged
    old_children, got = list(old_el), list(accepted)
    extra = 2 + (caption is not None)
    if len(got) != len(old_children) + extra:
        raise VerificationFailed(f"{OP}: the body does not hold exactly the new table, caption and paragraph")
    new_block = got[anchor_at + 1:anchor_at + 1 + extra]
    others = got[:anchor_at + 1] + got[anchor_at + 1 + extra:]
    source_ppr = old_children[anchor_at].find(W + "pPr")
    base = _base_ppr(source_ppr)
    for number, (before, after) in enumerate(zip(old_children, others)):
        if word_way and number == anchor_at:
            strip = lambda el: [canonical(child) for child in el if child.tag != W + "pPr"]  # noqa: E731
            if after.tag != W + "p" or strip(after) != strip(before):
                raise VerificationFailed(f"{OP}: the anchor paragraph's content changed")
            if _props(after.find(W + "pPr")) != _props(base):
                raise VerificationFailed(f"{OP}: the anchor paragraph's properties are not its own minus the section "
                                         "break, property change and mark revisions (Word's way)")
            continue
        if canonical(after) != canonical(before):
            raise VerificationFailed(f"{OP}: another body element changed (Accept view)")
    if caption is not None:
        paragraph = new_block[0]
        if paragraph.tag != W + "p" or _pstyle(paragraph) != caption_style_id:
            raise VerificationFailed(f"{OP}: the caption paragraph does not carry style {caption_style_id!r}")
        if any(child.tag not in (W + "pPr", W + "r") for child in paragraph):
            raise VerificationFailed(f"{OP}: the caption paragraph holds something other than runs")
        if "".join(ch for ch, _ in _et_final_chars(paragraph)) != caption:
            raise VerificationFailed(f"{OP}: the caption does not read as requested")
    _check_accepted_table(new_block[-2], grid, descriptive, merges)
    trailing = new_block[-1]
    if trailing.tag != W + "p" or any(child.tag != W + "pPr" for child in trailing):
        raise VerificationFailed(f"{OP}: the paragraph after the table is not empty")
    want = _props(source_ppr) if word_way else _props(base)
    if _props(trailing.find(W + "pPr")) != want:
        raise VerificationFailed(f"{OP}: the paragraph after the table does not take the anchor's properties")
    return added


# --------------------------------------------------------------------------
# The operation.
# --------------------------------------------------------------------------

def _body(xml: str) -> ET.Element:
    return ET.fromstring(xml.encode("utf-8")).find(W + "body")


def insert_table(source, output, *, after_paragraph_index: int, headers: Sequence[str],
                 rows: Sequence[Sequence[str]], descriptive_columns: Sequence[int],
                 merges: Sequence[MergeGroup] = (), caption: str | None = None, caption_style_id: str | None = None,
                 caption_style_name: str | None = None, tracked: bool, author: str = DEFAULT_AUTHOR,
                 now: dt.datetime | None = None) -> dict:
    """Insert a journal-neutral table (and caption) after body paragraph `after_paragraph_index` as direct XML;
    tracked, as Word records a tracked table insert, with one comment per merged block (see the module doc)."""

    from .audit import audit_docx, file_sha256
    from .comment_ops import add_comment
    from .sections import StreamEdit, _paragraph_starts, anchor_errors_for_edits, canonical, read_layout

    author = (author or "").strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{OP}: author must not be empty")
    merges = tuple(merges)
    grid = _validate(headers, rows, descriptive_columns, merges, caption)
    if caption is not None and not caption_style_id:
        raise RefusedUnsafe(f"{OP}: a caption needs a paragraph style id")
    output = Path(output)
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    descriptive = set(descriptive_columns)
    parts, doc = load(source)
    index = after_paragraph_index
    node = doc.paragraph(index, OP)
    para = doc.layout.para(index)
    if para.nested:
        raise RefusedUnsafe(f"{OP}: paragraph {index} is inside a nested table; a new table cannot be anchored "
                            "inside an existing table")
    if para.table is not None or any(item.name == "w:tc" for item in node.ancestors()):
        where = f"table {para.table + 1}" if para.table is not None else "a table"  # layout tables count from 0
        raise RefusedUnsafe(f"{OP}: paragraph {index} is inside {where}; a new table cannot be "
                            "anchored inside an existing table")
    if node.parent is None or node.parent.name != "w:body":
        raise RefusedUnsafe(f"{OP}: paragraph {index} is inside {node.parent.name if node.parent else 'nothing'} "
                            "(a content control or other container), not the document body; insert the table in Word")
    xml = doc.xml
    ppr = node.child("w:pPr")
    mark_rpr = ppr.child("w:rPr") if ppr is not None else None
    for item in (mark_rpr.children if mark_rpr is not None else []):
        if item.name in ("w:del", "w:moveFrom"):
            raise RefusedUnsafe(
                f"{OP}: paragraph {index}'s mark is a tracked deletion ({describe_wrapper(xml, item)}), so it "
                "merges with the next paragraph; accept or reject that revision first, or insert after "
                "another paragraph")
    before_audit = audit_docx(source)
    first_id = next_revision_id(parts)
    writer = _Writer(first_id, author, revision_stamp(xml, now), tracked)
    lines = _cell_lines(grid, merges)
    plan = _plan(doc, node, lines, descriptive, merges, caption, caption_style_id, writer, tracked)
    new_xml = xml[:plan.start] + plan.replacement + xml[plan.end:]
    checks = dict(first_id=first_id, author=author, tracked=tracked, word_way=plan.word_way, grid=grid,
                  descriptive=descriptive, merges=merges, caption=caption, caption_style_id=caption_style_id)
    _verify(xml, new_xml, doc, node, plan.start, plan.end, **checks)

    new_texts = expected_layout(grid, merges, caption)
    before_lengths = [len(item.text) for item in doc.layout.paragraphs]
    lengths = before_lengths[:index] + [len(text) for text in new_texts] + before_lengths[index:]
    position = (_paragraph_starts(before_lengths)[index] if index < len(before_lengths)
                else sum(length + 1 for length in before_lengths))
    stream = [StreamEdit(position, 0, sum(len(text) + 1 for text in new_texts), either_side=True)]

    work = Path(tempfile.mkdtemp(prefix="research_mcp_table_"))
    try:
        staged = work / "table.docx"
        rezip(source, staged, new_xml)
        _verify_written(source, staged, new_xml, first_id, OP)
        after_layout = read_layout(staged)
        texts = [item.text for item in after_layout.paragraphs]
        expected = [item.text for item in doc.layout.paragraphs]
        expected[index:index] = new_texts
        if texts != expected:
            raise VerificationFailed(f"{OP}: the document does not read as requested after the insert",
                                     expected=new_texts, actual=texts[index:index + len(new_texts)])
        block = after_layout.paragraphs[index:index + len(new_texts)]
        row_ends = [number for number, item in enumerate(block) if item.row_end]
        if len(row_ends) != len(grid) or any(item.table is None for item in block[(caption is not None):-1]):
            raise VerificationFailed(f"{OP}: the new table does not lay out as {len(grid)} rows")
        problems = anchor_errors_for_edits(doc.layout, after_layout, stream, lengths)
        if problems:
            raise VerificationFailed(f"{OP}: bookmark/comment anchors moved: " + "; ".join(problems[:5]))

        untracked: list[dict] = []
        current = staged
        if tracked:
            for number, group in enumerate(merges):
                paragraph_index = index + 1 + plan.merge_offsets[number]
                anchor = lines[group.start_row - 1][group.column - 1][0]
                text = MERGE_COMMENT.format(start=group.start_row, end=group.end_row, column=group.column)
                target = work / f"merge{number + 1}.docx"
                done = add_comment(current, target, paragraph_index=paragraph_index, anchor=anchor, text=text,
                                   author=author)
                untracked.append({"kind": "merge", "column": group.column, "start_row": group.start_row,
                                  "end_row": group.end_row, "label": group.label, "comment_id": done["comment_id"],
                                  "paragraph_index": paragraph_index, "text": text})
                current = target
        if current is not staged:
            final_xml = read_parts(current)["word/document.xml"].decode("utf-8")
            added = _verify(xml, final_xml, doc, node, plan.start, plan.end, final=True, **checks)
            # The real reject (resolve_revisions) of the written file removes the table and its merge comments
            # and gives the input body and comment count back.
            rejected = work / "rejected.docx"
            resolution = resolve_revisions(current, rejected, mode="reject",
                                           ids=[int(item.attr(final_xml, "w:id")) for item in added])
            gone = sorted(item["comment_id"] for item in resolution.get("comments_removed", []))
            rejected_xml = read_parts(rejected)["word/document.xml"].decode("utf-8")
            if (gone != sorted(item["comment_id"] for item in untracked)
                    or canonical(_body(rejected_xml)) != canonical(_body(xml))
                    or audit_docx(rejected).live_comment_count != before_audit.live_comment_count):
                raise VerificationFailed(f"{OP}: rejecting the new table and its merge comments does not give "
                                         "back the input")
        after_audit = audit_docx(current)
        protected = ("bookmark_names", "field_count", "zotero_item_fields", "zotero_bibliography_fields",
                     "drawing_count", "live_media_count")
        changed = {name: (getattr(before_audit, name), getattr(after_audit, name)) for name in protected
                   if getattr(before_audit, name) != getattr(after_audit, name)}
        if after_audit.table_count != before_audit.table_count + 1:
            changed["table_count"] = (before_audit.table_count, after_audit.table_count)
        if after_audit.live_comment_count != before_audit.live_comment_count + len(untracked):
            changed["live_comment_count"] = (before_audit.live_comment_count, after_audit.live_comment_count)
        if changed:
            raise VerificationFailed(f"{OP}: table insertion changed protected structure", expected="unchanged",
                                     actual=changed)
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists():
            raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
        shutil.copyfile(current, output)
    finally:
        shutil.rmtree(work, ignore_errors=True)

    table_index = 1 + sum(1 for item in doc.body.iter() if item.name == "w:tbl" and item.start < node.start
                          and not any(parent.name == "w:tbl" for parent in item.ancestors()))
    return {
        "source": str(Path(source).resolve()),
        "output": str(output.resolve()),
        "operation": OP,
        "engine": "direct_xml",
        "table_index": table_index,
        "rows": len(grid),
        "columns": len(grid[0]),
        "caption": caption,
        "caption_style": caption_style_name if caption is not None else None,
        "source_sha256": before_audit.sha256,
        "output_sha256": file_sha256(output),
        "paragraph_count": after_audit.paragraph_count,
        "index_delta": len(new_texts),
        "indices_after_shifted": True,
        "stale_from_index": index + 1,
        "new_paragraph_index": index + 1,
        "protected_structure_unchanged": True,
        "tracked": tracked,
        "author": author if tracked else None,
        "revisions_added": writer.added if tracked else [],
        "anchor_mark_moved": plan.word_way,
        "untracked": untracked,
    }
