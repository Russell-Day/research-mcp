"""Delete whole paragraphs, tracked or plain, in the body or a table cell (plan 864 S3c).

The `delete_paragraph` edit_document op. Direct OOXML, Word never opens, on documents with or without tracked
changes: `revisions.delete_paragraphs` (S3a's primitive) with `keep_deletions`, so a paragraph whose text is
already a tracked deletion (any author), but whose mark is not, can be deleted too: its old `w:del` stays
byte-identical, the live text around it goes into new `w:del`s by the op's author and the mark is deleted.

A range lies in one container (the body, or one table cell) and is consecutive. The last line(s) of a cell
are deleted Word's way (the mark of the line before goes, the cell's last mark stays); a cell keeps at least
one paragraph. Tracked: paragraph indices do not change until the deletion is accepted (`index_delta` 0,
`accepted_index_delta` negative). Untracked: the paragraphs are gone (`index_delta` negative) and later
paragraph-indexed ops in the batch are shifted.
"""

from __future__ import annotations

import datetime as dt

from . import revisions
from .errors import RefusedUnsafe

OP = "delete_paragraph"


def _where(layout, index: int) -> str:
    item = layout.para(index)
    if item.table is None:
        return "the body"
    return f"table {item.table + 1} row {item.row + 1} cell {item.cell + 1}"


def delete_paragraph(source, output, *, paragraph_index: int, through_paragraph_index: int | None = None,
                     tracked: bool = True, author: str = revisions.DEFAULT_AUTHOR,
                     now: dt.datetime | None = None) -> dict:
    first = int(paragraph_index)
    last = int(through_paragraph_index) if through_paragraph_index is not None else first
    if last < first:
        raise RefusedUnsafe(f"{OP}: through_paragraph_index ({last}) is before paragraph_index ({first})")
    _, doc = revisions.load(source)
    layout = doc.layout
    count = len(layout.paragraphs)
    for index in (first, last):
        if not 1 <= index <= count:
            raise RefusedUnsafe(f"{OP}: paragraph_index {index} is outside 1..{count}")
    for index in range(first, last + 1):
        item = layout.para(index)
        if item.row_end:
            raise RefusedUnsafe(f"{OP}: paragraph {index} is a table row-end marker, not a paragraph; a range must "
                                "lie in the body or in one table cell (give one op per cell)")
        if item.nested:
            raise RefusedUnsafe(f"{OP}: paragraph {index} is inside a nested table; delete it in Word")
    home = (layout.para(first).table, layout.para(first).row, layout.para(first).cell)
    for index in range(first + 1, last + 1):
        item = layout.para(index)
        if (item.table, item.row, item.cell) != home:
            raise RefusedUnsafe(
                f"{OP}: paragraphs {first}..{last} do not lie in one container: paragraph {index - 1} is in "
                f"{_where(layout, index - 1)} and paragraph {index} in {_where(layout, index)}; give one op per "
                "cell (or per run of body paragraphs)")
    trailing, previous = False, None
    if home[0] is not None:
        cell = layout.row_of(first).cells[home[2]]
        if first <= cell[0] and last >= cell[-1]:
            where = _where(layout, first)
            raise RefusedUnsafe(f"{OP}: paragraphs {first}..{last} are all of {where}'s paragraphs; a cell keeps at "
                                "least one paragraph (delete the row, or keep one line)")
        if last == cell[-1]:
            trailing, previous = True, first - 1
    try:
        result = revisions.delete_paragraphs(
            source, output, first=first, last=last, trailing=trailing, previous=previous, fix_format=trailing,
            tracked=tracked, author=author, now=now, keep_deletions=True)
    except RefusedUnsafe as error:
        if trailing and "(the line before, whose mark is deleted) holds a tracked revision" in str(error):
            # the last lines of a cell deleted one op at a time: the line before is already deleted
            raise RefusedUnsafe(f"{error}. To delete the last lines of a cell, give them in one op "
                                f"(paragraph_index..through_paragraph_index ending at {last}), not one op per "
                                "line") from error
        raise
    texts = result.pop("deleted_texts")
    removed = last - first + 1
    result.update(
        operation=OP,
        engine="direct_xml",
        deleted=[{"paragraph_index": index, "text": text} for index, text in zip(range(first, last + 1), texts)],
        accepted_index_delta=-removed,
        container=_where(layout, first),
    )
    if not tracked:
        # the paragraphs after the range move up; the deleted ones no longer exist (edit_batch refuses them)
        result["index_shifts"] = [[last + 1, -removed]]
        result["deleted_range"] = [first, last]
    return result


def apply_operation(source, output, operation) -> dict:
    return delete_paragraph(
        source, output, paragraph_index=operation.paragraph_index,
        through_paragraph_index=operation.through_paragraph_index,
        tracked=True if operation.tracked is None else operation.tracked,
        author=operation.author or revisions.DEFAULT_AUTHOR)
