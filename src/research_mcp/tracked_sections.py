"""Plan 864 S3a: row and section edits as direct XML, tracked or plain (decisions 5, 11, 12).

`replace_table_row` and `replace_section` plan as before (`sections.py`), into text, line-insert
and paragraph-delete steps. This module runs those steps on the XML instead of in Word, reusing the one-paragraph
primitives of `revisions.py`: a text step is `text_edit` (`tracked_replace_paragraph` / `replace_paragraph`), a
line insert is `insert_paragraphs` (before or after its anchor, the new text formatted like the anchor's
character Word would copy), a paragraph delete is `delete_paragraphs` (the last line of a cell Word's way). The
steps run in plan order (bottom-up, so the planner's indices hold) on staged copies with one revision stamp; the
new revision ids are then renumbered to ascend in document order, and the whole op is verified before anything
is published:

- bytes outside the touched rows (or paragraphs) identical to the input;
- every pre-existing revision byte-identical (own insertions edited in place aside, as `text_edit` does them);
- new ids unique in the package, by the author, ascending, one stamp;
- Reject of the new revisions gives the input back (canonical XML; own in-place insertions rejected on both sides);
- Accept of the new revisions (untracked: the result itself) passes the planner's own `verify_plan` and the
  batch consistency check (cell count, no emptied cell, dates).

Plan 864 S3b: a plan that adds or removes a table row runs here too. A `DeleteRowStep` is `delete_rows` (tracked:
Word's row deletion, the row mark and every mark and run deleted; untracked: the row removed). A `DuplicateRowStep`
is `insert_rows` as a plain copy of the template row; the plan's own steps for the copy (its cell lines rewritten,
added or removed) then run plain on it, and, tracked, the finished copy is recorded as one row insertion as Word
records it (`<w:ins/>` in its w:trPr, every paragraph mark and run inserted), so all of its text is inside its own
insertion. What `delete_rows` / `insert_rows` refuse (a revision in the row, merged cells, a nested table, an
object) is refused, named, with the way forward.
"""
from __future__ import annotations

import copy
import datetime as dt
import re
import tempfile
from collections import Counter
from pathlib import Path

from .errors import RefusedUnsafe, VerificationFailed

def _span(doc, layout, first: int, last: int) -> tuple[int, int]:
    """Start and end, in document.xml, of the rows (or paragraphs) laid out as paragraphs first..last."""

    start, end = None, None
    for index in range(first, last + 1):
        item = layout.para(index)
        if item.row_end:
            continue
        node = doc.paragraph(index, "verify")
        outer = node
        for ancestor in node.ancestors():
            if ancestor.name == "w:tr":
                outer = ancestor
        start = outer.start if start is None else min(start, outer.start)
        end = outer.end if end is None else max(end, outer.end)
    if start is None:
        raise VerificationFailed(f"paragraphs {first}..{last} hold no paragraph to edit")
    return start, end


def _context_end(xml: str, end: int) -> int:
    """End of the paragraph right after the span (its next sibling), else `end`.

    A paragraph added or deleted at the span's end has its mark resolved into that following paragraph
    (Reject of an inserted mark, Accept of a deleted one), so the Reject check reads it too; its bytes are
    still checked identical by the outside check."""

    from .revisions import parse

    at = end
    while at < len(xml) and xml[at] in " \t\r\n":
        at += 1
    if not re.match(r"<w:p[\s>/]", xml[at:at + 5]):
        return end
    node = next((item for item in parse(xml).iter() if item.start == at and item.name == "w:p"), None)
    return node.end if node is not None else end


def _renumber(xml: str, first_id: int, author: str) -> tuple[str, dict[str, str]]:
    """New revisions (id >= first_id, by `author`) renumbered first_id.. in document order."""

    from .revisions import REV_TAGS, parse

    root = parse(xml)
    fresh = [node for node in root.iter() if node.local in REV_TAGS and node.name.startswith("w:")
             and (node.attr(xml, "w:id") or "").isdigit() and int(node.attr(xml, "w:id")) >= first_id
             and node.attr(xml, "w:author") == author]
    mapping: dict[str, str] = {}
    pieces, cursor = [], 0
    for number, node in enumerate(fresh):
        old = node.attr(xml, "w:id")
        new = str(first_id + number)
        mapping[old] = new
        head = node.head(xml)
        replaced, count = re.subn(r'(\bw:id\s*=\s*)(["\'])' + re.escape(old) + r"\2", r"\g<1>\g<2>" + new + r"\2",
                                  head, count=1)
        if count != 1:
            raise VerificationFailed(f"could not renumber revision id {old}")
        pieces.append(xml[cursor:node.start] + replaced)
        cursor = node.head_end
    pieces.append(xml[cursor:])
    return "".join(pieces), mapping


def _step_inside(step, span: list[int]) -> bool:
    """Whether a plan step edits only paragraphs of the row copy laid out as paragraphs span[0]..span[1] (its
    row-end marker last): the planner's steps for a copy follow its DuplicateRowStep."""

    from .sections import DeleteParasStep, InsertLinesStep, TextStep

    first, end = span
    if isinstance(step, (TextStep, InsertLinesStep)):
        return first <= step.index < end
    if isinstance(step, DeleteParasStep):
        low = step.previous if step.trailing and step.previous is not None else step.first
        return first <= low and step.last < end
    return False


def _track_copy(source, output, index: int, author: str, now: dt.datetime) -> list[dict]:
    """Record the plain row copy holding paragraph `index` as a row insertion by `author`, as Word records an
    added row (`revisions._track_row`); written to `output`. The whole op is verified afterwards."""

    from . import revisions

    parts, doc = revisions.load(source)
    xml = doc.xml
    _row, node = revisions._table_row(doc, index, "insert_rows", "copy")
    writer = revisions._Writer(revisions.next_revision_id(parts), author, revisions.revision_stamp(xml, now), True)
    piece = revisions._track_row(xml[node.start:node.end], "ins", writer)
    revisions.rezip(source, output, xml[:node.start] + piece + xml[node.end:])
    return writer.added


def run_plan_direct(source, output, plan, *, tracked: bool, author: str, now: dt.datetime | None = None) -> dict:
    """Run a planned row/section edit as direct XML (see the module docstring); verified, then written."""

    from . import revisions
    from .audit import file_sha256
    from .sections import (DeleteParasStep, DeleteRowStep, DuplicateRowStep, InsertLinesStep, TextStep,
                           apply_edits, consistency_errors, read_layout, verify_plan)

    op = plan.op
    author = (author or "").strip() or revisions.DEFAULT_AUTHOR
    if not plan.steps:
        what = "row" if op == "replace_table_row" else "section"
        raise RefusedUnsafe(f"{op}: the {what} already reads as requested; nothing to change")
    now = now or dt.datetime.now(dt.timezone.utc)
    parts, doc = revisions.load(source)
    layout = doc.layout
    first_id = revisions.next_revision_id(parts)
    span_start, span_end = _span(doc, layout, plan.first_index, plan.last_index)
    added: list[dict] = []
    in_place: list[dict] = []
    steps_done: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="research-mcp-section-") as scratch:
        current = Path(source)
        staged_count = 0
        copy_range: list[int] | None = None  # [first, end] of the plain row copy whose own steps are running

        def stage() -> Path:
            nonlocal staged_count
            staged_count += 1
            return Path(scratch) / f"step{staged_count}.docx"

        def finish_copy() -> None:
            nonlocal current, copy_range
            if copy_range is not None and tracked:
                staged = stage()
                added.extend(_track_copy(current, staged, copy_range[0], author, now))
                current = staged
            copy_range = None

        for step in plan.steps:
            in_copy = copy_range is not None and _step_inside(step, copy_range)
            if not in_copy:
                finish_copy()
            step_tracked = tracked and not in_copy
            staged = stage()
            if isinstance(step, TextStep):
                text = apply_edits(step.text, [None] * len(step.text), step.caps, step.edits)[0]
                result = revisions.text_edit(
                    current, staged, op="tracked_replace_paragraph" if step_tracked else "replace_paragraph",
                    paragraph_index=step.index, tracked=step_tracked, author=author, text=text, now=now)
                in_place.extend(result.get("revised_in_place") or [])
                kind = "text"
            elif isinstance(step, InsertLinesStep):
                result = revisions.insert_paragraphs(
                    current, staged, after_paragraph_index=step.index, text="\n".join(step.lines), style_id=None,
                    tracked=step_tracked, author=author, now=now, before=step.where == "before", run_rpr=True)
                kind = f"insert {step.where}"
            elif isinstance(step, DeleteParasStep):
                result = revisions.delete_paragraphs(
                    current, staged, first=step.first, last=step.last, trailing=step.trailing,
                    previous=step.previous, fix_format=step.fix_format, tracked=step_tracked, author=author,
                    now=now)
                kind = "delete"
            elif isinstance(step, DuplicateRowStep):
                # A plain copy first; its own steps follow plain, and `finish_copy` tracks the finished row.
                result = revisions.insert_rows(
                    current, staged, template=step.template, before=step.before, append_after=step.append_after,
                    tracked=False, author=author, now=now)
                kind = "add row"
            elif isinstance(step, DeleteRowStep):
                result = revisions.delete_rows(current, staged, first_indices=[step.index], tracked=tracked,
                                               author=author, now=now)
                kind = "delete row"
            else:  # pragma: no cover - planner and executor are versioned together
                raise RefusedUnsafe(f"{op}: unknown plan step {type(step).__name__}")
            added.extend(result.get("revisions_added") or [])
            steps_done.append({"step": kind, "paragraph_index": getattr(step, "index", getattr(
                step, "first", getattr(step, "template", None)))})
            current = staged
            if isinstance(step, DuplicateRowStep):
                copy_range = [result["new_row"]["first_index"], result["new_row"]["end_index"]]
            elif in_copy:
                copy_range[1] += int(result.get("index_delta") or 0)
        finish_copy()
        staged_xml = revisions.read_parts(current)["word/document.xml"].decode("utf-8")
        new_xml, mapping = _renumber(staged_xml, first_id, author) if tracked else (staged_xml, {})
        for item in added:
            item["id"] = mapping.get(item["id"], item["id"])
        added.sort(key=lambda item: int(item["id"]))
        _verify_whole(op, doc.xml, new_xml, span_start, span_end, first_id, author, tracked,
                      {item["id"] for item in in_place}, len(added), _context_end(doc.xml, span_end))
        output = Path(output)
        revisions.rezip(source, output, new_xml)
        try:
            revisions._verify_written(source, output, new_xml, first_id, op)
            if tracked:
                accepted = Path(scratch) / "accepted.docx"
                revisions.resolve_revisions(output, accepted, mode="accept", ids=[item["id"] for item in added])
            else:
                accepted = output
            verification = verify_plan(source, accepted, plan)
            problems = consistency_errors(source, accepted)
            if problems:
                raise VerificationFailed(f"{op}: " + "; ".join(problems[:5]), actual=problems[:20])
            after_count = len(read_layout(output).paragraphs)
        except Exception:
            output.unlink(missing_ok=True)
            raise
    delta = after_count - len(layout.paragraphs)
    verification["compared_against"] = "source"
    verification["accept_view"] = tracked
    return {
        "source": str(Path(source).resolve()),
        "output": str(output),
        "operation": op,
        **plan.report,
        "engine": "direct_xml",
        "tracked": tracked,
        "author": author if tracked else None,
        "revisions_added": added,
        "revised_in_place": in_place,
        "steps": steps_done,
        "source_sha256": file_sha256(source),
        "output_sha256": file_sha256(output),
        "paragraph_count": after_count,
        "paragraph_delta": delta,
        "index_delta": delta,
        "indices_after_shifted": delta != 0,
        "shift_base": plan.last_index + 1,
        "stale_range": [plan.first_index, plan.last_index] if delta else None,
        "accepted_index_delta": plan.paragraph_delta,
        "protected_structure_unchanged": True,
        "verification": verification,
    }


def _verify_whole(op: str, xml: str, new_xml: str, start: int, end: int, first_id: int, author: str,
                  tracked: bool, in_place: set[str], count: int, context_end: int | None = None) -> None:
    """Decision 12 over the whole op, read from the input and the result only."""

    from .revisions import REV_TAGS, W, _check_stamps, _fragment, _revision_raw, _row_normalized, et_resolve, parse

    tail = len(xml) - end
    if new_xml[:start] != xml[:start] or new_xml[len(new_xml) - tail:] != xml[end:]:
        raise VerificationFailed(f"{op}: bytes outside the edited rows/paragraphs changed")
    old_box, new_box = xml[start:end], new_xml[start:len(new_xml) - tail]
    old_tree, new_tree = parse(old_box), parse(new_box)

    def is_rev(node) -> bool:
        return node.local in REV_TAGS and node.name.startswith("w:")

    def fresh(node) -> bool:
        number = node.attr(new_box, "w:id") or ""
        return number.isdigit() and int(number) >= first_id and node.attr(new_box, "w:author") == author

    def kept(box, tree, fresh_ok):
        return Counter(_revision_raw(box, node) for node in tree.iter()
                       if is_rev(node) and node.attr(box, "w:id") not in in_place and not (fresh_ok and fresh(node)))

    if kept(new_box, new_tree, True) != kept(old_box, old_tree, False):
        raise VerificationFailed(f"{op}: a pre-existing revision changed")
    new_nodes = [node for node in new_tree.iter() if is_rev(node) and fresh(node)]
    if not tracked and new_nodes:
        raise VerificationFailed(f"{op}: an untracked edit added revisions")
    if len(new_nodes) != count:
        raise VerificationFailed(f"{op}: {len(new_nodes)} new revisions in the result; the steps reported {count}")
    numbers = [int(node.attr(new_box, "w:id")) for node in new_nodes]
    if numbers != list(range(first_id, first_id + count)):
        raise VerificationFailed(f"{op}: the new revision ids do not ascend in document order", actual=numbers)
    if new_nodes:
        _check_stamps(new_box, new_nodes, author, op)
    if not tracked:
        return
    new_ids = {str(number) for number in numbers}
    mine = lambda ids: lambda el: el.get(W + "id") in ids and el.get(W + "author") == author  # noqa: E731
    # Reject is read over the span plus the paragraph after it (`_context_end`), unchanged by the outside check.
    context_end = end if context_end is None else context_end
    old_ctx, new_ctx = xml[start:context_end], new_xml[start:len(new_xml) - (len(xml) - context_end)]
    # Rows are read inside a table (rejecting an added row removes it from its table, as Word does).
    wrap = ("<w:tbl>", "</w:tbl>") if re.match(r"<w:tr[\s>]", old_ctx) else ("", "")
    old_el = _fragment(xml, f"<w:body>{wrap[0]}{old_ctx}{wrap[1]}</w:body>")
    new_el = _fragment(new_xml, f"<w:body>{wrap[0]}{new_ctx}{wrap[1]}</w:body>")
    rejected, original = copy.deepcopy(new_el), copy.deepcopy(old_el)
    et_resolve(rejected, mine(new_ids | in_place), "reject")
    if in_place:
        et_resolve(original, mine(in_place), "reject")
    if _row_normalized(rejected) != _row_normalized(original):
        raise VerificationFailed(f"{op}: rejecting the new revisions does not give back the input (D5)")
