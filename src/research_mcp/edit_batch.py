"""Verified-partial orchestration for bounded document edit batches."""

from __future__ import annotations

from contextlib import contextmanager
import difflib
import json
import os
from pathlib import Path
import shutil
import tempfile

from .audit import audit_docx, file_sha256, document_structure_errors, is_file_locked
from .edit_models import (
    AddCommentOperation,
    COMMENT_OPERATION_TYPES,
    EditBatchResult,
    EditDocumentRequest,
    EditOperation,
    InsertParagraphOperation,
    InsertTableOperation,
    InsertTextOperation,
    OperationError,
    OperationResult,
    GroupResult,
    ProtectedResult,
    READ_ONLY_OPERATION_TYPES,
    REVISION_OPERATION_TYPES,
    ReplaceParagraphOperation,
    ReplaceSectionOperation,
    ReplaceTableRowOperation,
    ReplaceTextOperation,
    ReplyCommentOperation,
    RetryInfo,
)
from .editing import (
    add_comment_to_version,
    insert_paragraph_to_version,
    insert_table_to_version,
    insert_text_to_version,
    replace_paragraph_to_version,
    replace_section_to_version,
    replace_table_row_to_version,
    replace_text_to_version,
    reply_comment_to_version,
)
from .errors import IntegrationRefused, RefusedUnsafe, VerificationFailed
from .tables import MergeGroup
from .publication import note_versioned, publish_inplace
from .ooxml import zotero_package_fields


@contextmanager
def _exclusive_output(output: Path, *, is_inplace: bool = False):
    """Reserve one output name for the whole batch without creating a DOCX."""

    output.parent.mkdir(parents=True, exist_ok=True)
    if not is_inplace and output.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    lock = output.with_name(output.name + ".research-mcp.lock")
    try:
        descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError as exc:
        raise RefusedUnsafe(f"output is reserved by another edit batch: {output}") from exc
    try:
        os.write(descriptor, f"pid={os.getpid()}\n".encode("ascii"))
        yield
    finally:
        if descriptor >= 0:
            os.close(descriptor)
        try:
            lock.unlink()
        except FileNotFoundError:
            pass


def _publish(source: Path, output: Path, *, is_inplace: bool = False, expected_sha256: str | None = None,
             op: str | None = None) -> dict:
    """Publish with create-new semantics or atomic in-place replacement (pre-edit state checkpointed)."""

    if is_inplace:
        if expected_sha256 is None:
            raise RefusedUnsafe("In-place publication requires the inspected source hash.")
        return publish_inplace(source, output, expected_sha256=expected_sha256, tool="edit_document", op=op)
    created = False
    try:
        with source.open("rb") as reader, output.open("xb") as writer:
            created = True
            shutil.copyfileobj(reader, writer)
    except Exception:
        if created and output.exists():
            output.unlink(missing_ok=True)
        raise
    if expected_sha256 is None:
        expected_sha256 = file_sha256(source)
    return note_versioned(expected_sha256, output)


def _error(exc: Exception) -> OperationError:
    return OperationError(
        message=str(exc) or exc.__class__.__name__,
        code=exc.__class__.__name__,
        retryable=False,
    )


def _mapped_index(value: int, shifts: list[list[tuple[int, int]]]) -> int:
    """Map a source paragraph index to the current staged document.

    `shifts` holds one transform per applied op, in batch order. Each transform's (base, delta) pairs are in
    that op's own input coordinates and apply simultaneously (an accept that merges three paragraphs gives
    three pairs, all against the same input), so the transforms are composed in order: each one sees the
    index as the previous ops left it."""

    for transform in shifts:
        value += sum(delta for base, delta in transform if base <= value)
    return value


def _mapped_operation(
    operation: EditOperation, shifts: list[list[tuple[int, int]]]
) -> EditOperation:
    if not any(shifts):
        return operation

    data = operation.model_dump()
    for field in ("paragraph_index", "after_paragraph_index", "through_paragraph_index"):
        if data.get(field) is not None:
            data[field] = _mapped_index(int(data[field]), shifts)
    return operation.__class__.model_validate(data)


def _clean_operation(operation: EditOperation, layout, current: Path) -> tuple[EditOperation, list[dict]]:
    """Plan 551: clean every string the operation writes into the document (textclean, honouring `keep`).

    Cleaned: inserted/replacement text, paragraph text, table cells and headers (and merge labels, alike),
    comment and reply text. Not cleaned: find/match/anchor/heading strings (they must match the document),
    authors and initials, style names. Replacing a known text (replace_text's `find`, a paragraph's current
    text) cleans only what differs from it, so characters echoed from the document stay as they are."""

    from .revisions import clean_new_text

    keep = tuple(getattr(operation, "keep", ()) or ())
    if operation.type not in _CLEANED_FIELDS:
        return operation, []
    data = operation.model_dump()
    found: list[list[dict]] = []

    def one(text, *, old=None, single=True):
        if not isinstance(text, str):
            return text
        new, changes = clean_new_text(text, keep, old=old, single_paragraph=single)
        found.append(changes)
        return new

    def cells(rows):
        return None if rows is None else [[one(cell, single=False) for cell in row] for row in rows]

    kind = operation.type
    if kind in ("replace_text", "tracked_replace_text"):
        data["replace"] = one(data["replace"], old=data["find"])
    elif kind in ("replace_paragraph", "tracked_replace_paragraph"):
        old = None
        try:
            old = (layout if layout is not None else _read_layout(current)).para(data["paragraph_index"]).text
        except Exception:
            old = None
        data["text"] = one(data["text"], old=old)
    elif kind in ("insert_text", "insert_paragraph"):
        data["text"] = one(data["text"])
    elif kind == "insert_table":
        data["headers"] = [one(item, single=False) for item in data["headers"]]
        data["rows"] = cells(data["rows"])
        for merge in data["merges"]:
            if isinstance(merge, dict) and "label" in merge:
                merge["label"] = one(merge["label"], single=False)
    elif kind == "add_comment":
        data["text"] = one(data["text"], single=False)
    elif kind == "reply_comment":
        data["reply_text"] = one(data["reply_text"], single=False)
    elif kind == "replace_table_row":
        data["cells"] = [one(item, single=False) for item in data["cells"]]
    elif kind == "replace_section":
        data["rows"] = cells(data["rows"])
        if data["paragraphs"] is not None:
            data["paragraphs"] = [one(item) for item in data["paragraphs"]]
    from .textclean import merge_changes

    changes = merge_changes(*found)
    if not changes:
        return operation, []
    return operation.__class__.model_validate(data), changes


def _read_layout(path: Path):
    from .sections import read_layout

    return read_layout(path)


_CLEANED_FIELDS = ("replace_text", "tracked_replace_text", "replace_paragraph", "tracked_replace_paragraph",
                   "insert_text", "insert_paragraph", "insert_table", "add_comment", "reply_comment",
                   "replace_table_row", "replace_section")


def _apply_operation(source: Path, output: Path, operation: EditOperation) -> dict[str, object]:
    from .operation_process import run_operation
    return run_operation("edit", {"source": str(source), "output": str(output), "operation": operation.model_dump(mode="json")})


def _apply_operation_local(
    source: Path, output: Path, operation: EditOperation
) -> dict[str, object]:
    if operation.type in REVISION_OPERATION_TYPES:
        # Plan 501: tracked changes are written as direct XML, never through Word (D4).
        from .revisions import apply_operation

        return apply_operation(source, output, operation)
    if isinstance(operation, ReplaceTextOperation):
        if not operation.match_case or operation.occurrence != 1 or operation.expected_count != 1:
            raise RefusedUnsafe("Only one unique, case-sensitive match is supported; inspect and use a unique anchor.")
        return replace_text_to_version(
            source,
            output,
            paragraph_index=operation.paragraph_index,
            find=operation.find,
            replace=operation.replace,
        )
    if isinstance(operation, ReplaceParagraphOperation):
        return replace_paragraph_to_version(
            source, output, paragraph_index=operation.paragraph_index, text=operation.text,
        )
    if isinstance(operation, ReplaceTableRowOperation):
        return replace_table_row_to_version(
            source, output, match=operation.match, cells=operation.cells, table_index=operation.table_index,
        )
    if isinstance(operation, ReplaceSectionOperation):
        return replace_section_to_version(
            source, output, heading=operation.heading, paragraphs=operation.paragraphs,
            rows=operation.rows, until_heading=operation.until_heading,
        )
    if isinstance(operation, InsertTextOperation):
        return insert_text_to_version(
            source,
            output,
            paragraph_index=operation.paragraph_index,
            position=operation.position,
            anchor=operation.anchor,
            text=operation.text,
        )
    if isinstance(operation, InsertParagraphOperation):
        return insert_paragraph_to_version(
            source,
            output,
            after_paragraph_index=operation.after_paragraph_index,
            text=operation.text,
            style=operation.style,
        )
    if isinstance(operation, InsertTableOperation):
        return insert_table_to_version(
            source,
            output,
            after_paragraph_index=operation.after_paragraph_index,
            headers=tuple(operation.headers),
            rows=tuple(tuple(row) for row in operation.rows),
            descriptive_columns=tuple(operation.descriptive_columns),
            merges=tuple(
                MergeGroup(
                    column=merge["column"],
                    start_row=merge["start_row"],
                    end_row=merge["end_row"],
                    label=merge["label"],
                )
                for merge in operation.merges
            ),
        )
    if operation.type in COMMENT_OPERATION_TYPES:
        # Plan 562 S1: comments are written as direct XML (comment_ops), never through Word, so they work on
        # documents with tracked changes and leave every other thread byte-identical.
        from .comment_ops import apply_operation as apply_comment_operation

        return apply_comment_operation(source, output, operation)
    raise RefusedUnsafe(f"unsupported operation type {operation.type!r}")


def _format_diff(operation: EditOperation) -> str:
    """Produce a concise, human-readable diff summary for an edit operation."""
    op_type = operation.type
    if op_type == "replace_text":
        return f"P{getattr(operation, 'paragraph_index', '?')}: - '{getattr(operation, 'find', '')}' -> + '{getattr(operation, 'replace', '')}'"
    elif op_type == "insert_text":
        return f"P{getattr(operation, 'paragraph_index', '?')}: inserted '{getattr(operation, 'text', '')}' at {getattr(operation, 'position', 'anchor')}"
    elif op_type == "insert_paragraph":
        return f"After P{getattr(operation, 'after_paragraph_index', '?')}: + paragraph '{getattr(operation, 'text', '')[:40]}...'"
    elif op_type == "insert_table":
        cols = len(getattr(operation, "headers", []))
        rows = len(getattr(operation, "rows", []))
        return f"After P{getattr(operation, 'after_paragraph_index', '?')}: + table ({rows} rows x {cols} cols)"
    elif op_type == "add_comment":
        return f"P{operation.paragraph_index}: added comment on '{operation.anchor}'"
    elif op_type == "replace_paragraph":
        return f"P{operation.paragraph_index}: = '{operation.text[:60]}'"
    elif op_type == "replace_table_row":
        return f"Row matching '{operation.match[:40]}': {len(operation.cells)} cells rewritten"
    elif op_type == "replace_section":
        count = len(operation.rows) if operation.rows is not None else len(operation.paragraphs or [])
        unit = "rows" if operation.rows is not None else "paragraphs"
        return f"Section '{operation.heading}': rewritten as {count} {unit}"
    elif op_type in ("tracked_replace_text", "tracked_replace_paragraph"):
        change = (f"- '{operation.find}' -> + '{operation.replace}'" if op_type == "tracked_replace_text"
                  else f"= '{operation.text[:60]}'")
        return f"P{operation.paragraph_index}: tracked ({operation.author}) {change}"
    elif op_type == "sweep_characters":
        span = (f"P{operation.paragraph_index or 1}-P{operation.through_paragraph_index or operation.paragraph_index}"
                if operation.paragraph_index or operation.through_paragraph_index else "whole document")
        verb = "previewed" if operation.preview_only else "tracked"
        return f"Characters swept ({span}): {verb} ({operation.author})"
    elif op_type in ("accept_revisions", "reject_revisions", "list_revisions"):
        scope = [f"ids {operation.ids}" if operation.ids else "",
                 f"author {operation.author!r}" if operation.author else "",
                 f"P{operation.paragraph_index or operation.through_paragraph_index}"
                 + (f"-P{operation.through_paragraph_index}" if operation.through_paragraph_index else "")
                 if operation.paragraph_index or operation.through_paragraph_index else "",
                 f"section '{operation.heading}'" if operation.heading else ""]
        verb = {"accept_revisions": "accepted", "reject_revisions": "rejected", "list_revisions": "listed"}[op_type]
        return f"Revisions {verb}: " + (", ".join(item for item in scope if item) or "all")
    elif op_type == "reply_comment":
        resolved = " and resolved" if getattr(operation, "resolve", False) else ""
        return (f"Comment #{getattr(operation, 'comment_id', '?')}: {operation.disposition} reply{resolved} "
                f"'{getattr(operation, 'reply_text', '')[:40]}...'")
    elif op_type in ("resolve_comment", "reopen_comment"):
        return f"Comment #{operation.comment_id}: {'resolved' if op_type == 'resolve_comment' else 'reopened'}"
    elif op_type == "delete_comment":
        what = "thread deleted with its replies" if operation.include_replies else "reply deleted"
        return f"Comment #{operation.comment_id}: {what}"
    return f"{op_type} applied"


class _Groups:
    """All-or-nothing groups (D10): explicit `group`, same table row, same section.

    A row or section is identified twice: in the source as inspected (the
    coordinate system of every requested paragraph_index) and in the staged
    document each operation actually runs against, where an earlier edit may
    have changed the text a later `match`/`heading` names. Staged rows carry
    their source identity forward (`_advance_row_ids`), so "Delta Park
    Center" matched after op 1 renamed it is still table 1 row 5. A row or
    section operation found in neither is `unresolved`: allow_partial then
    refuses to publish a partial batch rather than guess its group.
    """

    def __init__(self, source: Path, operations: list[EditOperation]) -> None:
        from .sections import read_layout

        self.operations = operations
        self.parent = list(range(len(operations)))
        self.labels: dict[int, list[str]] = {}
        self.ranges: dict[int, tuple[int, int]] = {}  # section/row ops: source paragraph span they rewrite
        self.keyed: dict[tuple[str, str], list[int]] = {}
        self.section_ops: set[int] = set()
        self.resolved: set[int] = set()  # row/section ops whose target has a source identity
        self.row_keys: dict[int, set[str]] = {}
        self.unplaced: set[int] = set()  # accept/reject ops selecting a revision with no paragraph
        try:
            self.layout = read_layout(source)
        except Exception:  # unreadable layout: explicit groups still apply
            self.layout = None
        for position, operation in enumerate(operations):
            if operation.group:
                self._key(position, "group", operation.group)
            if self.layout is not None:
                self._resolve(position, operation, self.layout, source_row_ids(self.layout), source=True)
                self._revision_targets(position, operation, source, self.layout, source_row_ids(self.layout))
                self._comment_targets(position, operation, source, self.layout, source_row_ids(self.layout),
                                      source=True)
        self._rebuild()

    @staticmethod
    def targets_rows(operation: EditOperation) -> bool:
        return isinstance(operation, (ReplaceTableRowOperation, ReplaceSectionOperation))

    def _key(self, position: int, kind: str, value: str) -> None:
        members = self.keyed.setdefault((kind, value), [])
        if position not in members:
            members.append(position)
        if kind == "row":
            self.row_keys.setdefault(position, set()).add(value)

    def _row(self, position: int, row, row_ids) -> bool:
        if row_ids is None or row.table >= len(row_ids) or row.row >= len(row_ids[row.table]):
            return False
        for identity in row_ids[row.table][row.row]:
            self._key(position, "row", identity)
        return True

    def _revision_targets(self, position: int, operation: EditOperation, path: Path, layout, row_ids) -> None:
        """accept/reject: every row and paragraph holding a selected revision joins this op's groups.

        The selectors (ids, author, heading, range) name revisions anywhere, so the op's own paragraph_index
        says little about what it changes. Move partners and the paragraph a removed mark merges into join
        as well. A table-wide change (tblPrChange, tblGridChange) joins every row of its table. Selection
        errors are left to the op itself."""

        if operation.type not in ("accept_revisions", "reject_revisions"):
            return
        from . import revisions

        try:
            _, doc = revisions.load(path)
            chosen, _ = revisions.select(
                doc, op=operation.type, ids=operation.ids, author=operation.author,
                paragraph_index=operation.paragraph_index, through_paragraph_index=operation.through_paragraph_index,
                heading=operation.heading, until_heading=operation.until_heading)
            mode = operation.type.split("_")[0]
            expanded = revisions._expand_moves(doc, chosen)[0]
            indices = revisions.affected_paragraphs(doc, chosen, mode)
        except Exception:
            return
        # Move partners and the paragraph a removed mark merges into join too (affected_paragraphs).
        if any(revision.paragraph_index is None for revision in expanded):
            self.unplaced.add(position)
        table_wide = {revision.paragraph_index for revision in expanded
                      if revision.type in ("tblPrChange", "tblGridChange")}
        for index in indices:
            if not 1 <= index <= len(layout.paragraphs):
                self.unplaced.add(position)
                continue
            self._key(position, "para", str(index))
            try:
                row = layout.row_of(index)
            except RefusedUnsafe:
                row = None
            if row is None:
                continue
            if index in table_wide and row_ids is not None and row.table < len(row_ids):
                for identities in row_ids[row.table]:
                    for identity in identities:
                        self._key(position, "row", identity)
            elif not self._row(position, row, row_ids):
                self.unplaced.add(position)

    def _comment_targets(self, position: int, operation: EditOperation, path: Path, layout, row_ids, *,
                         source: bool) -> None:
        """reply/resolve/reopen/delete_comment carry no paragraph_index: the rows (and paragraphs) the thread is
        anchored on join this op's groups, so a failed row edit rolls back a disposition of that row's comment.
        An unknown id or a thread with no body markers is left to the op itself."""

        if operation.type not in ("reply_comment", "resolve_comment", "reopen_comment", "delete_comment"):
            return
        from .comment_ops import comment_paragraphs

        indices = comment_paragraphs(path, operation.comment_id)
        for index in indices or ():
            if not 1 <= index <= len(layout.paragraphs):
                continue
            if source:
                self._key(position, "para", str(index))
            try:
                row = layout.row_of(index)
            except RefusedUnsafe:
                row = None
            if row is not None and not self._row(position, row, row_ids):
                self.unplaced.add(position)

    def _resolve(self, position: int, operation: EditOperation, layout, row_ids, *, source: bool,
                 index: int | None = None) -> None:
        from .sections import find_row, locate_section

        if index is None:
            index = getattr(operation, "paragraph_index", None)
            if index is None:
                index = getattr(operation, "after_paragraph_index", None) or None
        if index is not None and 1 <= index <= len(layout.paragraphs):
            if source:
                self._key(position, "para", str(index))
            try:
                row = layout.row_of(index)
            except RefusedUnsafe:
                row = None
            if row is not None:
                self._row(position, row, row_ids)
        try:
            if isinstance(operation, ReplaceTableRowOperation):
                row = find_row(layout, operation.match, table_index=operation.table_index)
                if self._row(position, row, row_ids):
                    self.resolved.add(position)
                if source:
                    self.ranges[position] = (row.first_index, row.end_index)
            elif isinstance(operation, ReplaceSectionOperation):
                section = locate_section(layout, operation.heading, until_heading=operation.until_heading)
                self.section_ops.add(position)
                self._key(position, "section", section.heading.displayed.strip())
                head_row = layout.row_of(section.heading.index)
                mapped = head_row is None or self._row(position, head_row, row_ids)
                for row in section.rows:
                    mapped = self._row(position, row, row_ids) and mapped
                if source:
                    # Body-paragraph members are addressable in source coordinates only.
                    self._key(position, "para", str(section.heading.index))
                    for item in section.paragraphs:
                        self._key(position, "para", str(item.index))
                    members = [row.first_index for row in section.rows] + [row.end_index for row in section.rows]
                    members += [item.index for item in section.paragraphs]
                    if members:
                        self.ranges[position] = (min(members), max(members))
                if mapped and (source or section.kind == "rows"):
                    self.resolved.add(position)
        except RefusedUnsafe:
            pass  # the operation itself will fail with its own message

    def staged(self, position: int, operation: EditOperation, mapped: EditOperation, layout, row_ids,
               path: Path | None = None) -> None:
        """Add the identities `mapped` resolves to in the staged document it runs against."""

        if layout is None:
            return
        index = getattr(mapped, "paragraph_index", None)
        if index is None:
            index = getattr(mapped, "after_paragraph_index", None) or None
        self._resolve(position, mapped, layout, row_ids, source=False, index=index)
        if path is not None:
            # Revisions an earlier op of this batch wrote exist only in the staged document.
            self._revision_targets(position, mapped, path, layout, row_ids)
            # A thread an earlier op of this batch added or moved exists only in the staged document.
            self._comment_targets(position, mapped, path, layout, row_ids, source=False)
        self._rebuild()

    def unresolved(self) -> list[int]:
        return [
            position for position, operation in enumerate(self.operations)
            if (self.targets_rows(operation) and position not in self.resolved) or position in self.unplaced
        ]

    def _rebuild(self) -> None:
        self.parent = list(range(len(self.operations)))
        self.labels = {}
        for (kind, value), members in self.keyed.items():
            if kind == "para" and not self.section_ops.intersection(members):
                continue
            for other in members[1:]:
                self._union(members[0], other)
            if kind != "para":
                label = value if kind != "section" else f"section {value!r}"
                if kind == "group":
                    label = f"group {value!r}"
                for member in members:
                    self.labels.setdefault(member, [])
                    if label not in self.labels[member]:
                        self.labels[member].append(label)

    def _find(self, position: int) -> int:
        while self.parent[position] != position:
            self.parent[position] = self.parent[self.parent[position]]
            position = self.parent[position]
        return position

    def _union(self, a: int, b: int) -> None:
        a, b = self._find(a), self._find(b)
        if a != b:
            self.parent[max(a, b)] = min(a, b)

    def components(self) -> dict[int, list[int]]:
        """Grouped operations only: components that carry a group label."""

        found: dict[int, list[int]] = {}
        for position in range(len(self.parent)):
            found.setdefault(self._find(position), []).append(position)
        return {
            root: members for root, members in found.items()
            if any(member in self.labels for member in members)
        }

    def label(self, members: list[int]) -> str:
        names: list[str] = []
        for member in members:
            for name in self.labels.get(member, []):
                if name not in names:
                    names.append(name)
        return " + ".join(names[:3]) + (" + …" if len(names) > 3 else "")


RowIds = list[list[frozenset[str]]]


def source_row_ids(layout) -> RowIds:
    return [
        [frozenset({f"table {table + 1} row {row + 1}"}) for row in range(len(rows))]
        for table, rows in enumerate(layout.tables)
    ]


def _advance_row_ids(before, ids: RowIds | None, after, added: frozenset[str]) -> RowIds | None:
    """Carry source row identities from one staged document to the next.

    Tables are paired by position when their number is unchanged, else by
    their table properties with only insertions allowed. A table whose row
    count changed keeps identities for its unchanged leading and trailing
    rows; every row in between takes the union of the replaced rows'
    identities plus `added` (the edit's own), so a later edit of any of them
    groups with that edit. None when the pairing is unclear.
    """

    if ids is None or before is None or after is None or len(ids) != len(before.tables):
        return None
    if len(before.tables) == len(after.tables):
        pairs = list(range(len(after.tables)))
    else:
        old_keys = [unit.key() for unit in before.units if unit.kind == "tbl"]
        new_keys = [unit.key() for unit in after.units if unit.kind == "tbl"]
        pairs = [None] * len(after.tables)
        matcher = difflib.SequenceMatcher(None, old_keys, new_keys, autojunk=False)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                for offset in range(i2 - i1):
                    pairs[j1 + offset] = i1 + offset
            elif tag != "insert":
                return None
    result: RowIds = []
    for number, old in enumerate(pairs):
        new_rows = after.tables[number]
        if old is None:
            result.append([frozenset({f"table {number + 1} row {row + 1} (inserted in staging)"})
                           for row in range(len(new_rows))])
            continue
        old_rows, old_ids = before.tables[old], ids[old]
        if len(old_rows) == len(new_rows):
            result.append(list(old_ids))
            continue
        same = lambda a, b: before.units[a.unit].key() == after.units[b.unit].key()  # noqa: E731
        head = 0
        while head < min(len(old_rows), len(new_rows)) and same(old_rows[head], new_rows[head]):
            head += 1
        tail = 0
        while tail < min(len(old_rows), len(new_rows)) - head and same(old_rows[-1 - tail], new_rows[-1 - tail]):
            tail += 1
        middle = frozenset().union(*old_ids[head:len(old_rows) - tail]) | added
        result.append(
            list(old_ids[:head]) + [middle] * (len(new_rows) - head - tail) + list(old_ids[len(old_rows) - tail:]))
    return result


def _result_counts(results: list[OperationResult]) -> dict[str, int]:
    counts = {name: sum(item.status == name for item in results) for name in ("applied", "failed", "skipped_dependency")}
    rolled_back = sum(item.status == "rolled_back" for item in results)
    if rolled_back:
        counts["rolled_back"] = rolled_back
    return counts


def _mark_unpublished(results: list[OperationResult], reason: str) -> None:
    """Distinguish successful staging from edits actually published to disk."""
    for result in results:
        if result.status == "applied":
            result.status = "rolled_back"
            result.summary = "Verified in staging but not published"
            result.details["published"] = False
            result.details["rollback_reason"] = reason


def edit_document(request: EditDocumentRequest) -> EditBatchResult:
    source = Path(request.source).resolve(strict=True)
    is_inplace = request.output is None or Path(request.output).resolve() == source
    output = source if is_inplace else Path(request.output).resolve()
    actual_source_sha = file_sha256(source)
    if request.source_sha256 and request.source_sha256.upper() != actual_source_sha:
        return EditBatchResult(
            status="refused",
            source=str(source),
            source_sha256=actual_source_sha,
            operations=[],
            counts={"applied": 0, "failed": 0, "skipped_dependency": 0},
            retry=RetryInfo(
                source=str(source),
                source_sha256=actual_source_sha,
                requires_reinspection=True,
            ),
            message="source_sha256 does not match the current source; inspect again",
        )

    results: list[OperationResult] = []
    read_only = False
    groups: _Groups | None = None
    consistency = None
    publication = None
    published_sha = None
    applied = failed = skipped = 0
    try:
        locked, lock_reason = is_file_locked(source)
        if locked:
            raise RefusedUnsafe(f"Source is locked: {lock_reason}. Save and close it, then inspect again.")
        problems = document_structure_errors(source)
        if problems:
            raise RefusedUnsafe(
                "DOCX preflight failed before starting Word: "
                + "; ".join(problems[:5])
                + ". Repair and reinspect the source before retrying."
            )
        groups = _Groups(source, list(request.operations))
        with _exclusive_output(output, is_inplace=is_inplace), tempfile.TemporaryDirectory(
            prefix="research-mcp-edit-", dir=output.parent
        ) as temporary:
            temporary_path = Path(temporary)
            # D10: under allow_partial a group never half-applies. When a group
            # has a failed member and applied members, the whole group is left
            # out and the batch is staged again from the source without it.
            excluded: dict[int, str] = {}
            first_errors: dict[int, OperationResult] = {}
            for attempt in range(len(request.operations) + 1):
                results, current, current_sha = _stage(
                    request, source, actual_source_sha, temporary_path / f"attempt{attempt}",
                    groups, excluded, first_errors,
                )
                if not request.allow_partial:
                    break
                mixed = []
                for root, members in groups.components().items():
                    states = {results[member].status for member in members}
                    if "applied" in states and states & {"failed", "skipped_dependency"}:
                        mixed.append(members)
                if not mixed:
                    break
                for members in mixed:
                    culprits = [
                        results[member].op_id for member in members
                        if results[member].status in ("failed", "skipped_dependency")
                    ]
                    reason = (f"rolled back with {groups.label(members)} because "
                              f"{', '.join(culprits)} did not apply")
                    for member in members:
                        if results[member].status in ("failed", "skipped_dependency"):
                            first_errors[member] = results[member]
                        excluded[member] = reason

            applied = sum(item.status == "applied" for item in results)
            failed = sum(item.status == "failed" for item in results)
            skipped = sum(item.status == "skipped_dependency" for item in results)
            unpublished = len(results) - applied
            unresolved = groups.unresolved()
            if request.allow_partial and applied and unpublished and unresolved:
                names = ", ".join(
                    f"{request.operations[p].type} {request.operations[p].op_id!r}" for p in unresolved)
                raise RefusedUnsafe(
                    f"allow_partial cannot keep {names} apart from the edits that applied: its table row or "
                    "section was found neither in the document as inspected nor in the staged document, so its "
                    "group (D10) is unknown and a partial publish could split a row or section; nothing was "
                    "published. Fix that operation's match/heading or drop it, then retry.")
            if applied and (failed or skipped) and not request.allow_partial:
                _mark_unpublished(results, "Atomic batch discarded because at least one operation failed or was skipped")
                applied = 0
            read_only = bool(applied) and all(
                request.operations[position].type in READ_ONLY_OPERATION_TYPES or item.details.get("read_only")
                for position, item in enumerate(results) if item.status == "applied")
            if applied and not read_only:
                # D11, batch level: no touched table row lost or gained a cell
                # and no date cell stopped reading as a date. No parameter skips it.
                from .sections import consistency_errors

                problems = consistency_errors(source, current)
                if problems:
                    raise RefusedUnsafe(
                        "Batch consistency check failed; nothing was published: " + "; ".join(problems[:5]))
                consistency = "passed"
                publication = _publish(
                    current, output, is_inplace=is_inplace, expected_sha256=actual_source_sha,
                    op=", ".join(f"{r.type} {r.op_id}" for r in results if r.status == "applied")[:200],
                )
                published_sha = file_sha256(output)
                if published_sha != current_sha:
                    raise VerificationFailed("published output hash differs from verified candidate")

    except RefusedUnsafe as exc:
        if getattr(exc, "changed_nothing", True) is False:
            raise  # published, then a later check failed: never report it as rolled back
        _mark_unpublished(results, str(exc))
        return EditBatchResult(
            groups=_group_results(groups, results, published=False),
            status="refused",
            source=str(source),
            source_sha256=actual_source_sha,
            operations=results,
            counts=_result_counts(results),
            retry=RetryInfo(
                source=str(source),
                source_sha256=actual_source_sha,
                requires_reinspection=True,
            ),
            message=str(exc),
            sha256_before=actual_source_sha,
        )

    if failed or skipped:
        status = "partial" if applied else "no_changes"
    else:
        status = "complete"
    retry = None
    if failed or skipped:
        retry = RetryInfo(
            source=str(output if applied else source),
            source_sha256=published_sha or actual_source_sha,
            requires_reinspection=True,
        )

    summary_diffs = [
        _format_diff(request.operations[i])
        for i, r in enumerate(results)
        if r.status == "applied"
    ]

    return EditBatchResult(
        status=status,
        source=str(source),
        source_sha256=actual_source_sha,
        output=str(output) if applied and not read_only else None,
        output_sha256=published_sha,
        operations=results,
        counts=_result_counts(results),
        retry=retry,
        protected=ProtectedResult(
            fields_preserved=True,
            comments_preserved=True,
            drawings_preserved=True,
            media_preserved=True,
            unrelated_tables_preserved=True,
        ) if applied and not read_only else None,
        summary_diffs=summary_diffs,
        message=("Read only; nothing was written." if read_only else _message(results, publication, is_inplace)),
        undo_token=publication["undo_token"] if publication else None,
        undo_call=(
            f"undo_document(path={json.dumps(str(output))}, undo_token={json.dumps(publication['undo_token'])})"
            if publication and publication.get("undo_token") else None
        ),
        groups=_group_results(groups, results, published=publication is not None),
        consistency_check=consistency,
        sha256_before=publication["sha256_before"] if publication else actual_source_sha,
        sha256_after=publication["sha256_after"] if publication else actual_source_sha,
        publish_warning=publication.get("publish_warning") if publication else None,
        cleaned=_merged_cleaned(results),
    )


def _merged_cleaned(results: list[OperationResult]) -> list[dict]:
    from .textclean import merge_changes

    return merge_changes(*(item.details.get("cleaned") or [] for item in results
                           if item.status == "applied" and isinstance(item.details.get("cleaned"), list)))


def _message(results: list[OperationResult], publication: dict | None, is_inplace: bool) -> str:
    rolled_back = [r.op_id for r in results if r.status == "rolled_back"]
    if not publication:
        if rolled_back:
            return "Atomic batch discarded; source unchanged." if not any(
                r.details.get("group_rolled_back") for r in results
            ) else "No operation group applied in full; source unchanged."
        return ""
    group_note = (f" Not published (rolled back with their group): {', '.join(rolled_back)}."
                  if rolled_back else "")
    if is_inplace and publication.get("undo_token"):
        return (f"Published in place. undo_document(path) restores the state before this batch "
                f"(undo_token {publication['undo_token']})." + group_note
                + (f" WARNING: {publication['publish_warning']}" if publication.get("publish_warning") else ""))
    return ("Published as a new file." + group_note) if group_note else ""


def _group_results(groups: _Groups | None, results: list[OperationResult], *, published: bool) -> list[GroupResult]:
    if groups is None or len(results) != len(groups.parent):
        return []
    found: list[GroupResult] = []
    for members in groups.components().values():
        states = [results[member].status for member in members]
        reasons = [results[member].details.get("rollback_reason", "") for member in members]
        if all(state == "applied" for state in states) and published:
            status, reason = "applied", ""
        elif any(results[member].details.get("group_rolled_back") for member in members):
            status, reason = "rolled_back", next((r for r in reasons if r), "")
        elif any(state == "rolled_back" for state in states):
            status, reason = "rolled_back", next((r for r in reasons if r), "batch not published")
        elif all(state in ("failed", "skipped_dependency") for state in states):
            status, reason = "failed", next((results[m].error.message for m in members if results[m].error), "")
        else:
            status, reason = "not_applied", ""
        found.append(GroupResult(
            group=groups.label(members), op_ids=[results[member].op_id for member in members],
            status=status, reason=reason,
        ))
    return found


_REVISION_OPS = ("accept_revisions", "reject_revisions")


def _zotero_views(path: Path, workdir: Path, tag: str) -> list:
    """The document's Zotero fields as read now, with every revision accepted and with every one rejected.

    The package reader sees only w:instrText/w:t, so a citation inside a w:del is invisible until a reject
    restores it; the reject view makes it visible."""

    import xml.etree.ElementTree as ET
    import zipfile

    from .ooxml import W

    views = [zotero_package_fields(path)]
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
    for name, drop, rename in (("accepted", ("del", "moveFrom"), {}),
                               ("rejected", ("ins", "moveTo"), {"delInstrText": "instrText", "delText": "t"})):
        root = ET.fromstring(xml)
        for parent in list(root.iter()):
            for child in list(parent):
                if child.tag in tuple(W + local for local in drop):
                    parent.remove(child)
        for node in root.iter():
            if node.tag.startswith(W) and node.tag[len(W):] in rename:
                node.tag = W + rename[node.tag[len(W):]]
        view = workdir / f"zotero-{tag}-{name}.docx"
        view.unlink(missing_ok=True)
        with zipfile.ZipFile(path) as source, zipfile.ZipFile(view, "w") as target:
            for info in source.infolist():
                data = ET.tostring(root) if info.filename == "word/document.xml" else source.read(info.filename)
                target.writestr(info, data, compress_type=info.compress_type)
        views.append(zotero_package_fields(view))
        view.unlink(missing_ok=True)
    return views


def _zotero_contained(candidate, views) -> bool:
    """Every citation identity and rendered text in the candidate already exists in some view of the current
    document: citations may be dropped or reordered, never altered or invented."""

    from collections import Counter

    for attribute in ("item_results", "item_keys", "bibliography_results"):
        allowed: Counter = Counter()
        for view in views:
            allowed |= Counter(getattr(view, attribute))
        if Counter(getattr(candidate, attribute)) - allowed:
            return False
    return True


def _stage(
    request: EditDocumentRequest, source: Path, source_sha: str, workdir: Path, groups: _Groups,
    excluded: dict[int, str], first_errors: dict[int, OperationResult],
) -> tuple[list[OperationResult], Path, str]:
    """Apply every operation not in `excluded` to successive verified candidates."""

    from .sections import read_layout

    workdir.mkdir(parents=True, exist_ok=True)
    staged_layout = groups.layout
    row_ids = source_row_ids(staged_layout) if staged_layout is not None else None
    results: list[OperationResult] = []
    statuses: dict[str, str] = {}
    shifts: list[list[tuple[int, int]]] = []  # one transform per applied op, see _mapped_index
    stale: list[tuple[int, int, str]] = []  # source paragraph spans rewritten with a count change
    current = source
    current_sha = source_sha
    integration_failure: OperationError | None = None
    for sequence, operation in enumerate(request.operations, start=1):
        position = sequence - 1
        if position in excluded:
            if position in first_errors:
                result = first_errors[position].model_copy(deep=True)
                result.details["group_rolled_back"] = True
            else:
                result = OperationResult(
                    op_id=operation.op_id, type=operation.type, status="rolled_back",
                    summary="verified in staging, then rolled back with its group",
                    details={"published": False, "rollback_reason": excluded[position], "group_rolled_back": True},
                )
            results.append(result)
            statuses[operation.op_id] = result.status
            continue
        if integration_failure is not None:
            result = OperationResult(
                op_id=operation.op_id,
                type=operation.type,
                status="failed",
                summary="not attempted after Word integration failure",
                error=OperationError(
                    code="IntegrationUnavailable",
                    message="Batch stopped to avoid repeated Word dialogs. " + integration_failure.message,
                    context={"attempted": False},
                ),
            )
            results.append(result)
            statuses[operation.op_id] = result.status
            continue
        dependencies = list(getattr(operation, "depends_on", ()))
        failed_dependencies = [
            dependency
            for dependency in dependencies
            if statuses.get(dependency) != "applied"
        ]
        if failed_dependencies:
            groups.staged(position, operation, _mapped_operation(operation, shifts), staged_layout, row_ids,
                          current)
            result = OperationResult(
                op_id=operation.op_id,
                type=operation.type,
                status="skipped_dependency",
                summary="dependency did not apply",
                failed_dependencies=failed_dependencies,
            )
            results.append(result)
            statuses[operation.op_id] = result.status
            continue

        mapped = _mapped_operation(operation, shifts)
        groups.staged(position, operation, mapped, staged_layout, row_ids, current)
        candidate = workdir / f"{sequence:02d}-{operation.op_id}.docx"
        before_count = audit_docx(current).paragraph_count
        try:
            requested = getattr(operation, "paragraph_index", None)
            for first, last, owner in stale:
                if requested is not None and first <= requested <= last:
                    raise RefusedUnsafe(
                        f"paragraph_index {requested} lies inside the rows/paragraphs that {owner} rewrote with a "
                        f"different paragraph count, so it no longer names a known paragraph; fold this edit "
                        f"into {owner} or address the text by section/row")
            mapped, cleaned = _clean_operation(mapped, staged_layout, current)
            details = _apply_operation(current, candidate, mapped)
            candidate_problems = document_structure_errors(candidate)
            if candidate_problems:
                raise VerificationFailed("Edited candidate failed structural preflight", actual=candidate_problems)
            if mapped.type in _REVISION_OPS:
                # Resolving revisions may drop a citation (accepting its deletion) or restore one (rejecting
                # it), but every identity and rendered text must already exist in the current document.
                if not _zotero_contained(zotero_package_fields(candidate),
                                         _zotero_views(current, workdir, f"{sequence:02d}")):
                    raise VerificationFailed(
                        f"{operation.type} changed or invented a live Zotero identity or rendered "
                        "citation/bibliography text")
            elif zotero_package_fields(current) != zotero_package_fields(candidate):
                raise VerificationFailed("Prose edit changed live Zotero identities or rendered citation/bibliography text")
            candidate_sha = file_sha256(candidate)
            op_details: dict[str, object] = {"candidate_sha": candidate_sha}
            if isinstance(details, dict):
                op_details.update(details)
            if cleaned:
                op_details["cleaned"] = cleaned
            req_idx = getattr(operation, "paragraph_index", getattr(operation, "after_paragraph_index", None))
            res_idx = getattr(mapped, "paragraph_index", getattr(mapped, "after_paragraph_index", None))
            if req_idx is not None and res_idx is not None and (any(shifts) or req_idx != res_idx):
                op_details["requested_index"] = req_idx
                op_details["resolved_index"] = res_idx
            group_label = groups.label([position]) if position in groups.labels else None
            if group_label:
                op_details["group"] = group_label
            result = OperationResult(
                op_id=operation.op_id,
                type=operation.type,
                status="applied",
                summary=f"verified {operation.type}",
                details=op_details,
            )
            results.append(result)
            statuses[operation.op_id] = "applied"
            current = candidate
            current_sha = candidate_sha
            try:
                next_layout = read_layout(candidate)
            except Exception:
                next_layout = None
            row_ids = _advance_row_ids(
                staged_layout, row_ids, next_layout, frozenset(groups.row_keys.get(position, ())))
            staged_layout = next_layout
            after_count = int(details.get("paragraph_count", before_count))
            # index_delta (layout indices, row-end markers included) is exact for the plan-based ops.
            index_delta = int(details["index_delta"]) if "index_delta" in details else after_count - before_count
            if details.get("index_shifts"):
                # accept/reject: one (first index after the merged paragraph, delta) per merge, all in this
                # op's input coordinates (one simultaneous transform); the op's paragraph_index is a
                # selector, not the shift base.
                shifts.append([(int(first), int(delta)) for first, delta in details["index_shifts"]])
                if position in groups.ranges:
                    first, last = groups.ranges[position]
                    stale.append((first, last, f"{operation.type} {operation.op_id!r}"))
            elif index_delta:
                base = getattr(
                    mapped,
                    "paragraph_index",
                    getattr(mapped, "after_paragraph_index", None),
                )
                if base is None:
                    shifts.append([(int(details.get("shift_base", 1)), index_delta)])
                else:
                    shifts.append([(base + 1, index_delta)])
                if position in groups.ranges:
                    first, last = groups.ranges[position]
                    stale.append((first, last, f"{operation.type} {operation.op_id!r}"))
        except Exception as exc:
            result = OperationResult(
                op_id=operation.op_id,
                type=operation.type,
                status="failed",
                summary=f"{operation.type} failed",
                error=_error(exc),
            )
            results.append(result)
            statuses[operation.op_id] = "failed"
            if isinstance(exc, IntegrationRefused):
                integration_failure = result.error
    return results, current, current_sha
