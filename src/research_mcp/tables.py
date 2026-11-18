"""Journal-neutral table formatting and vertical merge application."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, fields
import hashlib
from pathlib import Path
import tempfile
from typing import Sequence
import xml.etree.ElementTree as ET
import zipfile

from .audit import Audit, TableAudit, audit_docx
from .constants import (
    WD_ALIGN_PARAGRAPH_LEFT,
    WD_ALIGN_PARAGRAPH_RIGHT,
    WD_BORDER_BOTTOM,
    WD_BORDER_HORIZONTAL,
    WD_BORDER_LEFT,
    WD_BORDER_RIGHT,
    WD_BORDER_TOP,
    WD_BORDER_VERTICAL,
    WD_LINE_STYLE_NONE,
    WD_LINE_STYLE_SINGLE,
    WD_LINE_WIDTH_050_PT,
)
from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import NS, W, parse_table, row_cells
from .word import IsolatedWordDocument, staged_document
from .word_baseline import section_errors


@dataclass(frozen=True)
class MergeGroup:
    column: int
    start_row: int
    end_row: int
    label: str


@dataclass(frozen=True)
class TableFormatSpec:
    index: int
    descriptive_columns: tuple[int, ...]
    merges: tuple[MergeGroup, ...] = ()


@dataclass(frozen=True)
class TablesAudit(Audit):
    """format_to_version's audit of the saved result, with what it was verified against (plan 973 S1):
    ``"word_baseline"`` (Word's own save of the untouched document, see ``word_baseline``) or ``"source"``."""

    compared_against: str = "source"


# A table's real cells per row, as (first grid column, grid columns spanned), in the order Word's
# ``Table.Cell(row, n)`` numbers them (plan 973 S1). A row with a gridSpan cell has fewer cells than the table
# has grid columns, so ``Cell(row, column)`` by grid column raises "The requested member of the collection does
# not exist" or addresses the wrong cell. ``ooxml.row_cells`` reads them; ``parse_table`` uses the same map.
RowCells = tuple[tuple[int, int], ...]
TableLayout = tuple[RowCells, ...]


def _top_level_tables(node: ET.Element) -> list[ET.Element]:
    """The tables Word's ``Document.Tables`` lists: body tables (also inside content controls), not tables nested
    in a cell or inside a drawing or text box."""

    found = []
    for child in node:
        if child.tag == f"{W}tbl":
            found.append(child)
        elif child.tag not in (f"{W}drawing", f"{W}pict", f"{W}txbxContent", f"{W}object"):
            found.extend(_top_level_tables(child))
    return found


def table_layouts(path: str | Path) -> tuple[TableLayout, ...]:
    """Each top-level table's real cells per row, read from the saved XML."""

    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find(f"{W}body")
    tables = _top_level_tables(body) if body is not None else []
    return tuple(tuple(row_cells(row) for row in table.findall(f"{W}tr")) for table in tables)


def _layout_for(table, layout: TableLayout | None) -> TableLayout:
    """``layout`` when it fits Word's view of the table, else one single cell per grid column (the addressing
    used before plan 973)."""

    rows, columns = int(table.Rows.Count), int(table.Columns.Count)
    if layout is not None and len(layout) == rows:
        return layout
    return tuple(tuple((column, 1) for column in range(1, columns + 1)) for _ in range(rows))


def _grid_width(table, layout: TableLayout) -> int:
    widest = max((start + span - 1 for row in layout for start, span in row), default=0)
    return max(int(table.Columns.Count), widest)


def _columns_text(start: int, span: int) -> str:
    return f"column {start}" if span == 1 else f"columns {start}-{start + span - 1}"


def _cell_position(layout: TableLayout, table_index: int, row: int, column: int, why: str) -> int:
    """Word's ``Cell(row, n)`` number of the cell at grid ``column`` in ``row``; refuses when that cell spans
    several grid columns or there is none there."""

    for position, (start, span) in enumerate(layout[row - 1], start=1):
        if start <= column < start + span:
            if span > 1:
                raise RefusedUnsafe(
                    f"{why}: in table {table_index} row {row} the cell at column {column} spans grid "
                    f"{_columns_text(start, span)}; only cells one column wide can be merged. Split that cell in "
                    "Word or leave this merge out")
            return position
    raise RefusedUnsafe(f"{why}: table {table_index} row {row} has no cell at column {column}")


def check_spans(spec: TableFormatSpec, layout: TableLayout) -> None:
    """Refuse, naming the cell, what a spanned cell makes ambiguous: a data-row cell spanning both a descriptive
    and a data column (left or right alignment?), and a merge through a spanned cell. The header row is
    left-aligned whatever it spans."""

    descriptive = set(spec.descriptive_columns)
    for row, cells in enumerate(layout, start=1):
        if row == 1:
            continue
        for start, span in cells:
            covered = set(range(start, start + span))
            if span > 1 and covered & descriptive and covered - descriptive:
                raise RefusedUnsafe(
                    f"table {spec.index} row {row}: the cell spanning grid {_columns_text(start, span)} holds "
                    f"descriptive column(s) {sorted(covered & descriptive)} and data column(s) "
                    f"{sorted(covered - descriptive)}, so its alignment is ambiguous; split that cell in Word or "
                    "declare all of its columns descriptive or none")
    for group in spec.merges:
        for row in range(group.start_row, group.end_row + 1):
            if 1 <= row <= len(layout):
                _cell_position(layout, spec.index, row, group.column,
                               f"merge of column {group.column} rows {group.start_row}-{group.end_row}")


def _clear_all_borders(borders) -> None:
    for border_id in (
        WD_BORDER_LEFT,
        WD_BORDER_RIGHT,
        WD_BORDER_TOP,
        WD_BORDER_BOTTOM,
        WD_BORDER_HORIZONTAL,
        WD_BORDER_VERTICAL,
    ):
        try:
            borders(border_id).LineStyle = WD_LINE_STYLE_NONE
        except Exception:
            pass


def _apply_border(border, style=WD_LINE_STYLE_SINGLE, width=WD_LINE_WIDTH_050_PT) -> None:
    border.LineStyle = style
    border.LineWidth = width


def _normalize_whitespace(text: str) -> str:
    return " ".join(text.split())


def _safe_cell_text(cell) -> str:
    return _normalize_whitespace(cell.Range.Text.rstrip("\r\x07"))


def snapshot_live(document, layouts: Sequence[TableLayout] = ()) -> tuple[TableAudit, ...]:
    """Each table's cell texts by grid position: a spanned cell's text fills every grid column it covers (as the
    saved audit's ``parse_table`` does); a cell Word can't address (below the first of a vertical merge) or a grid
    column no cell covers hashes as ""."""
    audits = []
    empty = hashlib.sha256(b"").hexdigest()
    for index in range(1, int(document.Tables.Count) + 1):
        table = document.Tables(index)
        layout = _layout_for(table, layouts[index - 1] if index <= len(layouts) else None)
        rows = int(table.Rows.Count)
        columns = _grid_width(table, layout)
        hashes = [empty] * (rows * columns)
        for row in range(1, rows + 1):
            for position, (start, span) in enumerate(layout[row - 1], start=1):
                try:
                    text = _safe_cell_text(table.Cell(row, position))
                except Exception:
                    text = ""
                digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
                for column in range(start, min(start + span, columns + 1)):
                    hashes[(row - 1) * columns + column - 1] = digest
        audits.append(
            TableAudit(
                index=index,
                rows=rows,
                columns=columns,
                cell_hashes=tuple(hashes),
                vertical_merges=(),
                horizontal_spans=(),
            )
        )
    return tuple(audits)


def apply_journal_neutral(document, spec: TableFormatSpec, layout: TableLayout | None = None) -> None:
    table = document.Tables(spec.index)
    layout = _layout_for(table, layout)
    rows = int(table.Rows.Count)
    columns = _grid_width(table, layout)
    if rows < 2:
        raise RefusedUnsafe(f"table {spec.index} has fewer than two rows")

    descriptive = set(spec.descriptive_columns)
    if not descriptive:
        raise RefusedUnsafe(f"table {spec.index} has no descriptive columns declared")

    invalid_descriptive = [col for col in descriptive if col < 1 or col > columns]
    if invalid_descriptive:
        raise RefusedUnsafe(
            f"table {spec.index} descriptive columns {invalid_descriptive} out of bounds"
        )

    _clear_all_borders(table.Borders)
    _apply_border(table.Rows(1).Borders(WD_BORDER_TOP))
    _apply_border(table.Rows(1).Borders(WD_BORDER_BOTTOM))
    _apply_border(table.Rows(rows).Borders(WD_BORDER_BOTTOM))

    for col in range(1, columns + 1):
        try:
            table.Columns(col).Borders(WD_BORDER_VERTICAL).LineStyle = WD_LINE_STYLE_NONE
        except Exception:
            pass

    check_spans(spec, layout)
    for row in range(1, rows + 1):
        is_header = (row == 1)
        for position, (start, span) in enumerate(layout[row - 1], start=1):
            cell = table.Cell(row, position)
            if is_header or set(range(start, start + span)) & descriptive:
                alignment = WD_ALIGN_PARAGRAPH_LEFT
            else:
                alignment = WD_ALIGN_PARAGRAPH_RIGHT
            for p in range(1, int(cell.Range.Paragraphs.Count) + 1):
                cell.Range.Paragraphs(p).Alignment = alignment


def _validate_merge_group(table, group: MergeGroup, layout: TableLayout | None = None,
                          table_index: int = 1) -> tuple[int, ...]:
    """Checks ``group`` and returns, per row of it, Word's ``Cell(row, n)`` number of its cell."""
    layout = _layout_for(table, layout)
    rows = int(table.Rows.Count)
    columns = _grid_width(table, layout)
    if group.column < 1 or group.column > columns:
        raise RefusedUnsafe(f"merge column {group.column} out of bounds")
    if group.start_row < 2 or group.end_row > rows or group.start_row >= group.end_row:
        raise RefusedUnsafe(
            f"merge rows {group.start_row}..{group.end_row} invalid for table with {rows} rows"
        )
    why = f"merge of column {group.column} rows {group.start_row}-{group.end_row}"
    positions = tuple(_cell_position(layout, table_index, row, group.column, why)
                      for row in range(group.start_row, group.end_row + 1))
    texts = []
    for row, position in zip(range(group.start_row, group.end_row + 1), positions):
        cell = table.Cell(row, position)
        text = _safe_cell_text(cell)
        if not text:
            raise RefusedUnsafe(
                f"refusing merge over empty cell at row {row}, column {group.column}"
            )
        texts.append(text)

    if len(set(texts)) != 1 or _normalize_whitespace(texts[0]) != _normalize_whitespace(group.label):
        raise RefusedUnsafe(
            f"refusing merge: cells in column {group.column} rows {group.start_row}..{group.end_row} "
            f"do not uniformly contain label {group.label!r} (found: {texts})"
        )
    return positions


def merge_repeated_label(table, group: MergeGroup, *, keep_formatting: bool = False,
                         layout: TableLayout | None = None, table_index: int = 1) -> None:
    """Merge the cells of ``group`` into its first cell. By default the merged cell's text is then set to the
    label, which retypes it in one run. ``keep_formatting`` (the tracked path, plan 916 S4) leaves the first cell's
    runs as they are: Word's merge keeps them and adds nothing from the emptied cells (measured, Word 16), and a
    merge can't be tracked, so retyping would change character formatting that Reject All could not give back.
    It refuses if the merged cell then does not read as the label. Cells are addressed by their place in the row
    (``layout``), so a gridSpan cell earlier in a row does not shift the merge to the wrong cell."""
    positions = _validate_merge_group(table, group, layout, table_index)
    first_cell = table.Cell(group.start_row, positions[0])
    for row, position in zip(range(group.start_row + 1, group.end_row + 1), positions[1:]):
        other_cell = table.Cell(row, position)
        other_cell.Range.Text = ""
    first_cell.Merge(table.Cell(group.end_row, positions[-1]))
    if not keep_formatting:
        first_cell.Range.Text = group.label
        return
    merged_text = _safe_cell_text(first_cell)
    if merged_text != _normalize_whitespace(group.label):
        raise VerificationFailed(
            f"the merged cell at row {group.start_row}, column {group.column} does not read as its label "
            "after the merge",
            expected=_normalize_whitespace(group.label),
            actual=merged_text,
        )


def _verify_protected_structure(before: Audit, after: Audit) -> None:
    if before.table_count != after.table_count:
        raise VerificationFailed(
            "table count changed",
            expected=before.table_count,
            actual=after.table_count,
        )
    if before.paragraph_count != after.paragraph_count:
        raise VerificationFailed(
            "paragraph count changed during table formatting",
            expected=before.paragraph_count,
            actual=after.paragraph_count,
        )
    for field in (
        "field_count",
        "live_comment_count",
        "drawing_count",
        "live_media_count",
        "bookmark_names",
    ):
        if getattr(before, field) != getattr(after, field):
            raise VerificationFailed(
                f"{field} changed during table formatting",
                expected=getattr(before, field),
                actual=getattr(after, field),
            )


def verify_saved_journal_neutral(path: str | Path, specs: Sequence[TableFormatSpec]) -> None:
    docx_path = Path(path).resolve(strict=True)
    with zipfile.ZipFile(docx_path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    parsed_tables = [parse_table(tbl) for tbl in root.findall(f".//{W}tbl")]
    tbl_elements = root.findall(f".//{W}tbl")
    if len(parsed_tables) < len(specs):
        raise VerificationFailed(
            "document has fewer tables than formatted specs",
            expected=len(specs),
            actual=len(parsed_tables),
        )
    for spec in specs:
        tbl = parsed_tables[spec.index - 1]
        tbl_elem = tbl_elements[spec.index - 1]
        tbl_pr = tbl_elem.find(f"{W}tblPr")
        if tbl_pr is not None:
            tbl_borders = tbl_pr.find(f"{W}tblBorders")
            if tbl_borders is not None:
                for side in ("left", "right", "insideV"):
                    border = tbl_borders.find(f"{W}{side}")
                    if border is not None and border.get(f"{W}val") not in ("none", "nil"):
                        raise VerificationFailed(
                            f"table {spec.index} has prohibited border {side}",
                            expected="none",
                            actual=border.get(f"{W}val"),
                        )
        for group in spec.merges:
            matching = [
                m for m in tbl.vertical_merges
                if m[0] == group.start_row and m[1] == group.end_row and m[2] == group.column and m[3] == group.label
            ]
            if not matching:
                raise VerificationFailed(
                    f"table {spec.index} missing expected vertical merge {group}",
                    expected=group,
                    actual=tbl.vertical_merges,
                )


def _without_merge_continuations(table: TableAudit, spec: TableFormatSpec) -> tuple[str | None, ...]:
    """Live cell hashes with the cells below the first of each requested merge masked. After the merge Word no
    longer addresses those cells (``table.Cell`` raises, so ``snapshot_live`` hashes ""), while the saved audit
    repeats the label there; the saved check (``verify_saved_tables``) compares every cell, and what each
    continuation physically holds (``verify_saved_continuations``). Both paths mask them
    (plan 973 S1: before, the untracked path compared them and refused every merge as "logical data changed")."""
    hashes: list[str | None] = list(table.cell_hashes)
    for group in spec.merges:
        for row in range(group.start_row + 1, group.end_row + 1):
            position = (row - 1) * table.columns + (group.column - 1)
            if 0 <= position < len(hashes):
                hashes[position] = None
    return tuple(hashes)


def format_to_version(
    source: str | Path,
    output: str | Path | None = None,
    specs: Sequence[TableFormatSpec] = (),
    *,
    visible: bool = False,
    tracked: bool = False,
) -> TablesAudit:
    """``tracked`` (plan 916 S4): borders and alignment with Word's TrackRevisions on, merges with it off (Word
    records a tracked merge as deleted and retyped text) and keeping the first cell's runs; the document's own
    setting is restored before the save. ``tracked=False`` is the untracked pass.

    The saved result is verified against Word's own save of the untouched document when that save preserves the
    source (``word_baseline.choose_reference``; plan 973 S1: Word appends an empty paragraph after a body-final
    table), else against the source; the result's ``compared_against`` says which."""
    from .word_baseline import choose_reference

    before = audit_docx(source)
    if tuple(spec.index for spec in specs) != tuple(range(1, before.table_count + 1)):
        raise RefusedUnsafe("format specs must cover every table exactly once in index order")
    layouts = table_layouts(source)
    for spec in specs:
        if spec.index <= len(layouts):
            check_spans(spec, layouts[spec.index - 1])
    target_output = output if output is not None else source
    with tempfile.TemporaryDirectory(prefix="research-mcp-baseline-") as scratch, \
            staged_document(source, target_output) as output_path:
        with IsolatedWordDocument(output_path, visible=visible) as session:
            document = session.document
            baseline = session.save_as(Path(scratch) / "baseline.docx")
            before_live = snapshot_live(document, layouts)
            if tracked:
                from .word_tracked import set_word_tracking
                original_tracking = set_word_tracking(document, True)
            for spec in specs:
                layout = layouts[spec.index - 1] if spec.index <= len(layouts) else None
                if tracked:
                    set_word_tracking(document, True)
                apply_journal_neutral(document, spec, layout)
                table = document.Tables(spec.index)
                if tracked:
                    set_word_tracking(document, False)
                for group in spec.merges:
                    merge_repeated_label(table, group, keep_formatting=tracked, layout=layout,
                                         table_index=spec.index)
            if tracked:
                set_word_tracking(document, original_tracking)
            after_live = snapshot_live(document, layouts)
            for spec, old, new in zip(specs, before_live, after_live, strict=True):
                old_hashes = _without_merge_continuations(old, spec)
                new_hashes = _without_merge_continuations(new, spec)
                if (old.rows, old.columns, old_hashes) != (new.rows, new.columns, new_hashes):
                    raise VerificationFailed(
                        f"table {old.index} logical data changed",
                        expected=(old.rows, old.columns, old_hashes),
                        actual=(new.rows, new.columns, new_hashes),
                    )
            session.save_as(output_path)

        reference = choose_reference(source, baseline)
        reference_audit = before if reference.path == Path(source) else audit_docx(reference.path)
        after = verify_saved_tables(reference_audit, output_path, specs, reference=reference.path)
        return TablesAudit(**{item.name: getattr(after, item.name) for item in fields(Audit)},
                           compared_against=reference.compared_against)


def _continuations(path: str | Path) -> list[dict[tuple[int, int], str]]:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    return [{(row, column): text for row, column, text in parse_table(table).continuations}
            for table in root.findall(f".//{W}tbl")]


def verify_saved_continuations(reference: str | Path, path: str | Path, specs: Sequence[TableFormatSpec]) -> None:
    """What the cells below the first of each vertical merge physically hold (plan 973 S1). The saved audit reads
    them as the merge's label (``parse_table``), so a continuation holding other text would pass the cell hashes.
    A requested merge empties them (``merge_repeated_label`` clears them before merging); every other continuation
    holds what it holds in ``reference``."""

    before, after = _continuations(reference), _continuations(path)
    if len(before) != len(after):
        raise VerificationFailed("table count changed", expected=len(before), actual=len(after))
    emptied = {spec.index: {(row, group.column) for group in spec.merges
                            for row in range(group.start_row + 1, group.end_row + 1)} for spec in specs}
    for index, (old, new) in enumerate(zip(before, after), start=1):
        expected = {**old, **{cell: "" for cell in emptied.get(index, ())}}
        if new != expected:
            cells = sorted(set(expected) ^ set(new) | {cell for cell in set(expected) & set(new)
                                                       if expected[cell] != new[cell]})
            where = ", ".join(f"row {row} column {column}" for row, column in cells)
            raise VerificationFailed(
                f"saved table {index}: merged cells below the first of a merge hold other text than expected "
                f"({where})",
                expected={f"{row},{column}": text for (row, column), text in sorted(expected.items())},
                actual={f"{row},{column}": text for (row, column), text in sorted(new.items())},
            )


def verify_saved_tables(before: Audit, path: str | Path, specs: Sequence[TableFormatSpec], *,
                        reference: str | Path) -> Audit:
    """format_to_version's checks of a saved result (also run on Word's Accept All of a tracked one). ``reference``
    is Word's own save of the untouched document when adopted, else the source; ``before`` is its audit."""
    after = audit_docx(path)
    _verify_protected_structure(before, after)
    sections = section_errors(reference, path)
    if sections:
        raise VerificationFailed("formatting tables changed section properties: " + "; ".join(sections))
    for old, new in zip(before.tables, after.tables, strict=True):
        if (old.rows, old.columns, old.cell_hashes) != (new.rows, new.columns, new.cell_hashes):
            raise VerificationFailed(
                f"saved table {old.index} logical data changed",
                expected=(old.rows, old.columns, old.cell_hashes),
                actual=(new.rows, new.columns, new.cell_hashes),
            )
    expected_merges = tuple(
        (group.start_row, group.end_row, group.column, group.label)
        for spec in specs
        for group in spec.merges
    )
    actual_merges = tuple(merge for table in after.tables for merge in table.vertical_merges)
    if actual_merges != expected_merges:
        raise VerificationFailed("saved merge structure differs", expected=expected_merges, actual=actual_merges)
    verify_saved_journal_neutral(path, specs)
    verify_saved_continuations(reference, path, specs)
    return after
