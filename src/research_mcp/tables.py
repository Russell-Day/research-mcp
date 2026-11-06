"""Journal-neutral table formatting and vertical merge application."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
import hashlib
from pathlib import Path
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
from .ooxml import NS, W, parse_table
from .word import IsolatedWordDocument, staged_document


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


def snapshot_live(document) -> tuple[TableAudit, ...]:
    audits = []
    for index in range(1, int(document.Tables.Count) + 1):
        table = document.Tables(index)
        rows = int(table.Rows.Count)
        columns = int(table.Columns.Count)
        hashes = []
        for row in range(1, rows + 1):
            for column in range(1, columns + 1):
                try:
                    cell = table.Cell(row, column)
                    text = _safe_cell_text(cell)
                except Exception:
                    text = ""
                hashes.append(hashlib.sha256(text.encode("utf-8")).hexdigest())
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


def apply_journal_neutral(document, spec: TableFormatSpec) -> None:
    table = document.Tables(spec.index)
    rows = int(table.Rows.Count)
    columns = int(table.Columns.Count)
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

    for row in range(1, rows + 1):
        is_header = (row == 1)
        for col in range(1, columns + 1):
            cell = table.Cell(row, col)
            if is_header or (col in descriptive):
                alignment = WD_ALIGN_PARAGRAPH_LEFT
            else:
                alignment = WD_ALIGN_PARAGRAPH_RIGHT
            for p in range(1, int(cell.Range.Paragraphs.Count) + 1):
                cell.Range.Paragraphs(p).Alignment = alignment


def _validate_merge_group(table, group: MergeGroup) -> None:
    rows = int(table.Rows.Count)
    columns = int(table.Columns.Count)
    if group.column < 1 or group.column > columns:
        raise RefusedUnsafe(f"merge column {group.column} out of bounds")
    if group.start_row < 2 or group.end_row > rows or group.start_row >= group.end_row:
        raise RefusedUnsafe(
            f"merge rows {group.start_row}..{group.end_row} invalid for table with {rows} rows"
        )
    texts = []
    for row in range(group.start_row, group.end_row + 1):
        cell = table.Cell(row, group.column)
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


def merge_repeated_label(table, group: MergeGroup) -> None:
    _validate_merge_group(table, group)
    first_cell = table.Cell(group.start_row, group.column)
    for row in range(group.start_row + 1, group.end_row + 1):
        other_cell = table.Cell(row, group.column)
        other_cell.Range.Text = ""
    first_cell.Merge(table.Cell(group.end_row, group.column))
    first_cell.Range.Text = group.label


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


def format_to_version(
    source: str | Path,
    output: str | Path | None = None,
    specs: Sequence[TableFormatSpec] = (),
    *,
    visible: bool = False,
) -> Audit:
    before = audit_docx(source)
    if tuple(spec.index for spec in specs) != tuple(range(1, before.table_count + 1)):
        raise RefusedUnsafe("format specs must cover every table exactly once in index order")
    target_output = output if output is not None else source
    with staged_document(source, target_output) as output_path:
        with IsolatedWordDocument(output_path, visible=visible) as session:
            document = session.document
            before_live = snapshot_live(document)
            for spec in specs:
                apply_journal_neutral(document, spec)
                table = document.Tables(spec.index)
                for group in spec.merges:
                    merge_repeated_label(table, group)
            after_live = snapshot_live(document)
            for old, new in zip(before_live, after_live, strict=True):
                if (old.rows, old.columns, old.cell_hashes) != (new.rows, new.columns, new.cell_hashes):
                    raise VerificationFailed(
                        f"table {old.index} logical data changed",
                        expected=(old.rows, old.columns, old.cell_hashes),
                        actual=(new.rows, new.columns, new.cell_hashes),
                    )
            session.save()

        after = audit_docx(output_path)
        _verify_protected_structure(before, after)
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
        verify_saved_journal_neutral(output_path, specs)
        return after
