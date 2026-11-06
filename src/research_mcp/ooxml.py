"""Read-only OOXML parsing used by preconditions and postconditions."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"


@dataclass(frozen=True)
class ZoteroPackageFields:
    """Rendered Zotero field state read directly from a closed DOCX package."""

    item_results: tuple[tuple[str, str], ...]
    bibliography_results: tuple[str, ...]
    item_keys: tuple[tuple[str, ...], ...] = ()


def zotero_package_fields(path: str | Path) -> ZoteroPackageFields:
    """Read complex Zotero field codes/results without opening Word.

    Results may span runs and paragraphs (the bibliography normally does), so
    this uses the field begin/separate/end state machine rather than a regex.
    Text from nested fields contributes to every active parent result.
    """

    with zipfile.ZipFile(Path(path).resolve(strict=True)) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))

    active: list[dict[str, object]] = []
    completed: list[tuple[str, str]] = []
    for node in root.iter():
        if node.tag == W + "p":
            for field in active:
                if field["separated"] and field["result"]:
                    field["result"].append("\n")
            continue
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType", "")
            if kind == "begin":
                active.append({"instruction": [], "result": [], "separated": False})
            elif kind == "separate" and active:
                active[-1]["separated"] = True
            elif kind == "end" and active:
                field = active.pop()
                completed.append((
                    "".join(field["instruction"]).strip(),
                    "".join(field["result"]).strip(),
                ))
            continue

        if node.tag == W + "instrText" and active and not active[-1]["separated"]:
            active[-1]["instruction"].append(node.text or "")
            continue

        if not active:
            continue
        if node.tag == W + "t":
            value = node.text or ""
        elif node.tag == W + "tab":
            value = "\t"
        elif node.tag in (W + "br", W + "cr"):
            value = "\n"
        else:
            continue
        for field in active:
            if field["separated"]:
                field["result"].append(value)

    item_results: list[tuple[str, str]] = []
    item_keys: list[tuple[str, ...]] = []
    bibliography_results: list[str] = []
    decoder = json.JSONDecoder()
    for instruction, result in completed:
        if instruction.startswith("ADDIN ZOTERO_ITEM"):
            start = instruction.find("{")
            if start < 0:
                continue
            try:
                payload, _ = decoder.raw_decode(instruction[start:])
            except (json.JSONDecodeError, TypeError):
                continue
            citation_id = payload.get("citationID") if isinstance(payload, dict) else None
            if isinstance(citation_id, str) and citation_id:
                item_results.append((citation_id, result))
                keys: list[str] = []
                for item in payload.get("citationItems", []):
                    if not isinstance(item, dict):
                        continue
                    for uri in item.get("uris", []):
                        match = re.search(r"/items/([^/]+)$", str(uri))
                        if match:
                            keys.append(match.group(1))
                item_keys.append(tuple(keys))
        elif instruction.startswith("ADDIN ZOTERO_BIBL"):
            bibliography_results.append(result)

    return ZoteroPackageFields(
        item_results=tuple(item_results),
        bibliography_results=tuple(bibliography_results),
        item_keys=tuple(item_keys),
    )


def normalized_text(node: ET.Element) -> str:
    return " ".join("".join(t.text or "" for t in node.findall(".//w:t", NS)).split())


def body_text(root: ET.Element) -> str:
    body = root.find("w:body", NS)
    if body is None:
        return ""
    paragraphs: list[str] = []
    for paragraph in body.iterfind(".//w:p", NS):
        # Nested table paragraphs are intentionally included. This snapshot is
        # structural evidence, not a prose view.
        paragraphs.append("".join(t.text or "" for t in paragraph.findall(".//w:t", NS)))
    return "\n".join(paragraphs)


def top_level_prose_text(root: ET.Element) -> str:
    body = root.find("w:body", NS)
    if body is None:
        return ""
    return "\n".join(
        "".join(t.text or "" for t in paragraph.findall(".//w:t", NS))
        for paragraph in body.findall("w:p", NS)
    )


@dataclass(frozen=True)
class LogicalTable:
    rows: int
    columns: int
    cells: tuple[tuple[str, ...], ...]
    vertical_merges: tuple[tuple[int, int, int, str], ...]
    horizontal_spans: tuple[tuple[int, int, int], ...]
    cell_hashes: tuple[str, ...] = ()

    def __post_init__(self):
        if not self.cell_hashes and self.cells:
            hashes = tuple(
                hashlib.sha256(value.encode("utf-8")).hexdigest()
                for row in self.cells
                for value in row
            )
            object.__setattr__(self, "cell_hashes", hashes)


def parse_table(table: ET.Element) -> LogicalTable:
    grid = table.findall("w:tblGrid/w:gridCol", NS)
    declared_columns = len(grid)
    raw_rows: list[list[str | None]] = []
    active_vertical: dict[int, tuple[int, str]] = {}
    merge_starts: dict[tuple[int, int], str] = {}
    merge_ends: dict[tuple[int, int], int] = {}
    horizontal_spans: list[tuple[int, int, int]] = []
    max_columns = declared_columns

    for row_index, row in enumerate(table.findall("w:tr", NS), start=1):
        parsed_row: list[str | None] = []
        column_cursor = 1
        for cell in row.findall("w:tc", NS):
            text = "".join(t.text or "" for t in cell.findall(".//w:t", NS))
            grid_span_node = cell.find("w:tcPr/w:gridSpan", NS)
            span = int(grid_span_node.get(W + "val", "1")) if grid_span_node is not None else 1
            if span > 1:
                horizontal_spans.append((row_index, column_cursor, span))

            vmerge = cell.find("w:tcPr/w:vMerge", NS)
            if vmerge is not None:
                val = vmerge.get(W + "val")
                if val == "restart":
                    active_vertical[column_cursor] = (row_index, text)
                    merge_starts[(row_index, column_cursor)] = text
                    merge_ends[(row_index, column_cursor)] = row_index
                elif column_cursor in active_vertical:
                    start_row, label = active_vertical[column_cursor]
                    text = label
                    merge_ends[(start_row, column_cursor)] = row_index
            else:
                active_vertical.pop(column_cursor, None)

            for _ in range(span):
                parsed_row.append(text)
                column_cursor += 1

        max_columns = max(max_columns, len(parsed_row))
        raw_rows.append(parsed_row)

    normalized_rows: list[tuple[str, ...]] = []
    for row in raw_rows:
        normalized_rows.append(tuple((value or "") for value in row + [""] * (max_columns - len(row))))

    vertical_merges = tuple(
        (start_row, merge_ends.get((start_row, column), start_row), column, label)
        for (start_row, column), label in sorted(merge_starts.items())
    )
    return LogicalTable(
        rows=len(normalized_rows),
        columns=max_columns,
        cells=tuple(normalized_rows),
        vertical_merges=vertical_merges,
        horizontal_spans=tuple(horizontal_spans),
    )


def bookmark_references(root: ET.Element) -> dict[str, tuple[str, ...]]:
    references: dict[str, list[str]] = {}
    for hyperlink in root.findall(".//w:hyperlink", NS):
        anchor = hyperlink.get(W + "anchor")
        if anchor:
            references.setdefault(anchor, []).append("hyperlink")

    instructions = " ".join(t.text or "" for t in root.findall(".//w:instrText", NS))
    for kind, name in re.findall(r"\b(REF|PAGEREF)\s+([^\s\\]+)", instructions, flags=re.I):
        references.setdefault(name, []).append(kind.upper())
    return {name: tuple(kinds) for name, kinds in references.items()}
