"""Guards for offline paragraph addressing and the inspect_document window, on synthetic documents.

The load-bearing claim is that `read_paragraphs` emits the same coordinates `document.Paragraphs(i)` does.
If that drifts, an agent reads an index from `inspect_document` and `insert_zotero_citation` silently edits
a different paragraph -- a wrong-place mutation with no error anywhere. It is checked against
`tests/data/com_paragraphs_parity.json`, a capture of every paragraph of `tests/data/paragraph_parity.docx`
(text, style, in-table) recorded once from live Word; both come from
`tests/fixture_builders/paragraph_parity.py`. Comparing against the capture keeps this file offline: no Word,
no Zotero, no integration lease.

The window, outline and summary rules below hold for any document, so they run on a generated one: a
title, five headed sections and a table, more paragraphs than one window may return.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

import pytest

from research_mcp.paragraphs import (
    DEFAULT_WINDOW,
    MAX_WINDOW,
    _paragraph_text,
    _read_paragraphs,
    inspect_document,
    outline,
    read_paragraphs,
    read_snapshot,
)
from research_mcp.audit import file_sha256
from research_mcp.ooxml import NS

ROOT = Path(__file__).resolve().parents[1]
PARITY_DOCX = ROOT / "tests" / "data" / "paragraph_parity.docx"
PARITY_CAPTURE = ROOT / "tests" / "data" / "com_paragraphs_parity.json"
PARITY_BUILDER = ROOT / "tests" / "fixture_builders" / "paragraph_parity.py"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SECTIONS = 5
BODY_PER_SECTION = 59


def _p(text: str, style: str | None = None) -> str:
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def _table() -> str:
    cell = lambda text: f"<w:tc>{_p(text)}</w:tc>"  # noqa: E731
    rows = "".join(f"<w:tr>{cell(f'R{r}A')}{cell(f'R{r}B')}</w:tr>" for r in (1, 2))
    return f"<w:tbl>{rows}</w:tbl>"


@pytest.fixture
def long_docx(tmp_path: Path) -> Path:
    """Title, five Heading1 sections of plain body text, a 2x2 table in section 2: 307 Word paragraphs."""

    parts = [_p("A Synthetic Study of Window Sizes", "Title")]
    for section in range(1, SECTIONS + 1):
        parts.append(_p(f"{section}. Section {section}", "Heading1"))
        for n in range(1, BODY_PER_SECTION + 1):
            parts.append(_p(f"Section {section} sentence {n} describes one measured result in plain words."))
        if section == 2:
            parts.append(_table())
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W_NS}"><w:body>'
           + "".join(parts) + "<w:sectPr/></w:body></w:document>")
    path = tmp_path / "long.docx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml",
                         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        archive.writestr("word/document.xml", xml)
    return path


TOTAL = 1 + SECTIONS * (1 + BODY_PER_SECTION) + 2 * 2 + 2  # + table cells + row ends


# --------------------------------------------------------------------------
# The parity claim, against a recorded Word capture
# --------------------------------------------------------------------------


def _capture() -> list[dict]:
    return json.loads(PARITY_CAPTURE.read_text(encoding="utf-8"))


def test_the_committed_fixture_is_what_the_builder_makes(tmp_path):
    """The capture describes these exact bytes; an edited fixture needs a new capture, not a silent pass."""

    spec = importlib.util.spec_from_file_location("parity_builder", PARITY_BUILDER)
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    def parts(path):
        with zipfile.ZipFile(path) as archive:
            return {name: archive.read(name) for name in archive.namelist()}

    assert parts(builder.build(tmp_path / "rebuilt.docx")) == parts(PARITY_DOCX)


def test_paragraph_indices_match_word_exactly():
    """Same count, same order, same in-table status, for every paragraph Word reports."""

    captured = _capture()
    parsed = read_paragraphs(PARITY_DOCX)

    assert len(parsed) == len(captured) == 83
    assert [item.index for item in parsed] == [row["i"] for row in captured] == list(range(1, 84))
    for mine, theirs in zip(parsed, captured):
        assert mine.in_table == theirs["in_table"], mine.index


def test_a_flat_ooxml_walk_would_be_wrong_by_one_per_table_row():
    """Pins why the table row-end paragraph cannot be skipped.

    `.//w:p` is the obvious way to read paragraphs and it silently produces indices that drift once the
    first table starts: by one per row, nested tables included.
    """

    root = ET.fromstring(zipfile.ZipFile(PARITY_DOCX).read("word/document.xml"))
    body = root.find("w:body", NS)
    flat = body.findall(".//w:p", NS)
    rows = body.findall(".//w:tr", NS)
    assert len(flat) == 67
    assert len(rows) == 16
    assert len(read_paragraphs(PARITY_DOCX)) - len(flat) == len(rows)

    # The two agree up to the first row-end paragraph (Word's 14) and disagree from there on.
    word_texts = [row["text"] for row in _capture()]
    flat_texts = [_paragraph_text(node) for node in flat]
    first_row_end = 14
    assert word_texts[first_row_end - 1] == "" and _capture()[first_row_end - 1]["in_table"]
    assert flat_texts[: first_row_end - 1] == word_texts[: first_row_end - 1]
    assert flat_texts[first_row_end - 1] == "Alpha" != word_texts[first_row_end - 1]


def test_text_is_exact_wherever_it_claims_to_be():
    """`text_exact` is a promise: if True, the text is what Word reports.

    Over-flagging is acceptable and under-flagging is not, because an `anchor` matched against wrong text
    lands in the wrong place.
    """

    captured = _capture()
    parsed = read_paragraphs(PARITY_DOCX)

    for mine, theirs in zip(parsed, captured):
        if mine.text_exact:
            assert mine.text == theirs["text"], f"paragraph {mine.index} claimed exact"

    differing = [mine.index for mine, theirs in zip(parsed, captured) if mine.text != theirs["text"]]
    assert differing, "the capture should still exercise the inexact path"
    assert all(not parsed[index - 1].text_exact for index in differing)


def test_fields_breaks_and_tabs_read_as_word_reads_them():
    """Field results, page breaks (\\x0c), line breaks (\\v), tab characters and tab stops."""

    captured = _capture()
    parsed = read_paragraphs(PARITY_DOCX)
    by_kind = {
        "field": [item for item in parsed if item.has_fields],
        "page break": [item for item in parsed if item.text == "\x0c"],
        "line break": [item for item in parsed if "\v" in item.text],
        "leading tab": [item for item in parsed if item.text.startswith("\t")],
        "tab stop": [item for item in parsed if item.text == "Label\tValue"],
    }
    for kind, items in by_kind.items():
        assert items, kind
        for item in items:
            assert item.text == captured[item.index - 1]["text"], (kind, item.index)


def test_bibliography_entries_keep_the_tab_after_their_number():
    """The `w:t`-only join dropped these, one character per entry."""

    captured = _capture()
    entries = [item for item in read_paragraphs(PARITY_DOCX) if item.style == "Bibliography"]
    assert len(entries) == sum(1 for row in captured if row["style"] == "Bibliography") == 4
    for item in entries:
        assert "\t" in item.text
        assert item.text == captured[item.index - 1]["text"]


def test_outline_finds_the_headings_word_reports():
    captured = _capture()
    expected = [row["i"] for row in captured if row["style"].startswith(("Heading", "Title"))]
    assert [item.index for item in outline(PARITY_DOCX)] == expected


# --------------------------------------------------------------------------
# Structure
# --------------------------------------------------------------------------


def test_row_end_paragraph_follows_each_row(minimal_docx):
    """The fixture table is 3 rows x 2 columns, so 3 row-end paragraphs."""

    parsed = read_paragraphs(minimal_docx)
    in_table = [item for item in parsed if item.in_table]
    assert len(in_table) == 3 * 2 + 3

    texts = [item.text for item in in_table]
    assert texts == ["LoRA", "Early", "", "", "Middle", "", "", "Late", ""]


def test_indices_are_contiguous_and_one_based(minimal_docx):
    parsed = read_paragraphs(minimal_docx)
    assert [item.index for item in parsed] == list(range(1, len(parsed) + 1))


def test_inspection_hash_describes_the_same_package(minimal_docx):
    result = inspect_document(minimal_docx)
    assert result["source_sha256"] == file_sha256(minimal_docx)


def test_snapshot_remains_self_consistent_if_the_path_changes(minimal_docx, tmp_path):
    target = tmp_path / "changing.docx"
    target.write_bytes(minimal_docx.read_bytes())
    snapshot = read_snapshot(target)
    expected = _read_paragraphs(snapshot)

    # Replace the path after capture. Parsing the snapshot must still describe
    # the bytes whose SHA the caller received, not this later package.
    target.write_bytes(b"not a docx anymore")
    assert _read_paragraphs(snapshot) == expected
    assert snapshot.sha256 != file_sha256(target)


def test_synthetic_document_has_the_expected_paragraph_count(long_docx):
    parsed = read_paragraphs(long_docx)
    assert len(parsed) == TOTAL > MAX_WINDOW
    assert sum(1 for item in parsed if item.in_table) == 2 * 2 + 2


def test_outline_is_headings_only_and_far_smaller_than_the_document(long_docx):
    headings = outline(long_docx)
    assert len(headings) == 1 + SECTIONS
    assert all((item.style or "").lower().startswith(("heading", "title")) for item in headings)
    assert len(headings) < len(read_paragraphs(long_docx)) / 10


# --------------------------------------------------------------------------
# The tool result
# --------------------------------------------------------------------------


def test_default_call_returns_an_outline_not_the_document(long_docx):
    """Whole-document dumps are the payload failure this project just removed."""

    result = inspect_document(long_docx)
    assert result["paragraph_count"] == TOTAL
    assert "outline" in result
    assert "paragraphs" not in result

    body = json.dumps(result, ensure_ascii=False)
    full = json.dumps([item.to_dict() for item in read_paragraphs(long_docx)], ensure_ascii=False)
    assert len(body) < len(full) / 10, (len(body), len(full))


def test_a_window_returns_paragraphs_and_reports_its_own_bounds(long_docx):
    result = inspect_document(long_docx, first_paragraph=10, last_paragraph=14)
    assert result["first_paragraph"] == 10
    assert result["last_paragraph"] == 14
    assert [item["index"] for item in result["paragraphs"]] == [10, 11, 12, 13, 14]
    assert result["truncated"] is False


def test_an_open_ended_window_uses_the_default_size(long_docx):
    result = inspect_document(long_docx, first_paragraph=1)
    assert len(result["paragraphs"]) == DEFAULT_WINDOW


def test_an_oversized_window_is_clamped_and_says_so(long_docx):
    """A single call must not be able to dump the document."""

    result = inspect_document(long_docx, first_paragraph=1, last_paragraph=TOTAL)
    assert result["truncated"] is True
    assert len(result["paragraphs"]) == MAX_WINDOW
    assert result["last_paragraph"] == MAX_WINDOW


def test_a_window_running_past_the_end_stops_at_the_end(long_docx):
    result = inspect_document(long_docx, first_paragraph=TOTAL - 2, last_paragraph=TOTAL + 400)
    assert result["last_paragraph"] == TOTAL
    assert [item["index"] for item in result["paragraphs"]] == [TOTAL - 2, TOTAL - 1, TOTAL]


def test_a_window_starting_past_the_end_is_refused(long_docx):
    with pytest.raises(ValueError, match="past the end"):
        inspect_document(long_docx, first_paragraph=TOTAL + 400)


def test_a_reversed_window_is_refused(long_docx):
    with pytest.raises(ValueError, match="precedes"):
        inspect_document(long_docx, first_paragraph=50, last_paragraph=10)


def test_summary_counts_are_consistent_with_the_paragraphs(long_docx):
    result = inspect_document(long_docx)
    parsed = read_paragraphs(long_docx)
    assert result["table_paragraph_count"] == sum(1 for item in parsed if item.in_table)
    assert result["field_paragraph_count"] == sum(1 for item in parsed if item.has_fields)
    assert result["inexact_text_paragraph_count"] == sum(1 for item in parsed if not item.text_exact)


def test_outline_entries_carry_only_what_a_heading_can_be_acted_on_by(long_docx):
    """Guards against re-growing the payload this tool exists to keep small."""

    result = inspect_document(long_docx)
    assert result["outline"]
    for entry in result["outline"]:
        assert set(entry) == {"index", "style", "text"}, entry
    assert len(json.dumps(result, ensure_ascii=False)) < 4000
