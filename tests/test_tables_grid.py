"""format_tables on tables with gridSpan cells and on a body that ends with a table (plan 973 S1).

Word's ``Table.Cell(row, n)`` numbers the real cells of a row, so a row whose first cell spans the grid has one
cell, and ``Cell(1, 2)`` raises "The requested member of the collection does not exist". Cells are mapped to grid
columns from the XML (``table_layouts``). Word's save appends an empty paragraph after a body-final table; the
saved checks compare against Word's own save of the untouched document (``word_baseline``).

Unit parts use fake Word tables; live parts need Word (RESEARCH_MCP_WORD_LIVE=1). Fixtures are synthetic."""

from __future__ import annotations

import hashlib
import os
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest

from research_mcp import tables
from research_mcp.audit import audit_docx, file_sha256
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.ooxml import parse_table, row_cells
from research_mcp.tables import MergeGroup, TableFormatSpec, check_spans, table_layouts

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
live = pytest.mark.skipif(os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1", reason="requires Word")

LABEL = "Research Assistant"
HEADER = "Experience"
DATA = [[LABEL, "Lab A", "2024", "2025"], [LABEL, "Lab B", "2025", "2026"]]


def build_cv(path: Path, *, span_data_row: bool = False) -> Path:
    """A CV-like page whose body ends with a 3x4 table: row 1 is one cell spanning the 4 grid columns, rows 2-3
    repeat the label in column 1. ``span_data_row`` also merges columns 1-2 of row 3 into one cell."""

    from docx import Document

    document = Document()
    document.add_paragraph("Sample Person")
    document.add_paragraph("Summary line of a synthetic CV.")
    table = document.add_table(rows=3, cols=4)
    header = table.cell(0, 0).merge(table.cell(0, 3))
    header.text = HEADER
    for r, row in enumerate(DATA, start=1):
        for c, value in enumerate(row):
            table.cell(r, c).text = value
    if span_data_row:
        merged = table.cell(2, 0).merge(table.cell(2, 1))
        merged.text = LABEL
    document.save(path)
    return path


@pytest.fixture
def cv(tmp_path) -> Path:
    return build_cv(tmp_path / "cv.docx")


def _body_tags(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        body = ET.fromstring(archive.read("word/document.xml")).find(W + "body")
    return [child.tag.removeprefix(W) for child in body]


# -- the grid map --------------------------------------------------------------------------------------------------


def test_row_cells_reads_spans_and_grid_before():
    row = ET.fromstring(
        f'<w:tr xmlns:w="{W_NS}"><w:trPr><w:gridBefore w:val="1"/></w:trPr>'
        '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>')
    assert row_cells(row) == ((2, 2), (4, 1))


def test_table_layouts_of_a_spanned_header(cv):
    assert table_layouts(cv) == ((((1, 4),), ((1, 1), (2, 1), (3, 1), (4, 1)), ((1, 1), (2, 1), (3, 1), (4, 1))),)
    assert _body_tags(cv)[-2:] == ["tbl", "sectPr"]  # the body ends with the table


LAYOUT = (((1, 4),), ((1, 1), (2, 2), (4, 1)), ((1, 1), (2, 1), (3, 1), (4, 1)))


def test_merge_inside_a_spanned_cell_is_refused_by_name():
    spec = TableFormatSpec(index=1, descriptive_columns=(1,), merges=(MergeGroup(3, 2, 3, "x"),))
    with pytest.raises(RefusedUnsafe, match=r"merge of column 3 rows 2-3: in table 1 row 2 the cell at column 3 "
                                            r"spans grid columns 2-3; only cells one column wide can be merged"):
        check_spans(spec, LAYOUT)


def test_spanned_data_cell_mixing_descriptive_and_data_columns_is_refused_by_name():
    spec = TableFormatSpec(index=1, descriptive_columns=(1, 2))
    with pytest.raises(RefusedUnsafe, match=r"table 1 row 2: the cell spanning grid columns 2-3 holds descriptive "
                                            r"column\(s\) \[2\] and data column\(s\) \[3\]"):
        check_spans(spec, LAYOUT)


def test_header_span_and_unmixed_spans_are_fine():
    check_spans(TableFormatSpec(index=1, descriptive_columns=(1,)), LAYOUT)  # header spans all; 2-3 are data
    check_spans(TableFormatSpec(index=1, descriptive_columns=(1, 2, 3)), LAYOUT)  # 2-3 both descriptive
    check_spans(TableFormatSpec(index=1, descriptive_columns=(1,), merges=(MergeGroup(1, 2, 3, "x"),)), LAYOUT)


# -- fake Word: a table whose rows have fewer real cells than grid columns -----------------------------------------


class _Collection:
    def __init__(self, count: int, item):
        self.Count = count
        self._item = item

    def __call__(self, index):
        return self._item(index)


def _border(_id=None):
    return SimpleNamespace(LineStyle=None, LineWidth=None)


class _RaggedTable:
    """Word's view of a table with spans: ``Cell(row, n)`` numbers a row's real cells and raises past them."""

    def __init__(self, rows: list[list[str]], log: list, columns: int = 4):
        self.cells = rows
        self.log = log
        self.Rows = _Collection(len(rows), lambda index: SimpleNamespace(Borders=_border))
        self.Columns = _Collection(columns, lambda index: SimpleNamespace(Borders=_border))
        self.Borders = _border

    def Cell(self, row: int, position: int):
        if position > len(self.cells[row - 1]):
            raise RuntimeError("The requested member of the collection does not exist.")
        self.log.append(("cell", row, position))
        table, text = self, self.cells[row - 1][position - 1]

        class Paragraph:
            @property
            def Alignment(self):
                return None

            @Alignment.setter
            def Alignment(self, value):
                table.log.append(("align", row, position, value))

        paragraphs = _Collection(1, lambda index: Paragraph())
        return SimpleNamespace(Range=SimpleNamespace(Text=text + "\r\x07", Paragraphs=paragraphs))


def _document(table) -> SimpleNamespace:
    return SimpleNamespace(Tables=_Collection(1, lambda index: table))


CV_CELLS = [[HEADER], *DATA]


def test_snapshot_live_matches_the_saved_audit_on_a_spanned_header(cv):
    """The live snapshot fills a spanned cell's text into every grid column it covers, as the saved audit does, so
    the live and saved checks compare the same grid."""

    live_table = tables.snapshot_live(_document(_RaggedTable(CV_CELLS, [])), table_layouts(cv))[0]
    saved = audit_docx(cv).tables[0]
    assert (live_table.rows, live_table.columns, live_table.cell_hashes) == (saved.rows, saved.columns,
                                                                             saved.cell_hashes)
    header = hashlib.sha256(HEADER.encode()).hexdigest()
    assert live_table.cell_hashes[:4] == (header,) * 4


def test_apply_journal_neutral_addresses_only_real_cells(cv):
    log: list = []
    spec = TableFormatSpec(index=1, descriptive_columns=(1, 2))
    tables.apply_journal_neutral(_document(_RaggedTable(CV_CELLS, log)), spec, table_layouts(cv)[0])
    aligned = [entry[1:] for entry in log if entry[0] == "align"]
    left, right = tables.WD_ALIGN_PARAGRAPH_LEFT, tables.WD_ALIGN_PARAGRAPH_RIGHT
    assert aligned == [(1, 1, left),
                       (2, 1, left), (2, 2, left), (2, 3, right), (2, 4, right),
                       (3, 1, left), (3, 2, left), (3, 3, right), (3, 4, right)]


def test_apply_journal_neutral_without_the_grid_map_hits_words_error(cv):
    """What the replay hit before plan 973: addressing cells by grid column reaches Cell(1, 2) in a row with one
    cell."""

    spec = TableFormatSpec(index=1, descriptive_columns=(1, 2))
    with pytest.raises(RuntimeError, match="requested member of the collection does not exist"):
        tables.apply_journal_neutral(_document(_RaggedTable(CV_CELLS, [])), spec, None)


def test_merge_addresses_the_real_cell_after_a_span():
    """Rows whose first cell spans columns 1-2: grid column 3 is Word's Cell(row, 2)."""

    class Mergeable(_RaggedTable):
        def Cell(self, row, position):
            cell = super().Cell(row, position)
            cell.Merge = lambda other: self.log.append(("merge", row, position))
            return cell

    log: list = []
    table = Mergeable([[HEADER], ["Span", LABEL, "1"], ["Span", LABEL, "2"]], log)
    layout = (((1, 4),), ((1, 2), (3, 1), (4, 1)), ((1, 2), (3, 1), (4, 1)))
    tables.merge_repeated_label(table, MergeGroup(3, 2, 3, LABEL), layout=layout, keep_formatting=False)
    assert ("merge", 2, 2) in log and not any(entry[0] == "cell" and entry[2] == 3 for entry in log)


def test_format_to_version_refuses_a_span_before_word_starts(tmp_path, monkeypatch):
    source = build_cv(tmp_path / "spanned.docx", span_data_row=True)
    monkeypatch.setattr(tables, "IsolatedWordDocument", lambda *a, **k: pytest.fail("Word must not start"))
    out = tmp_path / "out.docx"
    spec = TableFormatSpec(index=1, descriptive_columns=(1,))
    with pytest.raises(RefusedUnsafe, match=r"table 1 row 3: the cell spanning grid columns 1-2 holds descriptive"):
        tables.format_to_version(source, out, [spec])
    assert not out.exists()


# -- the saved parse honors gridBefore/gridAfter (plan 973 S1 r1) -----------------------------------------------------

RAGGED_LABEL = "Grant"
RAGGED = [["Kind", "Name", "Amount"], [RAGGED_LABEL, "10"], [RAGGED_LABEL, "20"]]


def build_ragged(path: Path, *, after: bool = False) -> Path:
    """A 3x3 table whose rows 2-3 have two cells: ``w:gridBefore`` 1 (or ``w:gridAfter`` 1 with ``after``) stands
    for the missing grid column, so their cells sit at grid columns 2-3 (1-2)."""

    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = Document()
    document.add_paragraph("Funding")
    table = document.add_table(rows=3, cols=3)
    for column, value in enumerate(RAGGED[0]):
        table.cell(0, column).text = value
    for row_index in (1, 2):
        tr = table.rows[row_index]._tr
        cells = tr.findall(qn("w:tc"))
        tr.remove(cells[-1] if after else cells[0])
        for tc, value in zip(tr.findall(qn("w:tc")), RAGGED[row_index]):
            tc.findall(qn("w:p"))[0].append(_run(value))
        tr_pr = tr.get_or_add_trPr()
        skip = OxmlElement("w:gridAfter" if after else "w:gridBefore")
        skip.set(qn("w:val"), "1")
        tr_pr.append(skip)
    document.add_paragraph("End.")
    document.save(path)
    return path


def _run(text: str):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    run, node = OxmlElement("w:r"), OxmlElement("w:t")
    node.text = text
    run.append(node)
    return run


def _first_table(path: Path):
    with zipfile.ZipFile(path) as archive:
        return ET.fromstring(archive.read("word/document.xml")).find(f".//{W}tbl")


def test_parse_table_places_cells_after_grid_before_and_fills_grid_after(tmp_path):
    before = parse_table(_first_table(build_ragged(tmp_path / "before.docx")))
    assert before.cells == (("Kind", "Name", "Amount"), ("", RAGGED_LABEL, "10"), ("", RAGGED_LABEL, "20"))
    after = parse_table(_first_table(build_ragged(tmp_path / "after.docx", after=True)))
    assert after.cells == (("Kind", "Name", "Amount"), (RAGGED_LABEL, "10", ""), (RAGGED_LABEL, "20", ""))
    rows = _first_table(tmp_path / "before.docx").findall(f"{W}tr")
    assert [row_cells(row) for row in rows] == [((1, 1), (2, 1), (3, 1)), ((2, 1), (3, 1)), ((2, 1), (3, 1))]


def test_snapshot_live_matches_the_saved_audit_on_a_ragged_row(tmp_path):
    """Word's Cell(2, 1) is grid column 2 of a gridBefore row: the live snapshot and the saved audit agree."""

    source = build_ragged(tmp_path / "ragged.docx")
    live_table = tables.snapshot_live(_document(_RaggedTable(RAGGED, [], columns=3)), table_layouts(source))[0]
    assert live_table.cell_hashes == audit_docx(source).tables[0].cell_hashes


def test_cells_moved_from_grid_before_to_grid_after_are_a_changed_table(tmp_path):
    """Same text, other grid columns: the saved check sees the change."""

    source, moved = build_ragged(tmp_path / "before.docx"), build_ragged(tmp_path / "moved.docx", after=True)
    with pytest.raises(VerificationFailed, match="saved table 1 logical data changed"):
        tables.verify_saved_tables(audit_docx(source), moved, [TableFormatSpec(index=1, descriptive_columns=(1, 2))],
                                    reference=source)


# -- the saved check reads what merge continuations physically hold (plan 973 S1 r1) ---------------------------------


def _merged(source: Path, target: Path, continuation: str) -> Path:
    """``source`` (the CV) with column 1 rows 2-3 vertically merged as Word saves it; the continuation cell holds
    ``continuation``."""

    with zipfile.ZipFile(source) as archive:
        xml = archive.read("word/document.xml").decode()
    rows = xml.split("<w:tr")
    assert len(rows) == 4  # text before the table, then three rows

    def merge(row: str, value: str, text: str) -> str:
        head, tail = row.split("<w:tc>", 1)
        tc_pr, rest = tail.split("</w:tcPr>", 1)
        first_paragraph_end = rest.index("</w:p>")
        cell_end = rest.index("</w:tc>")
        assert first_paragraph_end < cell_end
        paragraph = rest[:rest.index("<w:r>")] if "<w:r>" in rest[:cell_end] else rest[:first_paragraph_end]
        run = f"<w:r><w:t>{text}</w:t></w:r>" if text else ""
        return (f"{head}<w:tc>{tc_pr}<w:vMerge{value}/></w:tcPr>{paragraph}{run}</w:p></w:tc>"
                f"{rest[cell_end + len('</w:tc>'):]}")

    rows[2] = merge(rows[2], ' w:val="restart"', LABEL)
    rows[3] = merge(rows[3], "", continuation)
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as new:
        for info in old.infolist():
            data = old.read(info.filename)
            new.writestr(info, "<w:tr".join(rows).encode() if info.filename == "word/document.xml" else data)
    return target


MERGE_SPEC = [TableFormatSpec(index=1, descriptive_columns=(1, 2), merges=(MergeGroup(1, 2, 3, LABEL),))]


def test_a_continuation_holding_text_is_refused(cv, tmp_path):
    ok, lost = _merged(cv, tmp_path / "ok.docx", ""), _merged(cv, tmp_path / "lost.docx", "lost data")
    # what the cell hashes cannot see: the saved audit reads both continuations as the label
    assert audit_docx(ok).tables[0].cell_hashes == audit_docx(lost).tables[0].cell_hashes
    assert parse_table(_first_table(lost)).continuations == ((3, 1, "lost data"),)
    tables.verify_saved_continuations(cv, ok, MERGE_SPEC)
    with pytest.raises(VerificationFailed, match=r"saved table 1: merged cells below the first of a merge hold "
                                                 r"other text than expected \(row 3 column 1\)"):
        tables.verify_saved_continuations(cv, lost, MERGE_SPEC)


def test_a_continuation_already_in_the_source_keeps_its_text(cv, tmp_path):
    """A merge the pass did not make: its continuation must hold what it held."""

    source = _merged(cv, tmp_path / "source.docx", "kept")
    plain = [TableFormatSpec(index=1, descriptive_columns=(1, 2))]
    tables.verify_saved_continuations(source, _merged(cv, tmp_path / "same.docx", "kept"), plain)
    with pytest.raises(VerificationFailed, match=r"\(row 3 column 1\)"):
        tables.verify_saved_continuations(source, _merged(cv, tmp_path / "emptied.docx", ""), plain)


# -- live: the CV-style body-final table ---------------------------------------------------------------------------

CV_SPEC = [{"index": 1, "descriptive_columns": [1, 2],
            "merges": [{"column": 1, "start_row": 2, "end_row": 3, "label": LABEL}]}]


def _dump(result):
    return result.model_dump() if hasattr(result, "model_dump") else result


@live
def test_live_cv_body_final_table_plain(cv, tmp_path):
    from research_mcp import server

    before = file_sha256(cv)
    out = tmp_path / "cv_plain.docx"
    result = _dump(server.format_tables(str(cv), str(out), CV_SPEC, tracked=False))
    assert result["tracking"] == "off" and result["compared_against"] == "word_baseline"
    saved = audit_docx(out)
    assert saved.tables[0].vertical_merges == ((2, 3, 1, LABEL),)
    assert saved.paragraph_count == audit_docx(cv).paragraph_count + 1  # Word's appended paragraph
    assert _body_tags(out)[-3:] == ["tbl", "p", "sectPr"]
    assert file_sha256(cv) == before


@live
def test_live_cv_body_final_table_tracked_reviewer_b(cv, tmp_path):
    from research_mcp import server, word_tracked
    from research_mcp.track_setting import write_setting

    source = tmp_path / "cv_on.docx"
    write_setting(cv, source, on=True, author="Reviewer B")
    out = tmp_path / "cv_tracked.docx"
    result = _dump(server.format_tables(str(source), str(out), CV_SPEC))
    assert result["tracking"] == "on" and result["tracking_author"] == "Reviewer B"
    assert result["compared_against"] == "word_baseline"
    assert word_tracked.revision_authors(out) == {"Reviewer B"}
    assert [item["kind"] for item in result["untracked"]] == ["merge"]
    check = result["tracked_check"]
    assert check["reject_all_matches_input"] and check["accept_all_verified"]
    assert audit_docx(out).tables[0].vertical_merges == ((2, 3, 1, LABEL),)


@live
@pytest.mark.parametrize("tracked", [False, True])
def test_live_style_on_a_body_final_table(cv, tmp_path, tracked):
    from research_mcp import server

    out = tmp_path / f"cv_style_{tracked}.docx"
    result = server.apply_document_style(str(cv), str(out), file_sha256(cv), profile="compact_single",
                                         tracked=tracked)
    assert result["tracking"] == ("on" if tracked else "off")
    assert result["compared_against"] == "word_baseline"
    assert result["protected_structure_unchanged"] is True


def _with_explicit_zeros(source: Path, target: Path) -> Path:
    """``source`` with paragraph properties Word drops on save: an explicit ``keepNext`` off and zero spacing with
    autospacing off on the first paragraph. Word's save of the untouched copy then differs from it beyond the
    appended paragraph, so the baseline is not adopted."""

    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as new:
        for info in old.infolist():
            data = old.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                paragraph = "<w:p><w:r><w:t>Sample Person</w:t></w:r></w:p>"
                assert xml.count(paragraph) == 1
                data = xml.replace(paragraph, '<w:p><w:pPr><w:keepNext w:val="0"/><w:spacing w:before="0" '
                                              'w:after="0" w:beforeAutospacing="0" w:afterAutospacing="0"/></w:pPr>'
                                              '<w:r><w:t>Sample Person</w:t></w:r></w:p>').encode()
            new.writestr(info, data)
    return target


@live
@pytest.mark.parametrize("tracked", [False, True])
def test_live_body_final_table_with_an_unfaithful_baseline(cv, tmp_path, tracked):
    """Word's save of the untouched document drops the explicit zeros, so it is not adopted; the source is then
    compared in Word's body-final form (the appended empty paragraph after the table), and the pass applies."""

    from research_mcp import server, word_tracked
    from research_mcp.track_setting import write_setting

    source = _with_explicit_zeros(cv, tmp_path / "zeros.docx")
    if tracked:
        write_setting(source, tmp_path / "zeros_on.docx", on=True, author="Reviewer B")
        source = tmp_path / "zeros_on.docx"
    before = file_sha256(source)
    out = tmp_path / f"zeros_{tracked}.docx"
    result = _dump(server.format_tables(str(source), str(out), CV_SPEC, tracked=None if tracked else False))
    assert result["compared_against"] == "source"
    assert audit_docx(out).tables[0].vertical_merges == ((2, 3, 1, LABEL),)
    assert _body_tags(out)[-3:] == ["tbl", "p", "sectPr"]
    if tracked:
        assert result["tracking_author"] == "Reviewer B" and word_tracked.revision_authors(out) == {"Reviewer B"}
        assert result["tracked_check"]["reject_all_matches_input"]
    assert file_sha256(source) == before


def _corrupt_continuation(path: Path) -> None:
    """Puts text into the first continuation cell (``<w:vMerge/>``) of the saved file."""

    with zipfile.ZipFile(path) as archive:
        parts = {info: archive.read(info.filename) for info in archive.infolist()}
    from lxml import etree

    for info, data in parts.items():
        if info.filename == "word/document.xml":
            root = etree.fromstring(data)
            merge = next(node for node in root.iter(W + "vMerge") if node.get(W + "val") in (None, "continue"))
            paragraph = merge.getparent().getparent().find(W + "p")  # vMerge -> tcPr -> tc
            run = etree.SubElement(paragraph, W + "r")
            etree.SubElement(run, W + "t").text = "lost data"
            parts[info] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in parts.items():
            archive.writestr(info, data)


def _change_page_width(path: Path) -> None:
    """Widens the first section's page in a saved file (``w:pgSz`` ``w:w``)."""

    with zipfile.ZipFile(path) as archive:
        parts = {info: archive.read(info.filename) for info in archive.infolist()}
    from lxml import etree

    for info, data in parts.items():
        if info.filename == "word/document.xml":
            root = etree.fromstring(data)
            size = next(root.iter(W + "pgSz"))
            size.set(W + "w", str(int(size.get(W + "w")) + 1440))
            parts[info] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in parts.items():
            archive.writestr(info, data)


def _corrupt_saved_output(monkeypatch, corrupt) -> list[Path]:
    """Runs ``corrupt(path)`` on Word's saved output (not its baseline) after Word closes it; returns the paths
    Word saved."""

    from research_mcp import word

    saved: list[Path] = []
    original_save_as, original_exit = word.IsolatedWordDocument.save_as, word.IsolatedWordDocument.__exit__

    def save_as(self, target):
        result = original_save_as(self, target)
        saved.append(Path(result))
        return result

    def exit_(self, *exc):
        outcome = original_exit(self, *exc)
        outputs = [path for path in saved if path.name != "baseline.docx"]
        if outputs and outputs[-1].exists():
            corrupt(outputs[-1])
        return outcome

    monkeypatch.setattr(word.IsolatedWordDocument, "save_as", save_as)
    monkeypatch.setattr(word.IsolatedWordDocument, "__exit__", exit_)
    return saved


def _with_document_xml(source: Path, target: Path, transform) -> Path:
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as new:
        for info in old.infolist():
            data = old.read(info.filename)
            if info.filename == "word/document.xml":
                data = transform(data.decode("utf-8")).encode("utf-8")
            new.writestr(info, data)
    return target


def _word_form(xml: str) -> str:
    """Word's body-final form: an empty paragraph between the final table and the body's sectPr."""

    at = xml.rindex("<w:sectPr")
    return xml[:at] + '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>' + xml[at:]


def _wider(xml: str) -> str:
    assert xml.count('<w:pgSz w:w="12240"') == 1
    return xml.replace('<w:pgSz w:w="12240"', '<w:pgSz w:w="13680"')


@pytest.mark.parametrize("compared_against", ["word_baseline", "source"])
def test_a_changed_section_is_refused_against_either_reference(cv, tmp_path, compared_against):
    """format_tables never changes section properties or header/footer content: a saved file whose page width
    differs from the selected reference is refused, whether the reference is Word's untouched save or the source."""

    from research_mcp.word_baseline import choose_reference

    if compared_against == "word_baseline":
        baseline = _with_document_xml(cv, tmp_path / "baseline.docx", _word_form)
    else:  # the lead's probe: the baseline itself has the appended paragraph and a changed page width
        baseline = _with_document_xml(cv, tmp_path / "baseline.docx", lambda xml: _wider(_word_form(xml)))
    reference = choose_reference(cv, baseline)
    assert reference.compared_against == compared_against
    changed = _with_document_xml(cv, tmp_path / "changed.docx", lambda xml: _wider(_word_form(xml)))
    with pytest.raises(VerificationFailed, match=r"formatting tables changed section properties: section 1 "
                                                 r"properties, headers or footers changed"):
        tables.verify_saved_tables(audit_docx(reference.path), changed, MERGE_SPEC, reference=reference.path)


@live
@pytest.mark.parametrize("compared_against", ["word_baseline", "source"])
def test_live_a_changed_saved_section_is_refused_and_not_published(cv, tmp_path, monkeypatch, compared_against):
    """Word's saved output gets a wider page after Word closes it: the saved check refuses against the selected
    reference and nothing is published."""

    source = cv if compared_against == "word_baseline" else _with_explicit_zeros(cv, tmp_path / "zeros.docx")
    from research_mcp import word_baseline

    chosen: list[str] = []
    original_choose = word_baseline.choose_reference

    def choose(*args, **kwargs):
        reference = original_choose(*args, **kwargs)
        chosen.append(reference.compared_against)
        return reference

    monkeypatch.setattr(word_baseline, "choose_reference", choose)
    saved = _corrupt_saved_output(monkeypatch, _change_page_width)
    before = file_sha256(source)
    out = tmp_path / "out.docx"
    spec = [TableFormatSpec(index=1, descriptive_columns=(1, 2), merges=(MergeGroup(1, 2, 3, LABEL),))]
    with pytest.raises(VerificationFailed, match=r"formatting tables changed section properties: section 1 "
                                                 r"properties, headers or footers changed"):
        tables.format_to_version(source, out, spec)
    assert chosen == [compared_against]
    assert not out.exists()
    assert file_sha256(source) == before
    assert any(path.name != "baseline.docx" for path in saved)


@live
def test_live_a_corrupted_saved_continuation_is_refused_and_not_published(cv, tmp_path, monkeypatch):
    """Word's saved output (not its baseline) gets text in the merge's continuation cell after Word closes it: the
    saved check refuses and nothing is published. ``verify_saved_tables`` runs for real."""

    saved = _corrupt_saved_output(monkeypatch, _corrupt_continuation)
    before = file_sha256(cv)
    out = tmp_path / "out.docx"
    spec = [TableFormatSpec(index=1, descriptive_columns=(1, 2), merges=(MergeGroup(1, 2, 3, LABEL),))]
    with pytest.raises(VerificationFailed, match=r"merged cells below the first of a merge hold other text than "
                                                 r"expected \(row 3 column 1\)"):
        tables.format_to_version(cv, out, spec)
    assert not out.exists()
    assert file_sha256(cv) == before
    assert any(path.name != "baseline.docx" for path in saved)


def _with_logo_and_page_number(source: Path, target: Path) -> Path:
    """``source`` with a picture in its header and a PAGE field in its footer."""

    import struct
    import zlib

    from docx import Document
    from docx.oxml import parse_xml

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    logo = target.with_name("logo.png")
    logo.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
                     + chunk(b"IDAT", zlib.compress(bytes([0, 40, 90, 160]))) + chunk(b"IEND", b""))
    document = Document(source)
    section = document.sections[0]
    section.header.paragraphs[0].add_run().add_picture(str(logo))
    footer = section.footer.paragraphs[0]
    footer.add_run("Page ")
    for xml in ('<w:r %s><w:fldChar w:fldCharType="begin"/></w:r>', '<w:r %s><w:instrText xml:space="preserve"> PAGE '
                '</w:instrText></w:r>', '<w:r %s><w:fldChar w:fldCharType="separate"/></w:r>', '<w:r %s><w:t>1</w:t>'
                '</w:r>', '<w:r %s><w:fldChar w:fldCharType="end"/></w:r>'):
        footer._p.append(parse_xml(xml % f'xmlns:w="{W_NS}"'))
    document.save(target)
    return target


@live
@pytest.mark.parametrize("tracked", [False, True])
def test_live_a_header_picture_and_footer_page_field_are_preserved(cv, tmp_path, tracked):
    """A logo in the header and a PAGE field in the footer: Word rewrites the header's drawing markup on save, which
    is not a change, so format_tables applies against Word's untouched save, plain and tracked."""

    from research_mcp import server, word_baseline
    from research_mcp.track_setting import write_setting

    source = _with_logo_and_page_number(cv, tmp_path / "logo.docx")
    if tracked:
        write_setting(source, tmp_path / "logo_on.docx", on=True, author="Reviewer B")
        source = tmp_path / "logo_on.docx"
    before = file_sha256(source)
    out = tmp_path / f"logo_{tracked}.docx"
    result = _dump(server.format_tables(str(source), str(out), CV_SPEC, tracked=None if tracked else False))
    assert result["tracking"] == ("on" if tracked else "off")
    assert result["compared_against"] == "word_baseline"
    assert audit_docx(out).tables[0].vertical_merges == ((2, 3, 1, LABEL),)
    assert word_baseline.section_errors(source, out) == []
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(out) as new:
        media = sorted(hashlib.sha256(old.read(n)).hexdigest() for n in old.namelist() if n.startswith("word/media/"))
        assert media and media == sorted(hashlib.sha256(new.read(n)).hexdigest()
                                         for n in new.namelist() if n.startswith("word/media/"))
    if tracked:
        assert result["tracked_check"]["reject_all_matches_input"]
    assert file_sha256(source) == before


@live
def test_live_merge_on_a_ragged_row(tmp_path):
    """Rows 2-3 start with ``w:gridBefore`` 1: the label at grid column 2 is Word's Cell(row, 1); the merge is
    applied there and the saved audit agrees."""

    from research_mcp import server

    source = build_ragged(tmp_path / "ragged.docx")
    before = file_sha256(source)
    out = tmp_path / "ragged_out.docx"
    spec = [{"index": 1, "descriptive_columns": [1, 2],
             "merges": [{"column": 2, "start_row": 2, "end_row": 3, "label": RAGGED_LABEL}]}]
    result = _dump(server.format_tables(str(source), str(out), spec, tracked=False))
    assert result["tracking"] == "off"
    saved = parse_table(_first_table(out))
    assert saved.vertical_merges == ((2, 3, 2, RAGGED_LABEL),)
    assert saved.cells == (("Kind", "Name", "Amount"), ("", RAGGED_LABEL, "10"), ("", RAGGED_LABEL, "20"))
    assert saved.continuations == ((3, 2, ""),)
    assert file_sha256(source) == before
