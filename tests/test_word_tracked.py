"""Plan 916 S4: apply_document_style and format_tables tracked with Word's own Track Changes.

Unit parts run without Word (relabel, the untracked-change notes and comments, the merged rows' cell changes, the
server wiring with the worker stubbed). Live parts need Word (RESEARCH_MCP_WORD_LIVE=1). Fixtures are synthetic."""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace

import pytest

from research_mcp import word_tracked
from research_mcp.audit import audit_docx, file_sha256
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.revisions import revision_count
from research_mcp.tables import MergeGroup, TableFormatSpec, _without_merge_continuations
from research_mcp.track_setting import NOT_TRACKABLE_TOOLS, write_setting

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
PEOPLE_XML = (b'<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">%s'
              b"</w15:people>")
live = pytest.mark.skipif(os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1", reason="requires Word")


# -- synthetic documents ---------------------------------------------------------------------------------------

def build_letter(path: Path, *, with_table: bool = True) -> Path:
    """1.25" margins, a header, Arial 12, a 4x3 table whose first column repeats "Alpha" in rows 2-4."""

    from docx import Document
    from docx.shared import Inches, Pt

    document = Document()
    section = document.sections[0]
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, side, Inches(1.25))
    section.header.paragraphs[0].text = "Header text"
    document.add_paragraph("A Synthetic Title")
    document.add_paragraph("Introduction")
    paragraph = document.add_paragraph("Body one with ")
    paragraph.add_run("bold").bold = True
    paragraph.add_run(" and plain text.")
    if with_table:
        table = document.add_table(rows=4, cols=3)
        for r, row in enumerate([["Group", "A", "B"], ["Alpha", "1", "2"], ["Alpha", "3", "4"], ["Alpha", "5", "6"]]):
            for c, value in enumerate(row):
                table.cell(r, c).text = value
    document.add_paragraph("Closing paragraph.")
    for para in document.paragraphs:
        for run in para.runs:
            run.font.name = "Arial"
            run.font.size = Pt(12)
    document.save(path)
    return path


def read_part(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read(name)


def rewrite_parts(path: Path, changes: dict[str, bytes | None]) -> None:
    """Replace, add (bytes) or drop (None) parts, keeping every other part's bytes."""

    temp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename in changes:
                if changes[info.filename] is not None:
                    target.writestr(info, changes[info.filename])
            else:
                target.writestr(info, source.read(info.filename))
        names = set(source.namelist())
        for name, data in changes.items():
            if name not in names and data is not None:
                target.writestr(name, data)
    temp.replace(path)


@pytest.fixture
def letter(tmp_path) -> Path:
    return build_letter(tmp_path / "letter.docx")


@pytest.fixture
def letter_on(letter, tmp_path) -> Path:
    on = tmp_path / "letter_on.docx"
    write_setting(letter, on, on=True, author="Claude")
    return on


def no_word(*args, **kwargs):
    raise AssertionError("Word was started")


# -- relabel ---------------------------------------------------------------------------------------------------

def test_relabel_xml_handles_attribute_order_quotes_and_range_markers():
    xml = (
        '<w:body><w:ins w:id="1" w:author="Word User" w:date="2026-11-17T10:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins>'
        "<w:rPrChange w:date='2026-11-17T10:00:00Z' w:author='Word User' w:id='2'/>"
        '<w:pPrChange\n  w:id="3"\n  w:author = "Word User"/>'
        '<w:moveFromRangeStart w:id="4" w:name="m1" w:author="Word User"/>'
        '<w:insideH w:val="single" w:author="Not A Revision"/>'
        '<w:delText w:author="Not A Revision">y</w:delText>'
        '<w:comment w:id="5" w:author="Commenter"/></w:body>'
    )
    new, count = word_tracked._relabel_xml(xml, "Claude")
    assert count == 4
    assert re.findall(r'w:author\s*=\s*"([^"]*)"', new) == ["Claude", "Claude", "Claude", "Claude",
                                                             "Not A Revision", "Not A Revision", "Commenter"]
    assert "w:date='2026-11-17T10:00:00Z'" in new and "w:id='2'" in new  # other attributes untouched
    odd = word_tracked._relabel_xml(xml, 'A "quoted" <author>')[0]
    root = ET.fromstring(odd.replace("<w:body>", f'<w:body xmlns:w="{W_NS}">', 1))
    assert [e.get(W + "author") for e in root.iter() if e.get(W + "author")][:4] == ['A "quoted" <author>'] * 4


def _with_word_revision(source: Path, target: Path, *, people: bytes | None) -> Path:
    shutil.copyfile(source, target)
    document = read_part(target, "word/document.xml").decode("utf-8")
    first = document.index("<w:r>")
    document = (document[:first] + '<w:ins w:id="90" w:author="Word User" w:date="2026-11-17T10:00:00Z">'
                + "<w:r><w:t>new</w:t></w:r></w:ins>" + document[first:])
    changes: dict[str, bytes | None] = {"word/document.xml": document.encode("utf-8")}
    if people is not None:
        rels = read_part(target, "word/_rels/document.xml.rels").decode("utf-8").replace(
            "</Relationships>", '<Relationship Id="rIdPeople" Type="http://schemas.microsoft.com/office/2011/'
            'relationships/people" Target="people.xml"/></Relationships>')
        types = read_part(target, "[Content_Types].xml").decode("utf-8").replace(
            "</Types>", '<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-'
            'officedocument.wordprocessingml.people+xml"/></Types>')
        changes.update({"word/people.xml": people, "word/_rels/document.xml.rels": rels.encode("utf-8"),
                        "[Content_Types].xml": types.encode("utf-8")})
    rewrite_parts(target, changes)
    return target


def test_relabel_sets_author_and_drops_people_part_word_added(letter, tmp_path):
    stage = _with_word_revision(letter, tmp_path / "stage.docx", people=PEOPLE_XML % b"")
    assert word_tracked.revision_authors(stage) == {"Word User"}
    assert word_tracked.relabel(stage, "Claude", letter) == 1
    assert word_tracked.revision_authors(stage) == {"Claude"}
    with zipfile.ZipFile(stage) as archive:
        assert "word/people.xml" not in archive.namelist()
        assert b"people.xml" not in archive.read("word/_rels/document.xml.rels")
        assert b"people.xml" not in archive.read("[Content_Types].xml")
    assert audit_docx(stage).structural_errors in ((), [])


def test_relabel_restores_source_people_part(letter, tmp_path):
    source = _with_word_revision(letter, tmp_path / "source_people.docx", people=PEOPLE_XML % b"<w15:person w15:author=\"source\"/>")
    # the source must be revision-free: take the people part alone into a clean copy
    clean = tmp_path / "clean.docx"
    shutil.copyfile(source, clean)
    rewrite_parts(clean, {"word/document.xml": read_part(letter, "word/document.xml")})
    stage = _with_word_revision(clean, tmp_path / "stage.docx", people=None)
    rewrite_parts(stage, {"word/people.xml": PEOPLE_XML % b"<w15:person w15:author=\"Word User\"/>"})
    word_tracked.relabel(stage, "Claude", clean)
    assert read_part(stage, "word/people.xml") == PEOPLE_XML % b"<w15:person w15:author=\"source\"/>"


def test_relabel_refuses_a_source_with_revisions(tmp_path):
    revised = create_tracked_fixture(tmp_path / "revised.docx")
    with pytest.raises(RefusedUnsafe, match="revision-free source"):
        word_tracked.relabel(revised, "Claude", revised)


# -- notes for what Word can't track -----------------------------------------------------------------------------

def _pg_mar(**sides: int) -> ET.Element:
    element = ET.Element(W + "pgMar")
    for side, twips in sides.items():
        element.set(W + side, str(twips))
    return element


def test_margin_note_names_old_and_new():
    old = _pg_mar(top=1800, bottom=1800, left=1800, right=1800)
    new = _pg_mar(top=1440, bottom=1440, left=1440, right=1440)
    note = word_tracked.margin_note(1, old, new)
    assert note.text == 'Page margins changed from 1.25" to 1" (not tracked: Word can\'t track page setup)'
    assert note.covers(("sect", 1, "pgMar")) and not note.covers(("sect", 2, "pgMar"))
    mixed = word_tracked.margin_note(2, old, _pg_mar(top=1800, bottom=1800, left=1440, right=1800))
    assert mixed.text == 'Page margins changed (left 1.25" to 1") (not tracked: Word can\'t track page setup)'
    assert word_tracked.margin_note(1, old, old) is None


def test_merge_note_text_and_excluded_cells(letter):
    notes = word_tracked.untracked_notes(letter, letter, [(1, MergeGroup(column=1, start_row=2, end_row=4,
                                                                         label="Alpha"))])
    assert [note.text for note in notes] == [
        "Cells in rows 2-4 of column 1 of table 1 were merged (not tracked: Word can't track cell merges)"]
    note = notes[0]
    assert note.cell == (1, 2, 1) and note.rows == (2, 4)
    # only the merge is left out of the Reject check: the continuation cells' paragraphs...
    assert note.covers(("t", (1, 3, 1), "p", 1)) and note.covers(("t", (1, 4, 1), "p", 1))
    # ...not the first merged cell's paragraphs, and no cell's properties (those drop only w:vMerge)
    assert not note.covers(("t", (1, 2, 1), "p", 1))
    assert not any(note.covers(("t", (1, r, 1), "cell")) for r in (2, 3, 4))
    assert not note.covers(("t", (1, 3, 2), "p", 1))
    vmerge = f"<w:tcW {W}w='1'></w:tcW><w:vMerge {W}val='restart'></w:vMerge>"
    assert note.normalize(("t", (1, 2, 1), "cell"), vmerge) == f"<w:tcW {W}w='1'></w:tcW>"
    assert note.normalize(("t", (1, 2, 2), "cell"), vmerge) == vmerge  # another column keeps it


def test_untracked_notes_skip_margins_word_tracked(letter, tmp_path):
    """Margins changed with a w:sectPrChange are tracked: no note. Without one: one note."""

    changed = tmp_path / "changed.docx"
    shutil.copyfile(letter, changed)
    xml = read_part(changed, "word/document.xml").decode("utf-8")
    plain = re.sub(r'w:(top|bottom|left|right)="1800"', r'w:\1="1440"', xml)
    rewrite_parts(changed, {"word/document.xml": plain.encode("utf-8")})
    assert [n.kind for n in word_tracked.untracked_notes(letter, changed)] == ["margins"]
    tracked = plain.replace("</w:sectPr>", '<w:sectPrChange w:id="7" w:author="Claude" w:date="2026-11-17T10:00:00Z">'
                            "<w:sectPr/></w:sectPrChange></w:sectPr>")
    rewrite_parts(changed, {"word/document.xml": tracked.encode("utf-8")})
    assert word_tracked.untracked_notes(letter, changed) == []


def test_untracked_comment_lands_on_the_first_merged_cell(letter, tmp_path):
    out = tmp_path / "commented.docx"
    shutil.copyfile(letter, out)
    notes = word_tracked.untracked_notes(letter, letter, [(1, MergeGroup(1, 2, 4, "Alpha"))])
    added = word_tracked.add_untracked_comments(out, notes, "Claude")
    assert len(added) == 1 and added[0]["kind"] == "merge" and re.fullmatch(r"[0-9A-F]{8}", added[0]["comment_id"])
    comments = ET.fromstring(read_part(out, "word/comments.xml"))
    entries = comments.findall(W + "comment")
    assert len(entries) == 1 and entries[0].get(W + "author") == "Claude"
    assert "".join(t.text or "" for t in entries[0].iter(W + "t")) == notes[0].text
    body = ET.fromstring(read_part(out, "word/document.xml")).find(W + "body")
    row2 = body.find(W + "tbl").findall(W + "tr")[1]
    first_cell = row2.findall(W + "tc")[0]
    assert first_cell.find(f".//{W}commentRangeStart") is not None  # on row 2, column 1
    assert revision_count(out) == 0


def test_formatting_differences_find_a_run_change_and_respect_exclusions(letter, tmp_path):
    changed = tmp_path / "changed.docx"
    shutil.copyfile(letter, changed)
    xml = read_part(changed, "word/document.xml").decode("utf-8")
    xml = xml.replace('<w:sz w:val="24"/>', '<w:sz w:val="22"/>', 1)
    rewrite_parts(changed, {"word/document.xml": xml.encode("utf-8")})
    before, after = word_tracked.formatting_snapshot(letter), word_tracked.formatting_snapshot(changed)
    diff = word_tracked.differences(before, after, lambda key: False)
    assert diff == [("p", 1)]
    assert "run properties" in word_tracked.describe_difference(diff[0], before, after)
    assert "Synthetic" not in word_tracked.describe_difference(diff[0], before, after)  # no document text
    assert word_tracked.differences(before, after, lambda key: key == ("p", 1)) == []


# -- the merged rows' cell changes Word drops ---------------------------------------------------------------------

def _merged_stage(letter: Path, target: Path) -> Path:
    """Like Word's output: rows 2-4 bordered, column 1 merged (vMerge), one cell with its tcPrChange kept."""

    shutil.copyfile(letter, target)
    xml = read_part(target, "word/document.xml").decode("utf-8")
    rows = re.findall(r"<w:tr\b.*?</w:tr>", xml, flags=re.S)
    for r, row in enumerate(rows[1:], start=2):
        cells = re.findall(r"<w:tc>.*?</w:tc>", row, flags=re.S)
        new_row = row
        for c, cell in enumerate(cells, start=1):
            merge = ('<w:vMerge w:val="restart"/>' if r == 2 else "<w:vMerge/>") if c == 1 else ""
            kept = ('<w:tcPrChange w:id="50" w:author="Claude" w:date="2026-11-17T10:00:00Z"><w:tcPr>'
                    '<w:tcW w:w="2880" w:type="dxa"/></w:tcPr></w:tcPrChange>') if (r, c) == (3, 3) else ""
            new_cell = re.sub(r"(<w:tcW [^>]*/>)", r"\1" + merge + '<w:tcBorders><w:top w:val="single" '
                              'w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>' + kept, cell, count=1)
            new_row = new_row.replace(cell, new_cell, 1)
        xml = xml.replace(row, new_row, 1)
    rewrite_parts(target, {"word/document.xml": xml.encode("utf-8")})
    return target


def test_restore_merge_row_changes_reinserts_tcprchange_keeping_the_merge(letter, tmp_path):
    stage = _merged_stage(letter, tmp_path / "stage.docx")
    before = read_part(stage, "word/document.xml")
    count = word_tracked.restore_merge_row_changes(stage, letter, [(1, MergeGroup(1, 2, 4, "Alpha"))], "Claude")
    assert count == 8  # 3 rows x 3 cells, less the one that kept its change
    after = read_part(stage, "word/document.xml")
    inserted = re.findall(rb"<w:tcPrChange w:id=\"(\d+)\" w:author=\"Claude\".*?</w:tcPrChange>", after, flags=re.S)
    assert len(inserted) == 9 and len(set(inserted)) == 9  # 8 new + the kept one, ids unique
    # every byte but the insertions is kept
    stripped = re.sub(rb"<w:tcPrChange w:id=\"(?!50\")\d+\"[^>]*>.*?</w:tcPrChange>", b"", after, flags=re.S)
    assert stripped == before
    rows = ET.fromstring(after).find(W + "body").find(W + "tbl").findall(W + "tr")
    for r, row in enumerate(rows[1:], start=2):
        for c, cell in enumerate(row.findall(W + "tc"), start=1):
            old = cell.find(f"{W}tcPr/{W}tcPrChange/{W}tcPr")
            assert old.find(W + "tcBorders") is None  # Reject removes the borders...
            assert (old.find(W + "vMerge") is not None) == (c == 1 and (r, c) != (3, 3))  # ...and keeps the merge
    assert word_tracked.restore_merge_row_changes(stage, letter, [(1, MergeGroup(1, 2, 4, "Alpha"))], "Claude") == 0


def test_live_check_masks_merge_continuation_cells():
    table = SimpleNamespace(columns=3, cell_hashes=tuple(f"h{n}" for n in range(12)))
    spec = TableFormatSpec(index=1, descriptive_columns=(1,), merges=(MergeGroup(1, 2, 4, "Alpha"),))
    masked = _without_merge_continuations(table, spec)
    assert [n for n, value in enumerate(masked) if value is None] == [6, 9]  # rows 3-4 of column 1
    assert masked[3] == "h3"  # the first merged cell keeps its hash


# -- the merged cell keeps its character formatting (review finding 1) --------------------------------------------

def _split_label(letter: Path, target: Path) -> Path:
    """The letter with "Alpha" in rows 2-4 of column 1 as a bold "Al" run and a plain "pha" run."""

    from docx import Document

    document = Document(letter)
    for row in (1, 2, 3):
        paragraph = document.tables[0].cell(row, 0).paragraphs[0]
        paragraph.clear()
        paragraph.add_run("Al").bold = True
        paragraph.add_run("pha")
    document.save(target)
    return target


def _merged_like_word(source: Path, target: Path, *, retype: bool) -> Path:
    """What a merge leaves: the cells below the first emptied; ``retype`` also rewrites the first cell's text as
    one run with its first run's formatting (``Range.Text = label``, the untracked pass)."""

    from docx import Document

    document = Document(source)
    for row in (2, 3):
        document.tables[0].cell(row, 0).paragraphs[0].clear()
    if retype:
        paragraph = document.tables[0].cell(1, 0).paragraphs[0]
        paragraph.clear()
        paragraph.add_run("Alpha").bold = True
    document.save(target)
    return target


MERGE_ALPHA = [(1, MergeGroup(1, 2, 4, "Alpha"))]


def test_merged_cell_errors_catch_retyped_formatting(letter, tmp_path):
    source = _split_label(letter, tmp_path / "split.docx")
    notes = word_tracked.untracked_notes(source, source, MERGE_ALPHA)
    expected = word_tracked.formatting_snapshot(source, "rejected")
    kept = word_tracked.formatting_snapshot(_merged_like_word(source, tmp_path / "kept.docx", retype=False),
                                            "rejected")
    assert word_tracked.merged_cell_errors(expected, kept, notes) == []
    retyped = word_tracked.formatting_snapshot(_merged_like_word(source, tmp_path / "retyped.docx", retype=True),
                                               "rejected")
    errors = word_tracked.merged_cell_errors(expected, retyped, notes)
    assert errors == ["the first merged cell (table 1, row 2, column 1) paragraph 1: the character formatting of 3 "
                      "of 5 characters differs (from character 3)"]
    assert not any("Alpha" in error or "pha" in error for error in errors)  # no document text
    # the old whole-cell exclusion let exactly this through the Reject check; now it is a difference there too
    assert word_tracked.differences(expected, retyped, lambda key: any(n.covers(key) for n in notes)) == [
        ("t", (1, 2, 1), "p", 1)]
    # cells below the first that still hold text
    assert word_tracked.merged_cell_errors(expected, expected, notes) == [
        "the merged cell at table 1, row 3, column 1 still holds text",
        "the merged cell at table 1, row 4, column 1 still holds text"]


def test_rejected_view_reads_old_run_and_paragraph_properties(letter, tmp_path):
    source = _split_label(letter, tmp_path / "split.docx")
    tracked = tmp_path / "tracked.docx"
    shutil.copyfile(source, tracked)
    xml = read_part(tracked, "word/document.xml").decode("utf-8")
    assert xml.count("<w:r><w:t>pha</w:t></w:r>") == 3
    date = 'w:author="Claude" w:date="2026-11-17T10:00:00Z"'
    xml = xml.replace("<w:r><w:t>pha</w:t></w:r>", f'<w:r><w:rPr><w:b/><w:rPrChange w:id="61" {date}><w:rPr/>'
                      "</w:rPrChange></w:rPr><w:t>pha</w:t></w:r>", 1)
    xml = xml.replace("<w:t>Closing paragraph.</w:t></w:r>", "<w:t>Closing paragraph.</w:t></w:r>"
                      f'<w:ins w:id="62" {date}><w:r><w:t> Added.</w:t></w:r></w:ins>'
                      f'<w:del w:id="63" {date}><w:r><w:delText> Gone.</w:delText></w:r></w:del>', 1)
    rewrite_parts(tracked, {"word/document.xml": xml.encode("utf-8")})
    before = word_tracked.formatting_snapshot(source)
    accepted, rejected = word_tracked.formatting_snapshot(tracked), word_tracked.formatting_snapshot(tracked, "rejected")
    cell = ("t", (1, 2, 1), "p", 1)
    assert [rpr != "" for _, rpr in accepted[cell][1]] == [True] * 5  # accepted: all bold
    assert rejected[cell] == before[cell]  # rejected: "Al" bold, "pha" plain, as the input
    closing = max(key for key in before if key[0] == "p")
    assert "".join(ch for ch, _ in accepted[closing][1]) == "Closing paragraph. Added."
    assert "".join(ch for ch, _ in rejected[closing][1]) == "Closing paragraph. Gone."


# -- tables.py merge on a fake Word table: untracked as at 510c797, tracked keeps the runs (finding 2) -----------

BASE_COMMIT = "510c797"


class _FakeTable:
    """Enough of Word's Table for merge_repeated_label / snapshot_live: after a merge the cells below the first
    can no longer be addressed (Word raises), and the first keeps its text unless ``merged_text`` says otherwise."""

    def __init__(self, rows: list[list[str]], log: list, merged_text: str | None = None):
        self.texts = {(r, c): text for r, row in enumerate(rows, 1) for c, text in enumerate(row, 1)}
        self.Rows = SimpleNamespace(Count=len(rows))
        self.Columns = SimpleNamespace(Count=len(rows[0]))
        self.gone: set[tuple[int, int]] = set()
        self.log = log
        self.merged_text = merged_text

    def Cell(self, row: int, column: int):
        if (row, column) in self.gone:
            raise RuntimeError("Cannot access individual cells in this collection")
        table = self

        class Range:
            @property
            def Text(self):
                return table.texts[(row, column)] + "\r\x07"

            @Text.setter
            def Text(self, value):
                table.log.append(("set_text", row, column, value))
                table.texts[(row, column)] = value

        def merge(other):
            table.log.append(("merge", (row, column), (other._row, other._column)))
            for r in range(row + 1, other._row + 1):
                table.gone.add((r, column))
            if table.merged_text is not None:
                table.texts[(row, column)] = table.merged_text

        return SimpleNamespace(Range=Range(), Merge=merge, _row=row, _column=column)


ROWS = [["Group", "A", "B"], ["Alpha", "1", "2"], ["Alpha", "3", "4"], ["Alpha", "5", "6"]]


def _tables_at_base(monkeypatch):
    """tables.py as committed at BASE_COMMIT (before plan 916 S4), loaded as a module of the package."""

    import subprocess
    import sys
    import types

    root = Path(__file__).resolve().parents[1]
    try:
        source = subprocess.run(["git", "show", f"{BASE_COMMIT}:src/research_mcp/tables.py"], cwd=root,
                                capture_output=True, check=True, timeout=60).stdout
    except (OSError, subprocess.SubprocessError):
        pytest.skip(f"git show {BASE_COMMIT} unavailable")
    module = types.ModuleType("research_mcp._tables_at_base")
    module.__package__ = "research_mcp"
    module.__file__ = f"tables.py@{BASE_COMMIT}"
    monkeypatch.setitem(sys.modules, module.__name__, module)
    exec(compile(source, module.__file__, "exec"), module.__dict__)
    return module


def test_untracked_merge_calls_word_as_before(monkeypatch):
    from research_mcp import tables

    base = _tables_at_base(monkeypatch)
    now_log, base_log = [], []
    tables.merge_repeated_label(_FakeTable(ROWS, now_log), MergeGroup(1, 2, 4, "Alpha"))
    base.merge_repeated_label(_FakeTable(ROWS, base_log), base.MergeGroup(1, 2, 4, "Alpha"))
    assert now_log == base_log == [("set_text", 3, 1, ""), ("set_text", 4, 1, ""), ("merge", (2, 1), (4, 1)),
                                   ("set_text", 2, 1, "Alpha")]


def test_tracked_merge_keeps_the_first_cell_and_checks_its_text():
    from research_mcp import tables
    from research_mcp.errors import VerificationFailed

    log: list = []
    tables.merge_repeated_label(_FakeTable(ROWS, log), MergeGroup(1, 2, 4, "Alpha"), keep_formatting=True)
    assert log == [("set_text", 3, 1, ""), ("set_text", 4, 1, ""), ("merge", (2, 1), (4, 1))]  # no retyping
    with pytest.raises(VerificationFailed, match="row 2, column 1 does not read as its label"):
        tables.merge_repeated_label(_FakeTable(ROWS, [], merged_text="Alpha\rAlpha"), MergeGroup(1, 2, 4, "Alpha"),
                                    keep_formatting=True)


def _fake_word(monkeypatch, module, log: list) -> SimpleNamespace:
    """Patch ``module``'s Word session with a fake document holding one _FakeTable; borders are a no-op."""

    table = _FakeTable(ROWS, log)
    document = SimpleNamespace(TrackRevisions=False, Tables=lambda index: table)
    document.Tables = type("Tables", (), {"Count": 1, "__call__": lambda self, index: table})()

    class Session:
        def __init__(self, path, visible=False):
            self.document = document
            self.path = Path(path)

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def save(self):
            log.append(("save",))

        def save_as(self, target):
            """As SaveAs2: writes the open document (unchanged here) to ``target`` and retargets the session."""
            target = Path(target)
            if target != self.path:
                shutil.copyfile(self.path, target)
            log.append(("save_as", target.name))
            self.path = target
            return target

    monkeypatch.setattr(module, "IsolatedWordDocument", Session)
    monkeypatch.setattr(module, "apply_journal_neutral",
                        lambda document, spec, layout=None: log.append(("borders", spec.index)))
    return document


def test_untracked_live_check_refused_the_merge_at_base(letter, tmp_path, monkeypatch):
    """At 510c797 the live check refused this merge: Word can't address the cells below the first, so they hash
    empty after it."""

    from research_mcp.errors import VerificationFailed

    base = _tables_at_base(monkeypatch)
    log: list = []
    _fake_word(monkeypatch, base, log)
    spec = base.TableFormatSpec(index=1, descriptive_columns=(1,), merges=(base.MergeGroup(1, 2, 4, "Alpha"),))
    out = tmp_path / "base.docx"
    with pytest.raises(VerificationFailed, match="table 1 logical data changed") as caught:
        base.format_to_version(letter, out, [spec])
    assert ("save",) not in log and not out.exists()
    empty = hashlib.sha256(b"").hexdigest()
    assert caught.value.actual[2][6] == caught.value.actual[2][9] == empty


def test_untracked_live_check_masks_only_continuations(letter, tmp_path, monkeypatch):
    """Plan 973 S1 (decision 6): tracked=False masks the merge's continuation cells in the live check too, so the
    same merge goes on to the saved checks (``verify_saved_tables``, faked here; what they read of a saved
    continuation is tested in ``test_tables_grid``)."""

    from research_mcp import tables

    log: list = []
    _fake_word(monkeypatch, tables, log)
    seen: list = []
    monkeypatch.setattr(tables, "verify_saved_tables",
                        lambda before, path, specs, reference: seen.append((before, Path(path))) or audit_docx(path))
    spec = TableFormatSpec(index=1, descriptive_columns=(1,), merges=(MergeGroup(1, 2, 4, "Alpha"),))
    out = tmp_path / "now.docx"
    result = tables.format_to_version(letter, out, [spec], tracked=False)
    assert log == [("save_as", "baseline.docx"), ("borders", 1), ("set_text", 3, 1, ""), ("set_text", 4, 1, ""),
                   ("merge", (2, 1), (4, 1)), ("set_text", 2, 1, "Alpha"), ("save_as", out.name)]
    assert out.exists() and [path for _, path in seen] == [out]
    assert result.compared_against == "word_baseline"


def test_tracked_live_check_masks_only_continuations(letter, tmp_path, monkeypatch):
    from research_mcp import tables

    log: list = []
    document = _fake_word(monkeypatch, tables, log)
    seen: list = []
    monkeypatch.setattr(tables, "verify_saved_tables",
                        lambda before, path, specs, reference: seen.append(Path(path)) or audit_docx(path))
    spec = TableFormatSpec(index=1, descriptive_columns=(1,), merges=(MergeGroup(1, 2, 4, "Alpha"),))
    out = tmp_path / "tracked.docx"
    result = tables.format_to_version(letter, out, [spec], tracked=True)
    assert log == [("save_as", "baseline.docx"), ("borders", 1), ("set_text", 3, 1, ""), ("set_text", 4, 1, ""),
                   ("merge", (2, 1), (4, 1)), ("save_as", out.name)]
    assert document.TrackRevisions is False and seen == [out] and result.compared_against == "word_baseline"


# -- effective comparison (plan 916 S6: a real table's explicit zeros) ---------------------------------------------

STYLES = (f'<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Spaced"><w:name w:val="Spaced"/>'
          '<w:pPr><w:spacing w:after="160"/><w:ind w:hanging="360"/><w:keepNext/></w:pPr><w:rPr><w:b/></w:rPr>'
          '</w:style></w:styles>')


def _el(xml: str) -> ET.Element:
    return ET.fromstring(xml.replace("<w:pPr>", f'<w:pPr xmlns:w="{W_NS}">', 1)
                         .replace("<w:rPr>", f'<w:rPr xmlns:w="{W_NS}">', 1))


def _view(styles: str = STYLES):
    return word_tracked._Inherited({"word/styles.xml": styles.encode("utf-8")}).view()


def _ppr(xml: str) -> str:
    return word_tracked.canonical(_view().ppr(_el(xml)))


def _rpr(ppr_xml: str | None, xml: str, styles: str = STYLES) -> str:
    view = _view(styles)
    return word_tracked.canonical(view.rpr(_el(ppr_xml) if ppr_xml else None, _el(xml)))


def test_effective_ppr_drops_only_values_the_paragraph_inherits_anyway():
    zeros = ('<w:pPr><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:spacing w:before="0" w:after="0" '
             'w:line="240" w:lineRule="auto" w:beforeAutospacing="0" w:afterAutospacing="0"/>'
             '<w:keepNext w:val="0"/></w:pPr>')
    assert _ppr(zeros) == _ppr('<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>')
    # the same zeros under a style that sets spacing after, a hanging indent and keepNext override it: kept
    styled = zeros.replace("<w:pPr>", '<w:pPr><w:pStyle w:val="Spaced"/>')
    kept = _ppr(styled)
    assert "}after='0'" in kept and "}firstLine='0'" in kept and "<w:keepNext" in kept
    assert "}left='0'" not in kept and "}before='0'" not in kept
    # a value that differs from the inherited one is a difference, and on means on however it is spelled
    assert _ppr('<w:pPr><w:spacing w:after="120"/></w:pPr>') != _ppr("<w:pPr></w:pPr>")
    assert _ppr('<w:pPr><w:keepNext w:val="1"/></w:pPr>') == _ppr("<w:pPr><w:keepNext/></w:pPr>")
    assert _ppr('<w:pPr><w:keepNext/></w:pPr>') != _ppr("<w:pPr></w:pPr>")


def test_effective_rpr_reads_toggles_and_inherited_fonts():
    assert _rpr(None, '<w:rPr><w:b w:val="1"/></w:rPr>') == _rpr(None, "<w:rPr><w:b/></w:rPr>")
    assert _rpr(None, '<w:rPr><w:b w:val="0"/></w:rPr>') == _rpr(None, "<w:rPr></w:rPr>")
    spaced = '<w:pPr><w:pStyle w:val="Spaced"/></w:pPr>'
    assert _rpr(spaced, '<w:rPr><w:b w:val="0"/></w:rPr>') != _rpr(spaced, "<w:rPr></w:rPr>")  # the style is bold
    default_font = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/></w:rPr>'
    assert _rpr(None, default_font) == _rpr(None, '<w:rPr><w:sz w:val="19"/></w:rPr>')
    other_font = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'
    assert _rpr(None, other_font) != _rpr(None, "<w:rPr></w:rPr>")
    themed = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:asciiTheme="minorHAnsi"/></w:rPr>'
    assert "Times New Roman" in _rpr(None, themed)
    # a theme color caches its RGB in w:val: a direct RGB equal to that cache is another color (review r3)
    accent = STYLES.replace("<w:rPrDefault><w:rPr>",
                            '<w:rPrDefault><w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/>')
    none = "<w:rPr></w:rPr>"
    assert _rpr(None, '<w:rPr><w:color w:val="FF0000"/></w:rPr>', accent) != _rpr(None, none, accent)
    same = '<w:rPr><w:color w:val="ff0000" w:themeColor="accent1"/></w:rPr>'
    assert _rpr(None, same, accent) == _rpr(None, none, accent)
    plain = STYLES.replace("<w:rPrDefault><w:rPr>", '<w:rPrDefault><w:rPr><w:color w:val="000000"/>')
    assert _rpr(None, '<w:rPr><w:color w:val="000000"/></w:rPr>', plain) == _rpr(None, none, plain)


def test_effective_snapshot_matches_a_reject_all_that_drops_inherited_zeros(tmp_path):
    """Word's Reject All gave back a caption and cells without their explicit zeros (live, plan 916 S6): the
    effective snapshot finds no difference there and still finds a real one."""

    cell = ('<w:p><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>x</w:t></w:r></w:p>')
    zeros = '<w:ind w:left="0" w:firstLine="0"/><w:spacing w:before="0" w:after="0"/>'
    font = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b w:val="0"/>'

    def body(ppr: str, rpr: str) -> str:
        p = cell.format(ppr=ppr, rpr=rpr)
        return p + f"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>{p}</w:tc></w:tr></w:tbl><w:p/>"

    paths = {}
    for name, ppr, rpr in (("input", zeros, font), ("rejected", "", ""), ("real", '<w:spacing w:after="120"/>', "")):
        paths[name] = tmp_path / f"{name}.docx"
        build_letter(paths[name], with_table=False)
        rewrite_parts(paths[name], {"word/document.xml": _doc(body(ppr, rpr)), "word/styles.xml": STYLES.encode()})
    same = lambda key: False  # noqa: E731
    as_written = word_tracked.differences(word_tracked.formatting_snapshot(paths["input"]),
                                          word_tracked.formatting_snapshot(paths["rejected"]), same)
    assert as_written  # the strict comparison refused the real table
    snap = {name: word_tracked.formatting_snapshot(path, effective=True) for name, path in paths.items()}
    assert word_tracked.differences(snap["input"], snap["rejected"], same) == []
    assert word_tracked.differences(snap["input"], snap["real"], same) == [("p", 1), ("t", (1, 1, 1), "p", 1)]


def test_effective_snapshot_compares_cells_of_a_conditional_table_style_as_written(tmp_path):
    styles = STYLES.replace("</w:styles>", '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/>'
                            '<w:tblStylePr w:type="firstRow"><w:pPr><w:spacing w:after="200"/></w:pPr>'
                            "</w:tblStylePr></w:style></w:styles>")
    paths = {}
    for name, ppr in (("input", '<w:spacing w:after="0"/>'), ("rejected", "")):
        paths[name] = tmp_path / f"{name}.docx"
        build_letter(paths[name], with_table=False)
        p = f"<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>"
        rewrite_parts(paths[name], {"word/document.xml": _doc(
            f'<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid>'
            f"<w:tr><w:tc>{p}</w:tc></w:tr></w:tbl><w:p/>"), "word/styles.xml": styles.encode()})
    snap = {name: word_tracked.formatting_snapshot(path, effective=True) for name, path in paths.items()}
    assert word_tracked.differences(snap["input"], snap["rejected"], lambda key: False) == [("t", (1, 1, 1), "p", 1)]


# -- byte spans and the paragraph-change repair (review optional item) ---------------------------------------------

def _doc(body: str) -> bytes:
    return f'<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>'.encode("utf-8")


def test_start_tag_end_detects_empty_tags_and_skips_quoted_gt():
    data = b'<w:p w:x="a>b" w:y=\'/>\'/><w:pPr>'
    end, empty = word_tracked._start_tag_end(data, 0)
    assert data[:end] == b'<w:p w:x="a>b" w:y=\'/>\'/>' and empty
    assert word_tracked._start_tag_end(data, end) == (len(data), False)


def test_paragraph_spans_self_closing_nested_and_missing():
    data = _doc('<w:p><w:pPr/><w:r><w:t>x</w:t></w:r></w:p>'
                '<w:p><w:pPr><w:jc w:val="left"/><w:pPrChange w:id="1" w:author="A"><w:pPr/></w:pPrChange></w:pPr></w:p>'
                '<w:p w14:paraId="1" xmlns:w14="urn:x"/>'
                '<w:p><w:r><w:t>y</w:t></w:r></w:p>')
    spans = word_tracked._paragraph_spans(data)
    assert len(spans) == 4
    first, nested, empty_p, bare = spans
    assert data[first["pPr"][0]:first["pPr"][1]] == b"<w:pPr/>" and first["pPr"][2]
    pr = data[nested["pPr"][0]:nested["pPr"][1]]
    assert pr.startswith(b"<w:pPr><w:jc") and pr.endswith(b"</w:pPrChange></w:pPr>") and not nested["pPr"][2]
    assert data[nested["old"][0]:nested["old"][1]] == b"<w:pPr/>" and nested["old"][2]
    assert empty_p["open"][2] and empty_p["pPr"] is None and data[empty_p["open"][0]:empty_p["open"][1]].endswith(b"/>")
    assert bare["pPr"] is None and not bare["open"][2] and data[bare["open_end"]:].startswith(b"<w:r>")


def test_cell_spans_self_closing_tcpr():
    data = _doc('<w:tbl><w:tr><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w="1"/></w:tcPr><w:p/></w:tc>'
                '<w:tc><w:p/></w:tc></w:tr></w:tbl>')
    (table,) = word_tracked._cell_spans(data)
    first, second, third = table[0]
    assert data[first["tcPr"][0]:first["tcPr"][1]] == b"<w:tcPr/>" and first["tcPr"][2]
    assert data[second["tcPr"][0]:second["tcPr"][1]] == b'<w:tcPr><w:tcW w:w="1"/></w:tcPr>'
    assert third["tcPr"] is None and data[third["tc_open_end"]:].startswith(b"<w:p/>")
    with pytest.raises(VerificationFailed, match="empty table cell"):
        word_tracked._cell_spans(_doc("<w:tbl><w:tr><w:tc/></w:tr></w:tbl>"))


SPACING = '<w:spacing w:before="0" w:after="0" w:beforeAutospacing="1" w:afterAutospacing="1"/>'
CHANGE = '<w:pPrChange w:id="{id}" w:author="Claude" w:date="2026-11-17T10:00:00Z">{old}</w:pPrChange>'


def _repair(tmp_path, source_body: str, stage_body: str) -> tuple[int, str]:
    source, stage = tmp_path / "source.docx", tmp_path / "stage.docx"
    build_letter(source, with_table=False)
    shutil.copyfile(source, stage)
    rewrite_parts(source, {"word/document.xml": _doc(source_body)})
    rewrite_parts(stage, {"word/document.xml": _doc(stage_body)})
    count = word_tracked.restore_paragraph_changes(stage, source, "Claude")
    ET.fromstring(read_part(stage, "word/document.xml"))  # still well-formed
    return count, read_part(stage, "word/document.xml").decode("utf-8")


def test_restore_paragraph_changes_cases(tmp_path):
    x = "<w:r><w:t>x</w:t></w:r>"

    def jc(value: str) -> str:
        return f'<w:pPr><w:jc w:val="{value}"/></w:pPr>'

    untouched = [f'<w:p><w:pPr><w:jc w:val="both"/>{CHANGE.format(id=6, old=jc("right"))}</w:pPr>{x}</w:p>',
                 f"<w:p><w:pPr/>{x}</w:p>"]
    pairs = [  # (source paragraph, Word's paragraph)
        # 1: Word's pPrChange dropped the autospacing flags: its old pPr is set to the source's
        (f"<w:p><w:pPr>{SPACING}</w:pPr>{x}</w:p>",
         '<w:p><w:pPr><w:spacing w:before="0" w:after="240"/>'
         + CHANGE.format(id=5, old='<w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>') + f"</w:pPr>{x}</w:p>"),
        # 2: self-closing <w:pPr/>, changed with no pPrChange
        (f"<w:p>{jc('left')}{x}</w:p>", f"<w:p><w:pPr/>{x}</w:p>"),
        # 3: no w:pPr at all
        (f"<w:p>{jc('right')}{x}</w:p>", f"<w:p>{x}</w:p>"),
        # 4: an empty <w:p/>
        (f"<w:p>{jc('center')}</w:p>", '<w:p w:rsidR="00AB"/>'),
        # 5: a paragraph in a table cell, changed with no pPrChange
        (f"<w:tbl><w:tr><w:tc><w:p>{jc('left')}{x}</w:p></w:tc></w:tr></w:tbl>",
         f"<w:tbl><w:tr><w:tc><w:p>{jc('center')}{x}</w:p></w:tc></w:tr></w:tbl>"),
        # untouched: a correct pPrChange; properties unchanged (no pPr -> empty pPr)
        (f'<w:p>{jc("right")}{x}</w:p>', untouched[0]),
        (f"<w:p>{x}</w:p>", untouched[1]),
    ]
    count, out = _repair(tmp_path, "".join(s for s, _ in pairs), "".join(w for _, w in pairs))
    assert count == 5
    for piece in untouched:
        assert piece in out  # byte-identical
    assert out.count(x) == 6
    source_ps = list(ET.fromstring(_doc("".join(s for s, _ in pairs))).iter(W + "p"))
    stage_ps = list(ET.fromstring(_doc("".join(w for _, w in pairs))).iter(W + "p"))
    out_ps = list(ET.fromstring(out).iter(W + "p"))
    assert len(out_ps) == len(source_ps) == 7
    canon = word_tracked.canonical
    for old, word, new in zip(source_ps, stage_ps, out_ps):
        new_pr = new.find(W + "pPr")
        assert canon(new_pr) == canon(word.find(W + "pPr"))  # accepted view as Word left it
        change = new_pr.find(W + "pPrChange") if new_pr is not None else None
        rejected = change.find(W + "pPr") if change is not None else new_pr
        assert canon(rejected) == canon(old.find(W + "pPr"))  # Reject gives the source's properties
    added = re.findall(r'<w:pPrChange w:id="(\d+)" w:author="Claude" w:date="2026-11-17T10:00:00Z">', out)
    assert sorted(map(int, added)) == [5, 6, 7, 8, 9, 10]  # four new ids after the highest, unique


def test_restore_paragraph_changes_refuses_a_count_mismatch(tmp_path):
    x = "<w:r><w:t>x</w:t></w:r>"
    stage_body = f'<w:p><w:pPr><w:jc w:val="left"/></w:pPr>{x}</w:p><w:p>{x}</w:p><w:p>{x}</w:p>'
    with pytest.raises(VerificationFailed, match="paragraph count differs"):
        _repair(tmp_path, f"<w:p>{x}</w:p><w:p>{x}</w:p>", stage_body)
    assert read_part(tmp_path / "stage.docx", "word/document.xml") == _doc(stage_body)  # nothing written


# -- server wiring (worker stubbed; no Word) ---------------------------------------------------------------------

def test_style_tools_not_refused_while_on():
    assert "apply_document_style" not in NOT_TRACKABLE_TOOLS
    assert "format_tables" not in NOT_TRACKABLE_TOOLS
    assert "clean_bookmarks" in NOT_TRACKABLE_TOOLS


@pytest.mark.parametrize("tool", ["apply_document_style", "format_tables"])
def test_refuses_other_authors_revisions_with_hint(tool, tmp_path, monkeypatch):
    from research_mcp import server

    monkeypatch.setattr("research_mcp.operation_process.run_operation", no_word)
    revised = tmp_path / "revised_on.docx"
    write_setting(create_tracked_fixture(tmp_path / "revised.docx"), revised, on=True, author="Claude")
    sha = file_sha256(revised)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe) as caught:
        if tool == "apply_document_style":
            server.apply_document_style(str(revised), str(out), sha)
        else:
            server.format_tables(str(revised), str(out), [{"index": 1, "descriptive_columns": [1], "merges": []}])
    message = str(caught.value)
    assert f"{tool}: source contains" in message and "style_text" in message and "format_paragraph" in message
    assert not out.exists() and file_sha256(revised) == sha


def test_preview_unchanged_while_on(letter, letter_on, tmp_path, monkeypatch):
    from research_mcp import server

    monkeypatch.setattr("research_mcp.operation_process.run_operation", no_word)
    sha = file_sha256(letter_on)
    out = tmp_path / "never.docx"
    on_plan = server.apply_document_style(str(letter_on), str(out), sha, profile="compact_single", preview_only=True)
    off_plan = server.apply_document_style(str(letter), str(out), file_sha256(letter), profile="compact_single",
                                           preview_only=True)

    def strip(plan: dict) -> dict:
        return {k: v for k, v in plan.items() if k not in ("source", "source_sha256", "path", "sha256")}

    assert strip(on_plan) == strip(off_plan)
    assert not out.exists() and file_sha256(letter_on) == sha


def _stub_worker(monkeypatch, calls: list):
    def run(kind, payload, **kwargs):
        calls.append((kind, dict(payload)))
        if kind == "style":
            raise RuntimeError("stub: worker reached")
        return {}

    monkeypatch.setattr("research_mcp.operation_process.run_operation", run)


def test_format_tables_tracked_false_skips_the_tracked_path(letter_on, tmp_path, monkeypatch):
    from research_mcp import server

    calls: list = []
    _stub_worker(monkeypatch, calls)
    monkeypatch.setattr(word_tracked, "finish", lambda *a, **k: pytest.fail("tracked path taken"))
    result = server.format_tables(str(letter_on), str(tmp_path / "plain.docx"),
                                  [{"index": 1, "descriptive_columns": [1], "merges": []}], tracked=False)
    result = result.model_dump() if hasattr(result, "model_dump") else result
    assert calls[0][0] == "tables" and calls[0][1]["tracked"] is False
    assert result["tracking"] == "off" and result["untracked"] == [] and result["revisions_added"] == 0


def test_format_tables_switch_on_takes_the_tracked_path(letter_on, tmp_path, monkeypatch):
    from research_mcp import server

    calls: list = []
    seen: dict = {}
    _stub_worker(monkeypatch, calls)

    def finish(stage, input_copy, *, author, verify_accept, merges=()):
        seen.update(author=author, merges=list(merges), input_exists=Path(input_copy).exists())
        return {"tracking": "on", "tracking_author": author, "revisions_added": 3, "untracked": [],
                "tracked_check": {"stub": True}}

    monkeypatch.setattr(word_tracked, "finish", finish)
    merge = {"column": 1, "start_row": 2, "end_row": 4, "label": "Alpha"}
    result = server.format_tables(str(letter_on), str(tmp_path / "tracked.docx"),
                                  [{"index": 1, "descriptive_columns": [1], "merges": [merge]}])
    result = result.model_dump() if hasattr(result, "model_dump") else result
    assert calls[0][1]["tracked"] is True
    assert seen == {"author": "Claude", "merges": [(1, MergeGroup(1, 2, 4, "Alpha"))], "input_exists": True}
    assert result["tracking"] == "on" and result["revisions_added"] == 3


@pytest.mark.parametrize("tracked, expected", [(None, True), (False, False)])
def test_apply_document_style_passes_tracking_to_the_worker(letter_on, tmp_path, monkeypatch, tracked, expected):
    from research_mcp import server

    calls: list = []
    _stub_worker(monkeypatch, calls)
    with pytest.raises(RuntimeError, match="stub: worker reached"):
        server.apply_document_style(str(letter_on), str(tmp_path / "out.docx"), file_sha256(letter_on),
                                    profile="compact_single", tracked=tracked)
    assert calls[0][0] == "style" and calls[0][1]["tracked"] is expected


# -- live (Word) -------------------------------------------------------------------------------------------------

def _accept_in_word(path: Path, target: Path) -> Path:
    from research_mcp.operation_process import run_operation

    run_operation("tracked_check", {"jobs": [{"action": "accept", "source": str(path), "target": str(target)}]},
                  timeout_seconds=300)
    return target


@live
def test_live_style_tracked(letter_on, tmp_path):
    from research_mcp import server

    out = tmp_path / "style_tracked.docx"
    result = server.apply_document_style(str(letter_on), str(out), file_sha256(letter_on), profile="compact_single")
    assert result["tracking"] == "on" and result["tracking_author"] == "Claude"
    assert result["revisions_added"] == revision_count(out) > 0
    assert word_tracked.revision_authors(out) == {"Claude"}
    kinds = word_tracked.revision_kinds(out)
    assert kinds.get("rPrChange") and kinds.get("pPrChange") and kinds.get("sectPrChange")  # margins tracked too
    check = result["tracked_check"]
    assert check["reject_all_matches_input"] and check["accept_all_verified"]
    assert check["word_commands"] == {"accept": "AcceptAllChangesInDoc", "reject": "RejectAllChangesInDoc"}
    assert check["resolved_one_by_one"] == {"accept": 0, "reject": 0}
    with zipfile.ZipFile(out) as archive:
        assert "word/people.xml" not in archive.namelist()
    # Accept All equals the untracked run of the same input
    plain = tmp_path / "style_plain.docx"
    plain_result = server.apply_document_style(str(letter_on), str(plain), file_sha256(letter_on),
                                               profile="compact_single", tracked=False)
    assert plain_result["tracking"] == "off" and revision_count(plain) == 0
    accepted = _accept_in_word(out, tmp_path / "accepted.docx")

    def visible(path: Path) -> dict:
        # the untracked run unlinks the stripped header; the tracked run deletes its text (Word can't track
        # unlinking a header), leaving an empty one: both show no header
        return {key: value for key, value in word_tracked.formatting_snapshot(path).items()
                if not (key[0] == "sect" and key[2] in ("header", "footer") and not value)}

    assert word_tracked.differences(visible(plain), visible(accepted), lambda key: False) == []


@live
def test_live_tables_tracked_with_merge(letter_on, tmp_path):
    from research_mcp import server

    out = tmp_path / "tables_tracked.docx"
    merge = {"column": 1, "start_row": 2, "end_row": 4, "label": "Alpha"}
    result = server.format_tables(str(letter_on), str(out), [{"index": 1, "descriptive_columns": [1],
                                                             "merges": [merge]}])
    result = result.model_dump() if hasattr(result, "model_dump") else result
    assert result["tracking"] == "on"
    assert word_tracked.revision_authors(out) == {"Claude"}
    kinds = word_tracked.revision_kinds(out)
    assert kinds.get("tcPrChange") and kinds.get("tblPrChange")
    assert [(item["kind"], item["text"]) for item in result["untracked"]] == [
        ("merge", "Cells in rows 2-4 of column 1 of table 1 were merged (not tracked: Word can't track cell merges)")]
    comments = ET.fromstring(read_part(out, "word/comments.xml")).findall(W + "comment")
    assert [c.get(W + "author") for c in comments] == ["Claude"]
    assert result["tracked_check"]["reject_all_matches_input"] and result["tracked_check"]["accept_all_verified"]
    assert audit_docx(out).tables[0].vertical_merges == ((2, 4, 1, "Alpha"),)


@live
def test_live_tracked_merge_keeps_the_first_cells_runs(letter, tmp_path):
    """Review finding 1: "Al" bold, "pha" plain in each merged cell. After the tracked merge, Word's Accept All and
    Reject All both leave the first cell's runs as they were (the merge itself is the commented, untracked part)."""

    from research_mcp import server
    from research_mcp.operation_process import run_operation

    source = tmp_path / "split_on.docx"
    write_setting(_split_label(letter, tmp_path / "split.docx"), source, on=True, author="Claude")
    out = tmp_path / "merged.docx"
    result = server.format_tables(str(source), str(out), [{"index": 1, "descriptive_columns": [1], "merges": [
        {"column": 1, "start_row": 2, "end_row": 4, "label": "Alpha"}]}])
    result = result.model_dump() if hasattr(result, "model_dump") else result
    assert result["tracking"] == "on" and [item["kind"] for item in result["untracked"]] == ["merge"]
    reject = tmp_path / "reject.docx"
    run_operation("tracked_check", {"jobs": [{"action": "reject", "source": str(out), "target": str(reject)}]},
                  timeout_seconds=300)
    accepted = _accept_in_word(out, tmp_path / "accepted.docx")
    key = ("t", (1, 2, 1), "p", 1)
    before = word_tracked.formatting_snapshot(source)[key][1]
    assert [rpr != "" for _, rpr in before] == [True, True, False, False, False]
    assert word_tracked.formatting_snapshot(reject)[key][1] == before
    assert word_tracked.formatting_snapshot(accepted)[key][1] == before
    assert audit_docx(accepted).tables[0].vertical_merges == ((2, 4, 1, "Alpha"),)


@live
def test_live_style_untracked_false(letter_on, tmp_path):
    from research_mcp import server

    out = tmp_path / "plain.docx"
    result = server.apply_document_style(str(letter_on), str(out), file_sha256(letter_on), profile="compact_single",
                                         tracked=False)
    assert result["tracking"] == "off" and result["untracked"] == []
    assert revision_count(out) == 0
    with zipfile.ZipFile(out) as archive:
        assert "word/comments.xml" not in archive.namelist()
