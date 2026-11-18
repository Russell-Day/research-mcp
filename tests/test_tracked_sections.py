"""Plan 864 S3a: replace_table_row / replace_section as direct XML, tracked (switch on) or plain on a revised document.

Group 1: `revisions.delete_paragraphs` (decision 11): every run of a deleted paragraph in one `w:del` by the author
plus `<w:del/>` in its mark's rPr; the last line(s) of a cell Word's way (the mark of the line before is deleted, the
cell's last mark stays, `fix_format` restores the line before's properties with a `w:pPrChange`); untracked = the
paragraphs removed. Group 2: the executor (`tracked_sections.py`) for plans of text, insert and delete steps, and
`replace_table_row` routed through it. Expected XML is spelled out; Reject and Accept are checked with the string
engine (`resolve_revisions`), not the ElementTree resolver the code verifies itself with.
"""
from __future__ import annotations

import datetime as dt
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.sections import canonical, read_layout
from test_revisions import W, document_xml, other_parts_identical, raw_revisions, run_batch, windows_only

NOW = dt.datetime(2026, 11, 16, 12, 0, 0, tzinfo=dt.timezone.utc)
LOCAL = NOW.astimezone().strftime("%Y-%m-%dT%H:%M:%SZ")
STAMP = f'w:date="{LOCAL}" w16du:dateUtc="2026-11-16T12:00:00Z"'
AUTHOR = "Reviewer B"
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
OTHER = 'w:author="Other Author" w:date="2026-10-10T17:37:00Z"'
OTHER_INS = f'<w:ins w:id="901" {OTHER}><w:r><w:t xml:space="preserve"> here</w:t></w:r></w:ins>'


def rev(kind: str, rev_id: int, inner: str = "") -> str:
    attrs = f'w:id="{rev_id}" w:author="{AUTHOR}" {STAMP}'
    return f"<w:{kind} {attrs}/>" if not inner else f"<w:{kind} {attrs}>{inner}</w:{kind}>"


def run(text: str, rpr: str = "") -> str:
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


def deleted_run(text: str, rpr: str = "") -> str:
    return (f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}'
            f'<w:delText xml:space="preserve">{text}</w:delText></w:r>')


def para(ppr: str, *content: str) -> str:
    return "<w:p>" + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + "".join(content) + "</w:p>"


LINE_PPR = '<w:spacing w:after="0"/>'
LAST_PPR = LINE_PPR


def fixture_body(last_ppr: str = LINE_PPR, line_two: str | None = None) -> str:
    """Fixture T: a revised intro paragraph, a table row with a three-line cell, and a body section."""

    cell1 = (para(LINE_PPR, run("Role A", "<w:b/>"))
             + (line_two if line_two is not None else para(LINE_PPR, run("Line two")))
             + para(last_ppr, run("Line three", "<w:i/>")))
    row = (f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{cell1}</w:tc>'
           f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{para("", run("2024"))}</w:tc></w:tr>')
    table = ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/></w:tblGrid>' + row + "</w:tbl>")
    return (para("", run("Intro paragraph"), OTHER_INS, run("."))
            + table
            + para('<w:pStyle w:val="Heading1"/>', run("Methods"))
            + para("", run("First body.")) + para("", run("Second body.")) + para("", run("Third body."))
            + para('<w:pStyle w:val="Heading1"/>', run("Results")))


TEXTS = ["Intro paragraph here.", "Role A", "Line two", "Line three", "2024", "", "Methods",
         "First body.", "Second body.", "Third body.", "Results"]


def make(tmp_path: Path, name: str, body: str | None = None) -> Path:
    """The tracked fixture's package with its body replaced (fixture T by default)."""

    base = tmp_path / f"base-{name}.docx"
    create_tracked_fixture(base)
    target = tmp_path / f"{name}.docx"
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(target, "x") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                start = xml.index("<w:body>") + len("<w:body>")
                data = (xml[:start] + (body if body is not None else fixture_body()) + SECT
                        + "</w:body></w:document>").encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)
    return target


def first_id(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return revisions.next_revision_id({name: archive.read(name) for name in archive.namelist()})


def body(path: Path) -> ET.Element:
    return ET.fromstring(document_xml(path)).find(W + "body")


def texts(path: Path) -> list[str]:
    return [item.text for item in read_layout(path).paragraphs]


def resolved(out: Path, tmp_path: Path, mode: str, ids) -> Path:
    """Accept or reject the given ids with the string engine."""

    target = tmp_path / f"{mode}-{out.stem}.docx"
    revisions.resolve_revisions(out, target, mode=mode, ids=[str(item) for item in ids])
    return target


def new_ids(result: dict) -> list[str]:
    return [item["id"] for item in result["revisions_added"]]


def delete(source: Path, out: Path, first: int, last: int | None = None, *, tracked: bool = True, **kw) -> dict:
    return revisions.delete_paragraphs(source, out, first=first, last=last if last is not None else first,
                                       tracked=tracked, author=AUTHOR, now=NOW, **kw)


def assert_reject_is_input(source: Path, out: Path, tmp_path: Path, result: dict) -> None:
    back = resolved(out, tmp_path, "reject", new_ids(result))
    assert canonical(body(back)) == canonical(body(source))


# -- Group 1: delete_paragraphs ---------------------------------------------------------------------------------

def test_delete_middle_paragraph_tracked(tmp_path):
    source = make(tmp_path, "mid")
    out = tmp_path / "mid-out.docx"
    n = first_id(source)
    result = delete(source, out, 9)
    xml = document_xml(out)
    want = (f"<w:p><w:pPr><w:rPr>{rev('del', n)}</w:rPr></w:pPr>"
            f"{rev('del', n + 1, deleted_run('Second body.'))}</w:p>")
    assert want in xml
    assert xml.replace(want, para("", run("Second body."))) == document_xml(source)
    assert result["index_delta"] == 0 and result["tracked"] is True and result["engine"] == "direct_xml"
    assert texts(out) == TEXTS[:8] + [""] + TEXTS[9:]
    assert raw_revisions(source)[0] in raw_revisions(out)
    assert other_parts_identical(source, out)
    assert_reject_is_input(source, out, tmp_path, result)
    accepted = resolved(out, tmp_path, "accept", new_ids(result))
    assert texts(accepted) == TEXTS[:8] + TEXTS[9:]


def test_delete_middle_paragraph_untracked(tmp_path):
    source = make(tmp_path, "mid-plain")
    out = tmp_path / "mid-plain-out.docx"
    result = delete(source, out, 9, tracked=False)
    assert document_xml(out) == document_xml(source).replace(para("", run("Second body.")), "", 1)
    assert result["index_delta"] == -1 and result["revisions_added"] == [] and result["tracked"] is False
    assert texts(out) == TEXTS[:8] + TEXTS[9:]


def test_delete_two_paragraphs_tracked_ids_ascend(tmp_path):
    source = make(tmp_path, "two")
    out = tmp_path / "two-out.docx"
    n = first_id(source)
    result = delete(source, out, 8, 9)
    xml = document_xml(out)
    want = (f"<w:p><w:pPr><w:rPr>{rev('del', n)}</w:rPr></w:pPr>{rev('del', n + 1, deleted_run('First body.'))}</w:p>"
            f"<w:p><w:pPr><w:rPr>{rev('del', n + 2)}</w:rPr></w:pPr>"
            f"{rev('del', n + 3, deleted_run('Second body.'))}</w:p>")
    assert want in xml
    assert new_ids(result) == [str(n + k) for k in range(4)]
    assert_reject_is_input(source, out, tmp_path, result)
    assert texts(resolved(out, tmp_path, "accept", new_ids(result))) == TEXTS[:7] + TEXTS[9:]


def test_delete_last_cell_line_tracked_words_way(tmp_path):
    """Case C (last line): the mark of "Line two" is deleted and "Line three"'s content; the cell's last mark stays."""

    source = make(tmp_path, "last")
    out = tmp_path / "last-out.docx"
    n = first_id(source)
    result = delete(source, out, 4, trailing=True, previous=3, fix_format=True)
    xml = document_xml(out)
    want = (f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr>{run('Line two')}</w:p>"
            f"<w:p><w:pPr>{LINE_PPR}</w:pPr>{rev('del', n + 1, deleted_run('Line three', '<w:i/>'))}</w:p>")
    assert want in xml
    assert result["trailing"] is True and result["format_restored"] is False and result["index_delta"] == 0
    assert texts(out) == TEXTS[:3] + [""] + TEXTS[4:]
    assert_reject_is_input(source, out, tmp_path, result)
    accepted = resolved(out, tmp_path, "accept", new_ids(result))
    assert texts(accepted) == TEXTS[:3] + TEXTS[4:]
    cell = body(accepted).find(f"{W}tbl/{W}tr/{W}tc")
    assert [canonical(p) for p in cell.findall(W + "p")][-1] == canonical(
        ET.fromstring(f'<w:p xmlns:w="{W[1:-1]}"><w:pPr>{LINE_PPR}</w:pPr>{run("Line two")}</w:p>'))


def test_delete_last_cell_line_tracked_fix_format(tmp_path):
    """The kept mark takes "Line two"'s properties with a w:pPrChange holding its own (Word's restore of the format)."""

    other = '<w:spacing w:after="240"/><w:jc w:val="right"/>'
    source = make(tmp_path, "fix", fixture_body(last_ppr=other))
    out = tmp_path / "fix-out.docx"
    n = first_id(source)
    result = delete(source, out, 4, trailing=True, previous=3, fix_format=True)
    xml = document_xml(out)
    want = (f"<w:p><w:pPr>{LINE_PPR}{rev('pPrChange', n + 1, '<w:pPr>' + other + '</w:pPr>')}</w:pPr>"
            f"{rev('del', n + 2, deleted_run('Line three', '<w:i/>'))}</w:p>")
    assert want in xml, xml
    assert result["format_restored"] is True
    assert_reject_is_input(source, out, tmp_path, result)
    accepted = resolved(out, tmp_path, "accept", new_ids(result))
    cell = body(accepted).find(f"{W}tbl/{W}tr/{W}tc")
    assert canonical(cell.findall(W + "p")[-1]) == canonical(
        ET.fromstring(f'<w:p xmlns:w="{W[1:-1]}"><w:pPr>{LINE_PPR}</w:pPr>{run("Line two")}</w:p>'))
    # without fix_format the last mark keeps its own properties
    plain = tmp_path / "fix-keep.docx"
    kept = delete(source, plain, 4, trailing=True, previous=3, fix_format=False)
    assert kept["format_restored"] is False and "pPrChange" not in document_xml(plain)


def test_delete_last_cell_line_untracked(tmp_path):
    source = make(tmp_path, "last-plain")
    out = tmp_path / "last-plain-out.docx"
    result = delete(source, out, 4, tracked=False, trailing=True, previous=3, fix_format=True)
    xml = document_xml(out)
    assert f"{para(LINE_PPR, run('Line two'))}</w:tc>" in xml and "Line three" not in xml
    assert result["index_delta"] == -1
    assert texts(out) == TEXTS[:3] + TEXTS[4:]


def test_delete_last_cell_line_without_trailing_falls_back(tmp_path):
    source = make(tmp_path, "fallback")
    out = tmp_path / "fallback-out.docx"
    result = delete(source, out, 4)
    assert result["trailing"] is True and result["fallback_trailing"] is True
    assert result["previous_paragraph_index"] == 3
    assert_reject_is_input(source, out, tmp_path, result)


def test_delete_line_with_other_authors_revision_refused(tmp_path):
    """Case H: nothing written, the refusal names the paragraph, the revision and the way forward."""

    line = para(LINE_PPR, run("Line"), f'<w:ins w:id="902" {OTHER}><w:r><w:t xml:space="preserve"> two</w:t></w:r></w:ins>')
    source = make(tmp_path, "other", fixture_body(line_two=line))
    out = tmp_path / "other-out.docx"
    with pytest.raises(RefusedUnsafe) as caught:
        delete(source, out, 3)
    message = str(caught.value)
    assert "paragraph 3" in message and "tracked revision" in message and "accept or reject it first" in message
    assert "w:ins id=902 by 'Other Author'" in message
    assert not out.exists()


@pytest.mark.parametrize("inner, named", [
    ('<w:bookmarkStart w:id="5" w:name="mark"/>' + run("Second body.") + '<w:bookmarkEnd w:id="5"/>', "a bookmark"),
    (run("Second ") + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>'
     '<w:r><w:fldChar w:fldCharType="end"/></w:r>', "a field"),
    ('<w:hyperlink w:anchor="x">' + run("Second body.") + "</w:hyperlink>", "a hyperlink"),
])
def test_delete_paragraph_with_object_refused(tmp_path, inner, named):
    source = make(tmp_path, "object", fixture_body().replace(para("", run("Second body.")), "<w:p>" + inner + "</w:p>"))
    out = tmp_path / "object-out.docx"
    with pytest.raises(RefusedUnsafe) as caught:
        delete(source, out, 9)
    assert f"paragraph 9 (deleted) holds {named}" in str(caught.value)
    assert not out.exists()


def test_delete_paragraph_with_section_break_refused(tmp_path):
    source = make(tmp_path, "sect", fixture_body().replace(para("", run("Second body.")), para(SECT, run("Second body."))))
    with pytest.raises(RefusedUnsafe, match="holds a section break"):
        delete(source, tmp_path / "sect-out.docx", 9)


def test_delete_across_containers_refused(tmp_path):
    source = make(tmp_path, "cross")
    with pytest.raises(RefusedUnsafe, match="not consecutive paragraphs of one container"):
        delete(source, tmp_path / "cross-out.docx", 4, 5)


def test_delete_trailing_with_deleted_previous_mark_refused(tmp_path):
    line = f'<w:p><w:pPr>{LINE_PPR}<w:rPr><w:del w:id="903" {OTHER}/></w:rPr></w:pPr>{run("Line two")}</w:p>'
    source = make(tmp_path, "prevdel", fixture_body(line_two=line))
    with pytest.raises(RefusedUnsafe) as caught:
        delete(source, tmp_path / "prevdel-out.docx", 4, trailing=True, previous=3)
    assert "paragraph 3 (the line before, whose mark is deleted) holds a tracked revision" in str(caught.value)


# -- Group 2: replace_table_row through the direct-XML executor ------------------------------------------------

ROW = ["Role A\nLine two\nLine three", "2024"]


def row_edit(source: Path, out: Path, first_cell: str, *, tracked: bool | None = True, **kw) -> dict:
    from research_mcp.editing import replace_table_row_to_version

    return replace_table_row_to_version(source, out, match="Role A", cells=[first_cell, "2024"], tracked=tracked,
                                        author=AUTHOR, **kw)


def cell_texts(path: Path) -> list[str]:
    cell = body(path).find(f"{W}tbl/{W}tr/{W}tc")
    return ["".join(t.text or "" for t in p.iter(W + "t")) for p in cell.findall(W + "p")]


def views(source: Path, out: Path, tmp_path: Path, result: dict) -> list[str]:
    """Reject = input (canonical); returns the first cell's lines in the Accept view."""

    assert_reject_is_input(source, out, tmp_path, result)
    return cell_texts(resolved(out, tmp_path, "accept", new_ids(result)))


def tracked_row(monkeypatch, source, out, first_cell):
    import research_mcp.tracked_sections as ts

    real = ts.run_plan_direct
    monkeypatch.setattr(ts, "run_plan_direct", lambda *a, **k: real(*a, **dict(k, now=NOW)))
    return row_edit(source, out, first_cell)


def test_row_tracked_change_one_line(tmp_path, monkeypatch):
    """Case A: only "Line two" changes, as a word diff by Reviewer B."""

    source = make(tmp_path, "rowA")
    out = tmp_path / "rowA-out.docx"
    n = first_id(source)
    result = tracked_row(monkeypatch, source, out, "Role A\nLine 2\nLine three")
    xml = document_xml(out)
    want = (f"<w:p><w:pPr>{LINE_PPR}</w:pPr>{run('Line ')}{rev('del', n, deleted_run('two'))}"
            f"{rev('ins', n + 1, run('2'))}</w:p>")
    assert xml == document_xml(source).replace(para(LINE_PPR, run("Line two")), want)
    assert result["tracked"] is True and result["author"] == AUTHOR and result["engine"] == "direct_xml"
    assert result["index_delta"] == 0
    assert views(source, out, tmp_path, result) == ["Role A", "Line 2", "Line three"]


def test_row_tracked_add_line_middle(tmp_path, monkeypatch):
    """Case B (middle): a new paragraph before "Line three", its pPr, its first character's run formatting."""

    source = make(tmp_path, "rowB")
    out = tmp_path / "rowB-out.docx"
    n = first_id(source)
    result = tracked_row(monkeypatch, source, out, "Role A\nLine two\nNew line\nLine three")
    new = (f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('ins', n)}</w:rPr></w:pPr>"
           f"{rev('ins', n + 1, run('New line', '<w:i/>'))}</w:p>")
    anchor = para(LINE_PPR, run("Line three", "<w:i/>"))
    assert document_xml(out) == document_xml(source).replace(anchor, new + anchor)
    assert result["index_delta"] == 1 and result["shift_base"] == 7
    assert views(source, out, tmp_path, result) == ["Role A", "Line two", "New line", "Line three"]


def test_row_tracked_add_line_end(tmp_path, monkeypatch):
    """Case B (end), Word's way: "Line three"'s mark becomes the inserted one; the new last line takes its pPr."""

    source = make(tmp_path, "rowB2")
    out = tmp_path / "rowB2-out.docx"
    n = first_id(source)
    result = tracked_row(monkeypatch, source, out, "Role A\nLine two\nLine three\nTail")
    anchor = para(LINE_PPR, run("Line three", "<w:i/>"))
    want = (f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('ins', n)}</w:rPr></w:pPr>{run('Line three', '<w:i/>')}</w:p>"
            f"<w:p><w:pPr>{LINE_PPR}</w:pPr>{rev('ins', n + 1, run('Tail', '<w:i/>'))}</w:p>")
    assert document_xml(out) == document_xml(source).replace(anchor, want)
    assert result["index_delta"] == 1
    assert views(source, out, tmp_path, result) == ["Role A", "Line two", "Line three", "Tail"]


def test_row_tracked_delete_middle_line(tmp_path, monkeypatch):
    """Case C (middle): "Line two"'s runs in one w:del, its mark deleted; index_delta 0."""

    source = make(tmp_path, "rowC")
    out = tmp_path / "rowC-out.docx"
    n = first_id(source)
    result = tracked_row(monkeypatch, source, out, "Role A\nLine three")
    want = f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr>{rev('del', n + 1, deleted_run('Line two'))}</w:p>"
    assert document_xml(out) == document_xml(source).replace(para(LINE_PPR, run("Line two")), want)
    assert result["index_delta"] == 0 and result["accepted_index_delta"] == -1
    assert views(source, out, tmp_path, result) == ["Role A", "Line three"]


def test_row_tracked_delete_last_line(tmp_path, monkeypatch):
    """Case C (last), Word's way: "Line two"'s mark deleted, "Line three"'s runs deleted, the last mark kept."""

    source = make(tmp_path, "rowC2")
    out = tmp_path / "rowC2-out.docx"
    n = first_id(source)
    result = tracked_row(monkeypatch, source, out, "Role A\nLine two")
    old = para(LINE_PPR, run("Line two")) + para(LINE_PPR, run("Line three", "<w:i/>"))
    want = (f"<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n)}</w:rPr></w:pPr>{run('Line two')}</w:p>"
            f"<w:p><w:pPr>{LINE_PPR}</w:pPr>{rev('del', n + 1, deleted_run('Line three', '<w:i/>'))}</w:p>")
    assert document_xml(out) == document_xml(source).replace(old, want)
    assert result["index_delta"] == 0
    assert views(source, out, tmp_path, result) == ["Role A", "Line two"]


def test_row_tracked_delete_last_line_fix_format(tmp_path, monkeypatch):
    other = '<w:spacing w:after="240"/><w:jc w:val="right"/>'
    source = make(tmp_path, "rowC3", fixture_body(last_ppr=other))
    out = tmp_path / "rowC3-out.docx"
    n = first_id(source)
    result = tracked_row(monkeypatch, source, out, "Role A\nLine two")
    want = (f"<w:p><w:pPr>{LINE_PPR}{rev('pPrChange', n + 1, '<w:pPr>' + other + '</w:pPr>')}</w:pPr>"
            f"{rev('del', n + 2, deleted_run('Line three', '<w:i/>'))}</w:p>")
    assert want in document_xml(out)
    assert views(source, out, tmp_path, result) == ["Role A", "Line two"]


def test_row_tracked_change_and_delete_ids_ascend(tmp_path, monkeypatch):
    """A changed line above a deleted last line: the steps run bottom-up, the ids are renumbered top-down."""

    source = make(tmp_path, "rowmix")
    out = tmp_path / "rowmix-out.docx"
    n = first_id(source)
    result = tracked_row(monkeypatch, source, out, "Role B\nLine two")
    assert new_ids(result) == [str(n + k) for k in range(len(result["revisions_added"]))]
    ids = [int(m.group(1)) for m in re.finditer(r'<w:(?:ins|del|pPrChange) w:id="(\d+)" w:author="Reviewer B"',
                                                document_xml(out))]
    assert ids == sorted(ids) and ids[0] == n
    assert views(source, out, tmp_path, result) == ["Role B", "Line two"]


def test_row_untracked_on_revised_doc(tmp_path):
    """Case E (row): switch off on a revised document: plain direct XML, no new revisions."""

    source = make(tmp_path, "rowE")
    out = tmp_path / "rowE-out.docx"
    result = row_edit(source, out, "Role A\nLine 2\nAdded", tracked=False)
    xml = document_xml(out)
    assert result["tracked"] is False and result["revisions_added"] == [] and result["engine"] == "direct_xml"
    assert raw_revisions(out) == raw_revisions(source)
    old = para(LINE_PPR, run("Line two")) + para(LINE_PPR, run("Line three", "<w:i/>"))
    assert xml == document_xml(source).replace(old, para(LINE_PPR, run("Line ") + run("2"))
                                                   + para(LINE_PPR, run("Added", "<w:i/>")))
    assert result["index_delta"] == 0
    assert cell_texts(out) == ["Role A", "Line 2", "Added"]


def test_row_delete_line_with_other_authors_revision_refused(tmp_path, monkeypatch):
    """Case H through replace_table_row: the deletion refuses, naming the paragraph and the way forward."""

    line = para(LINE_PPR, run("Line"), f'<w:ins w:id="902" {OTHER}><w:r><w:t xml:space="preserve"> two</w:t></w:r></w:ins>')
    source = make(tmp_path, "rowH", fixture_body(line_two=line))
    out = tmp_path / "rowH-out.docx"
    with pytest.raises(RefusedUnsafe) as caught:
        tracked_row(monkeypatch, source, out, "Role A\nLine three")
    message = str(caught.value)
    assert "paragraph 3" in message and "tracked revision" in message and "accept or reject it first" in message
    assert "w:ins id=902 by 'Other Author' (2026-10-10T17:37:00Z)" in message
    assert not out.exists()


def rewrite_document(path: Path, change) -> None:
    """Rewrite word/document.xml of a docx in place (fault injection)."""

    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w") as out:
        for info, data in items:
            if info.filename == "word/document.xml":
                data = change(data.decode("utf-8")).encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)


def tampering_text_edit(monkeypatch, old: str, new: str) -> None:
    """The executor's text step writes its result, then `old` becomes `new` in it (a faulty step)."""

    real = revisions.text_edit

    def faulty(source, output, **kwargs):
        result = real(source, output, **kwargs)
        rewrite_document(Path(output), lambda xml: xml.replace(old, new, 1))
        return result

    monkeypatch.setattr(revisions, "text_edit", faulty)


def test_executor_refuses_bytes_changed_outside_the_row(tmp_path, monkeypatch):
    """Decision 12: a step that changes anything outside the row fails the whole-op check; nothing written."""

    from research_mcp.errors import VerificationFailed

    source = make(tmp_path, "faultout")
    out = tmp_path / "faultout-out.docx"
    tampering_text_edit(monkeypatch, ">Intro paragraph<", ">Intro Paragraph<")
    with pytest.raises(VerificationFailed, match="bytes outside the edited rows/paragraphs changed"):
        tracked_row(monkeypatch, source, out, "Role A\nLine 2\nLine three")
    assert not out.exists()


def test_executor_refuses_when_reject_is_not_the_input(tmp_path, monkeypatch):
    """Decision 12: an untracked change inside the row (Reject no longer gives the input) fails; nothing written."""

    from research_mcp.errors import VerificationFailed

    source = make(tmp_path, "faultrej")
    out = tmp_path / "faultrej-out.docx"
    tampering_text_edit(monkeypatch, ">Role A<", ">Role X<")
    with pytest.raises(VerificationFailed, match="rejecting the new revisions does not give back the input"):
        tracked_row(monkeypatch, source, out, "Role A\nLine 2\nLine three")
    assert not out.exists()


def test_delete_refuses_when_reject_is_not_the_input(tmp_path, monkeypatch):
    """delete_paragraphs' own Reject check: deleted text that is not the paragraph's text fails; nothing written."""

    from research_mcp.errors import VerificationFailed

    real = revisions._deleted_content
    monkeypatch.setattr(revisions, "_deleted_content",
                        lambda *a, **k: real(*a, **k).replace("Second body.", "Second b0dy."))
    source = make(tmp_path, "faultdel")
    out = tmp_path / "faultdel-out.docx"
    with pytest.raises(VerificationFailed, match="rejecting the deletion does not give back the input"):
        delete(source, out, 9)
    assert not out.exists()


@windows_only
def test_batch_switch_on_and_tracked_false_override(tmp_path):
    """Case F (revised document): the switch on, `tracked: false` on the op: plain direct XML, tracking "off"."""

    source = make(tmp_path, "rowF")
    on = run_batch(source, {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR})
    assert on.status == "complete", on
    before = tmp_path / "rowF-before.docx"
    before.write_bytes(source.read_bytes())
    result = run_batch(source, {"op_id": "r", "type": "replace_table_row", "match": "Role A", "tracked": False,
                                "cells": ["Role A\nLine 2\nLine three", "2024"]})
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.tracking == "on"  # the document's switch stays on; the op itself ran untracked
    details = result.operations[0].details
    assert details["tracking"] == "off"
    assert details["engine"] == "direct_xml" and details["tracked"] is False and details["revisions_added"] == []
    assert raw_revisions(source) == raw_revisions(before)
    assert "Reviewer B" not in document_xml(source)
    assert document_xml(source) == document_xml(before).replace(para(LINE_PPR, run("Line two")),
                                                                 para(LINE_PPR, run("Line ") + run("2")))


@windows_only
def test_batch_switch_on_and_tracked_false_clean_doc_goes_to_word(tmp_path, monkeypatch):
    """Case F (revision-free document): `tracked: false` while on keeps the Word path (`_run_plan`)."""

    import research_mcp.editing as editing
    import research_mcp.tracked_sections as ts

    clean = fixture_body().replace(OTHER_INS, run(" here"))
    source = make(tmp_path, "rowF2", clean)
    on = run_batch(source, {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR})
    assert on.status == "complete", on
    calls = []

    def word_path(*args, **kwargs):
        calls.append(args[2])
        raise RefusedUnsafe("sentinel: Word path")

    import research_mcp.edit_batch as edit_batch

    # ops normally run in a child process; run them here so the patches below apply
    monkeypatch.setattr(edit_batch, "_apply_operation", edit_batch._apply_operation_local)
    monkeypatch.setattr(editing, "_run_plan", word_path)
    monkeypatch.setattr(ts, "run_plan_direct", lambda *a, **k: pytest.fail("direct XML path taken"))
    result = run_batch(source, {"op_id": "r", "type": "replace_table_row", "match": "Role A", "tracked": False,
                                "cells": ["Role A\nLine 2\nLine three", "2024"]})
    assert calls == ["replace_table_row"]
    assert "sentinel: Word path" in (result.operations[0].error.message if result.operations[0].error else "")


@windows_only
def test_cv_role_row_sibling(tmp_path):
    """Sibling: a revision-free CV, the switch on, a role row's lines changed, added and removed."""

    from research_mcp.fixtures import create_mixed_run_fixture

    cv = create_mixed_run_fixture(tmp_path / "cv.docx")
    before = tmp_path / "cv-before.docx"
    before.write_bytes(cv.read_bytes())
    on = run_batch(cv, {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR})
    assert on.status == "complete", on
    switched = tmp_path / "cv-switched.docx"
    switched.write_bytes(cv.read_bytes())
    layout = read_layout(cv)
    title = next(item for item in layout.paragraphs if item.text.endswith("Teaching Assistant"))
    row = next(r for r in layout.tables[title.table] if title.index in r.cells[0])
    date = "\n".join(layout.para(i).text for i in row.cells[1])
    result = run_batch(cv, {"op_id": "r", "type": "replace_table_row", "match": title.text,
                            "cells": [title.text + "\nField research on soil microbes and crop yields\nAdded line",
                                      date]})
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.tracking == "on" and result.tracking_author == AUTHOR
    xml = document_xml(cv)
    authors = {m.group(1) for m in re.finditer(r'<w:(?:ins|del)\b[^>]*w:author="([^"]+)"', xml)}
    assert authors == {AUTHOR}
    ids = [m.group(0) for m in re.finditer(r'<w:(?:ins|del)\b[^>]*>', xml)]
    rejected = tmp_path / "cv-rejected.docx"
    revisions.resolve_revisions(cv, rejected, mode="reject", author=AUTHOR)
    assert canonical(body(rejected)) == canonical(body(switched))
    accepted = tmp_path / "cv-accepted.docx"
    revisions.resolve_revisions(cv, accepted, mode="accept", author=AUTHOR)
    lines = [item.text for item in read_layout(accepted).paragraphs]
    at = lines.index(title.text)
    assert lines[at:at + 3] == [title.text, "Field research on soil microbes and crop yields", "Added line"]
    assert ids


# -- Group 3: replace_section through the direct-XML executor --------------------------------------------------

BODY = ["First body.", "Second body.", "Third body."]
FIRST, SECOND, THIRD = (para("", run(text)) for text in BODY)


def section_edit(monkeypatch, source: Path, out: Path, *, tracked: bool | None = True, **kw) -> dict:
    from research_mcp.editing import replace_section_to_version
    import research_mcp.tracked_sections as ts

    real = ts.run_plan_direct
    monkeypatch.setattr(ts, "run_plan_direct", lambda *a, **k: real(*a, **dict(k, now=NOW)))
    kw.setdefault("heading", "Methods")
    return replace_section_to_version(source, out, tracked=tracked, author=AUTHOR, **kw)


def accepted_texts(out: Path, tmp_path: Path, result: dict) -> list[str]:
    return texts(resolved(out, tmp_path, "accept", new_ids(result)))


def test_section_tracked_add(tmp_path, monkeypatch):
    """Case D1: a new paragraph between "Second body." and "Third body.", its mark and run inserted."""

    source = make(tmp_path, "secD1")
    out = tmp_path / "secD1-out.docx"
    n = first_id(source)
    result = section_edit(monkeypatch, source, out, paragraphs=BODY[:2] + ["Added body."] + BODY[2:])
    new = f"<w:p><w:pPr><w:rPr>{rev('ins', n)}</w:rPr></w:pPr>{rev('ins', n + 1, run('Added body.'))}</w:p>"
    assert document_xml(out) == document_xml(source).replace(SECOND + THIRD, SECOND + new + THIRD)
    assert result["tracked"] is True and result["engine"] == "direct_xml" and result["section_kind"] == "paragraphs"
    assert result["index_delta"] == 1 and result["accepted_index_delta"] == 1
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == TEXTS[:9] + ["Added body."] + TEXTS[9:]


def test_section_tracked_remove(tmp_path, monkeypatch):
    """Case D2: "Second body." runs in one w:del and its mark deleted; the heading untouched; index_delta 0."""

    source = make(tmp_path, "secD2")
    out = tmp_path / "secD2-out.docx"
    n = first_id(source)
    result = section_edit(monkeypatch, source, out, paragraphs=[BODY[0], BODY[2]])
    want = f"<w:p><w:pPr><w:rPr>{rev('del', n)}</w:rPr></w:pPr>{rev('del', n + 1, deleted_run('Second body.'))}</w:p>"
    assert document_xml(out) == document_xml(source).replace(SECOND, want)
    assert result["index_delta"] == 0 and result["accepted_index_delta"] == -1
    assert texts(out) == TEXTS[:8] + [""] + TEXTS[9:]
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == TEXTS[:8] + TEXTS[9:]


def test_section_tracked_remove_last_paragraph(tmp_path, monkeypatch):
    """The section's last paragraph: its mark deleted too; Accept keeps the next heading as it was."""

    source = make(tmp_path, "secD2b")
    out = tmp_path / "secD2b-out.docx"
    n = first_id(source)
    result = section_edit(monkeypatch, source, out, paragraphs=BODY[:2])
    want = f"<w:p><w:pPr><w:rPr>{rev('del', n)}</w:rPr></w:pPr>{rev('del', n + 1, deleted_run('Third body.'))}</w:p>"
    assert document_xml(out) == document_xml(source).replace(THIRD, want)
    assert_reject_is_input(source, out, tmp_path, result)
    accepted = resolved(out, tmp_path, "accept", new_ids(result))
    assert texts(accepted) == TEXTS[:9] + TEXTS[10:]
    heading = para('<w:pStyle w:val="Heading1"/>', run("Results"))
    assert f"{SECOND}{heading}" in document_xml(accepted)


def test_section_tracked_reword(tmp_path, monkeypatch):
    """Case D3: a word diff inside "Second body." only."""

    source = make(tmp_path, "secD3")
    out = tmp_path / "secD3-out.docx"
    n = first_id(source)
    result = section_edit(monkeypatch, source, out, paragraphs=[BODY[0], "Second body changed.", BODY[2]])
    want = f"<w:p>{run('Second body')}{rev('ins', n, run(' changed'))}{run('.')}</w:p>"
    assert document_xml(out) == document_xml(source).replace(SECOND, want)
    assert result["index_delta"] == 0 and new_ids(result) == [str(n)]
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == TEXTS[:8] + ["Second body changed."] + TEXTS[9:]


def test_section_untracked_on_revised_doc(tmp_path, monkeypatch):
    """Case E (section): switch off on a revised document: plain direct XML; reword, remove and add in one op."""

    source = make(tmp_path, "secE")
    out = tmp_path / "secE-out.docx"
    result = section_edit(monkeypatch, source, out, tracked=None,
                          paragraphs=["First body changed.", "Third body.", "Added body."])
    assert result["tracked"] is False and result["engine"] == "direct_xml" and result["revisions_added"] == []
    assert raw_revisions(out) == raw_revisions(source)
    assert "Second body." not in document_xml(out)
    assert texts(out) == TEXTS[:7] + ["First body changed.", "Third body.", "Added body."] + TEXTS[10:]
    assert result["index_delta"] == 0
    assert other_parts_identical(source, out)


def experience_doc(tmp_path: Path, name: str) -> Path:
    """Fixture T with a heading "Experience" right before the table, so the table is a rows section."""

    heading = para('<w:pStyle w:val="Heading1"/>', run("Experience"))
    intro = para("", run("Intro paragraph"), OTHER_INS, run("."))
    return make(tmp_path, name, fixture_body().replace(intro, intro + heading, 1))


def test_section_rows_tracked_without_row_count_change(tmp_path, monkeypatch):
    """rows=: the same rows with one line changed: case A's word diff, through replace_section."""

    source = experience_doc(tmp_path, "secR")
    out = tmp_path / "secR-out.docx"
    n = first_id(source)
    result = section_edit(monkeypatch, source, out, heading="Experience",
                          rows=[["Role A\nLine 2\nLine three", "2024"]])
    want = (f"<w:p><w:pPr>{LINE_PPR}</w:pPr>{run('Line ')}{rev('del', n, deleted_run('two'))}"
            f"{rev('ins', n + 1, run('2'))}</w:p>")
    assert document_xml(out) == document_xml(source).replace(para(LINE_PPR, run("Line two")), want)
    assert result["section_kind"] == "rows" and result["index_delta"] == 0
    assert views(source, out, tmp_path, result) == ["Role A", "Line 2", "Line three"]


@windows_only
def test_batch_switch_on_replace_section_tracked(tmp_path):
    """edit_document: replace_section is no longer refused while the switch is on; it runs tracked by its author."""

    source = make(tmp_path, "secBatch")
    on = run_batch(source, {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR})
    assert on.status == "complete", on
    switched = tmp_path / "secBatch-switched.docx"
    switched.write_bytes(source.read_bytes())
    result = run_batch(source, {"op_id": "r", "type": "replace_section", "heading": "Methods",
                                "paragraphs": [BODY[0], "Second body changed.", BODY[2], "Added body."]})
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    details = result.operations[0].details
    assert details["tracking"] == "on" and details["routed_as"] == "replace_section" and details["tracked"] is True
    rejected = tmp_path / "secBatch-rejected.docx"
    revisions.resolve_revisions(source, rejected, mode="reject", author=AUTHOR)
    assert canonical(body(rejected)) == canonical(body(switched))
    accepted = tmp_path / "secBatch-accepted.docx"
    revisions.resolve_revisions(source, accepted, mode="accept", author=AUTHOR)
    assert texts(accepted)[7:11] == [BODY[0], "Second body changed.", BODY[2], "Added body."]


def test_section_tracked_reword_and_add_at_end(tmp_path, monkeypatch):
    """A reworded paragraph and one added after the section's last paragraph: the new mark is inserted before the
    next heading, which the Reject check reads too (its inserted mark resolves into that heading); ids ascend."""

    source = make(tmp_path, "secRA")
    out = tmp_path / "secRA-out.docx"
    n = first_id(source)
    result = section_edit(monkeypatch, source, out, paragraphs=[BODY[0], "Second body changed.", BODY[2], "Added body."])
    reworded = f"<w:p>{run('Second body')}{rev('ins', n, run(' changed'))}{run('.')}</w:p>"
    added = f"<w:p><w:pPr><w:rPr>{rev('ins', n + 1)}</w:rPr></w:pPr>{rev('ins', n + 2, run('Added body.'))}</w:p>"
    assert document_xml(out) == document_xml(source).replace(SECOND + THIRD, reworded + THIRD + added)
    assert new_ids(result) == [str(n), str(n + 1), str(n + 2)] and result["index_delta"] == 1
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == TEXTS[:8] + ["Second body changed.", "Third body.", "Added body."] \
        + TEXTS[10:]


# -- Batches: later indices and allow_partial after a tracked or plain deletion (review r1) ---------------------
SECTION_DELETE = {"op_id": "sec", "type": "replace_section", "heading": "Methods",
                  "paragraphs": ["First body.", "Third body."]}
LATER_EDIT = {"op_id": "later", "type": "replace_text", "paragraph_index": 11, "find": "Results", "replace": "Findings"}


def all_accepted_texts(path: Path, tmp_path: Path) -> list[str]:
    target = tmp_path / f"accepted-{path.stem}.docx"
    revisions.resolve_revisions(path, target, mode="accept", author=AUTHOR)
    return texts(target)


def test_batch_tracked_delete_keeps_later_indices(tmp_path):
    """A tracked deletion leaves its paragraph in the layout until accepted (index_delta 0): an op later in the
    batch that names paragraph 11 ("Results") by its original index still edits that paragraph."""

    source = make(tmp_path, "batchT")
    result = run_batch(source, {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR},
                       SECTION_DELETE, LATER_EDIT)
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.operations[1].details["index_delta"] == 0
    assert texts(source)[8:11] == ["", "Third body.", "Findings"]  # paragraph 9 still there, its text deleted
    assert all_accepted_texts(source, tmp_path)[6:] == ["Methods", "First body.", "Third body.", "Findings"]


def test_batch_plain_delete_shifts_later_indices(tmp_path):
    """Switch off on a revised document: the deletion removes paragraph 9, so the later op's original index 11 is
    shifted to 10 and still edits "Results"."""

    source = make(tmp_path, "batchP")
    result = run_batch(source, SECTION_DELETE, LATER_EDIT)
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.operations[0].details["index_delta"] == -1
    assert texts(source)[6:] == ["Methods", "First body.", "Third body.", "Findings"]


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
def test_batch_allow_partial_rolls_back_a_deletion_with_its_group(tmp_path, tracked):
    """The deletion shares a group with an op that fails: both roll back, no index shift is left behind, and the
    later unrelated op (original index 11) still edits "Results"."""

    source = make(tmp_path, f"batchG-{tracked}")
    ops = [{"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR}] if tracked else []
    bad = {"op_id": "bad", "type": "replace_text", "paragraph_index": 1, "find": "no such text", "replace": "x",
           "group": "g"}
    result = run_batch(source, *ops, dict(SECTION_DELETE, group="g"), bad, LATER_EDIT, allow_partial=True)
    statuses = {o.op_id: o.status for o in result.operations}
    assert statuses["sec"] == "rolled_back" and statuses["bad"] == "failed" and statuses["later"] == "applied", statuses
    assert texts(source)[6:] == ["Methods", "First body.", "Second body.", "Third body.", "Findings"]
    assert "Second body.</w:delText>" not in document_xml(source)  # the rolled-back deletion left no w:del


# -- Plan 864 S3b group 3: replace_section(rows=) adding and removing table rows ----------------------------------
# Fixture X: a revised intro, heading "Experience", a 3-row table A / B / C, then heading "Methods" and body.
# Layout: 1 intro, 2 Experience, 3-7 row A (three lines, the date, the row-end marker), 8-10 row B, 11-13 row C,
# 14 Methods, 15-17 body, 18 Results.

W16DU = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
TCPR1 = '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
TCPR2 = '<w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
BOLD = "<w:b/>"
ROWS = [["Role A\nLine two\nLine three", "2024"], ["Role B", "2023"], ["Role C", "2022"]]


def table_row(cells: list[str], trpr: str = "", extra: str = "") -> str:
    first, *rest = cells[0].split("\n")
    cell1 = para(LINE_PPR, run(first, BOLD)) + "".join(para(LINE_PPR, run(line)) for line in rest)
    return (f"<w:tr>{trpr}<w:tc>{TCPR1}{cell1}</w:tc>"
            f'<w:tc>{TCPR2}{para("", run(cells[1]), extra)}</w:tc></w:tr>')


def rows_body(rows: list[str], *, revised: bool = True) -> str:
    table = ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/></w:tblGrid>' + "".join(rows) + "</w:tbl>")
    intro = para("", run("Intro paragraph"), OTHER_INS if revised else run(" here"), run("."))
    return (intro + para('<w:pStyle w:val="Heading1"/>', run("Experience")) + table
            + para('<w:pStyle w:val="Heading1"/>', run("Methods"))
            + para("", run("First body.")) + para("", run("Second body.")) + para("", run("Third body."))
            + para('<w:pStyle w:val="Heading1"/>', run("Results")))


def rows_doc(tmp_path: Path, name: str, rows: list[str] | None = None, *, revised: bool = True) -> Path:
    return make(tmp_path, name, rows_body(rows or [table_row(cells) for cells in ROWS], revised=revised))


HEAD = ["Intro paragraph here.", "Experience"]
TAIL = ["Methods", "First body.", "Second body.", "Third body.", "Results"]


def row_lines(*rows: list[str]) -> list[str]:
    out = []
    for cells in rows:
        out += cells[0].split("\n") + [cells[1], ""]
    return out


def rows_edit(monkeypatch, source: Path, out: Path, rows: list[list[str]], **kw) -> dict:
    return section_edit(monkeypatch, source, out, heading="Experience", rows=rows, **kw)


def tracked_rows_of(xml: str, kind: str) -> list[str]:
    return [m.group(0) for m in re.finditer(r"<w:tr>(?:(?!</w:tr>).)*?<w:trPr>(?:(?!</w:trPr>).)*?<w:" + kind
                                            + r" (?:(?!</w:tr>).)*</w:tr>", xml, re.S)]


def parse_row(xml: str) -> ET.Element:
    return ET.fromstring(f'<w:root xmlns:w="{W[1:-1]}" xmlns:w16du="{W16DU}">{xml}</w:root>')[0]


def test_rows_tracked_append(tmp_path, monkeypatch):
    """X1: a row added after the table's last row: a plain copy of row C rewritten to D, then recorded as Word
    records an added row; all of its text inside its own insertions, no deletion; ids ascend."""

    source = rows_doc(tmp_path, "X1")
    out = tmp_path / "X1-out.docx"
    n = first_id(source)
    result = rows_edit(monkeypatch, source, out, ROWS + [["Role D", "2021"]])
    xml = document_xml(out)
    new_row = (f"<w:tr><w:trPr>{rev('ins', n)}</w:trPr>"
               f"<w:tc>{TCPR1}<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('ins', n + 1)}</w:rPr></w:pPr>"
               f"{rev('ins', n + 2, run('Role ', BOLD) + run('D', BOLD))}</w:p></w:tc>"
               f"<w:tc>{TCPR2}<w:p><w:pPr><w:rPr>{rev('ins', n + 3)}</w:rPr></w:pPr>"
               f"{rev('ins', n + 4, run('202') + run('1'))}</w:p></w:tc></w:tr>")
    last = table_row(ROWS[2])
    assert xml == document_xml(source).replace(last, last + new_row)
    assert "<w:del " not in xml
    assert new_ids(result) == [str(n + k) for k in range(5)]
    assert result["engine"] == "direct_xml" and result["tracked"] is True
    assert result["index_delta"] == 3 and result["accepted_index_delta"] == 3
    assert [step["step"] for step in result["steps"]] == ["add row", "text", "text"]
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == HEAD + row_lines(*ROWS, ["Role D", "2021"]) + TAIL


def test_rows_tracked_insert_middle_fewer_lines(tmp_path, monkeypatch):
    """X2: a row between A and B whose first cell has two lines (its template A has three): the copy's third line
    is removed before the row is tracked, so the new row holds no deletion, and its lines read as requested."""

    source = rows_doc(tmp_path, "X2")
    out = tmp_path / "X2-out.docx"
    new = ["Role N\nNew line", "2025"]
    result = rows_edit(monkeypatch, source, out, [ROWS[0], new, ROWS[1], ROWS[2]])
    xml = document_xml(out)
    added = tracked_rows_of(xml, "ins")
    assert len(added) == 1 and "<w:del " not in xml
    row = parse_row(added[0])
    # every run of the new row inside a w:ins, every paragraph mark inserted
    assert all(r in [x for ins in row.iter(W + "ins") for x in ins.iter(W + "r")] for r in row.iter(W + "r"))
    assert all(p.find(f"{W}pPr/{W}rPr/{W}ins") is not None for p in row.iter(W + "p"))
    assert len(list(row.iter(W + "p"))) == 3
    assert result["index_delta"] == 4 and result["accepted_index_delta"] == 4
    assert texts(out) == HEAD + row_lines(ROWS[0], new, ROWS[1], ROWS[2]) + TAIL
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == HEAD + row_lines(ROWS[0], new, ROWS[1], ROWS[2]) + TAIL


def test_rows_tracked_two_above_first_row(tmp_path, monkeypatch):
    """X3: two rows added above the section's first row: both are copies of row A (its bold first line), in order;
    the second copy is not made from the first (the planner names the template where it now stands)."""

    source = rows_doc(tmp_path, "X3")
    out = tmp_path / "X3-out.docx"
    x, y = ["Role X\nLine x\nLine three", "2026"], ["Role Y\nLine y\nLine three", "2025"]
    result = rows_edit(monkeypatch, source, out, [x, y] + ROWS)
    added = tracked_rows_of(document_xml(out), "ins")
    assert len(added) == 2
    for row, title in zip(added, ["Role X", "Role Y"]):  # the first line bold, as row A's; X before Y
        element = parse_row(row)
        bold = "".join(t.text for r in element.iter(W + "r") if r.find(f"{W}rPr/{W}b") is not None
                       for t in r.iter(W + "t"))
        assert bold == title
    assert result["index_delta"] == 10
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == HEAD + row_lines(x, y, *ROWS) + TAIL
    ids = [int(i) for i in new_ids(result)]
    assert ids == list(range(first_id(source), first_id(source) + len(ids)))


def test_rows_tracked_remove(tmp_path, monkeypatch):
    """X4: row B removed: Word's row deletion (row mark, every mark and run deleted); its paragraphs stay in the
    layout reading "" until accepted (index_delta 0, accepted_index_delta -3)."""

    source = rows_doc(tmp_path, "X4")
    out = tmp_path / "X4-out.docx"
    n = first_id(source)
    result = rows_edit(monkeypatch, source, out, [ROWS[0], ROWS[2]])
    want = (f"<w:tr><w:trPr>{rev('del', n)}</w:trPr>"
            f"<w:tc>{TCPR1}<w:p><w:pPr>{LINE_PPR}<w:rPr>{rev('del', n + 1)}</w:rPr></w:pPr>"
            f"{rev('del', n + 2, deleted_run('Role B', BOLD))}</w:p></w:tc>"
            f"<w:tc>{TCPR2}<w:p><w:pPr><w:rPr>{rev('del', n + 3)}</w:rPr></w:pPr>"
            f"{rev('del', n + 4, deleted_run('2023'))}</w:p></w:tc></w:tr>")
    assert document_xml(out) == document_xml(source).replace(table_row(ROWS[1]), want)
    assert result["index_delta"] == 0 and result["accepted_index_delta"] == -3
    assert result["indices_after_shifted"] is False
    assert texts(out) == HEAD + row_lines(ROWS[0]) + ["", "", ""] + row_lines(ROWS[2]) + TAIL
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == HEAD + row_lines(ROWS[0], ROWS[2]) + TAIL


def test_rows_tracked_mixed(tmp_path, monkeypatch):
    """X5: one op rewords row A, removes row B and adds row D: ids ascend in document order."""

    source = rows_doc(tmp_path, "X5")
    out = tmp_path / "X5-out.docx"
    a = ["Role A\nLine 2\nLine three", "2024"]
    d = ["Role D", "2021"]
    result = rows_edit(monkeypatch, source, out, [a, ROWS[2], d])
    xml = document_xml(out)
    assert len(tracked_rows_of(xml, "ins")) == 1 and len(tracked_rows_of(xml, "del")) == 1
    numbers = [int(m.group(1)) for m in re.finditer(r'<w:(?:ins|del) w:id="(\d+)" w:author="Reviewer B"', xml)]
    assert numbers == sorted(numbers) == [int(i) for i in new_ids(result)]
    assert result["index_delta"] == 3 and result["accepted_index_delta"] == 0
    assert_reject_is_input(source, out, tmp_path, result)
    assert accepted_texts(out, tmp_path, result) == HEAD + row_lines(a, ROWS[2], d) + TAIL


def test_rows_plain_on_revised_doc(tmp_path, monkeypatch):
    """X6: switch off on a revised document: row B's w:tr removed, D a plain copy; no revision added."""

    source = rows_doc(tmp_path, "X6")
    out = tmp_path / "X6-out.docx"
    d = ["Role D", "2021"]
    result = rows_edit(monkeypatch, source, out, [ROWS[0], ROWS[2], d], tracked=None)
    assert result["tracked"] is False and result["revisions_added"] == []
    assert raw_revisions(out) == raw_revisions(source)
    assert table_row(ROWS[1]) not in document_xml(out)
    assert texts(out) == HEAD + row_lines(ROWS[0], ROWS[2], d) + TAIL
    assert result["index_delta"] == 0 and other_parts_identical(source, out)


def test_rows_remove_row_with_revision_refused(tmp_path, monkeypatch):
    """X7: the row to remove holds another author's insertion: refused, named, nothing written."""

    rows = [table_row(ROWS[0]), table_row(ROWS[1], extra=OTHER_INS), table_row(ROWS[2])]
    source = rows_doc(tmp_path, "X7", rows)
    out = tmp_path / "X7-out.docx"
    with pytest.raises(RefusedUnsafe) as caught:
        rows_edit(monkeypatch, source, out, [ROWS[0], ROWS[2]])
    message = str(caught.value)
    assert "table 1 row 2" in message and "Other Author" in message and "accept or reject it first" in message
    assert not out.exists()


def test_rows_tracked_remove_empty_trpr(tmp_path, monkeypatch):
    """X8: a removed row whose trPr is empty (`<w:trPr/>`): Reject drops the trPr the mark emptied, as Word does,
    and the whole-op check reads an empty trPr as absent."""

    rows = [table_row(ROWS[0]), table_row(ROWS[1], trpr="<w:trPr/>"), table_row(ROWS[2])]
    source = rows_doc(tmp_path, "X8", rows)
    out = tmp_path / "X8-out.docx"
    result = rows_edit(monkeypatch, source, out, [ROWS[0], ROWS[2]])
    assert len(tracked_rows_of(document_xml(out), "del")) == 1
    assert accepted_texts(out, tmp_path, result) == HEAD + row_lines(ROWS[0], ROWS[2]) + TAIL


# -- Batches with rows added or removed (copied from S3a's batch tests) --------------------------------------------
ROWS_DELETE = {"op_id": "sec", "type": "replace_section", "heading": "Experience", "rows": [ROWS[0], ROWS[2]]}
ROWS_ADD = {"op_id": "sec", "type": "replace_section", "heading": "Experience",
            "rows": ROWS + [["Role D", "2021"]]}
ROWS_LATER = {"op_id": "later", "type": "replace_text", "paragraph_index": 18, "find": "Results",
              "replace": "Findings"}
SWITCH_ON = {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR}


def test_batch_rows_tracked_delete_keeps_later_indices(tmp_path):
    source = rows_doc(tmp_path, "XbT")
    result = run_batch(source, SWITCH_ON, ROWS_DELETE, ROWS_LATER)
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.operations[1].details["index_delta"] == 0
    assert texts(source)[-1] == "Findings"
    assert all_accepted_texts(source, tmp_path) == HEAD + row_lines(ROWS[0], ROWS[2]) + TAIL[:-1] + ["Findings"]


def test_batch_rows_plain_delete_shifts_later_indices(tmp_path):
    source = rows_doc(tmp_path, "XbP")
    result = run_batch(source, ROWS_DELETE, ROWS_LATER)
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.operations[0].details["index_delta"] == -3
    assert texts(source) == HEAD + row_lines(ROWS[0], ROWS[2]) + TAIL[:-1] + ["Findings"]


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
def test_batch_rows_add_shifts_later_indices(tmp_path, tracked):
    source = rows_doc(tmp_path, f"XbA-{tracked}")
    result = run_batch(source, *([SWITCH_ON] if tracked else []), ROWS_ADD, ROWS_LATER)
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.operations[-2].details["index_delta"] == 3
    assert texts(source) == HEAD + row_lines(*ROWS, ["Role D", "2021"]) + TAIL[:-1] + ["Findings"]


@pytest.mark.parametrize("tracked", [True, False], ids=["tracked", "plain"])
def test_batch_rows_allow_partial_rolls_back_with_its_group(tmp_path, tracked):
    source = rows_doc(tmp_path, f"XbG-{tracked}")
    before = texts(source)
    bad = {"op_id": "bad", "type": "replace_text", "paragraph_index": 1, "find": "no such text", "replace": "x",
           "group": "g"}
    result = run_batch(source, *([SWITCH_ON] if tracked else []), dict(ROWS_ADD, group="g"), bad, ROWS_LATER,
                       allow_partial=True)
    statuses = {o.op_id: o.status for o in result.operations}
    assert statuses["sec"] == "rolled_back" and statuses["bad"] == "failed" and statuses["later"] == "applied", statuses
    assert texts(source) == before[:-1] + ["Findings"]
    assert "Role D" not in document_xml(source)


def test_batch_rows_sibling_revision_free(tmp_path):
    """X10: a revision-free document, the switch on, author "Reviewer B": a row added and a row removed in one
    replace_section; only Reviewer B's revisions; Reject All = the switched input; Accept All = as requested."""

    source = rows_doc(tmp_path, "Xsib", revised=False)
    on = run_batch(source, SWITCH_ON)
    assert on.status == "complete", on
    switched = tmp_path / "Xsib-switched.docx"
    switched.write_bytes(source.read_bytes())
    d = ["Role D\nLed a team", "2021"]
    result = run_batch(source, {"op_id": "r", "type": "replace_section", "heading": "Experience",
                                "rows": [ROWS[0], ROWS[2], d]})
    assert result.status == "complete", (result.message, [(o.status, o.error) for o in result.operations])
    assert result.tracking == "on" and result.tracking_author == AUTHOR
    xml = document_xml(source)
    assert {m.group(1) for m in re.finditer(r'<w:(?:ins|del)\b[^>]*w:author="([^"]+)"', xml)} == {AUTHOR}
    rejected = tmp_path / "Xsib-rejected.docx"
    revisions.resolve_revisions(source, rejected, mode="reject", author=AUTHOR)
    assert canonical(body(rejected)) == canonical(body(switched))
    assert all_accepted_texts(source, tmp_path) == HEAD + row_lines(ROWS[0], ROWS[2], d) + TAIL
