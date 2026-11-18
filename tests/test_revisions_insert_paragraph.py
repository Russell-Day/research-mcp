"""Plan 864 S2: insert_paragraph as direct XML, tracked (switch on) or plain on a revised document.

Decisions 9 (the new paragraph copies the anchor's pPr minus sectPr/pPrChange/mark revisions, runs take the
anchor mark's rPr unless a style is given, "\\n" makes several paragraphs), 10 (no following paragraph or an
anchor holding a section break: Word's way, the anchor's mark becomes the inserted one and the last new
paragraph takes the anchor's original pPr), 12 (verification) and 19 (style names resolved against
styles.xml). Expected XML is spelled out; Reject is checked with the string engine (`resolve_revisions`),
not with the ElementTree resolver the insert verifies itself with.
"""
from __future__ import annotations

import datetime as dt
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import revisions
from research_mcp.editing import insert_paragraph_to_version, resolve_paragraph_style
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.fixtures import create_mixed_run_fixture, create_tracked_fixture
from research_mcp.sections import canonical, read_layout
from test_revisions import W, document_xml, other_parts_identical, raw_revisions, run_batch, windows_only

NOW = dt.datetime(2026, 11, 16, 12, 0, 0, tzinfo=dt.timezone.utc)
LOCAL = NOW.astimezone().strftime("%Y-%m-%dT%H:%M:%SZ")
STAMP = f'w:date="{LOCAL}" w16du:dateUtc="2026-11-16T12:00:00Z"'
AUTHOR = "Reviewer B"
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
OTHER = 'w:author="Other Author" w:date="2026-10-10T17:37:00Z"'


def ins(rev_id: int, inner: str = "", author: str = AUTHOR, stamp: str = STAMP) -> str:
    attrs = f'w:id="{rev_id}" w:author="{author}" {stamp}'
    return f"<w:ins {attrs}/>" if not inner else f"<w:ins {attrs}>{inner}</w:ins>"


def run(text: str, rpr: str = "") -> str:
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(ppr: str, *content: str) -> str:
    return "<w:p>" + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + "".join(content) + "</w:p>"


STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:latentStyles><w:lsdException w:name="heading 4" w:uiPriority="9"/></w:latentStyles>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="BodyA"><w:name w:val="Body A"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="bodya2"><w:name w:val="body a"/></w:style>'
    '<w:style w:type="character" w:styleId="Heading3Char"><w:name w:val="Heading 3 Char"/></w:style>'
    "</w:styles>"
)


def make(tmp_path: Path, name: str, body: str | None = None, styles: str | None = None) -> Path:
    """The tracked fixture, with its body replaced by `body` and its styles by `styles` when given."""

    base = tmp_path / f"base-{name}.docx"
    create_tracked_fixture(base)
    target = tmp_path / f"{name}.docx"
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(target, "x") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml" and body is not None:
                xml = data.decode("utf-8")
                start = xml.index("<w:body>") + len("<w:body>")
                data = (xml[:start] + body + SECT + "</w:body></w:document>").encode("utf-8")
            elif info.filename == "word/styles.xml" and styles is not None:
                data = styles.encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)
    return target


def insert(source: Path, out: Path, after: int, text: str, *, tracked: bool = True, style_id=None, **kw):
    return revisions.insert_paragraphs(source, out, after_paragraph_index=after, text=text, style_id=style_id,
                                       tracked=tracked, author=AUTHOR, now=NOW, **kw)


def body(path: Path) -> ET.Element:
    return ET.fromstring(document_xml(path)).find(W + "body")


def rejected(source: Path, out: Path, tmp_path: Path, ids) -> Path:
    """Reject the given ids with the string engine (not the resolver the insert verified itself with)."""

    target = tmp_path / f"rejected-{out.stem}.docx"
    revisions.resolve_revisions(out, target, mode="reject", ids=[str(item) for item in ids])
    return target


def texts(path: Path) -> list[str]:
    return [item.text for item in read_layout(path).paragraphs]


def span_of(xml: str, text: str) -> tuple[int, int]:
    """Start and end of the one w:p holding `text`."""

    found = [m for m in re.finditer(r"<w:p>(?:(?!<w:p>).)*?</w:p>", xml, re.S) if text in m.group(0)]
    assert len(found) == 1, text
    return found[0].start(), found[0].end()


# -- Case A: tracked insert after a middle paragraph -------------------------------------------------------

def test_tracked_insert_middle_exact_xml(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    result = insert(src, out, 2, "New sentence.")

    xml = document_xml(src)
    at = span_of(xml, "Prior work ")[0]
    new = para(f"<w:rPr>{ins(31)}</w:rPr>", ins(32, run("New sentence.")))
    assert document_xml(out) == xml[:at] + new + xml[at:]
    assert other_parts_identical(src, out)
    assert all(raw in document_xml(out) for raw in raw_revisions(src))
    before = texts(src)
    assert texts(out) == before[:2] + ["New sentence."] + before[2:]
    assert result["index_delta"] == 1 and result["indices_after_shifted"] is True
    assert result["new_paragraph_index"] == 3 and result["stale_from_index"] == 3
    assert result["engine"] == "direct_xml" and result["tracked"] is True and result["author"] == AUTHOR
    assert [(item["id"], item["type"]) for item in result["revisions_added"]] == [("31", "ins"), ("32", "ins")]
    assert result["anchor_mark_moved"] is False
    back = rejected(src, out, tmp_path, [31, 32])
    assert canonical(body(back)) == canonical(body(src))


def test_tracked_insert_two_paragraphs(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    result = insert(src, out, 2, "First.\nSecond.")

    xml = document_xml(src)
    at = span_of(xml, "Prior work ")[0]
    new = (para(f"<w:rPr>{ins(31)}</w:rPr>", ins(32, run("First.")))
           + para(f"<w:rPr>{ins(33)}</w:rPr>", ins(34, run("Second."))))
    assert document_xml(out) == xml[:at] + new + xml[at:]
    before = texts(src)
    assert texts(out) == before[:2] + ["First.", "Second."] + before[2:]
    assert result["index_delta"] == 2 and result["new_paragraph_indices"] == [3, 4]
    assert result["paragraph_texts"] == ["First.", "Second."]
    back = rejected(src, out, tmp_path, [31, 32, 33, 34])
    assert canonical(body(back)) == canonical(body(src))


def test_new_paragraph_copies_anchor_ppr_and_mark_formatting(tmp_path):
    ppr = (f'<w:pStyle w:val="BodyA"/><w:jc w:val="both"/>'
           f'<w:rPr><w:ins w:id="5" {OTHER}/><w:i/><w:rPrChange w:id="6" {OTHER}><w:rPr/></w:rPrChange></w:rPr>'
           f'<w:pPrChange w:id="7" {OTHER}><w:pPr/></w:pPrChange>')
    src = make(tmp_path, "copy", para(ppr, run("Anchor.", "<w:b/>")) + para("", run("Next.")), STYLES)
    out = tmp_path / "out.docx"
    insert(src, out, 1, "New.")

    xml = document_xml(src)
    at = span_of(xml, "Next.")[0]
    new = para(f'<w:pStyle w:val="BodyA"/><w:jc w:val="both"/><w:rPr>{ins(8)}<w:i/></w:rPr>',
               ins(9, run("New.", "<w:i/>")))
    assert document_xml(out) == xml[:at] + new + xml[at:]
    back = rejected(src, out, tmp_path, [8, 9])
    assert canonical(body(back)) == canonical(body(src))


def test_body_level_marker_stays_before_the_new_paragraph(tmp_path):
    src = make(tmp_path, "marker", para("", '<w:bookmarkStart w:id="40" w:name="Spot"/>', run("Anchor."))
               + '<w:bookmarkEnd w:id="40"/>' + para("", run("Next.")))
    out = tmp_path / "out.docx"
    insert(src, out, 1, "New.")
    xml = document_xml(out)
    assert xml.index('<w:bookmarkEnd w:id="40"/>') < xml.index("New.") < xml.index("Next.")
    back = rejected(src, out, tmp_path, [41, 42])
    assert canonical(body(back)) == canonical(body(src))


# -- Case C: style names ----------------------------------------------------------------------------------

def test_style_resolver_names(tmp_path):
    src = make(tmp_path, "styles", para("", run("Anchor.")), STYLES)
    for name in ("Heading 3", "Heading3", "heading 3", "HEADING3", " heading  3 "):
        ref = resolve_paragraph_style(src, name)
        assert (ref.style_id, ref.name, ref.defined) == ("Heading3", "heading 3", True), name
    assert resolve_paragraph_style(src, "Normal").default is True
    assert resolve_paragraph_style(src, "Body A").style_id == "BodyA"  # exact name beats the loose match
    with pytest.raises(RefusedUnsafe, match="matches several paragraph styles"):
        resolve_paragraph_style(src, "BODY A")
    with pytest.raises(RefusedUnsafe, match=r"not a paragraph style.*near matches: 'heading 3' \(styleId Heading3\)"):
        resolve_paragraph_style(src, "Heading 33")
    with pytest.raises(RefusedUnsafe, match="not a paragraph style"):
        resolve_paragraph_style(src, "Heading 3 Char")  # a character style is not a paragraph style
    latent = resolve_paragraph_style(src, "Heading 4")
    assert (latent.style_id, latent.name, latent.defined) == (None, "heading 4", False)


def test_tracked_insert_heading_style(tmp_path):
    src = make(tmp_path, "heading", para('<w:jc w:val="both"/><w:rPr><w:i/></w:rPr>', run("Anchor.", "<w:i/>"))
               + para("", run("Next.")), STYLES)
    out = tmp_path / "out.docx"
    result = insert_paragraph_to_version(src, out, after_paragraph_index=1, text="Methods", style="heading 3",
                                         tracked=True, author=AUTHOR)

    added = re.search(r'<w:p><w:pPr><w:pStyle w:val="Heading3"/><w:jc w:val="both"/><w:rPr><w:ins w:id="1" '
                      r'w:author="Reviewer B" [^>]*/></w:rPr></w:pPr><w:ins w:id="2" w:author="Reviewer B" '
                      r'[^>]*><w:r><w:t xml:space="preserve">Methods</w:t></w:r></w:ins></w:p>', document_xml(out))
    assert added, document_xml(out)
    assert read_layout(out).para(2).style == "Heading3"
    assert (result["style"], result["style_id"]) == ("heading 3", "Heading3")
    # Plan 973 S3: Heading 4 is now added as Word adds it (tests/test_builtin_styles.py); a latent built-in
    # outside the captured set is still refused.
    src5 = make(tmp_path, "heading5", para("", run("Anchor.")), STYLES.replace("heading 4", "heading 5"))
    with pytest.raises(RefusedUnsafe, match="built-in this document has not used yet"):
        insert_paragraph_to_version(src5, tmp_path / "latent.docx", after_paragraph_index=1, text="X",
                                    style="Heading 5", tracked=True)
    assert not (tmp_path / "latent.docx").exists()


# -- Case D / H: untracked on a revised document ----------------------------------------------------------

def test_untracked_insert_on_revised_document(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    result = insert_paragraph_to_version(src, out, after_paragraph_index=2, text="Plain new.")

    xml = document_xml(src)
    at = span_of(xml, "Prior work ")[0]
    assert document_xml(out) == xml[:at] + para("", run("Plain new.")) + xml[at:]
    assert result["engine"] == "direct_xml" and result["tracked"] is False and result["revisions_added"] == []
    assert result["index_delta"] == 1
    assert revisions.revision_count(out) == revisions.revision_count(src)


@windows_only
def test_tracked_false_override(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    count = revisions.revision_count(src)
    on = run_batch(src, {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR})
    assert on.status == "complete", on
    plain = run_batch(src, {"op_id": "p", "type": "insert_paragraph", "after_paragraph_index": 2,
                            "text": "Plain new.", "tracked": False})
    assert plain.status == "complete", plain
    assert revisions.revision_count(src) == count
    assert plain.operations[0].details["tracking"] == "off"
    tracked = run_batch(src, {"op_id": "t", "type": "insert_paragraph", "after_paragraph_index": 3,
                              "text": "Tracked new."})
    assert tracked.status == "complete", tracked
    details = tracked.operations[0].details
    assert (details["tracking"], details["tracking_author"], details["routed_as"]) == ("on", AUTHOR, "insert_paragraph")
    assert revisions.revision_count(src) == count + 2
    assert texts(src)[2:4] == ["Plain new.", "Tracked new."]
    from research_mcp.publication import document_history

    ops = [entry["op"] for entry in document_history(src)["entries"] if entry["kind"] == "edit"]
    assert ops[0] == "insert_paragraph(tracked) t" and ops[1] == "insert_paragraph p"


# -- Cases E, F, G: Word's way ----------------------------------------------------------------------------

TABLE = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>'
         '<w:tr><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>' + para("", run("Cell.")) + "</w:tc></w:tr></w:tbl>")


def test_tracked_insert_before_table_word_way(tmp_path):
    anchor_ppr = '<w:jc w:val="center"/><w:rPr><w:b/></w:rPr>'
    src = make(tmp_path, "table", para(anchor_ppr, run("Intro.", "<w:b/>")) + TABLE + para("", run("After.")))
    out = tmp_path / "out.docx"
    result = insert(src, out, 1, "New.")

    xml = document_xml(src)
    start, end = span_of(xml, "Intro.")
    expected = (para(f'<w:jc w:val="center"/><w:rPr>{ins(1)}<w:b/></w:rPr>', run("Intro.", "<w:b/>"))
                + para(anchor_ppr, ins(2, run("New.", "<w:b/>"))))
    assert document_xml(out) == xml[:start] + expected + xml[end:]
    assert result["anchor_mark_moved"] is True
    assert texts(out)[:2] == ["Intro.", "New."]
    back = rejected(src, out, tmp_path, [1, 2])
    assert canonical(body(back)) == canonical(body(src))
    accepted = tmp_path / "accepted.docx"
    revisions.resolve_revisions(out, accepted, mode="accept", ids=["1", "2"])
    assert texts(accepted)[:2] == ["Intro.", "New."]
    assert revisions.revision_count(accepted) == 0


def test_tracked_insert_end_of_cell(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    result = insert(src, out, 12, "Cell line.")

    xml, new_xml = document_xml(src), document_xml(out)
    start, end = span_of(xml, "Cell B")
    expected = para(f"<w:rPr>{ins(31)}</w:rPr>", run("Cell B")) + para("", ins(32, run("Cell line.")))
    assert new_xml == xml[:start] + expected + xml[end:]
    before = texts(src)
    assert texts(out) == before[:12] + ["Cell line."] + before[12:]
    assert read_layout(out).para(13).cell == read_layout(out).para(12).cell
    assert read_layout(out).para(14).row_end
    assert result["index_delta"] == 1
    back = rejected(src, out, tmp_path, [31, 32])
    assert canonical(body(back)) == canonical(body(src))


def test_tracked_insert_end_of_body(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    insert(src, out, 14, "Last words.")

    xml, new_xml = document_xml(src), document_xml(out)
    start, end = span_of(xml, "Closing paragraph.")
    expected = para(f"<w:rPr>{ins(31)}</w:rPr>", run("Closing paragraph.")) + para("", ins(32, run("Last words.")))
    assert new_xml == xml[:start] + expected + xml[end:]
    assert texts(out)[-2:] == ["Closing paragraph.", "Last words."]
    back = rejected(src, out, tmp_path, [31, 32])
    assert canonical(body(back)) == canonical(body(src))


def test_tracked_insert_after_section_break(tmp_path):
    section = '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>'
    anchor_ppr = f'<w:ind w:left="720"/><w:rPr><w:i/></w:rPr>{section}'
    src = make(tmp_path, "section", para(anchor_ppr, run("End of part one.")) + para("", run("Part two.")))
    out = tmp_path / "out.docx"
    insert(src, out, 1, "Still part one.")

    xml = document_xml(src)
    start, end = span_of(xml, "End of part one.")
    expected = (para(f'<w:ind w:left="720"/><w:rPr>{ins(1)}<w:i/></w:rPr>', run("End of part one."))
                + para(anchor_ppr, ins(2, run("Still part one.", "<w:i/>"))))
    assert document_xml(out) == xml[:start] + expected + xml[end:]
    back = rejected(src, out, tmp_path, [1, 2])
    assert canonical(body(back)) == canonical(body(src))

    plain = tmp_path / "plain.docx"
    insert(src, plain, 1, "Still part one.", tracked=False)
    expected = (para('<w:ind w:left="720"/><w:rPr><w:i/></w:rPr>', run("End of part one."))
                + para(anchor_ppr, run("Still part one.", "<w:i/>")))
    assert document_xml(plain) == xml[:start] + expected + xml[end:]


def test_word_way_with_style_records_the_property_change(tmp_path):
    src = make(tmp_path, "styled", para('<w:jc w:val="right"/>', run("Intro.")) + TABLE, STYLES)
    out = tmp_path / "out.docx"
    insert(src, out, 1, "Heading text", style_id="Heading3")

    xml = document_xml(src)
    start, end = span_of(xml, "Intro.")
    expected = (para(f'<w:jc w:val="right"/><w:rPr>{ins(1)}</w:rPr>', run("Intro."))
                + para(f'<w:pStyle w:val="Heading3"/><w:jc w:val="right"/><w:pPrChange w:id="2" w:author="Reviewer B" '
                       f'{STAMP}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>',
                       ins(3, run("Heading text"))))
    assert document_xml(out) == xml[:start] + expected + xml[end:]
    back = rejected(src, out, tmp_path, [1, 2, 3])
    assert canonical(body(back)) == canonical(body(src))


# -- Case I: another author's mark revision on the anchor -------------------------------------------------

def test_tracked_insert_after_other_authors_mark(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    insert(src, out, 9, "After the inserted mark.")
    xml = document_xml(src)
    start, end = span_of(xml, "Inserted paragraph mark here.")
    new = para(f"<w:rPr>{ins(31)}</w:rPr>", ins(32, run("After the inserted mark.")))
    assert document_xml(out) == xml[:end] + new + xml[end:]  # the anchor, Other Author's mark included, untouched

    sam = f'<w:ins w:id="5" {OTHER}/>'
    src2 = make(tmp_path, "sam", para(f"<w:rPr>{sam}</w:rPr>", run("Sam's line.")) + TABLE)
    out2 = tmp_path / "out2.docx"
    insert(src2, out2, 1, "Mine.")
    xml2 = document_xml(src2)
    start, end = span_of(xml2, "Sam's line.")
    expected = (para(f"<w:rPr>{ins(6)}</w:rPr>", run("Sam's line."))
                + para(f"<w:rPr>{sam}</w:rPr>", ins(7, run("Mine."))))
    assert document_xml(out2) == xml2[:start] + expected + xml2[end:]  # Sam's mark moves byte-identical
    back = rejected(src2, out2, tmp_path, [6, 7])
    assert canonical(body(back)) == canonical(body(src2))


# -- refusals ---------------------------------------------------------------------------------------------

def test_refusals_write_nothing(tmp_path):
    src = create_tracked_fixture(tmp_path / "tracked.docx")
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="mark is a tracked deletion"):
        insert(src, out, 10, "X")  # paragraph 10's mark is Other Author's w:del id=12
    with pytest.raises(RefusedUnsafe, match="empty line"):
        insert(src, out, 2, "A\n\nB")
    with pytest.raises(RefusedUnsafe, match="row-end marker"):
        insert(src, out, 13, "X")
    with pytest.raises(RefusedUnsafe, match="outside 1..14"):
        insert(src, out, 0, "X")
    styled = make(tmp_path, "pprchange", para(f'<w:jc w:val="right"/><w:pPrChange w:id="9" {OTHER}><w:pPr/></w:pPrChange>',
                                               run("Intro.")) + TABLE, STYLES)
    with pytest.raises(RefusedUnsafe, match="already hold a tracked formatting change"):
        insert(styled, out, 1, "X", style_id="Heading3")
    assert not out.exists()


# -- F1 (review r1): verification is independent of the planner ------------------------------------------

REAL_PLAN = revisions._plan_insert


def faulty_planner(monkeypatch, mutate):
    """Replace the planner with one whose output `mutate(plan, kwargs)` corrupts."""

    def plan(doc, node, pieces, **kw):
        return mutate(REAL_PLAN(doc, node, pieces, **kw), kw)

    monkeypatch.setattr(revisions, "_plan_insert", plan)


def edited(plan, replacement):
    assert replacement != plan.replacement, "the fault did not apply"
    return revisions._InsertPlan(plan.start, plan.end, replacement, plan.added, plan.word_way)


def refused_and_nothing_written(src, out, match, **kw):
    before = src.read_bytes()
    with pytest.raises(VerificationFailed, match=match):
        insert(src, out, 1, kw.pop("text", "New."), **kw)
    assert not out.exists() and src.read_bytes() == before


@pytest.mark.parametrize("tracked", [True, False])
def test_fault_dropped_mark_formatting_refused(tmp_path, monkeypatch, tracked):
    src = make(tmp_path, "italic", para('<w:jc w:val="both"/><w:rPr><w:i/></w:rPr>', run("Anchor.", "<w:i/>"))
               + para("", run("Next.")))
    faulty_planner(monkeypatch, lambda plan, kw: edited(plan, plan.replacement.replace("<w:r><w:rPr><w:i/></w:rPr>",
                                                                                      "<w:r>")))
    refused_and_nothing_written(src, tmp_path / "out.docx", "not formatted like the anchor's paragraph mark",
                                tracked=tracked)


def test_fault_wrong_style_refused(tmp_path, monkeypatch):
    src = make(tmp_path, "style", para('<w:jc w:val="both"/>', run("Anchor.")) + para("", run("Next.")), STYLES)
    faulty_planner(monkeypatch, lambda plan, kw: edited(plan, plan.replacement.replace('w:val="Heading3"',
                                                                                      'w:val="BodyA"')))
    refused_and_nothing_written(src, tmp_path / "out.docx", "paragraph properties are not the anchor's",
                                style_id="Heading3")


@pytest.mark.parametrize("tracked", [True, False])
def test_fault_spliced_one_paragraph_late_refused(tmp_path, monkeypatch, tracked):
    src = make(tmp_path, "late", para("", run("Anchor.")) + para("", run("Next.")) + para("", run("Third.")))

    def late(plan, kw):
        at = kw["following"].end  # after the following paragraph instead of before it
        return revisions._InsertPlan(at, at, plan.replacement, plan.added, plan.word_way)

    faulty_planner(monkeypatch, late)
    refused_and_nothing_written(src, tmp_path / "out.docx", "not right after the anchor", tracked=tracked)


@pytest.mark.parametrize("tracked", [True, False])
def test_fault_section_break_kept_on_the_anchor_refused(tmp_path, monkeypatch, tracked):
    section = '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>'
    src = make(tmp_path, "sect", para(f'<w:ind w:left="720"/>{section}', run("End of part one."))
               + para("", run("Part two.")))

    def keep(plan, kw):
        text = plan.replacement.replace(section, "")
        at = text.index("</w:pPr>")  # the anchor's pPr comes first (Word's way)
        return edited(plan, text[:at] + section + text[at:])

    faulty_planner(monkeypatch, keep)
    out = tmp_path / "out.docx"
    if tracked:  # Reject would merge a paragraph holding a section break: the resolver refuses first
        with pytest.raises((RefusedUnsafe, VerificationFailed)):
            insert(src, out, 1, "Still part one.")
        assert not out.exists()
    else:
        refused_and_nothing_written(src, out, "anchor paragraph's properties", tracked=False, text="Still part one.")


# -- Sibling: a revision-free CV-style document with the switch on ----------------------------------------

@windows_only
def test_cv_heading_sibling(tmp_path):
    cv = create_mixed_run_fixture(tmp_path / "cv.docx")
    on = run_batch(cv, {"op_id": "s", "type": "set_track_changes", "on": True, "author": AUTHOR})
    assert on.status == "complete", on
    heading = texts(cv).index("Projects") + 1
    result = run_batch(cv, {"op_id": "p", "type": "insert_paragraph", "after_paragraph_index": heading,
                            "text": "New project line.", "style": "Normal"})
    assert result.status == "complete", result
    assert result.tracking == "on" and result.tracking_author == AUTHOR
    layout = read_layout(cv)
    assert layout.para(heading + 1).text == "New project line."
    assert layout.para(heading + 1).style in (None, "Normal")
    added = [el for el in ET.fromstring(document_xml(cv)).iter(W + "ins")]
    assert [el.get(W + "author") for el in added] == [AUTHOR, AUTHOR]


# -- F3: a multi-paragraph insert inside an edit_document batch ----------------------------------------------

NEW3 = "First.\nSecond.\nThird."


def batch_ops(*extra: dict) -> list[dict]:
    """Switch on, insert three paragraphs after paragraph 1, then `extra`."""

    return [{"op_id": "on", "type": "set_track_changes", "on": True},
            {"op_id": "p", "type": "insert_paragraph", "after_paragraph_index": 1, "text": NEW3}, *extra]


AFTER = {"op_id": "after", "type": "replace_text", "paragraph_index": 14, "find": "Closing", "replace": "Final"}
BEFORE = {"op_id": "before", "type": "replace_text", "paragraph_index": 1, "find": "opening", "replace": "first"}
BAD = {"op_id": "bad", "type": "replace_text", "paragraph_index": 5, "find": "Not in the document",
       "replace": "x"}


def statuses(result) -> dict[str, str]:
    return {item.op_id: item.status for item in result.operations}


@windows_only
def test_batch_edit_after_the_anchor_is_shifted_and_before_it_is_not(tmp_path):
    """Paragraph 14 ("Closing ...") is paragraph 17 once three paragraphs follow paragraph 1; paragraph 1, the
    anchor, keeps its index. Unshifted, op 'after' would look for "Closing" in a table cell and fail."""

    doc = create_tracked_fixture(tmp_path / "tracked.docx")
    before = texts(doc)
    result = run_batch(doc, *batch_ops(AFTER, BEFORE))
    assert result.status == "complete", result
    assert result.operations[1].details["index_delta"] == 3
    assert texts(doc) == (["Plain first paragraph with no revisions.", "First.", "Second.", "Third."]
                          + before[1:13] + ["Final paragraph."])


@windows_only
def test_batch_allow_partial_insert_group_fails_later_ops_unshifted(tmp_path):
    """The insert shares group 'g' with a failing op: both roll back, and the later edits land where they
    would without the insert (no shift is left behind by the rolled-back insert)."""

    doc = create_tracked_fixture(tmp_path / "tracked.docx")
    before = texts(doc)
    ops = batch_ops({**BAD, "group": "g"}, AFTER, BEFORE)
    ops[1] = {**ops[1], "group": "g"}
    result = run_batch(doc, *ops, allow_partial=True)
    assert statuses(result) == {"on": "applied", "p": "rolled_back", "bad": "failed", "after": "applied",
                                "before": "applied"}, result
    assert texts(doc) == ["Plain first paragraph with no revisions."] + before[1:13] + ["Final paragraph."]


@windows_only
def test_batch_allow_partial_other_group_fails_insert_still_shifts(tmp_path):
    """Another group ('g2': a valid edit of paragraph 1 and a failing op) rolls back; the insert publishes and
    the later edit after the anchor is still shifted by three."""

    doc = create_tracked_fixture(tmp_path / "tracked.docx")
    before = texts(doc)
    ops = batch_ops({**BEFORE, "group": "g2"}, {**BAD, "group": "g2"}, AFTER)
    result = run_batch(doc, *ops, allow_partial=True)
    assert statuses(result) == {"on": "applied", "p": "applied", "before": "rolled_back", "bad": "failed",
                                "after": "applied"}, result
    assert texts(doc) == ([before[0], "First.", "Second.", "Third."] + before[1:13] + ["Final paragraph."])
