"""Offline tests for the minimal-diff planner and verifier (plan 328, D8-D12)."""

from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

import pytest

from research_mcp import sections as S
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.fixtures import create_mixed_run_fixture


TITLE = "Delta Park Museum, Department of Collections, Outreach Volunteer"
NEW_TITLE = "Delta Park Museum, Department of Colonial History, Seasonal Outreach Volunteer"


@pytest.fixture
def mixed(tmp_path: Path) -> Path:
    return create_mixed_run_fixture(tmp_path / "mixed.docx")


def _runs(item: S.Para) -> list[tuple[str, str]]:
    runs: list[list[str]] = []
    for ch, fp in zip(item.text, item.fps):
        if runs and runs[-1][1] == fp:
            runs[-1][0] += ch
        else:
            runs.append([ch, fp])
    return [(text, S._short(fp)) for text, fp in runs]


def _rewrite(source: Path, target: Path, old: str, new: str) -> Path:
    """Copy a package, replacing one exact string in document.xml."""

    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    xml = parts["word/document.xml"].decode("utf-8")
    assert xml.count(old) == 1, old
    parts["word/document.xml"] = xml.replace(old, new).encode("utf-8")
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return target


def test_mixed_fixture_has_mixed_runs_caps_headings_and_inline_objects(mixed):
    layout = S.read_layout(mixed)
    assert _runs(layout.para(17)) == [
        ("Delta Park Museum", "[b]"),
        (", Department of Collections, ", "[plain]"),
        ("Outreach Volunteer", "[i]"),
    ]
    heading = layout.para(10)
    assert heading.text == "WorK Experience" and heading.displayed == "WORK EXPERIENCE"
    contexts = {fmt for _, fmt in _runs(layout.para(28))}
    assert {"[plain in field]", "[plain in content control]", "[u=single in hyperlink]"} <= contexts


def test_whole_title_replace_is_a_minimal_character_diff(mixed):
    layout = S.read_layout(mixed)
    plan = S.plan_replace_text(layout, "replace_text", 17, TITLE, NEW_TITLE)
    (step,) = plan.steps
    assert [edit.describe(TITLE) for edit in step.edits] == [
        "replace 'lections' with 'onial History' at characters 36-44",
        "insert 'Seasonal ' at character 46",
    ]
    item = layout.para(17)
    text, fps, _ = S.apply_edits(item.text, item.fps, item.caps, step.edits)
    assert text == NEW_TITLE
    # Unchanged characters keep their own formatting.
    assert fps[: len("Delta Park Museum")] == list(item.fps[: len("Delta Park Museum")])
    assert all(fp == item.fps[-1] for fp in fps[-len("Outreach Volunteer"):])
    # A replacement takes the first replaced character's formatting (plain here).
    replaced = NEW_TITLE.index("onial History")
    assert set(fps[replaced: replaced + len("onial History")]) == {item.fps[36]}
    # A word inserted as a prefix of the next word takes that word's formatting (italic).
    inserted = NEW_TITLE.index("Seasonal ")
    assert set(fps[inserted: inserted + len("Seasonal ")]) == {item.fps[46]}
    assert S._short(item.fps[46]) == "[i]"


def test_insertion_takes_previous_character_format_and_paragraph_start_takes_next():
    old = "abc def"
    fps = ["|A"] * 3 + ["|B"] * 4
    caps = [False] * len(old)
    (tail,) = S.plan_text_diff(old, "abc defX", caps, fps)
    assert (tail.mode, tail.source) == ("after", 6)
    (head,) = S.plan_text_diff(old, "Xabc def", caps, fps)
    assert (head.mode, head.source) == ("before", 0)
    (middle,) = S.plan_text_diff(old, "abcY def", caps, fps)
    assert (middle.mode, middle.source) == ("after", 2)


def test_deletion_and_identity_plans():
    old = "keep this text"
    caps = [False] * len(old)
    assert S.plan_text_diff(old, old, caps) == []
    (edit,) = S.plan_text_diff(old, "keep text", caps)
    assert (edit.mode, edit.describe(old)) == ("delete", "delete 'this ' at characters 5-10")


def test_caps_relaxes_case_only_inside_caps_runs():
    stored = "WorK Experience"
    caps = [True] * len(stored)
    assert S.caps_equal(stored, caps, "WORK EXPERIENCE")
    assert S.plan_text_diff(stored, "WORK EXPERIENCE", caps) == []
    (edit,) = S.plan_text_diff(stored, "WORK HISTORY", caps)
    assert edit.describe(stored) == "replace 'Experience' with 'HISTORY' at characters 5-15"
    plain = [False] * len(stored)
    assert not S.caps_equal(stored, plain, "WORK EXPERIENCE")
    assert S.caps_find("Hello hello", [False] * 11, "hello") == [6]
    assert S.caps_find("Hello hello", [True] * 5 + [False] * 6, "hello") == [0, 6]


def test_heading_lookup_is_caps_aware_and_does_not_loosen_plain_text(mixed):
    layout = S.read_layout(mixed)
    assert S.find_heading(layout, "WORK EXPERIENCE").index == 10
    assert S.find_heading(layout, "work experience").index == 10
    with pytest.raises(RefusedUnsafe, match="not found"):
        S.find_heading(layout, "JORDAN EXAMPLE")
    with pytest.raises(RefusedUnsafe, match="must equal one cell"):
        S.find_row(layout, TITLE.upper())
    assert S.find_row(layout, TITLE).row == 4


def test_section_locates_rows_up_to_the_next_heading(mixed):
    layout = S.read_layout(mixed)
    section = S.locate_section(layout, "WORK EXPERIENCE")
    assert section.kind == "rows"
    assert [row.row for row in section.rows] == [3, 4]
    assert section.end_heading.index == 22


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("03/2020 – Present ", True),
        ("06/2024 – 08/2024", True),
        ("May 2024", True),
        ("Summer 2023 - Fall 2024", True),
        ("2019", True),
        ("Expected May 2027", True),
        ("06/2024 – 08/2024\n05/2023 – 08/2023", True),
        ("Ongoing work", False),
        ("Outreach Volunteer", False),
        ("", False),
    ],
)
def test_date_grammar(text, expected):
    assert S.is_date_like(text) is expected


def test_table_row_plan_refuses_bad_dates_and_cell_count_changes(mixed):
    layout = S.read_layout(mixed)
    with pytest.raises(RefusedUnsafe, match="does not read as a date"):
        S.plan_table_row(layout, "replace_table_row", TITLE, [TITLE, "Ongoing work"])
    with pytest.raises(RefusedUnsafe, match="cell"):
        S.plan_table_row(layout, "replace_table_row", TITLE, [TITLE])
    with pytest.raises(RefusedUnsafe, match="cell"):
        S.plan_table_row(layout, "replace_table_row", TITLE, [TITLE, "", "03/2020"])


def test_table_row_plan_keeps_unchanged_lines_and_reports_counts(mixed):
    layout = S.read_layout(mixed)
    plan = S.plan_table_row(layout, "replace_table_row", TITLE, [
        NEW_TITLE + "\nContributed to collections review, donor records, cataloging\n"
        "Focused on: Visitor Programs, Place-Based Exhibition\nAdded line",
        "03/2020 – 08/2024",
    ])
    assert plan.paragraph_delta == 1
    (unit,) = plan.expected
    first_cell = unit.lines[0]
    assert first_cell[1].requested is None and first_cell[2].requested is None  # untouched lines: canonical equality
    assert first_cell[0].requested == NEW_TITLE and first_cell[3].requested == "Added line"


def test_section_plan_duplicates_a_template_row_for_new_entries(mixed):
    layout = S.read_layout(mixed)
    rows = [
        ["Lakeside University, College of Arts, Rivera Lab (Soil), Teaching Assistant\n"
         "Field research on soil microbes and crop yield", "01/2024 – Present "],
        [NEW_TITLE + "\nContributed to collections review", "03/2020 – Present "],
        ["New Org, Intern\nDid things", "06/2024 – 08/2024"],
    ]
    plan = S.plan_section(layout, "replace_section", "WORK EXPERIENCE", rows=rows)
    assert sum(isinstance(step, S.DuplicateRowStep) for step in plan.steps) == 1
    assert plan.report["rows_after"] == rows
    with pytest.raises(RefusedUnsafe, match="does not read as a date"):
        S.plan_section(layout, "replace_section", "WORK EXPERIENCE", rows=[rows[0], [NEW_TITLE, "soon"]])


def test_canonical_ignores_word_save_noise_only():
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    a = ET.fromstring(
        f'<w:p xmlns:w="{w}" w:rsidR="00AA"><w:r w:rsidR="1"><w:rPr><w:b/></w:rPr><w:t>He</w:t></w:r>'
        f'<w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">llo</w:t></w:r>'
        f'<w:sdt><w:sdtPr/><w:sdtEndPr/><w:sdtContent/></w:sdt></w:p>')
    b = ET.fromstring(
        f'<w:p xmlns:w="{w}"><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r>'
        f'<w:sdt><w:sdtPr/><w:sdtContent/></w:sdt></w:p>')
    assert S.canonical(a) == S.canonical(b)
    c = ET.fromstring(f'<w:p xmlns:w="{w}"><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r>'
                      f'<w:sdt><w:sdtPr/><w:sdtContent/></w:sdt></w:p>')
    assert S.canonical(b) != S.canonical(c)


def test_canonical_ignores_table_style_cache():
    """Word drops w:cnfStyle on save (JBJS Table 1); it formats nothing."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    a = ET.fromstring(
        f'<w:tcPr xmlns:w="{w}"><w:cnfStyle w:val="001000000000" w:firstColumn="1"/>'
        f'<w:tcW w:w="1342" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>')
    b = ET.fromstring(f'<w:tcPr xmlns:w="{w}"><w:tcW w:w="1342" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>')
    assert S.canonical(a) == S.canonical(b)
    c = ET.fromstring(f'<w:tcPr xmlns:w="{w}"><w:tcW w:w="1342" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>')
    assert S.canonical(b) != S.canonical(c)


def test_verify_plan_catches_changes_outside_the_edit(mixed, tmp_path):
    layout = S.read_layout(mixed)
    plan = S.plan_replace_text(layout, "replace_text", 17, "Collections", "Collections")
    assert S.verify_plan(mixed, mixed, plan)["consistency_check"] == "passed"
    after = _rewrite(mixed, tmp_path / "after.docx", "Plain closing paragraph.", "Plain closing paragraph!")
    with pytest.raises(VerificationFailed, match="content after the edit changed"):
        S.verify_plan(mixed, after, plan)
    before = _rewrite(mixed, tmp_path / "before.docx", "Jordan Example", "Jordan Exemplar")
    with pytest.raises(VerificationFailed, match="content before the edit changed"):
        S.verify_plan(mixed, before, plan)


def test_verify_plan_catches_formatting_lost_on_unchanged_characters(mixed, tmp_path):
    layout = S.read_layout(mixed)
    plan = S.plan_replace_text(layout, "replace_text", 17, TITLE, NEW_TITLE)
    item = layout.para(17)
    (step,) = plan.steps
    text, _, _ = S.apply_edits(item.text, item.fps, item.caps, step.edits)
    assert text == NEW_TITLE
    # A candidate whose text is right but whose bold run lost bold must fail.
    xml = zipfile.ZipFile(mixed).read("word/document.xml").decode("utf-8")
    start = xml.index("Delta Park Museum</w:t>")
    run_start = xml.rindex("<w:r>", 0, start)
    unbolded = xml[:run_start] + xml[run_start:start].replace("<w:b/>", "") + xml[start:]
    unbolded = unbolded.replace("Collections, </w:t>", "Colonial History, Seasonal </w:t>", 1)
    candidate = tmp_path / "unbolded.docx"
    with zipfile.ZipFile(mixed) as source, zipfile.ZipFile(candidate, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            target.writestr(name, unbolded if name == "word/document.xml" else source.read(name))
    with pytest.raises(VerificationFailed, match=r"character 0 \('D'\) has formatting \[plain\] instead of \[b\]"):
        S.verify_plan(mixed, candidate, plan)


def test_consistency_check_flags_broken_dates(mixed, tmp_path):
    assert S.consistency_errors(mixed, mixed) == []
    broken = _rewrite(mixed, tmp_path / "broken.docx", "03/2020 – Present ", "Ongoing work")
    (problem,) = S.consistency_errors(mixed, broken)
    assert "table 1 row 5 cell 2" in problem and "Ongoing work" in problem


def test_changed_span_keeps_unchanged_differently_formatted_characters():
    """Review r1 finding 1: one changed word holding several formats is diffed by character."""

    old, new = "ABCDEF", "AXCDYF"
    fps = ("|",) * 2 + ("|<b />",) * 2 + ("|",) * 2
    edits = S.plan_text_diff(old, new, [False] * 6, fps)
    assert [edit.describe(old) for edit in edits] == [
        "replace 'B' with 'X' at characters 1-2", "replace 'E' with 'Y' at characters 4-5"]
    text, got, _ = S.apply_edits(old, fps, [False] * 6, edits)
    assert text == new and got[2:4] == ["|<b />"] * 2
    # Several changes around bold and italic stretches of one token.
    old = "ABCDEFGH"
    fps = ("|",) * 2 + ("|<b />",) * 2 + ("|<i />",) * 2 + ("|",) * 2
    edits = S.plan_text_diff(old, "AxCDyEFzH", [False] * 8, fps)
    text, got, _ = S.apply_edits(old, fps, [False] * 8, edits)
    assert text == "AxCDyEFzH"
    assert [fp.partition("|")[2] or "plain" for fp in got] == [
        "plain", "plain", "<b />", "<b />", "<b />", "<i />", "<i />", "plain", "plain"]
    # A uniformly formatted word is still replaced in one step.
    assert len(S.plan_text_diff("Health", "Clinic", [False] * 6, ("|<b />",) * 6)) == 1


def _inner_fixture(mixed: Path, target: Path, runs: str) -> Path:
    from research_mcp.fixtures import _para, _run

    return _rewrite(mixed, target, _para(_run("Plain closing paragraph.")), runs)


def test_verify_checks_unchanged_characters_independently_of_the_plan(mixed, tmp_path, monkeypatch):
    from research_mcp.fixtures import _para, _run

    source = _inner_fixture(mixed, tmp_path / "inner.docx", _para(_run("AB"), _run("CD", "b"), _run("EF", "i")))
    index = 29
    assert S.read_layout(source).para(index).text == "ABCDEF"
    flattened = _rewrite(source, tmp_path / "flat.docx", _para(_run("AB"), _run("CD", "b"), _run("EF", "i")),
                         _para(_run("AXCDYF")))
    good = _rewrite(source, tmp_path / "good.docx", _para(_run("AB"), _run("CD", "b"), _run("EF", "i")),
                    _para(_run("AX"), _run("CD", "b"), _run("YF", "i")))
    plan = S.plan_paragraph(S.read_layout(source), "replace_paragraph", index, "AXCDYF")
    assert S.verify_plan(source, good, plan)["consistency_check"] == "passed"
    with pytest.raises(VerificationFailed, match=r"character 2 \('C'\) has formatting \[plain\] instead of \[b\]"):
        S.verify_plan(source, flattened, plan)
    # The r1 planner (one replace over the mixed span) expected a flat result; the
    # plan-independent check still refuses it.
    monkeypatch.setattr(S, "_refine", lambda shown, new, a, b, c, d: [(a, b, new[c:d])])
    legacy = S.plan_paragraph(S.read_layout(source), "replace_paragraph", index, "AXCDYF")
    assert len(legacy.steps[0].edits) == 1
    with pytest.raises(VerificationFailed, match=r"2 unchanged character\(s\) lost their formatting; e\.g\. "
                                                 r"character 2 \('C'\) was \[b\] and is now \[plain\]"):
        S.verify_plan(source, _rewrite(source, tmp_path / "legacy.docx",
                                       _para(_run("AB"), _run("CD", "b"), _run("EF", "i")),
                                       _para(_run("AXCDY"), _run("F", "i"))), legacy)


def test_unchanged_characters_are_unchanged_words_and_kept_letters_of_a_rewritten_word():
    assert S._kept_pairs("ABCDEF", "AXCDYF") == [(0, 0), (2, 2), (3, 3), (5, 5)]
    assert S._kept_pairs("Health Dept", "Clinic Dept")[-4:] == [(7, 7), (8, 8), (9, 9), (10, 10)]
    # A phrase rewritten into a different number of words shares no "unchanged" letters.
    assert S._kept_pairs("Delta Park Museum", "Org") == []
    # Word for word: the shared space, and the 'e' of Delta -> New, are unchanged.
    assert S._kept_pairs("Delta Park Museum", "New Org") == [(1, 1), (5, 3)]


# Bookmark and comment extents ------------------------------------------------

CLOSING = '<w:t xml:space="preserve">Plain closing paragraph.</w:t>'


def _anchored(mixed: Path, target: Path, kind: str, text: str = "Plain closing paragraph.",
              anchored: str = "closing") -> Path:
    """The fixture with `anchored` (inside the closing paragraph, spelled as `text`) under a bookmark, comment or both."""

    head, tail = text.split(anchored, 1)
    opens, closes = "", ""
    if kind in ("bookmark", "both"):
        opens += '<w:bookmarkStart w:id="7" w:name="KeepClosing"/>'
        closes += '<w:bookmarkEnd w:id="7"/>'
    if kind in ("comment", "both"):
        opens += '<w:commentRangeStart w:id="3"/>'
        closes += '<w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>'
    run = lambda value: f'<w:r><w:t xml:space="preserve">{value}</w:t></w:r>' if value else ""
    new = (f'<w:t xml:space="preserve">{head}</w:t></w:r>{opens}{run(anchored)}{closes}{run(tail)}'
           f'<w:r><w:t xml:space="preserve"></w:t>')
    return _rewrite(mixed, target, CLOSING, new)


def _closing_index(layout: S.Layout) -> int:
    return next(item.index for item in layout.paragraphs if item.text == "Plain closing paragraph.")


def _anchor_texts(path: Path) -> tuple[dict, list]:
    layout = S.read_layout(path)
    text = S.stream_text(layout)
    found = S.anchor_extents(layout)
    return ({name: text[start:end] for name, (start, end) in found["bookmarks"].items()},
            [text[start:end] for _, start, end in found["comments"]])


def test_map_anchor_point_follows_word_typing():
    insert = S.StreamEdit(10, 0, 3)
    assert S.map_anchor_point(9, "before", insert) == (9, 9)
    assert S.map_anchor_point(11, "after", insert) == (14, 14)
    assert S.map_anchor_point(10, "before", insert) == (10, 10)  # bookmark start takes it; an end does not
    assert S.map_anchor_point(10, "after", insert) == (13, 13)  # a comment start does not take it
    replace = S.StreamEdit(10, 2, 5)
    assert S.map_anchor_point(10, "after", replace) == (10, 10)
    assert S.map_anchor_point(12, "before", replace) == (15, 15)
    assert S.map_anchor_point(11, "before", replace) == (10, 15)
    assert S.map_anchor_point(10, "before", S.StreamEdit(10, 0, 4, either_side=True)) == (10, 14)


def test_expected_extent_never_drops_unchanged_anchored_text():
    # "closing" at 6..13; "!" typed at its end stays outside; typed at its start
    # joins a bookmark but not a comment; a replaced last letter stays inside.
    end = [S.StreamEdit(13, 0, 1)]
    assert S._expect(6, 13, "before", end) == S._Expected((6, 6), (13, 13))
    start = [S.StreamEdit(6, 0, 1)]
    assert S._expect(6, 13, "before", start) == S._Expected((6, 6), (14, 14))
    assert S._expect(6, 13, "after", start) == S._Expected((7, 7), (14, 14))
    last = [S.StreamEdit(12, 1, 1)]
    assert S._expect(6, 13, "after", last) == S._Expected((6, 6), (13, 13))
    assert S._expect(6, 13, "before", [S.StreamEdit(6, 7, 0)]).may_vanish
    assert not S._expect(6, 13, "before", [S.StreamEdit(6, 7, 3)]).may_vanish
    assert S._expect(13, 13, "after", end) == S._Expected((13, 13), (13, 13))  # collapsed stays before


def test_anchor_extents_measure_bookmarks_and_comments(mixed, tmp_path):
    both = _anchored(mixed, tmp_path / "both.docx", "both")
    assert _anchor_texts(both) == ({"KeepClosing": "closing"}, ["closing"])
    assert _anchor_texts(mixed) == ({}, [])


@pytest.mark.parametrize("kind", ["bookmark", "comment", "both"])
def test_verify_plan_refuses_a_shrunk_anchor_and_accepts_word_typing(mixed, tmp_path, kind):
    source = _anchored(mixed, tmp_path / "source.docx", kind)
    layout = S.read_layout(source)
    plan = S.plan_paragraph(layout, "replace_paragraph", _closing_index(layout), "Plain closing! paragraph.")
    good = _anchored(mixed, tmp_path / "good.docx", kind, "Plain closing! paragraph.")
    report = S.verify_plan(source, good, plan)
    assert report["consistency_check"] == "passed"
    assert report["anchors_verified"] == (2 if kind == "both" else 1)
    shrunk = _anchored(mixed, tmp_path / "shrunk.docx", kind, "Plain closing! paragraph.", "closin")
    with pytest.raises(VerificationFailed, match="would cover 'closin' instead of 'closing'"):
        S.verify_plan(source, shrunk, plan)
    grown = _anchored(mixed, tmp_path / "grown.docx", kind, "Plain closing! paragraph.", "closing!")
    with pytest.raises(VerificationFailed, match="would cover 'closing!' instead of 'closing'"):
        S.verify_plan(source, grown, plan)


def test_verify_plan_applies_the_boundary_rule_at_an_anchor_start(mixed, tmp_path):
    source = _anchored(mixed, tmp_path / "source.docx", "both")
    layout = S.read_layout(source)
    plan = S.plan_paragraph(layout, "replace_paragraph", _closing_index(layout), "Plain !closing paragraph.")
    # Word: the bookmark takes "!" typed at its start, the comment does not.
    text = "Plain !closing paragraph."
    candidate = _rewrite(_anchored(mixed, tmp_path / "c.docx", "comment", text),
                         tmp_path / "candidate.docx",
                         '<w:commentRangeStart w:id="3"/>',
                         '<w:bookmarkStart w:id="7" w:name="KeepClosing"/><w:r><w:t>!</w:t></w:r>'
                         '<w:commentRangeStart w:id="3"/>')
    candidate = _rewrite(candidate, tmp_path / "candidate2.docx",
                         '<w:t xml:space="preserve">Plain !</w:t>', '<w:t xml:space="preserve">Plain </w:t>')
    candidate = _rewrite(candidate, tmp_path / "candidate3.docx",
                         '<w:commentReference w:id="3"/></w:r>', '<w:commentReference w:id="3"/></w:r>'
                         '<w:bookmarkEnd w:id="7"/>')
    assert _anchor_texts(candidate) == ({"KeepClosing": "!closing"}, ["closing"])
    assert S.verify_plan(source, candidate, plan)["anchors_verified"] == 2
    unmoved = _anchored(mixed, tmp_path / "unmoved.docx", "both", text)
    with pytest.raises(VerificationFailed, match="bookmark 'KeepClosing' would cover 'closing' instead of '!closing'"):
        S.verify_plan(source, unmoved, plan)


def test_verify_plan_keeps_a_replaced_last_letter_inside_and_refuses_a_deleted_anchor(mixed, tmp_path):
    source = _anchored(mixed, tmp_path / "source.docx", "both")
    layout = S.read_layout(source)
    index = _closing_index(layout)
    plan = S.plan_paragraph(layout, "replace_paragraph", index, "Plain closink paragraph.")
    good = _anchored(mixed, tmp_path / "good.docx", "both", "Plain closink paragraph.", "closink")
    assert S.verify_plan(source, good, plan)["consistency_check"] == "passed"
    short = _anchored(mixed, tmp_path / "short.docx", "both", "Plain closink paragraph.", "closin")
    with pytest.raises(VerificationFailed, match="would cover 'closin' instead of 'closink'"):
        S.verify_plan(source, short, plan)
    gone_plan = S.plan_paragraph(layout, "replace_paragraph", index, "Plain  paragraph.")
    gone = _rewrite(mixed, tmp_path / "gone.docx", CLOSING, '<w:t xml:space="preserve">Plain  paragraph.</w:t>')
    with pytest.raises(VerificationFailed, match="bookmark 'KeepClosing' on 'closing' would be deleted"):
        S.verify_plan(source, gone, gone_plan)
    lost = _rewrite(mixed, tmp_path / "lost.docx", CLOSING, '<w:t xml:space="preserve">Plain closing! paragraph.</w:t>')
    lost_plan = S.plan_paragraph(layout, "replace_paragraph", index, "Plain closing! paragraph.")
    with pytest.raises(VerificationFailed, match="bookmark 'KeepClosing' on 'closing' was lost"):
        S.verify_plan(source, lost, lost_plan)
