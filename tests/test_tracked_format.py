"""Plan 864 S6: bold / italic on one phrase as a tracked formatting change (w:rPrChange) or plain, direct XML.

Synthetic fixtures only. Expectations are read back with this file's own ElementTree readers (Accept All /
Reject All of every revision, characters with their run formatting), never with revisions.py.

Input -> Expected output -> Checked by
--------------------------------------
* "Intro\\vOBJECTIVE. To test." (one run), tracked bold on "OBJECTIVE." -> the run is split at the label's
  edges, the label run gets <w:b/> and one w:rPrChange by "Claude" holding the old rPr; Accept All = the label
  bold, nothing else; Reject All = input -> ``check_tracked``.
* a phrase over 3 runs (plain, italic, underlined) -> 3 rPrChange, ids ascending, each holding its own run's
  old rPr; only the selected characters turn bold -> ``check_tracked`` + ``changes``.
* another author's w:ins (L2) -> formatted, rPrChange by "Claude"; the w:ins keeps its start tag and text.
* the author's own w:ins (L1) -> bold in place, no rPrChange, nothing added; ``styled_in_place`` names it;
  Reject All still = input (the insertion goes); Accept All = bold.
* a run with "Claude"'s own rPrChange (L3), whole run -> rPr updated, the rPrChange byte-identical; part of it
  -> split as Word does (each piece keeps the change, later pieces a copy with a fresh id); another author's
  rPrChange, whole or part -> refused naming id, author and date.
* refused, named, nothing written: a phrase touching a w:del, a w:moveTo, a field result, crossing a field, a
  footnote reference inside, a paragraph in a nested table, text occurring twice, already formatted.
* a phrase in a table cell; inside and across a hyperlink -> styled, hyperlink intact.
* bold=False on bold text -> <w:b w:val="0"/> + rPrChange; italic only -> w:b untouched.
* untracked on a revised document -> new rPr only, no revision added, every other revision byte-identical.
* sibling: a revision-free CV-style document, tracked bold by "Reviewer B" -> rPrChange by "Reviewer B".
"""
from __future__ import annotations

import datetime as dt
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.tracked_format import style_runs

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CLAUDE = "Claude"
OTHER = "Guest User"
DATE = "2026-11-01T09:00:00Z"
NOW = dt.datetime(2026, 11, 17, 12, 0, tzinfo=dt.timezone.utc)
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
SZ = '<w:rPr><w:sz w:val="22"/></w:rPr>'
ITAL = '<w:rPr><w:i/><w:sz w:val="22"/></w:rPr>'
UND = '<w:rPr><w:u w:val="single"/></w:rPr>'
BOLD = '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr>'
OWN_CHANGE = (f'<w:rPr><w:b/><w:rPrChange w:id="40" w:author="{CLAUDE}" w:date="{DATE}"><w:rPr/></w:rPrChange>'
              '</w:rPr>')
OTHER_CHANGE = (f'<w:rPr><w:i/><w:rPrChange w:id="41" w:author="{OTHER}" w:date="{DATE}"><w:rPr/></w:rPrChange>'
                '</w:rPr>')


def r(text: str, rpr: str = SZ) -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def ins(rid: int, author: str, *content: str) -> str:
    return f'<w:ins w:id="{rid}" w:author="{author}" w:date="{DATE}">' + "".join(content) + "</w:ins>"


def dele(rid: int, text: str) -> str:
    return (f'<w:del w:id="{rid}" w:author="{OTHER}" w:date="{DATE}"><w:r>{SZ}'
            f'<w:delText xml:space="preserve">{text}</w:delText></w:r></w:del>')


def p(*content: str) -> str:
    return "<w:p>" + "".join(content) + "</w:p>"


def field(result: str) -> str:
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF x </w:instrText>'
            '</w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r(result)
            + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def cell(*paragraphs: str) -> str:
    return "<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>" + "".join(paragraphs) + "</w:tc>"


def table(*cells: str) -> str:
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>'
            '</w:tblGrid><w:tr>' + "".join(cells) + "</w:tr></w:tbl>")


HYPER = '<w:hyperlink w:anchor="target">'

BODY = [
    ("label", p(f'<w:r>{SZ}<w:t>Intro</w:t><w:br/><w:t xml:space="preserve">OBJECTIVE. To test.</w:t></w:r>')),
    ("three", p(r("alpha "), r("beta ", ITAL), r("gamma delta", UND))),
    ("other_ins", p(r("Existing "), ins(18, OTHER, r("inserted words ", ITAL)), r("text."))),
    ("own_ins", p(r("Lead "), ins(13, CLAUDE, r("OWN LABEL. ")), r("body text."))),
    ("own_change", p(r("Some "), r("bold words ", OWN_CHANGE), r("here."))),
    ("other_change", p(r("More "), r("italic words ", OTHER_CHANGE), r("there."))),
    ("deleted", p(r("Keep "), dele(30, "gone "), r("rest of it."))),
    ("moved", p(r("Start "), f'<w:moveTo w:id="31" w:author="{OTHER}" w:date="{DATE}">{r("moved words ")}</w:moveTo>',
                r("end."))),
    ("field", p(r("See "), field("(1)"), r(" for details."))),
    ("footnote", p(r("Claim"), '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                                '<w:footnoteReference w:id="1"/></w:r>', r(" stands."))),
    ("twice", p(r("echo and echo again."))),
    ("bold_text", p(r("Strong words here.", BOLD))),
    ("hyperlink", p(r("Visit "), HYPER + r("the site here") + "</w:hyperlink>", r(" today."))),
    ("table", table(cell(p(r("Cell label: value"))))),
    ("nested", table(cell(table(cell(p(r("Inner words here")))), p(r(""))))),
    ("tail", p(r("Tail paragraph."))),
]
INDEX: dict[str, int] = {}


def _index() -> None:
    # layout numbering: paragraphs in document order, plus one row-end marker per table row
    number = 0
    for name, xml in BODY:
        paragraphs = len(re.findall(r"<w:p>", xml)) + len(re.findall(r"<w:tr>", xml))
        INDEX[name] = number + 1
        number += paragraphs


_index()


def build(tmp_path: Path, body: str, name: str = "doc.docx") -> Path:
    base = create_tracked_fixture(tmp_path / f"base-{name}")
    target = tmp_path / name
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(target, "x") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                start = xml.index("<w:body>") + len("<w:body>")
                data = (xml[:start] + body + SECT + "</w:body></w:document>").encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)
    return target


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    return build(tmp_path, "".join(xml for _, xml in BODY))


def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def paragraph_xml(path: Path, name: str) -> str:
    """Raw XML of the BODY entry `name` (its top-level paragraph or table)."""

    xml = document_xml(path)
    body = xml[xml.index("<w:body>") + len("<w:body>"):xml.rindex(SECT)]
    blocks, depth, start = [], 0, 0
    for tag in re.finditer(r"<(/?)w:(p|tbl)\b[^>]*?(/?)>", body):
        closing, selfclosing = tag.group(1), tag.group(3)
        if selfclosing:
            if depth == 0:
                blocks.append(tag.group(0))
            continue
        if not closing:
            if depth == 0:
                start = tag.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                blocks.append(body[start:tag.end()])
    return blocks[[key for key, _ in BODY].index(name)]


def entry(path: Path, name: str) -> ET.Element:
    root = ET.fromstring(document_xml(path))
    body = root.find(W + "body")
    return list(body)[[key for key, _ in BODY].index(name)]


# -- independent readers -----------------------------------------------------------------------------------

def resolved(element: ET.Element, mode: str) -> ET.Element:
    """Accept All ("accept") or Reject All ("reject") of every run-level and formatting revision."""

    import copy

    clone = copy.deepcopy(element)
    parents = {child: parent for parent in clone.iter() for child in parent}
    for rpr in list(clone.iter(W + "rPr")):
        change = rpr.find(W + "rPrChange")
        if change is None:
            continue
        if mode == "accept":
            rpr.remove(change)
        else:
            old = change.find(W + "rPr")
            rpr[:] = list(old) if old is not None else []
    for wrapper in [el for el in clone.iter() if el.tag in (W + "ins", W + "del", W + "moveTo", W + "moveFrom")]:
        parent = parents[wrapper]
        at = list(parent).index(wrapper)
        parent.remove(wrapper)
        gone = (wrapper.tag in (W + "ins", W + "moveTo")) == (mode == "reject")
        if gone:
            continue
        for offset, child in enumerate(list(wrapper)):
            for text in child.iter(W + "delText"):
                text.tag = W + "t"
            parent.insert(at + offset, child)
            parents[child] = parent
    return clone


def fmt(run: ET.Element) -> tuple:
    rpr = run.find(W + "rPr")
    if rpr is None:
        return ()
    out = []
    for child in rpr:
        if child.tag == W + "rPrChange":
            continue
        out.append((child.tag.replace(W, "w:"), tuple(sorted((k.replace(W, "w:"), v) for k, v in child.attrib.items()))))
    return tuple(sorted(out))


def chars(element: ET.Element) -> list[tuple[str, tuple]]:
    out = []
    for run in element.iter(W + "r"):
        for part in run:
            if part.tag == W + "t":
                out.extend((ch, fmt(run)) for ch in part.text or "")
            elif part.tag == W + "br":
                out.append(("\v", fmt(run)))
    return out


def text_of(element: ET.Element) -> str:
    return "".join(ch for ch, _ in chars(element))


def bold_of(run_fmt: tuple) -> bool | None:
    for name, attrs in run_fmt:
        if name == "w:b":
            return dict(attrs).get("w:val", "1") not in ("0", "false", "off")
    return None


def changes(element: ET.Element) -> list[ET.Element]:
    return list(element.iter(W + "rPrChange"))


def check_tracked(source: Path, output: Path, name: str, phrase: str, *, bold=None, italic=None):
    """Reject All of the output = Reject All of the input; Accept All = input's Accept All with b/i set on the
    phrase's characters only."""

    before, after = entry(source, name), entry(output, name)
    assert chars(resolved(after, "reject")) == chars(resolved(before, "reject"))
    old, new = chars(resolved(before, "accept")), chars(resolved(after, "accept"))
    text = "".join(ch for ch, _ in old)
    assert "".join(ch for ch, _ in new) == text
    start = text.index(phrase)
    for number, ((ch, old_fmt), (_, new_fmt)) in enumerate(zip(old, new)):
        if start <= number < start + len(phrase):
            want = dict(old_fmt)
            if bold is not None:
                want["w:b"] = () if bold else (("w:val", "0"),)
            if italic is not None:
                want["w:i"] = () if italic else (("w:val", "0"),)
            assert dict(new_fmt) == want, (ch, new_fmt)
        else:
            assert new_fmt == old_fmt, (ch, new_fmt)


def others_identical(source: Path, output: Path, *names: str) -> None:
    """Every BODY entry not in `names` is byte-identical, and so is everything outside the body."""

    a, b = document_xml(source), document_xml(output)
    for key, _ in BODY:
        if key not in names:
            assert paragraph_xml(source, key) == paragraph_xml(output, key), key
    assert a[:a.index("<w:body>")] == b[:b.index("<w:body>")]


def run(source: Path, tmp_path: Path, name: str, phrase: str, *, tracked=True, author=CLAUDE, out="out.docx",
        **kw):
    output = tmp_path / out
    result = style_runs(source, output, paragraph_index=INDEX[name], text=phrase, tracked=tracked, author=author,
                        now=NOW, **kw)
    return output, result


def refused(source: Path, tmp_path: Path, name: str, phrase: str, *needles: str, **kw) -> None:
    output = tmp_path / "refused.docx"
    with pytest.raises(RefusedUnsafe) as caught:
        style_runs(source, output, paragraph_index=INDEX[name], text=phrase, tracked=kw.pop("tracked", True),
                   author=kw.pop("author", CLAUDE), now=NOW, **({"bold": True} | kw))
    for needle in needles:
        assert needle in str(caught.value), str(caught.value)
    assert not output.exists()


# -- tracked -----------------------------------------------------------------------------------------------

def test_label_after_line_break_tracked_bold(doc, tmp_path):
    output, result = run(doc, tmp_path, "label", "OBJECTIVE.", bold=True)
    assert result["engine"] == "direct_xml" and result["tracked"] is True and result["author"] == CLAUDE
    assert result["selected_text"] == "OBJECTIVE." and result["styled_in_place"] == []
    assert [item["type"] for item in result["revisions_added"]] == ["rPrChange"]
    check_tracked(doc, output, "label", "OBJECTIVE.", bold=True)
    after = entry(output, "label")
    runs = after.findall(W + "r")
    assert [text_of(item) for item in runs] == ["Intro\v", "OBJECTIVE.", " To test."]
    found = changes(after)
    assert len(found) == 1 and found[0].get(W + "author") == CLAUDE
    assert ET.tostring(found[0].find(W + "rPr"), encoding="unicode").count("sz") == 1  # the old rPr, complete
    assert re.fullmatch(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ", found[0].get(W + "date"))
    others_identical(doc, output, "label")


def test_phrase_over_three_runs(doc, tmp_path):
    output, result = run(doc, tmp_path, "three", "pha beta gam", bold=True)
    check_tracked(doc, output, "three", "pha beta gam", bold=True)
    found = changes(entry(output, "three"))
    ids = [int(item.get(W + "id")) for item in found]
    assert len(ids) == 3 and ids == sorted(ids) and len(set(ids)) == 3
    olds = [fmt(ET.fromstring(ET.tostring(item))) for item in found]
    assert [sorted(child.tag.replace(W, "") for child in item.find(W + "rPr")) for item in found] == \
        [["sz"], ["i", "sz"], ["u"]], olds
    assert [item["text"] for item in result["revisions_added"]] == ["pha ", "beta ", "gam"]


def test_other_authors_insertion_L2(doc, tmp_path):
    output, result = run(doc, tmp_path, "other_ins", "inserted", bold=True)
    check_tracked(doc, output, "other_ins", "inserted", bold=True)
    after = entry(output, "other_ins")
    wrapper = after.find(W + "ins")
    assert wrapper.get(W + "id") == "18" and wrapper.get(W + "author") == OTHER and wrapper.get(W + "date") == DATE
    assert text_of(wrapper) == "inserted words "
    assert [item.get(W + "author") for item in changes(wrapper)] == [CLAUDE]
    assert result["styled_in_place"] == []


def test_own_insertion_L1_in_place(doc, tmp_path):
    output, result = run(doc, tmp_path, "own_ins", "OWN LABEL.", bold=True)
    after = entry(output, "own_ins")
    assert changes(after) == [] and result["revisions_added"] == []
    assert result["styled_in_place"] == [{"kind": "own_insertion", "id": "13", "author": CLAUDE, "date": DATE,
                                          "text": "OWN LABEL."}]
    wrapper = after.find(W + "ins")
    assert wrapper.get(W + "id") == "13" and text_of(wrapper) == "OWN LABEL. "
    assert [bold_of(fmt(item)) for item in wrapper.findall(W + "r")] == [True, None]
    check_tracked(doc, output, "own_ins", "OWN LABEL.", bold=True)


def test_own_format_change_L3(doc, tmp_path):
    before = paragraph_xml(doc, "own_change")
    old_change = re.search(r"<w:rPrChange .*?</w:rPrChange>", before).group(0)
    output, result = run(doc, tmp_path, "own_change", "bold words ", italic=True)
    after = paragraph_xml(output, "own_change")
    assert re.findall(r"<w:rPrChange .*?</w:rPrChange>", after) == [old_change]
    assert result["revisions_added"] == []
    assert result["styled_in_place"][0]["kind"] == "own_format_change" and result["styled_in_place"][0]["id"] == "40"
    new = entry(output, "own_change")
    runs = new.findall(W + "r")
    assert dict(fmt(runs[1])) == {"w:b": (), "w:i": (), "w:sz": (("w:val", "22"),)} or \
        dict(fmt(runs[1])) == {"w:b": (), "w:i": ()}
    # Reject All restores the state before the own revision; Accept All keeps bold + italic
    assert chars(resolved(new, "reject")) == chars(resolved(entry(doc, "own_change"), "reject"))


@pytest.mark.parametrize("phrase, pieces", [("bold", ["bold", " words "]), ("old wor", ["b", "old wor", "ds "])])
def test_own_format_change_partial_split_L3(doc, tmp_path, phrase, pieces):
    """L3 split, as Word splits it (live-checked): every piece keeps the w:rPrChange with its old rPr; the first
    piece the original (id 40, byte-identical), each later one a copy with a fresh id, ascending; nothing else is
    added; Reject All = input, Accept All = italic on the phrase only."""
    before = paragraph_xml(doc, "own_change")
    old_change = re.search(r"<w:rPrChange .*?</w:rPrChange>", before).group(0)
    output, result = run(doc, tmp_path, "own_change", phrase, italic=True)
    after = paragraph_xml(output, "own_change")
    found = re.findall(r"<w:rPrChange .*?</w:rPrChange>", after)
    assert len(found) == len(pieces) and found[0] == old_change
    ids = [int(re.search(r'w:id="(\d+)"', item).group(1)) for item in found]
    assert ids[0] == 40 and ids == sorted(ids) and len(set(ids)) == len(ids)
    assert all(item == old_change.replace('w:id="40"', f'w:id="{number}"') for item, number in zip(found, ids))
    assert result["revisions_added"] == []
    assert result["styled_in_place"] == [{"kind": "own_format_change", "id": "40", "author": CLAUDE, "date": DATE,
                                          "text": phrase, "split_copies": [str(number) for number in ids[1:]]}]
    styled = [text_of(run_) for run_ in entry(output, "own_change").findall(W + "r") if W + "rPrChange" in
              {child.tag for rpr in run_.findall(W + "rPr") for child in rpr}]
    assert styled == pieces
    check_tracked(doc, output, "own_change", phrase, italic=True)
    others_identical(doc, output, "own_change")


def test_other_authors_format_change_partial_refused(doc, tmp_path):
    refused(doc, tmp_path, "other_change", "italic", "id=41", repr(OTHER), "another author")


def test_other_authors_format_change_refused(doc, tmp_path):
    refused(doc, tmp_path, "other_change", "italic words", "id=41", repr(OTHER), DATE, "another author")


@pytest.mark.parametrize("name, phrase, needles", [
    ("deleted", "Keep rest", ("w:del id=30", "deleted")),
    ("moved", "moved", ("w:moveTo id=31",)),
    ("field", "(1)", ("field",)),
    ("field", "See (1) for", ("crosses a field",)),
    ("footnote", "Claim stands", ("w:footnoteReference",)),
    ("nested", "Inner", ("nested table",)),
    ("twice", "echo", ("occurs 2 times",)),
    ("bold_text", "words", ("already formatted",)),
    ("label", "Intro\vOBJECTIVE. To test.", ("whole-paragraph",)),
])
def test_refusals(doc, tmp_path, name, phrase, needles):
    refused(doc, tmp_path, name, phrase, *needles)


def test_table_cell_and_hyperlink(doc, tmp_path):
    output, _ = run(doc, tmp_path, "table", "Cell label:", bold=True)
    check_tracked(doc, output, "table", "Cell label:", bold=True)
    second, _ = run(output, tmp_path, "hyperlink", "the site", bold=True, out="two.docx")
    check_tracked(output, second, "hyperlink", "the site", bold=True)
    # styling part of the run just given Claude's rPrChange splits it (L3): "the" keeps the change, " site" a copy
    fourth, result = run(second, tmp_path, "hyperlink", "Visit the", italic=True, out="four.docx")
    check_tracked(second, fourth, "hyperlink", "Visit the", italic=True)
    assert len(result["revisions_added"]) == 1 and len(result["styled_in_place"][0]["split_copies"]) == 1
    third, _ = run(doc, tmp_path, "hyperlink", "Visit the", italic=True, out="three.docx")
    check_tracked(doc, third, "hyperlink", "Visit the", italic=True)
    for result in (second, third, fourth):
        link = entry(result, "hyperlink").find(W + "hyperlink")
        assert link.get(W + "anchor") == "target" and text_of(link) == "the site here"


def test_bold_false_and_italic_only(doc, tmp_path):
    output, _ = run(doc, tmp_path, "bold_text", "words", bold=False)
    check_tracked(doc, output, "bold_text", "words", bold=False)
    assert '<w:b w:val="0"/>' in paragraph_xml(output, "bold_text")
    second, _ = run(doc, tmp_path, "bold_text", "Strong", italic=True, out="two.docx")
    check_tracked(doc, second, "bold_text", "Strong", italic=True)
    runs = entry(second, "bold_text").findall(W + "r")
    assert [bold_of(fmt(item)) for item in runs] == [True, True]  # bold untouched


# -- untracked on a revised document -----------------------------------------------------------------------

def test_untracked_on_revised_document(doc, tmp_path):
    output, result = run(doc, tmp_path, "other_ins", "inserted", bold=True, tracked=False)
    assert result["tracked"] is False and result["revisions_added"] == [] and result["author"] is None
    after = entry(output, "other_ins")
    assert changes(after) == []
    old, new = chars(resolved(entry(doc, "other_ins"), "accept")), chars(resolved(after, "accept"))
    assert [bold_of(f) for _, f in new][9:17] == [True] * 8 and new[:9] == old[:9] and new[17:] == old[17:]
    others_identical(doc, output, "other_ins")
    third, _ = run(doc, tmp_path, "three", "beta", italic=False, tracked=False, out="plain.docx")
    assert changes(entry(third, "three")) == []
    assert '<w:i w:val="0"/>' in paragraph_xml(third, "three")
    others_identical(doc, third, "three")


# -- sibling: revision-free CV-style document, another reviewer ---------------------------------------------

def test_revision_free_cv_reviewer_b(tmp_path):
    body = (p(r("WORK EXPERIENCE", BOLD))
            + p(f'<w:r>{SZ}<w:t>Research Assistant, Example Lab</w:t><w:br/><w:t>Jan 2024 - Present</w:t></w:r>')
            + p(r("Built a pipeline for imaging data.")))
    source = build(tmp_path, body, "cv.docx")
    output = tmp_path / "cv-out.docx"
    result = style_runs(source, output, paragraph_index=2, text="Research Assistant", bold=True, tracked=True,
                        author="Reviewer B", now=NOW)
    assert [item["author"] for item in result["revisions_added"]] == ["Reviewer B"]
    root = ET.fromstring(document_xml(output)).find(W + "body")
    found = changes(root)
    assert len(found) == 1 and found[0].get(W + "author") == "Reviewer B"
    second = list(root)[1]
    accepted = chars(resolved(second, "accept"))
    assert [bold_of(f) for _, f in accepted][:18] == [True] * 18
    assert all(bold_of(f) is None for _, f in accepted[18:])
    assert chars(resolved(second, "reject")) == chars(list(ET.fromstring(document_xml(source)).find(W + "body"))[1])


# -- the style_text tool (plan 864 S6 group 2) -------------------------------------------------------------

FREE_BODY = (p(f'<w:r>{SZ}<w:t>Intro</w:t><w:br/><w:t xml:space="preserve">OBJECTIVE. To test.</w:t></w:r>')
             + p(r("Second paragraph.")))


def _no_word(monkeypatch):
    from research_mcp import style_tools

    def started(*args, **kwargs):
        raise AssertionError("Word was started")

    monkeypatch.setattr(style_tools, "run_operation", started)


def _word_called(monkeypatch) -> list:
    """Stand-in for the Word worker: records the call and writes nothing (the Word path is live-tested)."""
    from research_mcp import style_tools

    calls: list = []

    def fake(kind, payload, **kwargs):
        calls.append((kind, payload))
        raise RefusedUnsafe("word path reached")

    monkeypatch.setattr(style_tools, "run_operation", fake)
    return calls


def _sha(path: Path) -> str:
    from research_mcp.audit import file_sha256

    return file_sha256(path)


def test_tool_tracked_in_place_direct_xml_with_undo(tmp_path, monkeypatch):
    from research_mcp import server

    _no_word(monkeypatch)
    source = build(tmp_path, FREE_BODY, "free.docx")
    before = source.read_bytes()
    sha = _sha(source)
    shown = server.style_text(str(source), sha, 1, "OBJECTIVE.", bold=True, tracked=True)
    assert shown["status"] == "preview" and source.read_bytes() == before
    assert (shown["engine"], shown["tracked"], shown["author"]) == ("direct_xml", True, CLAUDE)
    assert (shown["tracking"], shown["tracking_author"]) == ("on", CLAUDE) and shown["revisions_to_add"] == 1
    result = server.style_text(str(source), sha, 1, "OBJECTIVE.", bold=True, tracked=True, author="Reviewer B",
                               preview_only=False)
    assert result["status"] == "complete" and result["engine"] == "direct_xml" and result["author"] == "Reviewer B"
    assert result["undo_token"] and result["sha256_before"] == sha.upper() and result["sha256_after"] == _sha(source)
    found = changes(ET.fromstring(document_xml(source)))
    assert [item.get(W + "author") for item in found] == ["Reviewer B"]
    undone = server.undo_document(str(source))
    assert source.read_bytes() == before, undone


def test_tool_routes_by_switch_and_tracked(tmp_path, monkeypatch):
    """Never set + revision-free: Word. Switch on: direct XML, tracked by the switch's author. Switch on with
    tracked=False: Word. Switch set off: Word. A revised document untracked: direct XML, plain."""
    from research_mcp import style_tools
    from research_mcp.track_setting import write_setting

    free = build(tmp_path, FREE_BODY, "free.docx")
    calls = _word_called(monkeypatch)
    shown = style_tools.style_text(str(free), _sha(free), 1, "OBJECTIVE.", bold=True)
    assert (shown["engine"], shown["tracked"], shown["tracking"]) == ("word", False, "off")
    with pytest.raises(RefusedUnsafe, match="word path reached"):
        style_tools.style_text(str(free), _sha(free), 1, "OBJECTIVE.", bold=True, preview_only=False)
    assert len(calls) == 1 and calls[0][0] == "inline_style"

    on = tmp_path / "on.docx"
    write_setting(free, on, on=True, author="Editor")
    shown = style_tools.style_text(str(on), _sha(on), 1, "OBJECTIVE.", bold=True)
    assert (shown["engine"], shown["tracked"], shown["author"], shown["tracking_author"]) == (
        "direct_xml", True, "Editor", "Editor")
    out = tmp_path / "on-out.docx"
    done = style_tools.style_text(str(on), _sha(on), 1, "OBJECTIVE.", bold=True, preview_only=False, output=str(out))
    assert [item.get(W + "author") for item in changes(ET.fromstring(document_xml(out)))] == ["Editor"]
    assert done["output_sha256"] == _sha(out)
    with pytest.raises(RefusedUnsafe, match="word path reached"):
        style_tools.style_text(str(on), _sha(on), 1, "OBJECTIVE.", bold=True, tracked=False, preview_only=False)
    off = tmp_path / "off.docx"
    write_setting(on, off, on=False)
    assert style_tools.style_text(str(off), _sha(off), 1, "OBJECTIVE.", bold=True)["engine"] == "word"
    assert len(calls) == 2

    revised = build(tmp_path, "".join(xml for _, xml in BODY), "revised.docx")
    _no_word(monkeypatch)
    out = tmp_path / "revised-out.docx"
    done = style_tools.style_text(str(revised), _sha(revised), INDEX["three"], "beta", bold=True,
                                  preview_only=False, output=str(out))
    assert (done["engine"], done["tracked"], done["revisions_added"]) == ("direct_xml", False, [])
    assert changes(entry(out, "three")) == [] and done["tracking"] == "off"
    others_identical(revised, out, "three")


def test_tool_refusals_write_nothing(tmp_path, monkeypatch):
    from research_mcp import style_tools

    _no_word(monkeypatch)
    source = build(tmp_path, "".join(xml for _, xml in BODY), "revised.docx")
    before = source.read_bytes()
    for name, phrase, needle in (("deleted", "Keep rest", "w:del id=30"), ("other_change", "italic", OTHER),
                                 ("twice", "echo", "exactly once")):
        for preview_only in (True, False):
            with pytest.raises(RefusedUnsafe, match=re.escape(needle)):
                style_tools.style_text(str(source), _sha(source), INDEX[name], phrase, bold=True, tracked=True,
                                       preview_only=preview_only)
    assert source.read_bytes() == before


# -- edit_document op and edit feedback (plan 864 S6 group 2) ----------------------------------------------

LABEL_BODY = (p(f'<w:r>{SZ}<w:t>Intro</w:t><w:br/><w:t xml:space="preserve">To test the method.</w:t></w:r>')
              + p(r("Second paragraph.")))


def batch(path: Path, *ops: dict, output: Path | None = None, **kw) -> dict:
    """edit_document as the MCP client calls it (envelope ops), in place unless `output` is given."""
    from research_mcp import server

    envelope = [server.MCPEditOperation(**op) for op in ops]
    return server.edit_document(str(path), _sha(path), envelope, str(output or path), **kw)


def feedback_events(path: Path) -> list[dict]:
    from research_mcp.edit_feedback_word import doc_log

    return doc_log(path).events()


def proposals(events: list[dict]) -> list[dict]:
    return [event for event in events if event.get("event") == "proposal"]


def outcomes(events: list[dict]) -> list[dict]:
    return [event for event in events if event.get("event") == "outcome"]


def test_batch_insert_label_then_bold_one_tracked_insertion(tmp_path, monkeypatch):
    """Plan 875: insert_text "OBJECTIVE. " after the line break, then style_text bold on it, in one tracked
    batch: one w:ins by Claude holding a bold run, no w:rPrChange; Reject All = input; Accept All = the label
    bold and nothing else changed."""
    _no_word(monkeypatch)
    source = build(tmp_path, LABEL_BODY, "label.docx")
    before = tmp_path / "label-before.docx"
    before.write_bytes(source.read_bytes())
    result = batch(source,
                   {"op_id": "label", "type": "insert_text", "paragraph_index": 1, "position": "before",
                    "anchor": "To test", "text": "OBJECTIVE. "},
                   {"op_id": "bold", "type": "style_text", "paragraph_index": 1, "text": "OBJECTIVE.", "bold": True},
                   track_changes=True)
    assert result["status"] == "complete", result
    styled = next(item for item in result["operations"] if item["op_id"] == "bold")
    assert styled["status"] == "applied" and styled["details"]["engine"] == "direct_xml"
    assert styled["details"]["revisions_added"] == [] and styled["details"]["styled_in_place"]
    root = ET.fromstring(document_xml(source)).find(W + "body")
    first, base = list(root)[0], list(ET.fromstring(document_xml(before)).find(W + "body"))[0]
    inserted = first.findall(W + "ins")
    assert len(inserted) == 1 and inserted[0].get(W + "author") == CLAUDE and changes(first) == []
    bold_runs = [run for run in inserted[0].findall(W + "r") if bold_of(fmt(run))]
    assert len(bold_runs) == 1 and text_of(bold_runs[0]) == "OBJECTIVE."
    assert text_of(inserted[0]) == "OBJECTIVE. "
    assert chars(resolved(first, "reject")) == chars(resolved(base, "reject"))
    accepted = chars(resolved(first, "accept"))
    assert "".join(ch for ch, _ in accepted) == "Intro\vOBJECTIVE. To test the method."
    assert [bold_of(f) for _, f in accepted] == [None] * 6 + [True] * 10 + [None] * 20
    # L5: the batch records one proposal (the insertion), pending; the formatting adds none
    events = feedback_events(source)
    assert len(proposals(events)) == 1 and outcomes(events) == []


def test_style_text_records_no_proposal_and_keeps_open_one_pending(tmp_path, monkeypatch):
    """L5: a style_text write (tool or op) records no text proposal, and an open tracked proposal of the same
    paragraph, its own insertion split and bolded (L1) or its neighbour formatted (rPrChange), stays pending."""
    from research_mcp import server

    _no_word(monkeypatch)
    source = build(tmp_path, LABEL_BODY, "label.docx")
    batch(source, {"op_id": "e1", "type": "replace_text", "paragraph_index": 1, "find": "the method",
                   "replace": "the new method"}, track_changes=True)
    events = feedback_events(source)
    assert len(proposals(events)) == 1 and outcomes(events) == []
    server.style_text(str(source), _sha(source), 1, "new", bold=True, preview_only=False)  # L1, in place
    server.style_text(str(source), _sha(source), 1, "Intro", italic=True, preview_only=False)  # rPrChange
    batch(source, {"op_id": "s", "type": "style_text", "paragraph_index": 1, "text": "test", "bold": True})
    events = feedback_events(source)
    assert len(proposals(events)) == 1 and outcomes(events) == [], [e for e in events if e.get("event") != "sync"]
    from research_mcp.edit_feedback import chains

    assert [record["pending"] for record in chains(events)] == [True]


def test_batch_style_text_switch_off_untracked_and_never_set(tmp_path, monkeypatch):
    """On a revised document: switch set off, tracked: false under the switch, and never set are all plain
    direct XML (no revision added, every other entry byte-identical); the switch on tracks by its author."""
    _no_word(monkeypatch)
    source = build(tmp_path, "".join(xml for _, xml in BODY), "revised.docx")
    op = {"op_id": "s", "type": "style_text", "paragraph_index": INDEX["three"], "text": "beta", "bold": True}
    plain = tmp_path / "never.docx"
    result = batch(source, op, output=plain)
    details = result["operations"][0]["details"]
    assert (details["engine"], details["tracked"], details["revisions_added"]) == ("direct_xml", False, [])
    assert changes(entry(plain, "three")) == [] and result["tracking"] == "off"
    others_identical(source, plain, "three")
    off = tmp_path / "off.docx"
    batch(source, op, output=off, track_changes=False)
    assert changes(entry(off, "three")) == []
    others_identical(source, off, "three")
    forced = tmp_path / "forced.docx"
    batch(source, {**op, "tracked": False}, output=forced, track_changes=True)
    assert changes(entry(forced, "three")) == []
    tracked = tmp_path / "on.docx"
    result = batch(source, op, output=tracked, track_changes=True)
    assert [item.get(W + "author") for item in changes(entry(tracked, "three"))] == [CLAUDE]
    check_tracked(source, tracked, "three", "beta", bold=True)
    assert result["operations"][-1]["details"]["tracked"] is True


def test_apply_operation_word_path_on_revision_free_untracked(tmp_path, monkeypatch):
    """Revision-free and untracked: tracked_format.apply_operation hands the op to Word's inline styling."""
    from research_mcp import inline_style
    from research_mcp.edit_models import StyleTextOperation
    from research_mcp.tracked_format import apply_operation

    calls = []

    def fake(source, output, index, text, **kw):
        calls.append((index, text, kw))
        Path(output).write_bytes(Path(source).read_bytes())
        return {"paragraph_index": index}

    monkeypatch.setattr(inline_style, "style_exact_text", fake)
    source = build(tmp_path, LABEL_BODY, "free.docx")
    op = StyleTextOperation(op_id="s", type="style_text", paragraph_index=1, text="Intro", bold=True,
                            tracked=False, author=CLAUDE)
    result = apply_operation(source, tmp_path / "out.docx", op)
    assert calls == [(1, "Intro", {"bold": True, "italic": None, "allow_whole_paragraph": False})]
    assert (result["engine"], result["tracked"], result["revisions_added"]) == ("word", False, [])
    tracked = apply_operation(source, tmp_path / "tracked.docx", op.model_copy(update={"tracked": True}))
    assert tracked["engine"] == "direct_xml" and len(calls) == 1


def test_batch_allow_partial_rolls_back_style_text_with_its_row(tmp_path, monkeypatch):
    _no_word(monkeypatch)
    source = build(tmp_path, "".join(xml for _, xml in BODY), "revised.docx")
    out = tmp_path / "partial.docx"
    result = batch(source,
                   {"op_id": "cell", "type": "style_text", "paragraph_index": INDEX["table"], "text": "Cell label:",
                    "bold": True},
                   {"op_id": "bad", "type": "replace_text", "paragraph_index": INDEX["table"], "find": "missing",
                    "replace": "x"},
                   {"op_id": "other", "type": "style_text", "paragraph_index": INDEX["three"], "text": "beta",
                    "italic": False},
                   output=out, allow_partial=True)
    status = {item["op_id"]: item["status"] for item in result["operations"]}
    assert status == {"cell": "rolled_back", "bad": "failed", "other": "applied"}, result
    assert paragraph_xml(out, "table") == paragraph_xml(source, "table")
    assert '<w:i w:val="0"/>' in paragraph_xml(out, "three")
    others_identical(source, out, "table", "three")


def test_style_text_op_contract_and_validation():
    from research_mcp import server

    required, optional = server._field_contract("style_text")
    assert any(name.startswith("text") for name in required) and "paragraph_index" in required
    with pytest.raises(RefusedUnsafe, match="text and find both given"):
        server._strict_operation_payload(server.MCPEditOperation(op_id="s", type="style_text", paragraph_index=1,
                                                                 text="a", find="a", bold=True))
    op = server._strict_operation_payload(server.MCPEditOperation(op_id="s", type="style_text", paragraph_index=1,
                                                                  find="a", bold=True))
    assert op.text == "a" and op.find is None


def test_format_only_writes_are_named():
    from research_mcp.edit_feedback_word import _format_only

    assert _format_only("style_text", None)
    assert _format_only("edit_document", "set_track_changes track_changes, style_text s")
    assert not _format_only("edit_document", "replace_text(tracked) e1, style_text s")
    assert _format_only("edit_document", "format_paragraph f1")
    assert _format_only("edit_document", "set_track_changes t, format_paragraph f1, style_text s")
    assert not _format_only("edit_document", "format_paragraph f1, replace_text e1")
    assert not _format_only("edit_document", "set_track_changes track_changes")
    assert not _format_only("edit_document", None)


def test_after_publish_drops_proposals_of_format_only_writes(tmp_path):
    """L5 wiring: after_publish records nothing for a format-only write even when the two files differ in text
    (so a missing guard cannot hide behind capture finding no hunk); the same pair under a text op records one."""
    from research_mcp.edit_feedback_word import after_publish

    before = build(tmp_path, p(r("Alpha beta.")), "before.docx")
    for tool, op, expected in (("style_text", None, 0), ("edit_document", "style_text s", 0),
                               ("edit_document", "replace_text e1", 1)):
        target = build(tmp_path, p(r("Alpha gamma.")), f"after-{expected}-{tool}-{bool(op)}.docx")
        result = after_publish(target, before_file=before, sha_before=_sha(before), sha_after=_sha(target),
                               undo_token=None, tool=tool, op=op)
        assert result["feedback_warning"] is None, result
        assert len(proposals(feedback_events(target))) == expected, (tool, op)


# -- review round 1 (plan 864 S6) ----------------------------------------------------------------------------

@pytest.mark.parametrize("phrase", ["bold", "bold words"])
def test_untracked_on_a_format_change_refused(doc, tmp_path, phrase):
    """Untracked styling of text carrying a tracked formatting change (the author's own too, partial or whole run)
    is refused: a split would add revision ids, and rejecting that change would undo the plain formatting."""
    refused(doc, tmp_path, "own_change", phrase, "id=40", "untracked", tracked=False, italic=True, bold=None)


def test_tool_takes_find_as_another_name_for_text(tmp_path, monkeypatch):
    from research_mcp import server

    _no_word(monkeypatch)
    source = build(tmp_path, FREE_BODY, "free.docx")
    before = source.read_bytes()
    sha = _sha(source)
    shown = server.style_text(str(source), sha, 1, find="OBJECTIVE.", bold=True, tracked=True)
    assert shown["status"] == "preview" and shown["selected_text"] == "OBJECTIVE."
    with pytest.raises(RefusedUnsafe, match="not both"):
        server.style_text(str(source), sha, 1, "OBJECTIVE.", find="OBJECTIVE.", bold=True, tracked=True)
    with pytest.raises(RefusedUnsafe, match="give text"):
        server.style_text(str(source), sha, 1, bold=True, tracked=True)
    assert source.read_bytes() == before


def test_long_style_text_op_id_does_not_hide_a_text_op_from_feedback(tmp_path, monkeypatch):
    """The op names edit feedback reads are not cut: a style_text op with a long op_id followed by a tracked text
    op still records the text op's proposal."""
    _no_word(monkeypatch)
    source = build(tmp_path, LABEL_BODY, "label.docx")
    result = batch(source,
                   {"op_id": "s" * 240, "type": "style_text", "paragraph_index": 2, "text": "Second", "bold": True},
                   {"op_id": "e1", "type": "replace_text", "paragraph_index": 1, "find": "the method",
                    "replace": "the new method"},
                   track_changes=True)
    assert result["status"] == "complete", result
    assert len(proposals(feedback_events(source))) == 1
