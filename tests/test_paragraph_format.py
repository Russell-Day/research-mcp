"""Plan 916 S3: format_paragraph (font, size, indents, spacing, numbering; tracked or plain, direct XML) and the
format_note on a filled blank paragraph.

Synthetic fixtures only. Expectations are read back with this file's own ElementTree readers (Accept All /
Reject All of the formatting revisions, raw pPr/rPr children), never with paragraph_format's own resolver.

Input -> Expected output -> Checked by
--------------------------------------
* replace_paragraph fills an empty numbered paragraph whose mark is 6 pt -> the op result's format_note names
  "numbered", "6 pt" and format_paragraph(paragraph_index=N, like_paragraph=M); nothing else changes ->
  ``test_format_note_on_numbered_small_blank_fill`` (no note for a matching blank or a non-empty paragraph).
* like_paragraph=body + remove_numbering on a style-numbered list paragraph -> numId 0, the body's spacing,
  every run and the mark in the body's font and size, text unchanged, no revision ->
  ``test_like_paragraph_removes_numbering_and_matches``; a direct numPr is removed outright.
* explicit size/left_indent with like_paragraph -> the explicit values win -> ``test_explicit_fields_override_like``.
* left 36 pt, first line -18 pt -> <w:ind w:left="720" w:hanging="360"/> -> ``test_hanging_indent``.
* a range of three paragraphs -> every run and mark Arial 11 pt; other paragraphs byte-identical ->
  ``test_range_formats_every_paragraph``.
* tracked (switch on) -> one w:pPrChange per paragraph holding the old pPr, one w:rPrChange per run and on the
  mark, ids ascending; Reject All (this file's reader and MCP reject_revisions) = input; Accept All = requested
  -> ``test_tracked_switch_writes_ppr_and_rpr_changes``, ``test_tracked_reject_restores_input``.
* another author's w:ins / w:del -> runs in the w:ins formatted (+ rPrChange tracked), the w:del and its run
  byte-identical, the w:ins start tag unchanged -> ``test_revised_document_keeps_other_revisions``; field-result
  runs keep their rPr -> ``test_deleted_and_field_runs_keep_rpr``.
* own w:ins (L1) -> formatted in place, no rPrChange; own rPrChange (L3) kept byte-identical.
* refused, named, nothing written -> ``test_refusals_*``; bold/italic refused with a pointer to style_text.
* edit feedback: no proposal; an open one stays pending -> ``test_edit_feedback_stays_pending``.
"""
from __future__ import annotations

import copy
import datetime as dt
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.paragraph_format import blank_fill_note, format_paragraphs

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CLAUDE = "Claude"
OTHER = "Reviewer B"
DATE = "2026-11-01T09:00:00Z"
NOW = dt.datetime(2026, 11, 17, 12, 0, tzinfo=dt.timezone.utc)
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
SZ24 = '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
SZ20 = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>'
ITAL = '<w:rPr><w:i/><w:sz w:val="20"/></w:rPr>'
BODY_PPR = '<w:pPr><w:spacing w:after="120"/></w:pPr>'

NUMBERING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
    '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    '</w:numbering>')
LIST_STYLE = ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="NumberedItem"><w:name w:val="Numbered Item"/>'
              '<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>')


def r(text: str, rpr: str = SZ24) -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def ins(rid: int, author: str, *content: str) -> str:
    return f'<w:ins w:id="{rid}" w:author="{author}" w:date="{DATE}">' + "".join(content) + "</w:ins>"


def dele(rid: int, text: str) -> str:
    return (f'<w:del w:id="{rid}" w:author="{OTHER}" w:date="{DATE}"><w:r>{SZ20}'
            f'<w:delText xml:space="preserve">{text}</w:delText></w:r></w:del>')


def p(*content: str, ppr: str = "") -> str:
    return "<w:p>" + ppr + "".join(content) + "</w:p>"


def field(result: str) -> str:
    return (f'<w:r>{SZ20}<w:fldChar w:fldCharType="begin"/></w:r><w:r>{SZ20}<w:instrText xml:space="preserve"> '
            f'REF x </w:instrText></w:r><w:r>{SZ20}<w:fldChar w:fldCharType="separate"/></w:r>' + r(result, SZ20)
            + f'<w:r>{SZ20}<w:fldChar w:fldCharType="end"/></w:r>')


def cell(*paragraphs: str) -> str:
    return '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + "".join(paragraphs) + "</w:tc>"


def table(*cells: str) -> str:
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>'
            '</w:tblGrid><w:tr>' + "".join(cells) + "</w:tr></w:tbl>")


OTHER_PPR = (f'<w:pPr><w:spacing w:after="60"/><w:pPrChange w:id="41" w:author="{OTHER}" w:date="{DATE}">'
             '<w:pPr/></w:pPrChange></w:pPr>')
OTHER_RPR = (f'<w:rPr><w:sz w:val="20"/><w:rPrChange w:id="42" w:author="{OTHER}" w:date="{DATE}"><w:rPr/>'
             '</w:rPrChange></w:rPr>')
OWN_RPR = (f'<w:rPr><w:sz w:val="20"/><w:rPrChange w:id="43" w:author="{CLAUDE}" w:date="{DATE}"><w:rPr>'
           '<w:sz w:val="18"/></w:rPr></w:rPrChange></w:rPr>')

BODY = [
    ("intro", p(r("Question one asks something."), ppr=BODY_PPR)),
    ("body", p(r("Reference answer text."), ppr=BODY_PPR)),
    ("blank", p(ppr='<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
                    '<w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr></w:pPr>')),
    ("blank_ok", p(ppr='<w:pPr><w:spacing w:after="120"/><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>')),
    ("list", p(r("Listed ", SZ20), r("item words.", ITAL), ppr='<w:pPr><w:pStyle w:val="NumberedItem"/>'
                                                       '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>')),
    ("small1", p(r("Small one.", SZ20))),
    ("small2", p(r("Small two.", SZ20), r(" more", ITAL))),
    ("small3", p(r("Small three.", SZ20))),
    ("revised", p(r("Keep ", SZ20), ins(18, OTHER, r("inserted ", SZ20)), dele(30, "gone "), r("end.", SZ20))),
    ("field", p(r("See ", SZ20), field("(1)"), r(" more.", SZ20))),
    ("other_ppr", p(r("Other para change.", SZ20), ppr=OTHER_PPR)),
    ("other_rpr", p(r("Plain ", SZ20), r("changed run.", OTHER_RPR))),
    ("own", p(r("Lead ", SZ20), ins(13, CLAUDE, r("own words ", SZ20)), r("kept.", OWN_RPR))),
    ("table", table(cell(p(r("Cell one.", SZ20)), p(r("Cell two.", SZ20))))),
    ("nested", table(cell(table(cell(p(r("Inner words here", SZ20)))), p(r(""))))),
    ("tail", p(r("Tail paragraph.", SZ20))),
]
INDEX: dict[str, int] = {}


def _index() -> None:
    number = 0
    for name, xml in BODY:
        INDEX[name] = number + 1
        number += len(re.findall(r"<w:p>", xml)) + len(re.findall(r"<w:tr>", xml))


_index()


def build(tmp_path: Path, body: str | None = None, name: str = "doc.docx") -> Path:
    """The tracked fixture's package with `body`, a numbering part and a numbered paragraph style."""

    body = body if body is not None else "".join(xml for _, xml in BODY)
    base = create_tracked_fixture(tmp_path / f"base-{name}")
    target = tmp_path / name
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(target, "x", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                xml = data.decode("utf-8")
                start = xml.index("<w:body>") + len("<w:body>")
                data = (xml[:start] + body + SECT + "</w:body></w:document>").encode("utf-8")
            elif info.filename == "word/styles.xml":
                data = data.decode("utf-8").replace("</w:styles>", LIST_STYLE + "</w:styles>").encode("utf-8")
            elif info.filename == "[Content_Types].xml":
                data = data.decode("utf-8").replace(
                    "</Types>", '<Override PartName="/word/numbering.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>').encode("utf-8")
            elif info.filename == "word/_rels/document.xml.rels":
                data = data.decode("utf-8").replace(
                    "</Relationships>", '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
                ).encode("utf-8")
            out.writestr(info, data, compress_type=info.compress_type)
        info = zipfile.ZipInfo("word/numbering.xml", date_time=(2026, 10, 30, 0, 0, 0))
        out.writestr(info, NUMBERING, compress_type=zipfile.ZIP_DEFLATED)
    return target


# -- independent readers -----------------------------------------------------------------------------------

def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def blocks(path: Path) -> list[str]:
    """Raw XML of each BODY entry (top-level paragraph or table), in order."""

    xml = document_xml(path)
    body = xml[xml.index("<w:body>") + len("<w:body>"):xml.rindex("<w:sectPr>")]
    out, depth, start = [], 0, 0
    for tag in re.finditer(r"<(/?)w:(p|tbl)\b[^>]*?(/?)>", body):
        closing, selfclosing = tag.group(1), tag.group(3)
        if selfclosing:
            if depth == 0:
                out.append(tag.group(0))
            continue
        if not closing:
            if depth == 0:
                start = tag.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                out.append(body[start:tag.end()])
    return out


def raw(path: Path, name: str) -> str:
    return blocks(path)[[key for key, _ in BODY].index(name)]


def entry(path: Path, name: str) -> ET.Element:
    body = ET.fromstring(document_xml(path)).find(W + "body")
    return list(body)[[key for key, _ in BODY].index(name)]


def others_identical(source: Path, output: Path, *names: str) -> None:
    a, b = blocks(source), blocks(output)
    for number, (key, _) in enumerate(BODY):
        if key not in names:
            assert a[number] == b[number], key
    sa, sb = document_xml(source), document_xml(output)
    assert sa[:sa.index("<w:body>")] == sb[:sb.index("<w:body>")]


def canonical(element: ET.Element) -> str:
    """Canonical XML; an empty w:pPr or w:rPr (what rejecting a change on an element that had none leaves)
    reads as absent, as it does for Word."""

    clone = copy.deepcopy(element)
    for _ in range(2):
        for parent in list(clone.iter()):
            for child in list(parent):
                if child.tag in (W + "pPr", W + "rPr") and len(child) == 0 and not child.attrib:
                    parent.remove(child)
    return ET.canonicalize(ET.tostring(clone, encoding="unicode"), strip_text=False)


def resolved(element: ET.Element, mode: str) -> ET.Element:
    """Accept All or Reject All of the formatting changes (w:pPrChange, w:rPrChange) in `element`."""

    clone = copy.deepcopy(element)
    for tag, holder in (("rPrChange", "rPr"), ("pPrChange", "pPr")):
        for props in list(clone.iter(W + holder)):
            change = props.find(W + tag)
            if change is None:
                continue
            if mode == "accept":
                props.remove(change)
                continue
            old = change.find(W + holder)
            kept = [child for child in props if holder == "pPr" and child.tag == W + "rPr"]
            props[:] = (list(old) if old is not None else []) + kept
    return clone


def run_props(element: ET.Element) -> list[tuple[str, dict]]:
    """(text, {child local name: attrs}) per run with text (w:t or w:delText), rPrChange left out."""

    out = []
    for run in element.iter(W + "r"):
        text = "".join(node.text or "" for node in run if node.tag in (W + "t", W + "delText"))
        if not text:
            continue
        rpr = run.find(W + "rPr")
        props = {} if rpr is None else {child.tag.replace(W, ""): dict(child.attrib) for child in rpr
                                        if child.tag != W + "rPrChange"}
        out.append((text, props))
    return out


def ppr_of(element: ET.Element) -> dict:
    ppr = element.find(W + "pPr")
    return {} if ppr is None else {child.tag.replace(W, ""): child for child in ppr}


def size_of(props: dict) -> str | None:
    return props.get("sz", {}).get(W + "val")


def text_of(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.iter(W + "t"))


def _sha(path: Path) -> str:
    from research_mcp.audit import file_sha256

    return file_sha256(path)


def fmt(source: Path, tmp_path: Path, name: str, *, out: str = "out.docx", tracked=False, author=CLAUDE, **kw):
    output = tmp_path / out
    result = format_paragraphs(source, output, paragraph_index=INDEX[name], tracked=tracked, author=author,
                               now=NOW, **kw)
    return output, result


def refused(source: Path, tmp_path: Path, name: str, *needles: str, **kw) -> None:
    output = tmp_path / "refused.docx"
    with pytest.raises(RefusedUnsafe) as caught:
        format_paragraphs(source, output, paragraph_index=INDEX[name], tracked=kw.pop("tracked", True),
                          author=kw.pop("author", CLAUDE), now=NOW, **kw)
    for needle in needles:
        assert needle in str(caught.value), str(caught.value)
    assert not output.exists()


def batch(path: Path, *ops: dict, output: Path | None = None, **kw) -> dict:
    from research_mcp import server

    envelope = [server.MCPEditOperation(**op) for op in ops]
    return server.edit_document(str(path), _sha(path), envelope, str(output or path), **kw)


def _no_word(monkeypatch):
    from research_mcp import operation_process

    original = operation_process.run_operation

    def guarded(kind, *args, **kwargs):
        if kind != "edit":
            raise AssertionError(f"Word was started ({kind})")
        return original(kind, *args, **kwargs)

    monkeypatch.setattr(operation_process, "run_operation", guarded)


# -- format_note (decision 6) ------------------------------------------------------------------------------

def test_format_note_on_numbered_small_blank_fill(tmp_path):
    source = build(tmp_path)
    blank, body = INDEX["blank"], INDEX["body"]
    note = blank_fill_note(source, blank)
    assert note == (f"Paragraph {blank} is numbered and 6 pt (paragraph {body} is 12 pt); "
                    f"format_paragraph(paragraph_index={blank}, like_paragraph={body}) matches it to paragraph "
                    f"{body}.")
    assert blank_fill_note(source, INDEX["blank_ok"]) is None
    assert blank_fill_note(source, INDEX["body"]) is None  # not empty

    result = batch(source, {"op_id": "fill", "type": "replace_paragraph", "paragraph_index": blank,
                            "text": "Answer"})
    assert result["status"] == "complete", result
    details = result["operations"][0]["details"]
    assert details["format_note"] == note
    filled = entry(source, "blank")
    assert text_of(filled) == "Answer" and "numPr" in ppr_of(filled)  # nothing changed by itself

    result = batch(source, {"op_id": "ok", "type": "replace_paragraph", "paragraph_index": INDEX["blank_ok"],
                            "text": "Matching answer"},
                   {"op_id": "full", "type": "replace_text", "paragraph_index": body, "find": "Reference",
                    "replace": "Model"})
    assert result["status"] == "complete", result
    assert all("format_note" not in item["details"] for item in result["operations"])


TITLE_PPR = '<w:pPr><w:jc w:val="center"/></w:pPr>'
TITLE_RPR = '<w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
SMALL_BLANK = p(ppr='<w:pPr><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr></w:pPr>')


def test_format_note_skips_a_title_made_by_direct_formatting(tmp_path):
    """A centered bold 18 pt line (style None) between an answer and a blank is no reference: the note names
    the answer, and following it gives the blank the answer's 12 pt, not 18 pt."""

    body = (p(r("Question text."), ppr=BODY_PPR) + p(r("Answer text."), ppr=BODY_PPR)
            + p(r("Part Two", TITLE_RPR), ppr=TITLE_PPR) + SMALL_BLANK + p(r("Closing words."), ppr=BODY_PPR))
    source = build(tmp_path, body=body)
    assert blank_fill_note(source, 4) == ("Paragraph 4 is 6 pt (paragraph 5 is 12 pt); "
                                          "format_paragraph(paragraph_index=4, like_paragraph=5) matches it to "
                                          "paragraph 5.")
    # without the closing line the nearest body paragraph is the answer before the title
    body = p(r("Question text."), ppr=BODY_PPR) + p(r("Answer text."), ppr=BODY_PPR) + \
        p(r("Part Two", TITLE_RPR), ppr=TITLE_PPR) + SMALL_BLANK
    source = build(tmp_path, body=body, name="answer.docx")
    note = blank_fill_note(source, 4)
    assert note == ("Paragraph 4 is 6 pt (paragraph 2 is 12 pt); format_paragraph(paragraph_index=4, "
                    "like_paragraph=2) matches it to paragraph 2.")
    # each title signal alone is enough: centered, mostly bold, 3 pt or more above the body size
    for name, title in (("centered", p(r("Part Two", SZ24), ppr=TITLE_PPR)),
                        ("bold", p(r("Part Two", '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>'))),
                        ("large", p(r("Part Two", '<w:rPr><w:sz w:val="30"/></w:rPr>')))):
        body = (p(r("Question text."), ppr=BODY_PPR) + p(r("Answer text."), ppr=BODY_PPR) + title
                + SMALL_BLANK)
        assert "like_paragraph=2)" in blank_fill_note(build(tmp_path, body=body, name=f"{name}.docx"), 4), name
    # a bold answer that is indented, or longer than a title line, is body text: the note names it
    indented = '<w:pPr><w:ind w:left="360"/></w:pPr>'
    for name, answer in (("indented", p(r("Bold answer.", '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>'), ppr=indented)),
                         ("long", p(r("A bold answer that runs on for well over eighty characters, as a full "
                                      "sentence of an answer does when the student writes it out in bold.", '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>')))):
        body = p(r("Question text.")) + answer + SMALL_BLANK
        assert "like_paragraph=2)" in blank_fill_note(build(tmp_path, body=body, name=f"bold-{name}.docx"), 3), name
    # a line 2 pt above the body size is still body text
    body = p(r("Answer text.", SZ20)) + p(r("Bigger answer.")) + SMALL_BLANK
    assert "like_paragraph=2)" in blank_fill_note(build(tmp_path, body=body, name="near.docx"), 3)


def test_format_note_without_a_body_paragraph_names_no_like_call(tmp_path):
    numbered_blank = p(ppr='<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
                           '<w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr></w:pPr>')
    body = p(r("Part Two", TITLE_RPR), ppr=TITLE_PPR) + numbered_blank
    source = build(tmp_path, body=body)
    note = blank_fill_note(source, 2)
    assert note == ("Paragraph 2 is numbered and 6 pt (most text is 18 pt); no paragraph reads as body text "
                    "to match it to, so give format_paragraph the font, size and indents it needs.")
    assert "like_paragraph" not in note
    # only headings around a blank of the heading's size: nothing to report
    body = p(r("Part Two", TITLE_RPR), ppr=TITLE_PPR) + p(ppr='<w:pPr><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr>')
    assert blank_fill_note(build(tmp_path, body=body, name="same.docx"), 2) is None


def test_size_sets_complex_script_size_too(tmp_path):
    """A run at the requested w:sz but another w:szCs, set or inherited (the fixture's docDefaults are 11 pt),
    is formatted, as Word's size box sets both; a run with both already at the size is left alone."""

    body = (p(r("Mixed size.", '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>'
                               '<w:szCs w:val="16"/></w:rPr>'), r(" plain", '<w:rPr><w:sz w:val="24"/></w:rPr>'),
              r(" both", SZ24), ppr=BODY_PPR))
    source = build(tmp_path, body=body)
    output = tmp_path / "out.docx"
    result = format_paragraphs(source, output, paragraph_index=1, size=12, tracked=False, author=CLAUDE, now=NOW)
    assert result["paragraphs"][0]["runs_formatted"] == 2, result
    runs = run_props(list(ET.fromstring(document_xml(output)).find(W + "body"))[0])
    assert [(size_of(run), run.get("szCs", {}).get(W + "val")) for _, run in runs] == [
        ("24", "24"), ("24", "24"), ("24", "24")]
    with pytest.raises(RefusedUnsafe, match="already formatted as requested"):
        format_paragraphs(output, tmp_path / "again.docx", paragraph_index=1, size=12, tracked=False,
                          author=CLAUDE, now=NOW)


# -- plain -------------------------------------------------------------------------------------------------

def test_like_paragraph_removes_numbering_and_matches(tmp_path):
    source = build(tmp_path)
    output, result = fmt(source, tmp_path, "list", like_paragraph=INDEX["body"], remove_numbering=True)
    para = entry(output, "list")
    props = ppr_of(para)
    numpr = props["numPr"]  # the style numbers it: switched off
    assert [(child.tag.replace(W, ""), child.get(W + "val")) for child in numpr] == [("numId", "0")]
    assert props["spacing"].get(W + "after") == "120"
    assert props["pStyle"].get(W + "val") == "NumberedItem"
    assert size_of({child.tag.replace(W, ""): child.attrib for child in props["rPr"]}) == "24"
    runs = run_props(para)
    assert [text for text, _ in runs] == ["Listed ", "item words."]
    for _, run in runs:
        assert size_of(run) == "24" and run["szCs"][W + "val"] == "24"
    # the Arial run takes the body's Calibri; the italic run already shows Calibri (document default)
    assert {key: runs[0][1]["rFonts"][W + key] for key in ("ascii", "hAnsi", "cs")} == dict.fromkeys(
        ("ascii", "hAnsi", "cs"), "Calibri")
    assert "rFonts" not in runs[1][1] and "i" in runs[1][1]  # italic kept
    assert "rPrChange" not in raw(output, "list") and "pPrChange" not in raw(output, "list")
    assert result["revisions_added"] == [] and result["tracked"] is False
    assert result["requested"]["numbering"] == "none" and result["requested"]["size"] == 12
    others_identical(source, output, "list")

    output, _ = fmt(source, tmp_path, "blank", out="blank.docx", like_paragraph=INDEX["body"],
                    remove_numbering=True)
    props = ppr_of(entry(output, "blank"))
    assert "numPr" not in props and props["spacing"].get(W + "after") == "120"
    assert props["rPr"].find(W + "sz").get(W + "val") == "24"
    others_identical(source, output, "blank")


def test_explicit_fields_override_like(tmp_path):
    source = build(tmp_path)
    output, result = fmt(source, tmp_path, "small1", like_paragraph=INDEX["body"], size=14, left_indent=36)
    para = entry(output, "small1")
    props = ppr_of(para)
    assert props["ind"].get(W + "left") == "720" and props["spacing"].get(W + "after") == "120"
    assert [size_of(run) for _, run in run_props(para)] == ["28"]
    assert result["requested"]["size"] == 14 and result["requested"]["left_indent"] == 36


def test_hanging_indent(tmp_path):
    source = build(tmp_path)
    output, _ = fmt(source, tmp_path, "small1", left_indent=36, first_line_indent=-18)
    ind = ppr_of(entry(output, "small1"))["ind"]
    assert dict(ind.attrib) == {W + "left": "720", W + "hanging": "360"}
    assert run_props(entry(output, "small1")) == run_props(entry(source, "small1"))
    others_identical(source, output, "small1")


def test_range_formats_every_paragraph(tmp_path):
    source = build(tmp_path)
    output, result = fmt(source, tmp_path, "small1", through_paragraph_index=INDEX["small3"], font="Georgia",
                         size=11)
    for name in ("small1", "small2", "small3"):
        para = entry(output, name)
        for _, run in run_props(para):
            assert size_of(run) == "22" and run["szCs"][W + "val"] == "22"
            assert {key: run["rFonts"][W + key] for key in ("ascii", "hAnsi", "cs")} == dict.fromkeys(
                ("ascii", "hAnsi", "cs"), "Georgia")
        mark = ppr_of(para)["rPr"]  # the default size is already 11 pt: only the font is written
        assert mark.find(W + "sz") is None and mark.find(W + "rFonts").get(W + "ascii") == "Georgia"
    assert [item["paragraph_index"] for item in result["paragraphs"]] == [INDEX[name] for name in
                                                                          ("small1", "small2", "small3")]
    others_identical(source, output, "small1", "small2", "small3")


# -- tracked -----------------------------------------------------------------------------------------------

def test_tracked_switch_writes_ppr_and_rpr_changes(tmp_path, monkeypatch):
    _no_word(monkeypatch)
    source = build(tmp_path)
    output = tmp_path / "tracked.docx"
    result = batch(source, {"op_id": "f", "type": "format_paragraph", "paragraph_index": INDEX["list"],
                            "like_paragraph": INDEX["body"], "remove_numbering": True},
                   output=output, track_changes=True)
    assert result["status"] == "complete", result
    details = result["operations"][-1]["details"]
    assert details["tracked"] is True and details["author"] == CLAUDE
    para = entry(output, "list")
    ppr_changes = list(para.iter(W + "pPrChange"))
    rpr_changes = list(para.iter(W + "rPrChange"))
    assert len(ppr_changes) == 1 and len(rpr_changes) == 3  # the mark and two runs
    assert {item.get(W + "author") for item in ppr_changes + rpr_changes} == {CLAUDE}
    ids = [int(item.get(W + "id")) for item in para.iter() if item.tag in (W + "pPrChange", W + "rPrChange")]
    assert len(set(ids)) == len(ids) and ids == sorted(ids)
    assert [int(item["id"]) for item in details["revisions_added"]] == sorted(ids)
    old_ppr = ppr_changes[0].find(W + "pPr")
    assert old_ppr.find(W + "pStyle").get(W + "val") == "NumberedItem" and old_ppr.find(W + "numPr") is None
    # Reject All = input; Accept All = the plain result
    assert canonical(resolved(para, "reject")) == canonical(entry(source, "list"))
    plain, _ = fmt(source, tmp_path, "list", out="plain.docx", like_paragraph=INDEX["body"], remove_numbering=True)
    assert canonical(resolved(para, "accept")) == canonical(entry(plain, "list"))
    others_identical(source, output, "list")


def test_tracked_reject_restores_input(tmp_path, monkeypatch):
    _no_word(monkeypatch)
    source = build(tmp_path)
    output = tmp_path / "tracked.docx"
    batch(source, {"op_id": "f", "type": "format_paragraph", "paragraph_index": INDEX["small1"],
                   "through_paragraph_index": INDEX["small2"], "font": "Arial", "size": 12, "space_after": 6},
          output=output, track_changes=True)
    rejected = tmp_path / "rejected.docx"
    result = batch(output, {"op_id": "r", "type": "reject_revisions", "author": CLAUDE}, output=rejected)
    assert result["status"] == "complete", result
    for name in ("small1", "small2"):
        assert canonical(entry(rejected, name)) == canonical(entry(source, name)), name
    accepted = tmp_path / "accepted.docx"
    result = batch(output, {"op_id": "a", "type": "accept_revisions", "author": CLAUDE}, output=accepted)
    assert result["status"] == "complete", result
    for name in ("small1", "small2"):
        para = entry(accepted, name)
        assert "rPrChange" not in ET.tostring(para, encoding="unicode")
        assert ppr_of(para)["spacing"].get(W + "after") == "120"
        assert {size_of(run) for _, run in run_props(para)} == {"24"}


def test_revised_document_keeps_other_revisions(tmp_path):
    source = build(tmp_path)
    for tracked in (False, True):
        output, result = fmt(source, tmp_path, "revised", out=f"rev-{tracked}.docx", size=12, tracked=tracked)
        before, after = raw(source, "revised"), raw(output, "revised")
        deleted = re.search(r"<w:del\b.*?</w:del>", before).group(0)
        assert deleted in after  # the w:del and its run byte-identical
        assert f'<w:ins w:id="18" w:author="{OTHER}" w:date="{DATE}">' in after
        para = entry(output, "revised")
        sizes = {text: size_of(run) for text, run in run_props(para)}
        assert sizes == {"Keep ": "24", "inserted ": "24", "gone ": "20", "end.": "24"}
        inserted_run = para.find(W + "ins").find(W + "r")
        has_change = inserted_run.find(f"{W}rPr/{W}rPrChange") is not None
        assert has_change is tracked
        assert len(list(para.iter(W + "rPrChange"))) == (4 if tracked else 0)  # mark + 3 live runs
        others_identical(source, output, "revised")
        if tracked:
            assert canonical(resolved(para, "reject")) == canonical(entry(source, "revised"))


def test_deleted_and_field_runs_keep_rpr(tmp_path):
    """Field code and result runs (here 10 pt Arial) keep their rPr; the text around the field is formatted."""
    source = build(tmp_path)
    for tracked in (False, True):
        output, result = fmt(source, tmp_path, "field", out=f"field-{tracked}.docx", size=12, tracked=tracked)
        before, after = entry(source, "field"), entry(output, "field")
        runs = list(after.iter(W + "r"))
        old_runs = list(before.iter(W + "r"))
        field_runs = range(1, 6)  # begin, instr, separate, result, end
        for number in field_runs:
            assert canonical(runs[number]) == canonical(old_runs[number]), number
        sizes = {text: size_of(run) for text, run in run_props(after)}
        assert sizes["See "] == "24" and sizes[" more."] == "24" and sizes["(1)"] == "20"
        assert result["paragraphs"][0]["runs_formatted"] == 2


def test_own_insertion_and_own_change_formatted_in_place(tmp_path):
    """L1: a run in Claude's own w:ins takes the size with no rPrChange; L3: a run with Claude's own rPrChange
    keeps that change byte-identical (its old rPr is still the state before the revision)."""
    source = build(tmp_path)
    output, result = fmt(source, tmp_path, "own", size=12, tracked=True)
    para = entry(output, "own")
    inserted = para.find(W + "ins").find(W + "r")
    assert inserted.find(f"{W}rPr/{W}rPrChange") is None and size_of(dict(run_props(para))["own words "]) == "24"
    own_change = re.search(r"<w:rPrChange w:id=\"43\".*?</w:rPrChange>", raw(source, "own")).group(0)
    assert own_change in raw(output, "own")
    kinds = sorted(item["kind"] for item in result["paragraphs"][0]["formatted_in_place"])
    assert kinds == ["own_format_change", "own_insertion"]
    assert {size_of(run) for _, run in run_props(para)} == {"24"}


def test_table_cell_range(tmp_path):
    source = build(tmp_path)
    output, _ = fmt(source, tmp_path, "table", through_paragraph_index=INDEX["table"] + 1, size=12)
    assert {size_of(run) for _, run in run_props(entry(output, "table"))} == {"24"}
    others_identical(source, output, "table")


# -- refusals ----------------------------------------------------------------------------------------------

def test_refusals_other_authors_changes(tmp_path):
    source = build(tmp_path)
    refused(source, tmp_path, "other_ppr", "w:pPrChange id=41", OTHER, "another author", space_after=12)
    refused(source, tmp_path, "other_rpr", "w:rPrChange id=42", OTHER, "another author", size=12)
    refused(source, tmp_path, "other_rpr", "untracked formatting is refused", size=12, tracked=False)
    # the other author's change is untouched when the requested property needs no change there
    output, _ = fmt(source, tmp_path, "other_rpr", out="font.docx", left_indent=12, tracked=True)
    assert re.search(r"<w:rPrChange w:id=\"42\".*?</w:rPrChange>", raw(source, "other_rpr")).group(0) in \
        raw(output, "other_rpr")


def test_refusals_containers_and_ranges(tmp_path):
    source = build(tmp_path)
    refused(source, tmp_path, "nested", "nested table", size=12)
    refused(source, tmp_path, "small3", "do not lie in one container", "give one op per cell",
            through_paragraph_index=INDEX["table"], size=12)
    refused(source, tmp_path, "table", "row-end marker", through_paragraph_index=INDEX["table"] + 2, size=12)
    refused(source, tmp_path, "small1", "outside 1..", through_paragraph_index=999, size=12)


def test_refusals_nothing_to_change_and_like(tmp_path):
    source = build(tmp_path)
    refused(source, tmp_path, "small1", "already formatted as requested", space_after=0)
    refused(source, tmp_path, "small1", "like_paragraph 999 is outside", like_paragraph=999)
    refused(source, tmp_path, "small1", "one of the paragraphs to format", like_paragraph=INDEX["small1"])
    refused(source, tmp_path, "small1", "size must be", size=11.3)


def test_refusals_bold_italic_point_to_style_text(tmp_path):
    from pydantic import ValidationError

    from research_mcp.edit_models import FormatParagraphOperation

    with pytest.raises(ValidationError, match="style_text"):
        FormatParagraphOperation(op_id="f", type="format_paragraph", paragraph_index=1, size=12, bold=True)
    with pytest.raises(ValidationError, match="give at least one|at least one"):
        FormatParagraphOperation(op_id="f", type="format_paragraph", paragraph_index=1)
    source = build(tmp_path)
    before = source.read_bytes()
    with pytest.raises(RefusedUnsafe, match="use style_text"):
        batch(source, {"op_id": "f", "type": "format_paragraph", "paragraph_index": 1, "size": 12, "bold": True})
    assert source.read_bytes() == before


def test_batch_refusal_writes_nothing(tmp_path, monkeypatch):
    _no_word(monkeypatch)
    source = build(tmp_path)
    before = source.read_bytes()
    result = batch(source, {"op_id": "f", "type": "format_paragraph", "paragraph_index": INDEX["other_rpr"],
                            "size": 12, "tracked": True})
    assert result["status"] != "complete" and source.read_bytes() == before
    assert "another author" in str(result)


# -- edit feedback (E11) -----------------------------------------------------------------------------------

def test_edit_feedback_stays_pending(tmp_path, monkeypatch):
    """A format_paragraph write records no proposal, and an open tracked proposal of the same paragraph stays
    pending. Needs "format_paragraph" in edit_feedback_word._FORMAT_ONLY_OPS (owned by another stage)."""
    from research_mcp.edit_feedback import chains
    from research_mcp.edit_feedback_word import doc_log

    _no_word(monkeypatch)
    source = build(tmp_path)
    name = INDEX["small1"]
    batch(source, {"op_id": "e1", "type": "replace_text", "paragraph_index": name, "find": "Small one.",
                   "replace": "Small first one."}, track_changes=True)
    events = doc_log(source).events()
    proposals = [event for event in events if event.get("event") == "proposal"]
    assert len(proposals) == 1
    result = batch(source, {"op_id": "f", "type": "format_paragraph", "paragraph_index": name, "size": 12})
    assert result["status"] == "complete", result
    events = doc_log(source).events()
    assert len([event for event in events if event.get("event") == "proposal"]) == 1
    assert [event for event in events if event.get("event") == "outcome"] == []
    assert [record["pending"] for record in chains(events)] == [True]
