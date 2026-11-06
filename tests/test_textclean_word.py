"""Plan 551 S2: clean on write, the tracked character sweep, and inspect_document's count.

Synthetic fixtures only (built here, under tmp_path).
Expectations are checked with this file's own ElementTree reader, not with revisions.py's resolver.
"""
from __future__ import annotations

import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp.audit import file_sha256
from research_mcp.edit_batch import _clean_operation, edit_document
from research_mcp.edit_models import EditDocumentRequest
from research_mcp.fixtures import _MIXED_CONTENT_TYPES, _MIXED_DOCUMENT_RELS, _MIXED_RELS, _MIXED_STYLES
from research_mcp.paragraphs import inspect_document

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
)
OTHER = "Coauthor"
NBSP, ZWSP, THIN = "\u00a0", "\u200b", "\u2009"
MANUSCRIPT = "x \u2192 y \u2026 \u201cq\u201d \u2014 z"


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _r(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'


MATH = (f'<m:oMath><m:r><m:t>a{THIN}+{ZWSP}b</m:t></m:r></m:oMath>')
ZOTERO = (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"id":"x\u00a0y"} </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    f'<w:r><w:t xml:space="preserve">(Smith,{NBSP}2020)</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
)
#  1  plain text with NBSP and ZWSP
#  2  OMML math holding U+2009 / U+200B, NBSP in the text around it
#  3  a Zotero citation whose result and code hold NBSP; NBSP in the text after it
#  4  manuscript punctuation textclean keeps
#  5  another author's tracked insertion holding an NBSP
PARAGRAPHS = [
    _r(f"Plain 10{NBSP}mm text a{ZWSP}b."),
    _r(f"Let{NBSP}") + MATH + _r(f" be{NBSP}small."),
    _r("Prior work ") + ZOTERO + _r(f" says{NBSP}so."),
    _r(MANUSCRIPT),
    _r("Kept ") + f'<w:ins w:id="1" w:author="{OTHER}" w:date="2026-01-01T00:00:00Z">'
    + _r(f"their{NBSP}words") + "</w:ins>" + _r(" end."),
]


def build(path: Path, paragraphs=PARAGRAPHS) -> Path:
    body = "".join(f"<w:p>{item}</w:p>" for item in paragraphs)
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
                '<w:sectPr/></w:body></w:document>')
    with zipfile.ZipFile(path, "x", zipfile.ZIP_DEFLATED) as archive:
        for name, data in (("[Content_Types].xml", _MIXED_CONTENT_TYPES), ("_rels/.rels", _MIXED_RELS),
                           ("word/document.xml", document), ("word/_rels/document.xml.rels", _MIXED_DOCUMENT_RELS),
                           ("word/styles.xml", _MIXED_STYLES)):
            archive.writestr(name, data)
    return path


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    return build(tmp_path / "clean.docx")


def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def view(paragraph: ET.Element, *, accept_new: bool, new_author: str = "Claude") -> str:
    """Text with every revision accepted, except those by `new_author`: accepted (or rejected) as asked."""
    out: list[str] = []

    def walk(element, inserted_by=None, deleted_by=None):
        for child in element:
            if child.tag in (W + "pPr", W + "rPr", W + "instrText"):
                continue
            if child.tag == W + "ins":
                walk(child, child.get(W + "author"), deleted_by)
            elif child.tag == W + "del":
                walk(child, inserted_by, child.get(W + "author"))
            elif child.tag in (W + "t", W + "delText"):
                keep_ins = inserted_by is None or inserted_by != new_author or accept_new
                keep_del = deleted_by is None or (deleted_by == new_author and not accept_new)
                if keep_ins and keep_del:
                    out.append(child.text or "")
            elif child.tag == W + "br":
                out.append("\v")
            elif child.tag.endswith("}oMath"):
                continue
            else:
                walk(child, inserted_by, deleted_by)

    walk(paragraph)
    return "".join(out)


def body_paragraphs(path: Path) -> list[ET.Element]:
    return list(ET.fromstring(document_xml(path)).iter(W + "p"))


def request(path: Path, *operations, output: Path | None = None) -> EditDocumentRequest:
    return EditDocumentRequest(source=str(path), source_sha256=file_sha256(path),
                               output=str(output) if output else None, operations=list(operations))


def codepoints(changes) -> set[str]:
    return {item["codepoint"] for item in changes}


def captured(monkeypatch):
    """Replace the Word worker with a copy that records the (cleaned) operation it was given."""
    seen = []

    def fake(source, output, operation):
        seen.append(operation)
        shutil.copy2(source, output)
        from research_mcp.audit import audit_docx

        return {"paragraph_count": audit_docx(output).paragraph_count}

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake)
    return seen


# -- clean on write -------------------------------------------------------------------------------------


def test_insert_cleans_nbsp_and_zero_width(monkeypatch, doc, tmp_path):
    seen = captured(monkeypatch)
    result = edit_document(request(doc, dict(op_id="i", type="insert_text", paragraph_index=4, position="after",
                                             anchor="z", text=f" 10{NBSP}mm a{ZWSP}b"),
                                   output=tmp_path / "out.docx"))
    assert result.status == "complete", result.message
    assert seen[0].text == " 10 mm ab"
    assert codepoints(result.cleaned) == {"U+00A0", "U+200B"}
    assert codepoints(result.operations[0].details["cleaned"]) == {"U+00A0", "U+200B"}


def test_insert_cleans_through_tracked_xml(doc, tmp_path):
    """End to end without Word: the tracked insertion on disk holds the cleaned text."""
    out = tmp_path / "out.docx"
    result = edit_document(request(doc, dict(op_id="t", type="tracked_replace_text", paragraph_index=4,
                                             find="z", replace=f"z 10{NBSP}mm a{ZWSP}b"), output=out))
    assert result.status == "complete", result.message
    accepted = view(body_paragraphs(out)[3], accept_new=True)
    assert accepted == MANUSCRIPT + " 10 mm ab"
    assert NBSP not in accepted and ZWSP not in accepted
    assert codepoints(result.cleaned) == {"U+00A0", "U+200B"}


def test_keep_holds_nbsp(monkeypatch, doc, tmp_path):
    seen = captured(monkeypatch)
    result = edit_document(request(doc, dict(op_id="i", type="insert_text", paragraph_index=4, position="after",
                                             anchor="z", text=f" 10{NBSP}mm a{ZWSP}b", keep=[NBSP]),
                                   output=tmp_path / "out.docx"))
    assert seen[0].text == f" 10{NBSP}mm ab"
    assert codepoints(result.cleaned) == {"U+200B"}
    seen.clear()
    edit_document(request(doc, dict(op_id="i", type="insert_text", paragraph_index=4, position="after",
                                    anchor="z", text=f" 10{NBSP}mm", keep=["U+00A0"]),
                          output=tmp_path / "out2.docx"))
    assert seen[0].text == f" 10{NBSP}mm"


def test_manuscript_style_unchanged(monkeypatch, doc, tmp_path):
    seen = captured(monkeypatch)
    result = edit_document(request(doc, dict(op_id="p", type="insert_paragraph", after_paragraph_index=4,
                                             text=MANUSCRIPT), output=tmp_path / "out.docx"))
    assert seen[0].text == MANUSCRIPT
    assert result.cleaned == []
    assert "cleaned" not in result.operations[0].details


def test_echoed_text_is_left_alone(monkeypatch, doc, tmp_path):
    """replace_paragraph / replace_text clean only what differs from the text being replaced."""
    seen = captured(monkeypatch)
    edit_document(request(doc, dict(op_id="p", type="replace_paragraph", paragraph_index=3,
                                    text=f"Earlier work (Smith,{NBSP}2020) says{NBSP}so{ZWSP}."),
                          output=tmp_path / "out.docx"))
    # the NBSPs sit inside text kept from the paragraph; only the new ZWSP goes
    assert seen[0].text == f"Earlier work (Smith,{NBSP}2020) says{NBSP}so."
    seen.clear()
    edit_document(request(doc, dict(op_id="r", type="replace_text", paragraph_index=1,
                                    find=f"10{NBSP}mm", replace=f"20{NBSP}mm{NBSP}wide"),
                          output=tmp_path / "out2.docx"))
    assert seen[0].replace == f"20{NBSP}mm wide"
    assert seen[0].find == f"10{NBSP}mm"  # match strings are never cleaned


def test_every_write_field_is_cleaned(doc):
    from pydantic import TypeAdapter

    from research_mcp.edit_models import EditOperation

    adapter = TypeAdapter(EditOperation)
    dirty = f"a{NBSP}b{ZWSP}"
    cases = {
        "insert_table": (dict(after_paragraph_index=4, headers=[dirty], rows=[[dirty]], descriptive_columns=[1]),
                         lambda op: (op.headers, op.rows)),
        "add_comment": (dict(paragraph_index=1, anchor="Plain", text=f"{dirty}\nline"),
                        lambda op: op.text),
        "reply_comment": (dict(comment_id="0", reply_text=dirty), lambda op: op.reply_text),
        "replace_table_row": (dict(match="x", cells=[dirty]), lambda op: op.cells),
        "replace_section": (dict(heading="H", paragraphs=[dirty]), lambda op: op.paragraphs),
        "tracked_replace_paragraph": (dict(paragraph_index=4, text=MANUSCRIPT + dirty), lambda op: op.text),
    }
    for kind, (fields, get) in cases.items():
        operation = adapter.validate_python({"op_id": kind, "type": kind, **fields})
        cleaned, changes = _clean_operation(operation, None, doc)
        assert codepoints(changes) == {"U+00A0", "U+200B"}, kind
        assert NBSP not in repr(get(cleaned)) and ZWSP not in repr(get(cleaned)), kind
    comment = adapter.validate_python({"op_id": "c", "type": "add_comment", "paragraph_index": 1,
                                       "anchor": f"Plain{NBSP}", "text": "ok"})
    assert _clean_operation(comment, None, doc) == (comment, [])


def test_line_separator_becomes_manual_break(doc, tmp_path):
    out = tmp_path / "out.docx"
    result = edit_document(request(doc, dict(op_id="t", type="tracked_replace_text", paragraph_index=4,
                                             find="z", replace="z\u2028next"), output=out))
    assert result.status == "complete", result.message
    assert view(body_paragraphs(out)[3], accept_new=True) == MANUSCRIPT + "\vnext"
    assert result.cleaned[0]["to"] == "\v"


# -- sweep ----------------------------------------------------------------------------------------------


def math_raw(xml: str) -> list[str]:
    return re.findall(r"<m:oMath>.*?</m:oMath>", xml, re.S)


def field_raw(xml: str) -> list[str]:
    return re.findall(r'<w:r><w:fldChar w:fldCharType="begin"/>.*?<w:fldChar w:fldCharType="end"/></w:r>', xml,
                      re.S)


def sweep(doc: Path, out: Path, **fields):
    return edit_document(request(doc, dict(op_id="s", type="sweep_characters", **fields), output=out))


def test_sweep_skips_math(doc, tmp_path):
    out = tmp_path / "out.docx"
    result = sweep(doc, out, paragraph_index=2)
    assert result.status == "complete", result.message
    before, after = document_xml(doc), document_xml(out)
    assert math_raw(after) == math_raw(before) == [MATH]
    paragraph = body_paragraphs(out)[1]
    assert view(paragraph, accept_new=True) == "Let  be small."  # the math sits between the spaces
    assert view(paragraph, accept_new=False) == f"Let{NBSP} be{NBSP}small."
    ins = [el for el in paragraph.iter(W + "ins")]
    dels = [el for el in paragraph.iter(W + "del")]
    assert len(ins) == len(dels) == 2 and {el.get(W + "author") for el in ins + dels} == {"Claude"}
    details = result.operations[0].details
    assert details["skipped"]["math"] == 2 and details["skipped_total"] == 2
    assert codepoints(details["cleaned"]) == {"U+00A0"}
    # other paragraphs untouched
    assert body_paragraphs(out)[0].__len__() == body_paragraphs(doc)[0].__len__()
    assert NBSP in view(body_paragraphs(out)[0], accept_new=True)


def test_sweep_skips_zotero_field(doc, tmp_path):
    out = tmp_path / "out.docx"
    result = sweep(doc, out, paragraph_index=3)
    assert result.status == "complete", result.message
    assert field_raw(document_xml(out)) == field_raw(document_xml(doc)) and len(field_raw(document_xml(doc))) == 1
    assert view(body_paragraphs(out)[2], accept_new=True) == f"Prior work (Smith,{NBSP}2020) says so."
    details = result.operations[0].details
    assert details["skipped"]["field"] == 2  # result NBSP and code NBSP
    assert details["revisions_added"] == 2


def test_sweep_skips_field_spanning_paragraphs(tmp_path):
    """A Zotero bibliography is one field over many paragraphs; none of it is swept."""
    bibliography = [
        _r(f"Intro{NBSP}text.") + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + _r(f"1. Smith{NBSP}A."),
        _r(f"2. Jones{NBSP}B."),
        _r(f"3. Lee{NBSP}C.") + '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        _r(f"After{NBSP}it."),
    ]
    source = build(tmp_path / "bib.docx", bibliography)
    out = tmp_path / "out.docx"
    result = sweep(source, out)
    assert result.status == "complete", result.message
    texts = [view(item, accept_new=True) for item in body_paragraphs(out)]
    assert texts == ["Intro text.1. Smith\u00a0A.", f"2. Jones{NBSP}B.", f"3. Lee{NBSP}C.", "After it."]
    details = result.operations[0].details
    assert details["skipped"] == {"field": 3} and details["paragraphs_changed"] == [1, 4]
    assert inspect_document(source)["unusual_characters"]["total"] == 2


def _fld(kind: str) -> str:
    return f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'


def _instr(code: str) -> str:
    return f'<w:r><w:instrText xml:space="preserve">{_esc(code)}</w:instrText></w:r>'


def _deleted(inner: str, rid: int) -> str:
    return f'<w:del w:id="{rid}" w:author="{OTHER}" w:date="2026-01-01T00:00:00Z">{inner}</w:del>'


_DELETED_BEGIN = _deleted(_fld("begin") + '<w:r><w:delInstrText xml:space="preserve"> EQ x </w:delInstrText></w:r>'
                          + _fld("separate"), 8)
_FIELD_CODES = [" EQ x\\to(y) ", ' ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY ']


def _field_result(path: Path) -> list[str]:
    """Raw XML from the first field begin through the last field end: field code and result together."""
    xml = document_xml(path)
    return [xml[xml.index('w:fldCharType="begin"'):xml.rindex('w:fldCharType="end"')]]


@pytest.mark.parametrize("code", _FIELD_CODES)
def test_sweep_keeps_field_with_revised_boundaries_across_paragraphs(tmp_path, code):
    """Reviewer regression: a deleted field end in P2 and a deleted begin/code/separate in P4, before the live
    end. Stored markers balance, but with the deletions accepted P2..P4 are still inside the field; rejected,
    P3 would be outside it. The field's text must stay byte-identical; only P5 is swept."""
    paragraphs = [
        _fld("begin") + _instr(code) + _fld("separate") + _r(f"one{NBSP}field"),
        _deleted(_fld("end"), 7) + _r(f"two{NBSP}field"),
        _r(f"three{NBSP}field"),
        _DELETED_BEGIN + _r(f"four{NBSP}field") + _fld("end"),
        _r(f"outside{NBSP}text"),
    ]
    source = build(tmp_path / "revised.docx", paragraphs)
    out = tmp_path / "out.docx"
    result = sweep(source, out)  # public edit_document route, Word worker subprocess not patched
    assert result.status == "complete", result.message
    before, after = document_xml(source), document_xml(out)
    assert _field_result(out) == _field_result(source)
    head = after[:after.rindex('w:fldCharType="end"')]
    assert head == before[:before.rindex('w:fldCharType="end"')]
    texts = [view(item, accept_new=True) for item in body_paragraphs(out)]
    assert texts[:4] == [view(item, accept_new=True) for item in body_paragraphs(source)][:4]
    assert texts[2] == f"three{NBSP}field" and texts[4] == "outside text"
    details = result.operations[0].details
    assert details["paragraphs_changed"] == [5] and details["skipped"]["field"] == 4
    assert inspect_document(source)["unusual_characters"]["total"] == 1


@pytest.mark.parametrize("code", _FIELD_CODES)
def test_sweep_keeps_field_with_revised_boundaries_in_one_paragraph(tmp_path, code):
    raw = (_fld("begin") + _instr(code) + _fld("separate") + _r("one") + _deleted(_fld("end"), 7)
           + _r(f"two{NBSP}field") + _DELETED_BEGIN + _r("three") + _fld("end"))
    source = build(tmp_path / "revised.docx", [raw + _r(f"outside{NBSP}text")])
    out = tmp_path / "out.docx"
    result = sweep(source, out)
    assert result.status == "complete", result.message
    assert _field_result(out) == _field_result(source)
    assert view(body_paragraphs(out)[0], accept_new=True) == f"onetwo{NBSP}fieldthreeoutside text"
    assert result.operations[0].details["skipped"]["field"] == 1


def test_sweep_cleans_prose_after_a_cross_paragraph_field_ends(tmp_path):
    source = build(tmp_path / "end.docx", [_fld("begin") + _instr(" QUOTE ") + _fld("separate") + _r(f"in{NBSP}it"),
                                           _r(f"still{NBSP}in") + _fld("end") + _r(f"after{NBSP}it")])
    out = tmp_path / "out.docx"
    result = sweep(source, out)
    assert result.status == "complete", result.message
    assert [view(item, accept_new=True) for item in body_paragraphs(out)] == [f"in{NBSP}it", f"still{NBSP}inafter it"]
    assert result.operations[0].details["skipped"] == {"field": 2}


@pytest.mark.parametrize("operation", [
    dict(op_id="r", type="tracked_replace_text", paragraph_index=4, find="z", replace="zA\vB"),
    dict(op_id="p", type="tracked_replace_paragraph", paragraph_index=4, text=f"one\vtwo{NBSP}three"),
])
def test_manual_line_break_survives_clean_on_write(doc, tmp_path, operation):
    """U+000B is Word's own manual line break (<w:br/>); char_map strips it for other MCPs, never here."""
    out = tmp_path / "out.docx"
    result = edit_document(request(doc, operation, output=out))
    assert result.status == "complete", result.message
    accepted = view(body_paragraphs(out)[3], accept_new=True)
    assert accepted == (MANUSCRIPT[:-1] + "zA\vB" if operation["type"] == "tracked_replace_text" else "one\vtwo three")
    assert "<w:br/>" in document_xml(out)
    assert "U+000B" not in codepoints(result.cleaned)


def test_manual_line_break_kept_in_insert_text(monkeypatch, doc, tmp_path):
    seen = captured(monkeypatch)
    result = edit_document(request(doc, dict(op_id="i", type="insert_text", paragraph_index=4, position="after",
                                             anchor="z", text=f" one\vtwo{NBSP}"),
                                   dict(op_id="p", type="insert_paragraph", after_paragraph_index=1, text="a\vb"),
                                   output=tmp_path / "out.docx"))
    assert result.status == "complete", result.message
    assert [item.text for item in seen] == [" one\vtwo ", "a\vb"]
    assert codepoints(result.cleaned) == {"U+00A0"}


def test_sweep_whole_document(doc, tmp_path):
    out = tmp_path / "out.docx"
    result = sweep(doc, out)
    assert result.status == "complete", result.message
    texts = [view(item, accept_new=True) for item in body_paragraphs(out)]
    assert texts == ["Plain 10 mm text ab.", "Let  be small.", f"Prior work (Smith,{NBSP}2020) says so.",
                     MANUSCRIPT, f"Kept their{NBSP}words end."]
    rejected = [view(item, accept_new=False) for item in body_paragraphs(out)]
    assert rejected == [view(item, accept_new=False) for item in body_paragraphs(doc)]
    details = result.operations[0].details
    assert details["skipped"] == {"math": 2, "field": 2, "revision": 1}
    assert codepoints(result.cleaned) == {"U+00A0", "U+200B"}
    assert details["paragraphs_changed"] == [1, 2, 3]
    # the other author's insertion is byte-identical
    assert re.findall(r'<w:ins w:id="1".*?</w:ins>', document_xml(out), re.S) == \
        re.findall(r'<w:ins w:id="1".*?</w:ins>', document_xml(doc), re.S)


def test_sweep_keep(doc, tmp_path):
    out = tmp_path / "out.docx"
    result = sweep(doc, out, paragraph_index=1, keep=["U+00A0"])
    assert view(body_paragraphs(out)[0], accept_new=True) == f"Plain 10{NBSP}mm text ab."
    assert codepoints(result.cleaned) == {"U+200B"}


def test_sweep_preview_and_inspect_count(doc, tmp_path):
    original = doc.read_bytes()
    result = sweep(doc, tmp_path / "out.docx", preview_only=True)
    assert result.status == "complete", result.message
    assert result.output is None and doc.read_bytes() == original
    assert not (tmp_path / "out.docx").exists()
    details = result.operations[0].details
    assert details["read_only"] and details["preview_only"] and details["total"] == 5
    report = inspect_document(doc)["unusual_characters"]
    assert report["total"] == 5
    assert report["by_codepoint"]["U+00A0"]["count"] == 4
    assert report["by_codepoint"]["U+200B"]["count"] == 1
    assert report["skipped_total"] == 5
    # a clean document needs nothing
    clean = build(tmp_path / "plain.docx", [_r(MANUSCRIPT)])
    assert inspect_document(clean)["unusual_characters"]["total"] == 0
    nothing = sweep(clean, tmp_path / "out2.docx")
    assert nothing.output is None and nothing.operations[0].details["revisions_added"] == 0


def test_sweep_in_place_is_checkpointed(doc):
    result = sweep(doc, None)
    assert result.status == "complete", result.message
    assert result.undo_token
    assert view(body_paragraphs(doc)[0], accept_new=True) == "Plain 10 mm text ab."


def test_mcp_envelope_accepts_sweep_and_keep():
    from research_mcp.server import MCPEditOperation, _field_contract, _strict_operation_payload

    op = _strict_operation_payload(MCPEditOperation(op_id="s", type="sweep_characters", preview_only=True))
    assert op.paragraph_index is None and op.preview_only and op.author == "Claude"
    op = _strict_operation_payload(MCPEditOperation(op_id="s", type="sweep_characters", paragraph_index=3,
                                                    through_paragraph_index=5, keep=["U+00A0"]))
    assert (op.paragraph_index, op.through_paragraph_index, op.keep) == (3, 5, ["U+00A0"])
    op = _strict_operation_payload(MCPEditOperation(op_id="i", type="insert_text", anchor="a", text="b",
                                                    keep=["\u00a0"]))
    assert op.keep == ["\u00a0"]
    assert "keep" in _field_contract("sweep_characters")[1]
    assert "keep" in _field_contract("replace_text")[1]
    assert "keep" not in _field_contract("accept_revisions")[1]
