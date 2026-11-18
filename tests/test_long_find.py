"""Plan 864 decision 14: find/anchor strings longer than Word Find's 255 characters.

The schema takes up to FIND_MAX; XML paths need nothing more; Word paths find the first FIND_LIMIT characters,
extend the range to the full length and compare Range.Text (counting every match). Word itself is mocked here:
a fake document whose positions are characters of a raw string (hidden field code included, as Word's Range
positions are), so a span that crosses a field reads differently from the visible text."""

from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path

import pytest
from pydantic import ValidationError

from research_mcp.audit import file_sha256
from research_mcp.edit_models import (
    FIND_MAX,
    AddCommentOperation,
    EditDocumentRequest,
    InsertTextOperation,
    ReplaceTextOperation,
    TrackedReplaceTextOperation,
)
from research_mcp.editing import FIND_LIMIT, LONG_FIND_CROSSES_FIELD, _locate, locate_anchor
from research_mcp.errors import RefusedUnsafe

windows_only = pytest.mark.skipif(os.name != "nt", reason="undo store writes need Windows Python")

LONG = ("The cohort included consecutive adult patients scanned between January and June at two sites, and "
        "every examination was read independently by two board-certified readers who were blinded to the "
        "clinical outcome, the reference standard and each other's reads; disagreements were settled by a "
        "third reader with ten years of experience in the same subspecialty.")
assert 300 < len(LONG) < 500


# -- schema ------------------------------------------------------------------------------------------------

def test_schema_takes_long_find_and_anchor_up_to_the_cap():
    ReplaceTextOperation(op_id="a", type="replace_text", paragraph_index=1, find=LONG, replace="x")
    TrackedReplaceTextOperation(op_id="a", type="tracked_replace_text", paragraph_index=1, find=LONG, replace="x")
    InsertTextOperation(op_id="a", type="insert_text", paragraph_index=1, position="after", anchor=LONG, text="x")
    AddCommentOperation(op_id="a", type="add_comment", paragraph_index=1, anchor=LONG, text="note")
    ReplaceTextOperation(op_id="a", type="replace_text", paragraph_index=1, find="x" * FIND_MAX, replace="y")
    with pytest.raises(ValidationError, match=f"at most {FIND_MAX} characters"):
        ReplaceTextOperation(op_id="a", type="replace_text", paragraph_index=1, find="x" * (FIND_MAX + 1),
                             replace="y")


def test_server_envelope_takes_a_long_find():
    from research_mcp.server import MCPEditOperation, _strict_operation_payload

    op = _strict_operation_payload(MCPEditOperation(op_id="a", type="replace_text", paragraph_index=2, find=LONG,
                                                    replace="Short."))
    assert op.find == LONG
    anchor = _strict_operation_payload(MCPEditOperation(op_id="b", type="insert_text", paragraph_index=2,
                                                        position="after", anchor=LONG, text=" More."))
    assert anchor.anchor == LONG


def test_style_text_validation_accepts_long_text():
    from research_mcp.inline_style import _validate_request

    _validate_request("Intro. " + LONG + " End.", LONG, allow_whole_paragraph=False)
    with pytest.raises(RefusedUnsafe, match=f"at most {FIND_MAX}"):
        _validate_request("x" * (FIND_MAX + 1), "x" * (FIND_MAX + 1), allow_whole_paragraph=True)


# -- Word path, mocked -------------------------------------------------------------------------------------

class _Find:
    def __init__(self, rng):
        self._range = rng
        self.Text = ""

    def ClearFormatting(self):
        pass

    def Execute(self):
        if len(self.Text) > FIND_LIMIT:
            raise AssertionError("Word's Find was given more than 255 characters")
        doc, start, end = self._range.doc, self._range.Start, self._range.End
        at = doc.raw.find(self.Text, start, end)
        if at < 0:
            return False
        self._range.Start, self._range.End = at, at + len(self.Text)
        return True


class _Range:
    def __init__(self, doc, start, end):
        self.doc, self.Start, self.End = doc, start, end
        self.Find = _Find(self)

    @property
    def Text(self):
        return self.doc.raw[self.Start:self.End]


class _Paragraphs:
    def __init__(self, doc):
        self.doc = doc

    def __call__(self, index):
        assert index == 1
        return type("P", (), {"Range": _Range(self.doc, 0, len(self.doc.raw))})()


class FakeDocument:
    def __init__(self, raw: str):
        self.raw = raw
        self.Paragraphs = _Paragraphs(self)

    def Range(self, start, end):
        return _Range(self, start, end)


def test_long_find_is_found_by_its_start_and_checked_in_full():
    raw = "Intro. " + LONG + " End.\r"
    found = _locate(FakeDocument(raw), 1, LONG)
    assert (found.Start, found.End, found.Text) == (7, 7 + len(LONG), LONG)
    assert locate_anchor(FakeDocument(raw), 1, LONG).Text == LONG


def test_long_find_counts_every_full_match():
    raw = LONG + " / " + LONG + "\r"
    with pytest.raises(RefusedUnsafe, match="occurs 2 times in paragraph 1; it must occur exactly once"):
        _locate(FakeDocument(raw), 1, LONG)


def test_prefix_twice_but_full_text_once_is_one_match():
    other = LONG[:FIND_LIMIT] + " but then it ends differently."
    found = _locate(FakeDocument(other + " " + LONG + "\r"), 1, LONG)
    assert found.Text == LONG and found.Start == len(other) + 1


def test_long_find_across_hidden_field_code_refuses_naming_replace_paragraph():
    cut = 300  # hidden field code sits inside the span, so the extended range reads differently
    raw = LONG[:cut] + "\x13 ADDIN ZOTERO_ITEM {} \x14" + LONG[cut:] + "\r"
    with pytest.raises(RefusedUnsafe) as caught:
        _locate(FakeDocument(raw), 1, LONG)
    assert str(caught.value) == LONG_FIND_CROSSES_FIELD.format(length=len(LONG))
    assert "replace_paragraph with paragraph_index" in str(caught.value)


def test_long_find_not_present_says_so():
    with pytest.raises(RefusedUnsafe, match=r"could not find the \d+-character text"):
        _locate(FakeDocument("Something else entirely.\r"), 1, LONG)


# -- XML paths (documents with revisions, tracked ops, comments) --------------------------------------------

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def make_docx(path: Path, *paragraph_texts: str, revision: bool = True) -> Path:
    runs = "".join(f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>' for text in paragraph_texts)
    if revision:  # one other author's insertion, so the document "has revisions"
        runs += ('<w:p><w:ins w:id="1" w:author="Sam" w:date="2026-11-01T10:00:00Z"><w:r><w:t>Added.</w:t></w:r>'
                 "</w:ins></w:p>")
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}"><w:body>'
                f"{runs}<w:sectPr/></w:body></w:document>")
    with zipfile.ZipFile(path, "x", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml",
                         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
                         'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                         '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
                         'officedocument.wordprocessingml.document.main+xml"/></Types>')
        archive.writestr("_rels/.rels",
                         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                         'relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        archive.writestr("word/_rels/document.xml.rels",
                         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        archive.writestr("word/document.xml", document)
    return path


def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def test_tracked_replace_with_a_long_find(tmp_path):
    from research_mcp.revisions import text_edit

    src = make_docx(tmp_path / "in.docx", "Intro. " + LONG + " End.")
    out = tmp_path / "out.docx"
    text_edit(src, out, op="tracked_replace_text", paragraph_index=1, tracked=True, author="Reviewer B",
              find=LONG, replace="Two readers read every scan.")
    xml = document_xml(out)
    deleted = "".join(re.findall(r"<w:delText[^>]*>([^<]*)<", xml))
    inserted = "".join("".join(re.findall(r"<w:t[^>]*>([^<]*)<", body)) for body in
                       re.findall(r'<w:ins w:id="\d+" w:author="Reviewer B"[^>]*>(.*?)</w:ins>', xml, re.S))
    assert re.search(r'<w:del w:id="\d+" w:author="Reviewer B"', xml)
    # a word diff: deleted + kept words read as LONG, kept + inserted words read as the replacement
    kept = "".join(re.findall(r"<w:t[^>]*>([^<]*)<", re.sub(r"<w:(ins|del)\b.*?</w:\1>", "", xml, flags=re.S)))
    assert kept.startswith("Intro. ") and " End." in kept  # untouched words stay outside the revisions
    assert len(deleted) > 200 and inserted  # the long span really was found and rewritten
    from research_mcp.paragraphs import read_paragraphs
    assert read_paragraphs(out)[0].text == "Intro. Two readers read every scan. End."
    first = xml[xml.index("<w:p>"):xml.index("</w:p>")]  # Reject: drop the insertions, revive the deletions
    rejected = re.sub(r"<w:ins\b.*?</w:ins>", "", first, flags=re.S)
    assert "".join(re.findall(r"<w:(?:t|delText)\b[^>]*>([^<]*)<", rejected)) == "Intro. " + LONG + " End."


@windows_only
def test_edit_document_plain_replace_long_find_on_a_revised_document(tmp_path):
    from research_mcp.edit_batch import edit_document
    from research_mcp.paragraphs import read_paragraphs

    src = make_docx(tmp_path / "in.docx", "Intro. " + LONG + " End.")
    result = edit_document(EditDocumentRequest(source=str(src), source_sha256=file_sha256(src), output=str(src),
                                               operations=[{"op_id": "a", "type": "replace_text",
                                                            "paragraph_index": 1, "find": LONG,
                                                            "replace": "Two readers read every scan."}]))
    assert result.status == "complete", result
    assert read_paragraphs(src)[0].text == "Intro. Two readers read every scan. End."


@windows_only
def test_add_comment_on_a_long_anchor(tmp_path):
    from research_mcp.edit_batch import edit_document

    src = make_docx(tmp_path / "in.docx", "Intro. " + LONG + " End.")
    result = edit_document(EditDocumentRequest(source=str(src), source_sha256=file_sha256(src), output=str(src),
                                               operations=[{"op_id": "c", "type": "add_comment",
                                                            "paragraph_index": 1, "anchor": LONG,
                                                            "text": "Too long; split it."}]))
    assert result.status == "complete", result.model_dump_json(include={"status", "message", "operations"})
    xml = document_xml(src)
    start, end = xml.index("<w:commentRangeStart"), xml.index("<w:commentRangeEnd")
    assert "".join(re.findall(r"<w:t[^>]*>([^<]*)<", xml[start:end])) == LONG
