from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile


from research_mcp.comments import inspect_comments
from research_mcp.paragraphs import inspect_document, read_snapshot
from research_mcp.ooxml import NS, W


W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
W16CID = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"


def test_document_without_comments_returns_an_empty_thread_list(minimal_docx):
    snapshot = read_snapshot(minimal_docx)
    assert inspect_comments(snapshot) == ()
    result = inspect_document(minimal_docx, include_comments=True)
    assert result["comments"] == []
    assert result["comment_thread_count"] == 0


def _comment_thread_fixture(source: Path, output: Path) -> Path:
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}

    document = ET.fromstring(parts["word/document.xml"])
    paragraph = document.find(".//w:body/w:p", NS)
    first_run = paragraph.find("w:r", NS)
    start = ET.Element(W + "commentRangeStart", {W + "id": "5"})
    end = ET.Element(W + "commentRangeEnd", {W + "id": "5"})
    reference_run = ET.Element(W + "r")
    ET.SubElement(reference_run, W + "commentReference", {W + "id": "5"})
    position = list(paragraph).index(first_run)
    paragraph.insert(position, start)
    paragraph.insert(position + 2, end)
    paragraph.insert(position + 3, reference_run)
    parts["word/document.xml"] = ET.tostring(document, encoding="utf-8", xml_declaration=True)

    comments = ET.Element(W + "comments")
    top = ET.SubElement(
        comments,
        W + "comment",
        {W + "id": "5", W + "author": "Reviewer", W + "initials": "R"},
    )
    top_p = ET.SubElement(top, W + "p", {W14 + "paraId": "AAAA0001"})
    top_r = ET.SubElement(top_p, W + "r")
    ET.SubElement(top_r, W + "t").text = "Please clarify."
    reply = ET.SubElement(
        comments,
        W + "comment",
        {W + "id": "6", W + "author": "Author", W + "initials": "A"},
    )
    reply_p = ET.SubElement(reply, W + "p", {W14 + "paraId": "BBBB0002"})
    reply_r = ET.SubElement(reply_p, W + "r")
    ET.SubElement(reply_r, W + "t").text = "Clarified in the revision."
    parts["word/comments.xml"] = ET.tostring(comments, encoding="utf-8", xml_declaration=True)

    extended = ET.Element(W15 + "commentsEx")
    ET.SubElement(extended, W15 + "commentEx", {W15 + "paraId": "AAAA0001", W15 + "done": "0"})
    ET.SubElement(
        extended,
        W15 + "commentEx",
        {W15 + "paraId": "BBBB0002", W15 + "paraIdParent": "AAAA0001", W15 + "done": "0"},
    )
    parts["word/commentsExtended.xml"] = ET.tostring(extended, encoding="utf-8", xml_declaration=True)

    ids = ET.Element(W16CID + "commentsIds")
    ET.SubElement(ids, W16CID + "commentId", {W16CID + "paraId": "AAAA0001", W16CID + "durableId": "11111111"})
    ET.SubElement(ids, W16CID + "commentId", {W16CID + "paraId": "BBBB0002", W16CID + "durableId": "22222222"})
    parts["word/commentsIds.xml"] = ET.tostring(ids, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return output


def test_thread_parser_links_replies_to_durable_parent(minimal_docx, tmp_path):
    fixture = _comment_thread_fixture(minimal_docx, tmp_path / "comments.docx")
    threads = inspect_comments(read_snapshot(fixture))
    assert len(threads) == 1
    assert threads[0].comment_id == "11111111"
    assert threads[0].legacy_id == "5"
    assert threads[0].anchor_paragraph_index == 1
    assert threads[0].anchor_text == "Fixture paragraph."
    assert len(threads[0].replies) == 1
    assert threads[0].replies[0].comment_id == "22222222"
    assert threads[0].replies[0].text == "Clarified in the revision."
