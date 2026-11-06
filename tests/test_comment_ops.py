"""Plan 562 S1: Word comments as direct XML (add, reply, resolve, reopen, delete) on revised documents.

Every test drives the operations the way edit_document does (validated EditOperation ->
edit_batch._apply_operation_local), so on the old code add/reply on a revised document are refused and the
new operation types do not validate: each test fails there.
"""

from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path

import pytest
from pydantic import TypeAdapter, ValidationError

from research_mcp.audit import file_sha256
from research_mcp.comments import inspect_comments
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_tracked_fixture
from research_mcp.paragraphs import read_snapshot

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
COMMENT_PARTS = ("word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml",
                 "word/commentsExtensible.xml")
REV = re.compile(r"<w:(ins|del|moveFrom|moveTo)\b[^>]*?(?:/>|>.*?</w:\1>)|<w:(rPrChange|pPrChange|tcPrChange)\b"
                 r".*?</w:\2>", re.S)
HEX8 = re.compile(r"^[0-9A-F]{8}$")


# -- helpers ------------------------------------------------------------------------------------------------

def apply(source: Path, output: Path, **op) -> dict:
    from research_mcp.edit_batch import _apply_operation_local
    from research_mcp.edit_models import EditOperation

    op.setdefault("op_id", "c1")
    if op.get("type") == "reply_comment" and op.get("disposition", "addressed") == "addressed":
        op.setdefault("depends_on", [])
    operation = TypeAdapter(EditOperation).validate_python(op)
    return _apply_operation_local(Path(source), Path(output), operation)


def parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def text(path: Path, name: str) -> str:
    return parts(path)[name].decode("utf-8")


def threads(path: Path) -> dict[str, dict]:
    return {thread.legacy_id: thread.to_dict() for thread in inspect_comments(read_snapshot(path))}


def revisions_raw(path: Path) -> list[str]:
    return [match.group(0) for match in REV.finditer(text(path, "word/document.xml"))]


def strip_comment(xml: str, legacy: str) -> str:
    xml = re.sub(r'<w:commentRange(?:Start|End) w:id="%s"/>' % legacy, "", xml)
    return re.sub(r'<w:r>(?:<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>)?<w:commentReference '
                  r'w:id="%s"/></w:r>' % legacy, "", xml)


def visible_text(xml: str) -> str:
    """Accepted-view text of document.xml: w:t outside w:del/w:moveFrom (run splits do not show)."""

    xml = re.sub(r"<w:(del|moveFrom)\b[^>]*>.*?</w:\1>", "", xml, flags=re.S)
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", xml))


def new_legacy(source: Path, output: Path) -> str:
    added = set(threads(output)) - set(threads(source))
    assert len(added) == 1, added
    return added.pop()


@pytest.fixture
def tracked(tmp_path: Path) -> Path:
    return create_tracked_fixture(tmp_path / "tracked.docx")


ROOT_A, REPLY_A, ROOT_B, REPLY_B = "1A000001", "1A000002", "1B000001", "1B000002"
DURABLE = {"0": "2A000001", "41": "2A000002", "42": "2B000001", "43": "2B000002"}


def make_rich(path: Path) -> Path:
    """Word-style comment package on the tracked fixture: open thread A (legacy 0) with reply 41, resolved
    thread B (legacy 42) with reply 43 on table cell "Cell B"; all four comment parts; reply 43 has no
    commentsExtensible entry (as in real documents)."""

    base = create_tracked_fixture(path.with_name(path.stem + "_base.docx"))
    data = parts(base)
    doc = data["word/document.xml"].decode()
    doc = doc.replace('<w:commentRangeStart w:id="0"/>', '<w:commentRangeStart w:id="0"/><w:commentRangeStart '
                      'w:id="41"/>')
    doc = doc.replace('<w:r><w:commentReference w:id="0"/></w:r>',
                      '<w:r><w:commentReference w:id="0"/></w:r><w:commentRangeEnd w:id="41"/>'
                      '<w:r><w:commentReference w:id="41"/></w:r>')
    doc = doc.replace('<w:p><w:r><w:t xml:space="preserve">Cell B</w:t></w:r></w:p>',
                      '<w:p><w:commentRangeStart w:id="42"/><w:commentRangeStart w:id="43"/><w:r><w:t '
                      'xml:space="preserve">Cell B</w:t></w:r><w:commentRangeEnd w:id="42"/><w:r>'
                      '<w:commentReference w:id="42"/></w:r><w:commentRangeEnd w:id="43"/><w:r>'
                      '<w:commentReference w:id="43"/></w:r></w:p>')
    ns = ('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
          'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          f'xmlns:w14="{W14}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
          'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
          'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
          'mc:Ignorable="w14 w15 w16cid w16cex"')

    def comment(legacy, author, body, *paras):
        ps = "".join(f'<w:p w14:paraId="{para}" w14:textId="77777777"><w:r><w:t>{line}</w:t></w:r></w:p>'
                     for para, line in zip(paras, body))
        return f'<w:comment w:id="{legacy}" w:author="{author}" w:date="2026-10-10T17:37:00Z" w:initials="OA">{ps}</w:comment>'

    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    comments = (head + f"<w:comments {ns}>"
                + comment("0", "Other Author", ["Check this phrase."], ROOT_A)
                + comment("41", "Second Reviewer", ["First line &amp; more", "Agree."], "1A0000F1", REPLY_A)
                + comment("42", "Other Author", ["Cell wording?"], ROOT_B)
                + comment("43", "Second Reviewer", ["Fixed."], REPLY_B) + "</w:comments>")
    extended = (head + f"<w15:commentsEx {ns}>"
                f'<w15:commentEx w15:paraId="{ROOT_A}" w15:done="0"/>'
                f'<w15:commentEx w15:paraId="{REPLY_A}" w15:paraIdParent="{ROOT_A}" w15:done="0"/>'
                f'<w15:commentEx w15:paraId="{ROOT_B}" w15:done="1"/>'
                f'<w15:commentEx w15:paraId="{REPLY_B}" w15:paraIdParent="{ROOT_B}" w15:done="1"/>'
                "</w15:commentsEx>")
    ids = (head + f"<w16cid:commentsIds {ns}>"
           + "".join(f'<w16cid:commentId w16cid:paraId="{para}" w16cid:durableId="{DURABLE[legacy]}"/>'
                     for legacy, para in (("0", ROOT_A), ("41", REPLY_A), ("42", ROOT_B), ("43", REPLY_B)))
           + "</w16cid:commentsIds>")
    extensible = (head + f"<w16cex:commentsExtensible {ns}>"
                  + "".join(f'<w16cex:commentExtensible w16cex:durableId="{DURABLE[legacy]}" '
                            f'w16cex:dateUtc="2026-10-10T21:37:00Z"/>' for legacy in ("0", "41", "42"))
                  + "</w16cex:commentsExtensible>")
    types = data["[Content_Types].xml"].decode().replace("</Types>", "".join(
        f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        f'wordprocessingml.{name}+xml"/>' for name in ("commentsExtended", "commentsIds", "commentsExtensible"))
        + "</Types>")
    rels = data["word/_rels/document.xml.rels"].decode().replace("</Relationships>", (
        '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" '
        'Target="commentsExtended.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" '
        'Target="commentsIds.xml"/>'
        '<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2018/08/relationships/'
        'commentsExtensible" Target="commentsExtensible.xml"/></Relationships>'))
    data.update({"word/document.xml": doc.encode(), "word/comments.xml": comments.encode(),
                 "word/commentsExtended.xml": extended.encode(), "word/commentsIds.xml": ids.encode(),
                 "word/commentsExtensible.xml": extensible.encode(), "[Content_Types].xml": types.encode(),
                 "word/_rels/document.xml.rels": rels.encode()})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in data.items():
            archive.writestr(name, value)
    return path


@pytest.fixture
def rich(tmp_path: Path) -> Path:
    return make_rich(tmp_path / "rich.docx")


def entries(xml: str) -> list[str]:
    """Child elements of the part root, raw."""

    body = re.sub(r"^.*?<(\w+:\w+)\b[^>]*>(.*)</\1>\s*$", r"\2", xml, flags=re.S)
    return re.findall(r"<(\w+:\w+)\b[^>]*?(?:/>|>.*?</\1>)", body, flags=re.S) and [
        match.group(0) for match in re.finditer(r"<(\w+:\w+)\b[^>]*?(?:/>|>.*?</\1>)", body, flags=re.S)]


# -- B1: comments on documents with tracked changes ---------------------------------------------------------

def test_add_comment_across_insertion_and_deletion(tracked, tmp_path):
    out = tmp_path / "out.docx"
    before_doc = text(tracked, "word/document.xml")
    result = apply(tracked, out, type="add_comment", paragraph_index=2, anchor="highly accurate on",
                   text="Is this the held-out set?")
    legacy = new_legacy(tracked, out)
    thread = threads(out)[legacy]
    assert (thread["anchor_paragraph_index"], thread["anchor_text"]) == (2, "highly accurate on")
    assert (thread["author"], thread["initials"], thread["done"]) == ("Test Author", "TA", False)
    assert thread["text"] == "Is this the held-out set?" and HEX8.match(thread["comment_id"])
    assert result["comment_id"] == thread["comment_id"] and result["thread_open"] is True
    # every revision byte-identical, visible text unchanged, nothing but the markers added
    assert revisions_raw(out) == revisions_raw(tracked)
    after_doc = text(out, "word/document.xml")
    assert visible_text(after_doc) == visible_text(before_doc)
    assert "w:ins" not in after_doc[after_doc.index(f'w:commentReference w:id="{legacy}"') - 200:][:200] or \
        after_doc.rfind("</w:ins>", 0, after_doc.index(f'w:commentReference w:id="{legacy}"')) > \
        after_doc.rfind("<w:ins ", 0, after_doc.index(f'w:commentReference w:id="{legacy}"'))
    # the start marker sits before the w:ins (edge at the insertion's first character is lifted out)
    start = after_doc.index(f'<w:commentRangeStart w:id="{legacy}"/>')
    assert after_doc[start:].split(">", 1)[1].startswith('<w:ins w:id="1"')


def test_add_comment_inside_insertion(tracked, tmp_path):
    out = tmp_path / "out.docx"
    apply(tracked, out, type="add_comment", paragraph_index=2, anchor="ighly", text="Hedge?")
    legacy = new_legacy(tracked, out)
    assert threads(out)[legacy]["anchor_text"] == "ighly"
    after = text(out, "word/document.xml")
    ins = re.search(r'<w:ins w:id="1"[^>]*>.*?</w:ins>', after, re.S).group(0)
    old_ins = re.search(r'<w:ins w:id="1"[^>]*>.*?</w:ins>', text(tracked, "word/document.xml"), re.S).group(0)
    # markers inside the insertion (as Word places them), same revision record and text
    assert f'<w:commentRangeStart w:id="{legacy}"/>' in ins and f'<w:commentRangeEnd w:id="{legacy}"/>' in ins
    assert ins.split(">", 1)[0] == old_ins.split(">", 1)[0]
    assert "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", ins)) == "highly "
    # the reference run is outside every revision wrapper
    assert "commentReference" not in ins
    assert revisions_raw(out)[1:] == revisions_raw(tracked)[1:]
    from research_mcp.revisions import revision_count
    assert revision_count(out) == revision_count(tracked)
    # the one exception to raw identity: the insertion holding the edges differs only by the comment markers
    # and the run split at them (same w:id/author/date, same text and run formatting)
    assert ins != old_ins
    assert merge_runs(strip_comment(ins, legacy)) == merge_runs(old_ins)


def test_add_keeps_every_revision_byte_identical_in_a_table_cell(tracked, tmp_path):
    out = tmp_path / "out.docx"
    apply(tracked, out, type="add_comment", paragraph_index=11, anchor="Cell added A", text="Table note")
    legacy = new_legacy(tracked, out)
    assert threads(out)[legacy]["anchor_text"] == "Cell added A"
    assert revisions_raw(out) == revisions_raw(tracked)


# -- B2: files with resolved threads ------------------------------------------------------------------------

def test_add_keeps_resolved_threads_byte_identical(rich, tmp_path):
    out = tmp_path / "out.docx"
    before = threads(rich)
    apply(rich, out, type="add_comment", paragraph_index=1, anchor="opening", text="New point")
    after = threads(out)
    legacy = new_legacy(rich, out)
    assert {key: value for key, value in after.items() if key != legacy} == before
    assert after["42"]["done"] is True and after["42"]["replies"][0]["done"] is True
    for name in COMMENT_PARTS:
        old, new = entries(text(rich, name)), entries(text(out, name))
        assert new[: len(old)] == old, name  # every old entry byte-identical, same order; new ones appended
        assert len(new) == len(old) + 1, name
    assert revisions_raw(out) == revisions_raw(rich)


# -- B3: parts created when missing -------------------------------------------------------------------------

def test_add_creates_missing_parts(tracked, tmp_path):
    out = tmp_path / "out.docx"
    old_comment = re.search(r"<w:comment\b.*?</w:comment>", text(tracked, "word/comments.xml"), re.S).group(0)
    result = apply(tracked, out, type="add_comment", paragraph_index=1, anchor="Plain", text="Note")
    package = parts(out)
    for name in COMMENT_PARTS:
        assert name in package
    types = package["[Content_Types].xml"].decode()
    rels = package["word/_rels/document.xml.rels"].decode()
    for name in ("commentsExtended", "commentsIds", "commentsExtensible"):
        assert f'PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.' \
               f'wordprocessingml.{name}+xml"' in types
        assert f'Target="{name}.xml"' in rels
    comments = package["word/comments.xml"].decode()
    assert f'xmlns:w14="{W14}"' in comments and old_comment in comments
    legacy = new_legacy(tracked, out)
    thread = threads(out)[legacy]
    assert HEX8.match(thread["comment_id"]) and result["comment_id"] == thread["comment_id"]
    assert threads(out)["0"] == threads(tracked)["0"]  # the old comment still reads as legacy:0
    assert set(result["parts_created"]) == set(COMMENT_PARTS[1:])


def test_add_on_document_without_comments_part(tracked, tmp_path):
    bare = tmp_path / "bare.docx"
    data = parts(tracked)
    del data["word/comments.xml"]
    data["word/document.xml"] = strip_comment(data["word/document.xml"].decode(), "0").encode()
    data["word/_rels/document.xml.rels"] = re.sub(rb'<Relationship Id="rId2"[^>]*/>', b"",
                                                  data["word/_rels/document.xml.rels"])
    data["[Content_Types].xml"] = re.sub(rb'<Override PartName="/word/comments.xml"[^>]*/>', b"",
                                         data["[Content_Types].xml"])
    with zipfile.ZipFile(bare, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in data.items():
            archive.writestr(name, value)
    out = tmp_path / "out.docx"
    apply(bare, out, type="add_comment", paragraph_index=4, anchor="Commented phrase",
          text="First line\nSecond line\n\nLast paragraph")
    [thread] = threads(out).values()
    assert thread["text"] == "First line\nSecond line\n\nLast paragraph"
    assert thread["anchor_text"] == "Commented phrase"
    assert "/word/comments.xml" in text(out, "[Content_Types].xml")


# -- B4/B5: replies -----------------------------------------------------------------------------------------

@pytest.mark.parametrize(("disposition", "label"), [("addressed", "Addressed"), ("rejected", "Rejected"),
                                                    ("no_change", "No change"),
                                                    ("needs_information", "Needs information")])
def test_reply_each_disposition(rich, tmp_path, disposition, label):
    out = tmp_path / "out.docx"
    result = apply(rich, out, type="reply_comment", comment_id=DURABLE["0"], disposition=disposition,
                   reply_text="  Reworded the sentence.  ")
    thread = threads(out)["0"]
    assert thread["done"] is False and result["thread_open"] is True
    assert len(thread["replies"]) == 2
    reply = thread["replies"][-1]
    assert reply["text"] == f"{label}: Reworded the sentence." == result["reply_text"]
    assert reply["author"] == "Test Author" and reply["done"] is False
    ext = text(out, "word/commentsExtended.xml")
    assert re.search(r'w15:paraIdParent="%s" w15:done="0"/>' % ROOT_A, ext.split(REPLY_A)[-1])
    # Word's layout: the reply's markers next to its thread's, after the last reply's reference
    doc = text(out, "word/document.xml")
    new = result["reply_legacy_id"]
    assert doc.index(f'<w:commentRangeStart w:id="{new}"/>') > doc.index('<w:commentRangeStart w:id="41"/>')
    assert doc.index(f'<w:commentReference w:id="{new}"/>') > doc.index('<w:commentReference w:id="41"/>')
    assert threads(out)["42"] == threads(rich)["42"]
    assert revisions_raw(out) == revisions_raw(rich)


def test_reply_on_revised_anchor_keeps_revisions(tracked, tmp_path):
    first, second = tmp_path / "a.docx", tmp_path / "b.docx"
    added = apply(tracked, first, type="add_comment", paragraph_index=2, anchor="ighly", text="Hedge?")
    apply(first, second, type="reply_comment", comment_id=added["comment_id"], disposition="no_change",
          reply_text="Kept.")
    assert threads(second)[added["legacy_id"]]["replies"][0]["text"] == "No change: Kept."
    doc = text(second, "word/document.xml")
    ins = re.search(r'<w:ins w:id="1"[^>]*>.*?</w:ins>', doc, re.S).group(0)
    assert "commentReference" not in ins
    assert revisions_raw(second)[1:] == revisions_raw(tracked)[1:]


def test_reply_and_resolve(rich, tmp_path):
    out = tmp_path / "out.docx"
    result = apply(rich, out, type="reply_comment", comment_id=DURABLE["0"], disposition="rejected",
                   reply_text="Out of scope.", resolve=True)
    thread = threads(out)["0"]
    assert thread["done"] is True and all(reply["done"] for reply in thread["replies"])
    assert thread["replies"][-1]["text"] == "Rejected: Out of scope."
    assert result["thread_open"] is False and result["resolved"] is True
    assert threads(out)["42"] == threads(rich)["42"]


def test_reply_to_resolved_thread_needs_reopen(rich, tmp_path):
    with pytest.raises(RefusedUnsafe, match="resolved; reopen_comment"):
        apply(rich, tmp_path / "out.docx", type="reply_comment", comment_id=DURABLE["42"],
              disposition="no_change", reply_text="x")
    assert not (tmp_path / "out.docx").exists()


# -- B6: resolve / reopen -----------------------------------------------------------------------------------

def test_resolve_then_reopen(rich, tmp_path):
    resolved, reopened = tmp_path / "r.docx", tmp_path / "o.docx"
    apply(rich, resolved, type="resolve_comment", comment_id=DURABLE["0"])
    thread = threads(resolved)["0"]
    assert thread["done"] is True and thread["replies"][0]["done"] is True
    before, after = parts(rich), parts(resolved)
    assert {name for name in before if before[name] != after[name]} == {"word/commentsExtended.xml"}
    old_ext, new_ext = entries(text(rich, "word/commentsExtended.xml")), entries(text(resolved,
                                                                                   "word/commentsExtended.xml"))
    assert old_ext[2:] == new_ext[2:]  # thread B's entries untouched
    assert new_ext[0] == f'<w15:commentEx w15:paraId="{ROOT_A}" w15:done="1"/>'
    with pytest.raises(RefusedUnsafe, match="already resolved"):
        apply(resolved, tmp_path / "again.docx", type="resolve_comment", comment_id=DURABLE["0"])
    apply(resolved, reopened, type="reopen_comment", comment_id=DURABLE["0"])
    assert threads(reopened) == threads(rich)
    assert parts(reopened) == parts(rich)


def test_resolve_reply_id_names_root(rich, tmp_path):
    with pytest.raises(RefusedUnsafe, match=f"reply in thread '{DURABLE['0']}'"):
        apply(rich, tmp_path / "out.docx", type="resolve_comment", comment_id=DURABLE["41"])


def test_reopen_resolved_thread(rich, tmp_path):
    out = tmp_path / "out.docx"
    apply(rich, out, type="reopen_comment", comment_id=DURABLE["42"])
    thread = threads(out)["42"]
    assert thread["done"] is False and thread["replies"][0]["done"] is False
    assert threads(out)["0"] == threads(rich)["0"]


# -- B7: delete ---------------------------------------------------------------------------------------------

def test_delete_thread_with_replies(rich, tmp_path):
    out = tmp_path / "out.docx"
    result = apply(rich, out, type="delete_comment", comment_id=DURABLE["0"])
    assert set(threads(out)) == {"42"} and threads(out)["42"] == threads(rich)["42"]
    doc = text(out, "word/document.xml")
    assert doc == strip_comment(strip_comment(text(rich, "word/document.xml"), "0"), "41")
    for name in COMMENT_PARTS:
        part = text(out, name)
        for value in (ROOT_A, REPLY_A, DURABLE["0"], DURABLE["41"]):
            assert value not in part, (name, value)
    assert entries(text(out, "word/commentsExtended.xml")) == entries(text(rich, "word/commentsExtended.xml"))[2:]
    assert sorted(result["deleted_legacy_ids"], key=int) == ["0", "41"]
    assert revisions_raw(out) == revisions_raw(rich)


def test_delete_single_reply(rich, tmp_path):
    out = tmp_path / "out.docx"
    apply(rich, out, type="delete_comment", comment_id=DURABLE["41"], include_replies=False)
    thread = threads(out)["0"]
    before = threads(rich)["0"]
    assert tuple(thread["replies"]) == () and {k: v for k, v in thread.items() if k != "replies"} == \
        {k: v for k, v in before.items() if k != "replies"}
    assert text(out, "word/document.xml") == strip_comment(text(rich, "word/document.xml"), "41")
    assert threads(out)["42"] == threads(rich)["42"]


def test_delete_reply_id_refused(rich, tmp_path):
    with pytest.raises(RefusedUnsafe, match=f"reply in thread '{DURABLE['0']}'.*include_replies=false"):
        apply(rich, tmp_path / "out.docx", type="delete_comment", comment_id=DURABLE["41"])


def test_delete_root_without_replies_flag_refused(rich, tmp_path):
    with pytest.raises(RefusedUnsafe, match="has 1 replies"):
        apply(rich, tmp_path / "out.docx", type="delete_comment", comment_id=DURABLE["0"],
              include_replies=False)
    assert not (tmp_path / "out.docx").exists()


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NONBODY_MARKERS = {
    "single_quotes": ("<w:commentRangeStart w:id='0' /><w:r><w:t>Story target</w:t></w:r><w:commentRangeEnd "
                      "w:id='0' /><w:r><w:commentReference w:id='0' /></w:r>"),
    "spaces_and_end_tags": ('<w:commentRangeStart  w:id = "0" ></w:commentRangeStart><w:r><w:t>Story target</w:t>'
                            '</w:r><w:commentRangeEnd\tw:id="0"/><w:r><w:commentReference w:id="0" /></w:r>'),
    "attribute_order": ('<w:commentRangeStart w:displacedByCustomXml="next" w:id="0"/><w:r><w:t>Story target</w:t>'
                        '</w:r><w:commentRangeEnd w:displacedByCustomXml="prev" w:id="0"/><w:r>'
                        '<w:commentReference w:customMarkFollows="0" w:id="0"/></w:r>'),
    "namespace_alias": (f'<x:commentRangeStart xmlns:x="{W_NS}" x:id="0"/><w:r><w:t>Story target</w:t></w:r>'
                        f'<x:commentRangeEnd xmlns:x="{W_NS}" x:id="0"/><w:r><x:commentReference xmlns:x="{W_NS}" '
                        'x:id="0"/></w:r>'),
}
STORIES = {
    "footnotes": ("word/footnotes.xml", "footnotes", '<w:footnotes xmlns:w="%s"><w:footnote w:id="2"><w:p>{}</w:p>'
                  "</w:footnote></w:footnotes>" % W_NS, "footnotes+xml"),
    "endnotes": ("word/endnotes.xml", "endnotes", '<w:endnotes xmlns:w="%s"><w:endnote w:id="2"><w:p>{}</w:p>'
                 "</w:endnote></w:endnotes>" % W_NS, "endnotes+xml"),
    "header": ("word/header7.xml", "header", '<w:hdr xmlns:w="%s"><w:p>{}</w:p></w:hdr>' % W_NS, "header+xml"),
    "footer": ("word/footer7.xml", "footer", '<w:ftr xmlns:w="%s"><w:p>{}</w:p></w:ftr>' % W_NS, "footer+xml"),
}


def nonbody_anchor(tmp_path: Path, story: str, spelling: str) -> Path:
    """sol61's footnote fixture, generalised: comment 0's only markers live in another story, spelled validly
    but not as `<w:commentX w:id="0"/>`."""
    name, rel, root, content_type = STORIES[story]
    data = parts(create_tracked_fixture(tmp_path / f"base-{story}-{spelling}.docx"))
    xml = data["word/document.xml"].decode()
    xml = re.sub(r"<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*/>", "", xml)
    if story in ("footnotes", "endnotes"):
        kind = "footnote" if story == "footnotes" else "endnote"
        xml = xml.replace("</w:body>", f'<w:p><w:r><w:{kind}Reference w:id="2"/></w:r></w:p></w:body>')
    data["word/document.xml"] = xml.encode()
    data[name] = root.format(NONBODY_MARKERS[spelling]).encode()
    data["word/_rels/document.xml.rels"] = data["word/_rels/document.xml.rels"].replace(
        b"</Relationships>", f'<Relationship Id="rIdStoryProbe" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/{rel}" Target="{name[5:]}"/></Relationships>'.encode())
    data["[Content_Types].xml"] = data["[Content_Types].xml"].replace(
        b"</Types>", f'<Override PartName="/{name}" ContentType="application/vnd.openxmlformats-'
        f'officedocument.wordprocessingml.{content_type}"/></Types>'.encode())
    path = tmp_path / f"{story}-{spelling}.docx"
    with zipfile.ZipFile(path, "w") as archive:
        for key, value in data.items():
            archive.writestr(key, value)
    return path


@pytest.mark.parametrize("spelling", sorted(NONBODY_MARKERS))
@pytest.mark.parametrize("story", sorted(STORIES))
def test_delete_comment_refuses_nonbody_anchor_whatever_the_spelling(tmp_path, story, spelling):
    """sol61 final review P2: a comment anchored in a footnote (or endnote, header, footer) is refused before
    anything is written, however its markers are spelled; deleting it would orphan those markers."""
    doc = nonbody_anchor(tmp_path, story, spelling)
    before = doc.read_bytes()
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=rf"comment 'legacy:0' is anchored in {STORIES[story][0]}; "
                                            "delete it in Word"):
        apply(doc, out, type="delete_comment", comment_id="legacy:0")
    assert not out.exists() and doc.read_bytes() == before


@pytest.mark.skipif(os.name != "nt", reason="undo store writes need Windows Python")
def test_delete_comment_refuses_nonbody_anchor_regardless_of_xml_quotes(tmp_path):
    """sol61's regression through the real batch API: refused, source bytes unchanged."""
    from test_revisions import run_batch

    doc = nonbody_anchor(tmp_path, "footnotes", "single_quotes")
    before = doc.read_bytes()
    result = run_batch(doc, {"op_id": "delete", "type": "delete_comment", "comment_id": "legacy:0"})
    assert result.status != "complete", "Nonbody reference was orphaned by a successfully published delete"
    assert "anchored in word/footnotes.xml; delete it in Word" in str(result.operations[0].error), result
    assert doc.read_bytes() == before


def test_comment_markers_found_structurally_in_every_story():
    from research_mcp.comment_ops import _story_markers

    package = {"word/document.xml": f'<w:document xmlns:w="{W_NS}"><w:body/></w:document>'.encode()}
    for story, (name, _, root, _) in STORIES.items():
        for spelling, markers in NONBODY_MARKERS.items():
            package[f"{name}.{spelling}.xml"] = root.format(markers).encode()
    package["word/styles.xml"] = f"<w:styles xmlns:w='{W_NS}'><w:style w:styleId='0'/></w:styles>".encode()
    found = _story_markers(package, {"0"})
    assert sorted(found) == sorted(f"{name}.{spelling}.xml" for name, *_ in STORIES.values()
                                   for spelling in NONBODY_MARKERS)
    assert all(found[name] == 3 for name in found)
    assert _story_markers(package, {"1"}) == {}


# -- B8: legacy ids -----------------------------------------------------------------------------------------

@pytest.mark.parametrize("comment_id", ["legacy:0", "0"])
def test_legacy_id_accepted(tracked, tmp_path, comment_id):
    out = tmp_path / "out.docx"
    result = apply(tracked, out, type="resolve_comment", comment_id=comment_id)
    thread = threads(out)["0"]
    assert thread["done"] is True and HEX8.match(thread["comment_id"])
    assert result["comment_id"] == thread["comment_id"]
    assert thread["text"] == "Check this phrase." and thread["anchor_text"] == "Commented phrase"
    assert text(out, "word/document.xml") == text(tracked, "word/document.xml")


def test_unknown_id_names_known_ids(tracked, tmp_path):
    with pytest.raises(RefusedUnsafe, match=r"matched no comment.*legacy:0"):
        apply(tracked, tmp_path / "out.docx", type="resolve_comment", comment_id="99")


# -- B9: refusals -------------------------------------------------------------------------------------------

@pytest.mark.parametrize(("index", "anchor", "message"), [
    (2, "overall", "deleted \\(tracked\\) text"),
    (3, "1). The", "cuts a field"),
    (3, "work (1", "cuts a field"),
    (2, "e", "not unique .*matches"),
    (2, "nowhere", "not found"),
    (5, "Bold", "tracked formatting change"),
])
def test_anchor_refusals(tracked, tmp_path, index, anchor, message):
    before = file_sha256(tracked)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=message):
        apply(tracked, out, type="add_comment", paragraph_index=index, anchor=anchor, text="x")
    assert not out.exists() and file_sha256(tracked) == before


def test_anchor_on_whole_field_result_keeps_the_field(tracked, tmp_path):
    out = tmp_path / "out.docx"
    apply(tracked, out, type="add_comment", paragraph_index=3, anchor="(1)", text="Which reference?")
    legacy = new_legacy(tracked, out)
    assert threads(out)[legacy]["anchor_text"] == "(1)"
    doc = text(out, "word/document.xml")
    start = doc.index(f'<w:commentRangeStart w:id="{legacy}"/>')
    end = doc.index(f'<w:commentRangeEnd w:id="{legacy}"/>')
    assert 'w:fldCharType="begin"' in doc[start:end] and 'w:fldCharType="end"' in doc[start:end]


def rewrite(source: Path, output: Path, change) -> Path:
    data = parts(source)
    change(data)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in data.items():
            archive.writestr(name, value)
    return output


OPENING_RUN = '<w:r><w:t xml:space="preserve">opening</w:t></w:r>'
SIMPLE_FIELD = '<w:fldSimple w:instr=" REF keep ">' + OPENING_RUN + '</w:fldSimple>'


@pytest.fixture
def simple_field(tracked: Path, tmp_path: Path) -> Path:
    def change(data):
        raw = data["word/document.xml"].decode()
        assert raw.count(OPENING_RUN) == 1
        data["word/document.xml"] = raw.replace(OPENING_RUN, SIMPLE_FIELD).encode()
    return rewrite(tracked, tmp_path / "simple.docx", change)


def test_anchor_on_whole_fldsimple_result_puts_markers_outside(simple_field, tmp_path):
    out = tmp_path / "out.docx"
    apply(simple_field, out, type="add_comment", paragraph_index=1, anchor="opening", text="Which section?")
    legacy = new_legacy(simple_field, out)
    assert threads(out)[legacy]["anchor_text"] == "opening"
    doc = text(out, "word/document.xml")
    # the field XML is unchanged and both markers sit outside it, around the whole field
    assert doc.count(SIMPLE_FIELD) == 1
    assert f'<w:commentRangeStart w:id="{legacy}"/>{SIMPLE_FIELD}<w:commentRangeEnd w:id="{legacy}"/>' in doc
    assert strip_comment(doc, legacy) == text(simple_field, "word/document.xml")


def test_anchor_on_part_of_fldsimple_result_refused(simple_field, tmp_path):
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="cuts a field"):
        apply(simple_field, out, type="add_comment", paragraph_index=1, anchor="penin", text="x")
    assert not out.exists()


def test_tracked_formatting_refusal_suggests_whole_run_anchor(tracked, tmp_path):
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="tracked formatting change") as refused:
        apply(tracked, out, type="add_comment", paragraph_index=5, anchor="Bold", text="x")
    assert "e.g. 'Bold words'" in str(refused.value)
    before = re.findall(r"<w:rPrChange\b.*?</w:rPrChange>", text(tracked, "word/document.xml"), re.S)
    apply(tracked, out, type="add_comment", paragraph_index=5, anchor="Bold words", text="Whole run")
    assert threads(out)[new_legacy(tracked, out)]["anchor_text"] == "Bold words"
    assert re.findall(r"<w:rPrChange\b.*?</w:rPrChange>", text(out, "word/document.xml"), re.S) == before
    assert revisions_raw(out) == revisions_raw(tracked)


# -- B10: addressed without depends_on ----------------------------------------------------------------------

def test_addressed_without_dependency_suggests_alternatives(tracked):
    from research_mcp.edit_models import EditDocumentRequest

    with pytest.raises(ValidationError) as caught:
        EditDocumentRequest(source=str(tracked), source_sha256="0" * 64, operations=[
            {"op_id": "r1", "type": "reply_comment", "comment_id": "A", "disposition": "addressed",
             "reply_text": "Done"}])
    message = str(caught.value)
    assert "requires" in message and "no_change" in message and "rejected" in message


# -- B11: author ---------------------------------------------------------------------------------------------

def test_author_default_and_override(tracked, tmp_path):
    default, named = tmp_path / "d.docx", tmp_path / "n.docx"
    apply(tracked, default, type="add_comment", paragraph_index=1, anchor="Plain", text="One")
    comments = text(default, "word/comments.xml")
    assert 'w:author="Test Author"' in comments and 'w:initials="TA"' in comments
    apply(default, named, type="add_comment", paragraph_index=1, anchor="opening", text="Two",
          author="Reviewer X", initials="RX")
    legacy = new_legacy(default, named)
    thread = threads(named)[legacy]
    assert (thread["author"], thread["initials"]) == ("Reviewer X", "RX")


def test_reply_author_override_written_to_comments_xml(rich, tmp_path):
    out = tmp_path / "out.docx"
    result = apply(rich, out, type="reply_comment", comment_id=DURABLE["0"], disposition="no_change",
                   reply_text="Kept as is.", author="Reviewer X", initials="RX")
    comments = text(out, "word/comments.xml")
    reply = re.search(r'<w:comment w:id="%s"[^>]*>' % result["reply_legacy_id"], comments).group(0)
    assert 'w:author="Reviewer X"' in reply and 'w:initials="RX"' in reply
    assert 'w:author="Test Author"' not in comments
    assert threads(out)["0"]["replies"][-1]["author"] == "Reviewer X"


def test_server_envelope_defaults():
    from research_mcp.server import MCPEditOperation, _strict_operation_payload

    add = _strict_operation_payload(MCPEditOperation(op_id="a", type="add_comment", paragraph_index=2,
                                                     anchor="x", text="y"))
    assert (add.author, add.initials) == (None, None)  # resolved when the op runs
    for given, expected in (("accepted", "addressed"), ("no_change", "no_change"),
                            ("needs_information", "needs_information"), ("rejected", "rejected")):
        reply = _strict_operation_payload(MCPEditOperation(op_id="r", type="reply_comment", comment_id="A",
                                                           disposition=given, reply_text="t", resolve=True))
        assert (reply.disposition, reply.resolve, reply.author) == (expected, True, None)
    delete = _strict_operation_payload(MCPEditOperation(op_id="d", type="delete_comment", comment_id="A",
                                                        include_replies=False))
    assert (delete.type, delete.include_replies) == ("delete_comment", False)
    for kind in ("resolve_comment", "reopen_comment"):
        assert _strict_operation_payload(MCPEditOperation(op_id="s", type=kind, comment_id="A")).type == kind
    tracked_edit = _strict_operation_payload(MCPEditOperation(op_id="e", type="tracked_replace_text",
                                                              paragraph_index=2, find="a", replace="b"))
    assert tracked_edit.author == "Claude"


# -- B12: publish / undo ------------------------------------------------------------------------------------

windows_only = pytest.mark.skipif(os.name != "nt", reason="undo store writes need Windows Python")


@windows_only
def test_edit_document_comment_ops_undo(rich):
    from research_mcp.edit_batch import edit_document
    from research_mcp.edit_models import EditDocumentRequest
    from research_mcp.publication import undo_document

    before_sha, before_threads = file_sha256(rich), threads(rich)
    result = edit_document(EditDocumentRequest(source=str(rich), source_sha256=before_sha, output=str(rich),
                                               operations=[
        {"op_id": "a", "type": "add_comment", "paragraph_index": 2, "anchor": "highly accurate", "text": "New"},
        {"op_id": "r", "type": "reply_comment", "comment_id": DURABLE["0"], "disposition": "no_change",
         "reply_text": "Kept", "resolve": True},
        {"op_id": "o", "type": "reopen_comment", "comment_id": DURABLE["42"]},
        {"op_id": "d", "type": "delete_comment", "comment_id": DURABLE["43"], "include_replies": False},
    ]))
    assert result.status == "complete" and result.undo_token, result
    after = threads(rich)
    assert after["0"]["done"] is True and after["0"]["replies"][-1]["text"] == "No change: Kept"
    assert after["42"]["done"] is False and tuple(after["42"]["replies"]) == ()
    assert len(after) == 3
    undone = undo_document(str(rich))
    assert undone["restored"] is True and file_sha256(rich) == before_sha and threads(rich) == before_threads


def _partial_batch(source: Path, comment_op: dict):
    from research_mcp.edit_batch import edit_document
    from research_mcp.edit_models import EditDocumentRequest

    return edit_document(EditDocumentRequest(source=str(source), source_sha256=file_sha256(source),
                                             output=str(source), allow_partial=True, operations=[
        {"op_id": "comment", **comment_op},
        # fails: an edit in the table row that holds thread B ("Cell B", paragraph 12)
        {"op_id": "row", "type": "add_comment", "paragraph_index": 12, "anchor": "not here", "text": "bad"},
        {"op_id": "other", "type": "add_comment", "paragraph_index": 1, "anchor": "opening", "text": "unrelated"},
    ]))


@windows_only
@pytest.mark.parametrize("comment_op", [
    {"type": "resolve_comment"},
    {"type": "delete_comment"},
    {"type": "reply_comment", "disposition": "no_change", "reply_text": "The table is correct."},
], ids=["resolve", "delete", "reply"])
def test_failed_row_edit_rolls_back_that_rows_comment_op(rich, tmp_path, comment_op):
    source = tmp_path / "open.docx"
    apply(rich, source, type="reopen_comment", comment_id=DURABLE["42"])
    before = threads(source)
    result = _partial_batch(source, {**comment_op, "comment_id": DURABLE["42"]})
    statuses = {item.op_id: item.status for item in result.operations}
    assert result.status == "partial" and result.undo_token, result
    assert statuses == {"comment": "rolled_back", "row": "failed", "other": "applied"}
    after = threads(source)
    assert after["42"] == before["42"]  # thread B untouched: not resolved, not deleted, no reply
    added = [value for key, value in after.items() if key not in before]
    assert [item["text"] for item in added] == ["unrelated"]  # the unrelated op still published


@windows_only
def test_comment_op_outside_the_failed_row_still_publishes(rich, tmp_path):
    before = threads(rich)
    result = _partial_batch(rich, {"type": "reply_comment", "comment_id": DURABLE["0"], "disposition": "no_change",
                                   "reply_text": "Fine.", "resolve": False})
    statuses = {item.op_id: item.status for item in result.operations}
    assert statuses == {"comment": "applied", "row": "failed", "other": "applied"}
    assert threads(rich)["0"]["replies"][-1]["text"] == "No change: Fine."
    assert threads(rich)["42"] == before["42"]


# -- run merging -------------------------------------------------------------------------------------------

RUN = re.compile(r'<w:r((?: [^>]*)?)>((?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?)<w:t(?: xml:space="preserve")?>([^<]*)'
                 r'</w:t></w:r>', re.S)


def merge_runs(xml: str) -> str:
    """Merge adjacent text runs with identical attributes and properties (the split an anchor edge makes)."""

    while True:
        out, last, changed = [], 0, False
        matches = list(RUN.finditer(xml))
        i = 0
        while i < len(matches):
            first = matches[i]
            text_ = first.group(3)
            j = i
            while j + 1 < len(matches) and matches[j + 1].start() == matches[j].end() and \
                    matches[j + 1].group(1, 2) == first.group(1, 2):
                j += 1
                text_ += matches[j].group(3)
            # every text run is written in one canonical form, so xml:space differences do not count
            out.append(xml[last:first.start()])
            out.append(f'<w:r{first.group(1)}>{first.group(2)}<w:t xml:space="preserve">{text_}</w:t></w:r>')
            last, changed = matches[j].end(), changed or j > i
            i = j + 1
        out.append(xml[last:])
        xml = "".join(out)
        if not changed:
            return xml


# -- final fix r2: single-quoted comment metadata (recheck sol61 P2) ------------------------------------------

def single_quoted(source: Path, out: Path) -> Path:
    """`source` with every attribute in the four comment parts written with single quotes (valid XML)."""

    data = parts(source)
    for name in COMMENT_PARTS:
        data[name] = re.sub(rb'="([^"]*)"', rb"='\1'", data[name])
    with zipfile.ZipFile(out, "w") as archive:
        for name, value in data.items():
            archive.writestr(name, value)
    return out


def test_prefixed_values_reads_either_quote_style():
    from research_mcp.comment_ops import _prefixed_values

    raw = """<x w15:paraId='1A000001' w16cid:paraId = "1B000001" w15:paraIdParent='1C000001' w14:textId="77777777"/>"""
    assert _prefixed_values(raw, "paraId") == ["1A000001", "1B000001"]
    assert _prefixed_values(raw, "paraIdParent") == ["1C000001"]
    assert _prefixed_values(raw, "textId") == ["77777777"]


def test_single_quoted_metadata_ids_are_reserved(rich, tmp_path):
    from research_mcp.comment_ops import _Ids, _Package

    ids = _Ids(_Package(single_quoted(rich, tmp_path / "sq.docx"), "test"))
    assert {ROOT_A, REPLY_A, ROOT_B, REPLY_B, "1A0000F1"} <= ids.para
    assert set(DURABLE.values()) <= ids.durable


def test_single_quoted_metadata_full_workflow(rich, tmp_path):
    """resolve, reopen, reply and delete work on a package whose comment parts use single quotes, and give
    the same threads as on the double-quoted original."""

    sq = single_quoted(rich, tmp_path / "sq.docx")
    before = sq.read_bytes()
    assert threads(sq) == threads(rich)

    resolved, resolved_dq = tmp_path / "resolved.docx", tmp_path / "resolved_dq.docx"
    apply(sq, resolved, type="resolve_comment", comment_id=DURABLE["0"])
    apply(rich, resolved_dq, type="resolve_comment", comment_id=DURABLE["0"])
    assert threads(resolved)["0"]["done"] is True and threads(resolved)["0"]["replies"][0]["done"] is True
    assert threads(resolved) == threads(resolved_dq)

    reopened = tmp_path / "reopened.docx"
    apply(resolved, reopened, type="reopen_comment", comment_id=DURABLE["0"])
    assert threads(reopened) == threads(rich)

    replied = tmp_path / "replied.docx"
    apply(sq, replied, type="reply_comment", comment_id=DURABLE["0"], disposition="no_change",
          reply_text="Checked")
    replies = threads(replied)["0"]["replies"]
    assert len(replies) == 2 and replies[1]["text"] == "No change: Checked"
    assert replies[1]["comment_id"] not in set(DURABLE.values())
    assert threads(replied)["42"] == threads(rich)["42"]

    deleted = tmp_path / "deleted.docx"
    result = apply(sq, deleted, type="delete_comment", comment_id=DURABLE["0"])
    assert set(threads(deleted)) == {"42"} and threads(deleted)["42"] == threads(rich)["42"]
    assert sorted(result["deleted_legacy_ids"], key=int) == ["0", "41"]
    for name in COMMENT_PARTS:
        for value in (ROOT_A, REPLY_A, DURABLE["0"], DURABLE["41"]):
            assert value not in text(deleted, name), (name, value)
    assert sq.read_bytes() == before
