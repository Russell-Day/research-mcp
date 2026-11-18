"""Plan 864 decision 16: insert_zotero_citation on documents with tracked revisions, tracked or plain.

Synthetic docx. Word and Zotero are faked: the fake Word splices an unrendered citation field into the
staged copy before the anchor (as `insert_unrendered` does through COM) and the fake refresh renders it. The
revision guard, the anchor guard and the tracked `w:ins` wrap are the real code. Expectations are read back
with regexes and through the MCP's own accept/reject (`revisions.resolve_revisions`).
A gated live test (RESEARCH_MCP_WORD_LIVE=1 and RESEARCH_MCP_ZOTERO_LIVE=1) is at the end.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path

import pytest

from research_mcp import citation_ops
from research_mcp.citation_ops import (
    TRACKED_BIBLIOGRAPHY_WARNING, insert_citation_to_version, refuse_anchor_in_revision, wrap_citation_insertion,
)
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.revisions import resolve_revisions
from research_mcp.word import file_sha256

from test_citation_delete import DEL, INS, doc_xml, field, instr_payloads, make_docx, para_texts, rev_elements, t

FMT = '<w:rPr><w:b/><w:rPrChange w:id="103" w:author="Sam" w:date="2026-11-01T10:00:00Z"><w:rPr/></w:rPrChange></w:rPr>'
OPENING = t("Prior work ") + field("cA", ["K1"], "(1)") + t(" agrees.")
TARGET = t("Results agree ") + t("closely.") + INS          # anchor "closely." is plain text, then Sam's insertion
REVISED = t("Methods were ") + DEL + t("x") + t("bold", FMT) + t(" kept.")
NEW_ID = "cNEW"


@pytest.fixture
def revised(tmp_path: Path) -> Path:
    return make_docx(tmp_path / "revised.docx", OPENING, TARGET, REVISED)


@pytest.fixture
def plain(tmp_path: Path) -> Path:
    """Revision-free, with package rels (the tracking switch is a custom property part)."""
    path = make_docx(tmp_path / "plain.docx", OPENING, t("Results agree ") + t("closely."))
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("_rels/.rels",
                         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                         'relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    return path


def has_ins(path: Path) -> bool:
    return re.search(r"<w:ins[ >]", doc_xml(path)) is not None


# -- fake Word and Zotero --------------------------------------------------------------------------------------

def _rewrite(path: Path, change) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    entries["word/document.xml"] = change(entries["word/document.xml"].decode("utf-8")).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)


class FakeReference:
    citation_id = NEW_ID


@pytest.fixture
def fake_word(monkeypatch):
    """Patch Word/Zotero in citation_ops. `state["on_save"]` may alter the saved XML (a misbehaving Word)."""
    state: dict = {"on_save": None, "track_during_insert": None, "track_at_save": None, "opened": []}

    class FakeDocument:
        def __init__(self, path: Path):
            self.path = path
            self.TrackRevisions = state.get("word_tracking", False)

    class FakeIsolated:
        def __init__(self, path, *args, **kwargs):
            self.document = FakeDocument(Path(path))
            state["opened"].append(Path(path))

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def save(self):
            state["track_at_save"] = self.document.TrackRevisions
            if state["on_save"]:
                _rewrite(self.document.path, state["on_save"])
            if state.get("on_save_path"):
                state["on_save_path"](self.document.path)

    def fake_insert(document, *, donor_path, paragraph_index, anchor, item_keys, zotero):
        state["track_during_insert"] = document.TrackRevisions
        state["keys"] = item_keys
        anchor_run = f'<w:r><w:t xml:space="preserve">{anchor}</w:t></w:r>'

        def splice(xml: str) -> str:
            assert xml.count(anchor_run) == 1
            return xml.replace(anchor_run, field(NEW_ID, list(item_keys), "{Citation}") + anchor_run)

        _rewrite(document.path, splice)
        return FakeReference()

    def fake_refresh(path, citation_ids):
        _rewrite(Path(path), lambda xml: xml.replace("<w:t>{Citation}</w:t>", "<w:t>(2)</w:t>"))
        return {NEW_ID: "(2)"}, [], {"seconds": 1.5, "zotero_refresh_seconds": 1.0}  # as the real worker

    monkeypatch.setattr(citation_ops, "IsolatedWordDocument", FakeIsolated)
    monkeypatch.setattr(citation_ops, "renew_zotero_session", lambda document: None)
    monkeypatch.setattr(citation_ops, "insert_unrendered", fake_insert)
    monkeypatch.setattr(citation_ops, "ZoteroClient", lambda: None)
    monkeypatch.setattr(citation_ops, "refresh_file_and_wait", fake_refresh)
    return state


def _insert(source: Path, out: Path, **kwargs):
    return insert_citation_to_version(source, out, paragraph_index=2, anchor="closely.", item_keys=["K2"],
                                      item_donor=source, **kwargs)


def _new_ins(xml: str, author: str) -> re.Match:
    found = re.findall(rf'<w:ins w:id="(\d+)" w:author="{re.escape(author)}" w:date="[^"]+">(.*?)</w:ins>', xml, re.S)
    assert len(found) == 1, found
    return found[0]


# -- tracked insert on a revised document ----------------------------------------------------------------------

def test_tracked_insert_on_revised_doc_is_one_ins_around_the_whole_field(revised, tmp_path, fake_word):
    """chk_70a6dab208b44277: one w:ins by the author holds the whole field; other revisions byte-identical."""
    out = tmp_path / "out.docx"
    before = rev_elements(revised)
    result = _insert(revised, out, tracked=True, author="Claude")

    xml = doc_xml(out)
    new_id, inner = _new_ins(xml, "Claude")
    assert result["tracked"] is True and result["revision_id"] == new_id and result["author"] == "Claude"
    assert inner.count('w:fldCharType="begin"') == 1 and inner.count('w:fldCharType="end"') == 1
    assert inner.startswith('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
    assert inner.endswith('<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    assert NEW_ID in inner and "<w:t>(2)</w:t>" in inner
    assert 'closely.' not in inner  # the anchor stays outside the insertion
    assert int(new_id) > 103  # unique in the package
    assert [r for r in rev_elements(out) if f'w:id="{new_id}"' not in r] == before
    assert result["revisions_checked"] == 3
    assert result["warnings"] == [TRACKED_BIBLIOGRAPHY_WARNING]  # timings from the refresh are not warnings
    assert "if the citation is rejected, run refresh_citations after resolving it" in TRACKED_BIBLIOGRAPHY_WARNING


def test_tracked_insert_reject_gives_the_input_paragraph_and_accept_keeps_the_citation(revised, tmp_path, fake_word):
    out = tmp_path / "out.docx"
    result = _insert(revised, out, tracked=True, author="Claude")
    rejected, accepted = tmp_path / "rejected.docx", tmp_path / "accepted.docx"
    resolve_revisions(out, rejected, mode="reject", ids=[result["revision_id"]])
    resolve_revisions(out, accepted, mode="accept", ids=[result["revision_id"]])

    assert para_texts(rejected) == para_texts(revised)
    assert doc_xml(rejected) == doc_xml(revised)  # byte-for-byte the input document.xml
    assert para_texts(accepted)[1] == "Results agree (2)closely. added"
    assert [p for p in instr_payloads(accepted) if NEW_ID in p]  # the live field, with its JSON
    assert len(instr_payloads(accepted)) == 2 and len(instr_payloads(rejected)) == 1


def test_untracked_insert_on_revised_doc_is_plain_and_keeps_every_revision(revised, tmp_path, fake_word):
    out = tmp_path / "out.docx"
    before = rev_elements(revised)
    result = _insert(revised, out, tracked=False)

    assert result["tracked"] is False and result["revision_id"] is None and result["author"] is None
    assert rev_elements(out) == before
    assert result["warnings"] == []
    xml = doc_xml(out)
    assert '<w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {&quot;citationID&quot;:&quot;cNEW' in xml
    assert para_texts(out)[1] == "Results agree (2)closely. added"


def test_words_own_track_changes_is_off_for_the_insert_and_back_on_for_the_save(revised, tmp_path, fake_word):
    fake_word["word_tracking"] = True
    _insert(revised, tmp_path / "out.docx", tracked=True)
    assert fake_word["track_during_insert"] is False and fake_word["track_at_save"] is True


# -- the revision guard -----------------------------------------------------------------------------------------

def test_a_word_save_that_alters_a_revision_refuses_and_writes_nothing(revised, tmp_path, fake_word):
    """Word renaming the author of Sam's insertion on save: refused, no output, source untouched."""
    sha = file_sha256(revised)
    fake_word["on_save"] = lambda xml: xml.replace('w:id="101" w:author="Sam"', 'w:id="101" w:author="Word User"')
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match=r"insert_zotero_citation \(Word's save\).*id=101 by 'Sam'"):
        _insert(revised, out, tracked=True)
    assert not out.exists() and file_sha256(revised) == sha


def test_a_word_save_that_adds_a_revision_refuses(revised, tmp_path, fake_word):
    fake_word["on_save"] = lambda xml: xml.replace(
        t(" kept."), f'<w:ins w:id="900" w:author="Word User" w:date="2026-11-15T00:00:00Z">{t(" kept.")}</w:ins>')
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="tracked revision count changed from 3 to 4"):
        _insert(revised, out, tracked=False)
    assert not out.exists()


SECONDS = "2026-11-15T21:49:42Z"
MCP_MADE = (t("Results agree ") + t("closely.")
            + f'<w:del w:id="10" w:author="Reviewer A" w:date="{SECONDS}"><w:r><w:delText>old</w:delText></w:r></w:del>'
            + f'<w:ins w:id="11" w:author="Reviewer A" w:date="{SECONDS}"><w:r><w:t>new</w:t></w:r></w:ins>')


def test_words_save_renumbering_ids_and_dropping_seconds_is_accepted(tmp_path, fake_word):
    """Live finding (manuscript copy): Word's save renumbered the MCP's revision ids (10, 11 -> 6, 7), cut
    w:date to the minute (21:49:42 -> 21:49:00) while keeping w16du:dateUtc, and dropped xml:space="preserve"
    from text with no outer spaces. That is Word's serialization, not a change: accepted."""
    source = make_docx(tmp_path / "in.docx", OPENING, MCP_MADE.replace("<w:delText>old", '<w:delText xml:space="preserve">old'))
    fake_word["on_save"] = lambda xml: (xml.replace('w:id="10"', 'w:id="6"').replace('w:id="11"', 'w:id="7"')
                                        .replace(SECONDS, "2026-11-15T21:49:00Z")
                                        .replace('<w:delText xml:space="preserve">old', "<w:delText>old"))
    result = _insert(source, tmp_path / "out.docx", tracked=True)
    assert result["revisions_checked"] == 2
    new_id, _ = _new_ins(doc_xml(tmp_path / "out.docx"), "Claude")
    assert new_id not in ("6", "7")


SPLIT_DELETION = (t("Results agree ") + t("closely.")
                  + f'<w:del w:id="10" w:author="Reviewer A" w:date="{SECONDS}"><w:r><w:delText>shar</w:delText></w:r></w:del>'
                  + f'<w:del w:id="11" w:author="Reviewer A" w:date="{SECONDS}"><w:r><w:delText>ed</w:delText></w:r></w:del>')


def test_words_save_merging_split_deletions_is_accepted(tmp_path, fake_word):
    """Plan 864 S4 fix (live check 6): the MCP deletes a word stored in two runs as two touching w:del of one
    author and date; Word's save merges them into one. Accepted, and reported as merged."""
    source = make_docx(tmp_path / "in.docx", OPENING, SPLIT_DELETION)
    merged = ('<w:del w:id="3" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r><w:delText>shared</w:delText>'
              "</w:r></w:del>")
    fake_word["on_save"] = lambda xml: re.sub(r"<w:del w:id=\"10\".*?<w:del w:id=\"11\".*?</w:del>", merged, xml)
    result = _insert(source, tmp_path / "out.docx", tracked=True)
    assert result["revisions_checked"] == 2
    assert result["revision_check"]["word_merged"] == 1 and result["revision_check"]["word_merged_revisions"] == 2

@pytest.mark.parametrize("save, pattern", [
    (lambda xml: xml.replace(SECONDS, "2026-11-15T21:48:00Z", 1), "1 tracked revision"),   # another minute
    (lambda xml: xml.replace(SECONDS, "2026-11-15T21:49:41Z", 1), "1 tracked revision"),   # another second
    (lambda xml: xml.replace('w:id="11"', 'w:id="10"'), "no longer map one to one"),       # two ids merged
    (lambda xml: xml.replace("<w:t>new</w:t>", '<w:t xml:space="preserve">new </w:t>'), "1 tracked revision"),
    (lambda xml: xml.replace('w:author="Reviewer A" w:date', 'w:author="Word" w:date', 1), "1 tracked revision"),
])
def test_other_changes_in_words_save_still_refuse(tmp_path, fake_word, save, pattern):
    source = make_docx(tmp_path / "in.docx", OPENING, MCP_MADE)
    fake_word["on_save"] = save
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match=pattern):
        _insert(source, out, tracked=True)
    assert not out.exists()


NESTED_INS = (t("Results agree ") + t("closely.")
              + f'<w:ins w:id="11" w:author="Reviewer A" w:date="{SECONDS}"><w:bookmarkStart w:id="0" w:name="Kept"/>'
              '<w:r><w:t>new</w:t></w:r><w:bookmarkEnd w:id="0"/></w:ins>')
MOVED = (f'<w:moveFromRangeStart w:id="20" w:author="Reviewer A" w:date="{SECONDS}" w:name="move1"/>'
         f'<w:moveFrom w:id="21" w:author="Reviewer A" w:date="{SECONDS}"><w:r><w:t>gone</w:t></w:r></w:moveFrom>'
         '<w:moveFromRangeEnd w:id="20"/>' + t(" stays."))


@pytest.mark.parametrize("save, pattern", [
    # Word never drops a revision's id: an id that disappears is a change, not a renumbering.
    (lambda xml: xml.replace(' w:id="11"', ""), "1 tracked revision"),
    # A bookmark inside a revision renumbered at one end only: its start and end no longer pair.
    (lambda xml: xml.replace('<w:bookmarkStart w:id="0"', '<w:bookmarkStart w:id="5"'), "1 tracked revision"),
    # A move range renumbered at its start but kept at its end (one shared id, two records).
    (lambda xml: xml.replace('<w:moveFromRangeStart w:id="20"', '<w:moveFromRangeStart w:id="30"'),
     "keeps an id that Word's save renumbered"),
    # Two bookmarks in revisions given one id.
    (lambda xml: xml.replace('w:id="0"', 'w:id="9"').replace('w:id="1"', 'w:id="9"'),
     "no longer map one to one|partner marker"),
])
def test_words_save_tolerance_needs_every_id_present_and_mapped_one_to_one(tmp_path, fake_word, save, pattern):
    """terra61 S5 review P1: the Word-save tolerance dropped every w:id before comparing and mapped only revision
    ids, so a missing id or a renumbered nested bookmark passed. Each id must still be there, keep one new id
    everywhere (nested markers included) and not share it."""
    second = (NESTED_INS.split("closely.</w:t></w:r>", 1)[1]
              .replace('w:id="11"', 'w:id="12"').replace('w:id="0"', 'w:id="1"').replace("Kept", "Also"))
    source = make_docx(tmp_path / "in.docx", OPENING, NESTED_INS + second, MOVED)
    fake_word["on_save"] = save
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match=pattern):
        _insert(source, out, tracked=False)
    assert not out.exists()


def test_words_save_renumbering_every_id_consistently_is_accepted(tmp_path, fake_word):
    source = make_docx(tmp_path / "in.docx", OPENING, NESTED_INS, MOVED)
    fake_word["on_save"] = lambda xml: (xml.replace('w:id="11"', 'w:id="4"').replace('w:id="0"', 'w:id="7"')
                                        .replace('w:id="20"', 'w:id="9"').replace('w:id="21"', 'w:id="10"')
                                        .replace(SECONDS, "2026-11-15T21:49:00Z"))
    result = _insert(source, tmp_path / "out.docx", tracked=False)
    assert result["revisions_checked"] == 4


def test_a_citation_with_a_bibliography_returns_the_rendered_bibliography(revised, tmp_path, fake_word, monkeypatch):
    """terra61 S5 review P1: with both bibliography arguments the result called .get() on the refresh's list of
    bibliography results and crashed after the refresh. The result now carries that list."""
    monkeypatch.setattr(citation_ops, "insert_bibliography_unrendered",
                        lambda document, *, donor_path, paragraph_index: fake_word.setdefault("bib", paragraph_index))
    real_refresh = citation_ops.refresh_file_and_wait

    def refresh_with_bibliography(path, citation_ids):
        rendered, _, extra = real_refresh(path, citation_ids)
        return rendered, ["1. Title K1\n2. Title K2"], extra

    monkeypatch.setattr(citation_ops, "refresh_file_and_wait", refresh_with_bibliography)
    monkeypatch.setattr(citation_ops, "audit_docx", _counting_audit(citation_ops.audit_docx, bibliographies=1))
    result = _insert(revised, tmp_path / "out.docx", bibliography_donor=revised, bibliography_paragraph_index=3)
    assert fake_word["bib"] == 3
    assert result["bibliography_rendered"] == ["1. Title K1\n2. Title K2"]


def _counting_audit(audit, *, bibliographies: int):
    """The real audit, with the output's bibliography count raised as Word's inserted field would raise it."""
    calls = {"n": 0}

    def wrapped(path):
        found = audit(path)
        calls["n"] += 1
        if calls["n"] == 1:
            return found
        from dataclasses import replace
        return replace(found, zotero_bibliography_fields=found.zotero_bibliography_fields + bibliographies)

    return wrapped


SPLIT_BOOKMARK = (t("Results agree ") + t("closely.")
                  + f'<w:ins w:id="11" w:author="Reviewer A" w:date="{SECONDS}"><w:bookmarkStart w:id="3" w:name="Out"/>'
                  '<w:r><w:t>new</w:t></w:r></w:ins>' + t(" tail") + '<w:bookmarkEnd w:id="3"/>')
FOOTNOTE = (t("Results agree ") + t("closely.")
            + f'<w:ins w:id="11" w:author="Reviewer A" w:date="{SECONDS}"><w:r><w:t>new</w:t></w:r>'
            '<w:r><w:footnoteReference w:id="1"/></w:r></w:ins>')


@pytest.mark.parametrize("body, save, pattern", [
    # r2 probe: a bookmark starting inside an insertion and ending outside it; Word's save renumbers the start
    # only, so the end no longer pairs with it.
    (SPLIT_BOOKMARK, lambda xml: xml.replace('<w:bookmarkStart w:id="3"', '<w:bookmarkStart w:id="8"'),
     "lost or gained a partner marker"),
    # r2 probe: a footnote reference inside an insertion pointing at another footnote. Word does not renumber
    # those; only revision, bookmark and comment ids may change.
    (FOOTNOTE, lambda xml: xml.replace('<w:footnoteReference w:id="1"', '<w:footnoteReference w:id="2"'),
     "1 tracked revision"),
])
def test_words_save_tolerance_checks_partner_markers_and_other_references(tmp_path, fake_word, body, save, pattern):
    source = make_docx(tmp_path / "in.docx", OPENING, body)
    fake_word["on_save"] = save
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match=pattern):
        _insert(source, out, tracked=False)
    assert not out.exists()


def test_words_save_renumbering_a_bookmark_at_both_ends_across_a_revision_edge_is_accepted(tmp_path, fake_word):
    source = make_docx(tmp_path / "in.docx", OPENING, SPLIT_BOOKMARK)
    fake_word["on_save"] = lambda xml: xml.replace('w:id="3"', 'w:id="8"').replace('w:id="11"', 'w:id="4"')
    assert _insert(source, tmp_path / "out.docx", tracked=False)["revisions_checked"] == 1


COMMENTED = (t("Results agree ") + t("closely.")
             + f'<w:ins w:id="11" w:author="Reviewer A" w:date="{SECONDS}"><w:commentRangeStart w:id="1"/>'
             '<w:r><w:t>new</w:t></w:r><w:commentRangeEnd w:id="1"/>'
             '<w:r><w:commentReference w:id="1"/></w:r></w:ins>')
COMMENTS_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:comment w:id="1" w:author="Reviewer B" w:initials="RB"><w:p><w:r><w:t>Why?</w:t></w:r></w:p>'
                '</w:comment></w:comments>')


def _with_comments(path: Path) -> Path:
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("word/comments.xml", COMMENTS_XML)
    return path


def _renumber_comments_part(path: Path, new_id: str) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    entries["word/comments.xml"] = entries["word/comments.xml"].replace(b'w:id="1"', f'w:id="{new_id}"'.encode())
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)


def test_words_save_renumbering_comment_markers_without_their_definition_refuses(tmp_path, fake_word):
    """terra61 S5 r3 probe: the markers in document.xml renumbered 1 -> 9 while comments.xml keeps w:comment 1
    detaches the comment, so it refuses."""
    source = _with_comments(make_docx(tmp_path / "in.docx", OPENING, COMMENTED))
    fake_word["on_save"] = lambda xml: xml.replace('w:id="1"/>', 'w:id="9"/>')
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="partner marker"):
        _insert(source, out, tracked=False)
    assert not out.exists()


def test_words_save_renumbering_comment_markers_with_their_definition_is_accepted(tmp_path, fake_word):
    source = _with_comments(make_docx(tmp_path / "in.docx", OPENING, COMMENTED))
    fake_word["on_save"] = lambda xml: xml.replace('w:id="1"/>', 'w:id="9"/>')
    fake_word["on_save_path"] = lambda path: _renumber_comments_part(path, "9")  # Word renumbers both parts
    assert _insert(source, tmp_path / "out.docx", tracked=False)["revisions_checked"] == 1


# -- the anchor guard -------------------------------------------------------------------------------------------

@pytest.mark.parametrize("paragraph, anchor, pattern", [
    (t("Results ") + INS, " added", r"lies inside tracked revision w:ins id=101 by 'Sam'"),
    (t("Results ") + INS, "Results  add", r"crosses tracked revision w:ins id=101 by 'Sam'"),
    (t("Methods ") + DEL + t("x"), "old x", r"lies inside|crosses"),
    (t("Methods ") + DEL + t("x"), "Methods old", r"crosses tracked revision w:del id=102 by 'Reviewer'"),
    (t("A ") + t("bold", FMT), "bold", r"lies inside a tracked formatting change \(w:rPrChange id=103 by 'Sam'\)"),
])
def test_anchor_inside_or_crossing_a_revision_is_refused_and_named(tmp_path, fake_word, paragraph, anchor, pattern):
    source = make_docx(tmp_path / "in.docx", OPENING, paragraph)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=pattern) as caught:
        insert_citation_to_version(source, out, paragraph_index=2, anchor=anchor, item_keys=["K2"], item_donor=source)
    assert "pick anchor text outside tracked changes" in str(caught.value)
    assert not out.exists() and fake_word["opened"] == []  # refused before Word starts


def test_anchor_outside_revisions_passes_the_guard(revised):
    refuse_anchor_in_revision(revised, 2, "closely.")
    refuse_anchor_in_revision(revised, 3, " kept.")
    refuse_anchor_in_revision(revised, 2, "missing")  # no single match: left to Word, which refuses it


def test_tab_stops_in_paragraph_properties_are_not_text(tmp_path):
    ppr = '<w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>'
    source = make_docx(tmp_path / "in.docx", ppr + t("Lead") + INS)
    refuse_anchor_in_revision(source, 1, "Lead")


# -- the wrap on its own ----------------------------------------------------------------------------------------

def test_wrap_refuses_a_citation_that_is_already_inside_a_revision(tmp_path):
    inside = f'<w:ins w:id="7" w:author="Sam" w:date="2026-11-01T10:00:00Z">{field(NEW_ID, ["K2"], "(2)")}</w:ins>'
    source = make_docx(tmp_path / "in.docx", t("Lead ") + inside + t("tail"))
    with pytest.raises(RefusedUnsafe, match="insert_zotero_citation"):
        wrap_citation_insertion(source, citation_id=NEW_ID, author="Claude")


def test_wrap_needs_the_citation_exactly_once(revised):
    with pytest.raises(VerificationFailed, match="not in the document exactly once"):
        wrap_citation_insertion(revised, citation_id="nope", author="Claude")


# -- sibling: revision-free document, switch on, author "Reviewer B" (server routing) ---------------------------

def test_switch_on_routes_a_revision_free_insert_as_tracked_by_the_switch_author(plain, tmp_path, fake_word):
    from research_mcp import server
    from research_mcp.track_setting import write_setting

    switched = tmp_path / "switched.docx"
    write_setting(plain, switched, on=True, author="Reviewer B")
    out = tmp_path / "out.docx"
    result = server.insert_zotero_citation(str(switched), str(out), 2, "closely.", ["K2"], str(switched))

    assert result["tracking"] == "on" and result["tracking_author"] == "Reviewer B"
    new_id, inner = _new_ins(doc_xml(out), "Reviewer B")
    assert result["revision_id"] == new_id and NEW_ID in inner
    assert TRACKED_BIBLIOGRAPHY_WARNING in result["warnings"]
    rejected = tmp_path / "rejected.docx"
    resolve_revisions(out, rejected, mode="reject", ids=[new_id])
    assert para_texts(rejected) == para_texts(plain)


def test_tracked_false_overrides_the_switch(plain, tmp_path, fake_word):
    from research_mcp import server
    from research_mcp.track_setting import write_setting

    switched = tmp_path / "switched.docx"
    write_setting(plain, switched, on=True, author="Reviewer B")
    out = tmp_path / "out.docx"
    result = server.insert_zotero_citation(str(switched), str(out), 2, "closely.", ["K2"], str(switched),
                                           tracked=False)
    assert result["tracking"] == "off" and result["revision_id"] is None
    assert not has_ins(out)


def test_switch_off_on_a_revision_free_doc_is_a_plain_insert(plain, tmp_path, fake_word):
    from research_mcp import server

    out = tmp_path / "out.docx"
    result = server.insert_zotero_citation(str(plain), str(out), 2, "closely.", ["K2"], str(plain))
    assert result["tracking"] == "off" and result["warnings"] == []
    assert not has_ins(out)


# -- live (gated) -----------------------------------------------------------------------------------------------

LIVE = os.environ.get("RESEARCH_MCP_WORD_LIVE") == "1" and os.environ.get("RESEARCH_MCP_ZOTERO_LIVE") == "1"


@pytest.mark.skipif(not LIVE, reason="needs Word and Zotero (RESEARCH_MCP_WORD_LIVE=1, RESEARCH_MCP_ZOTERO_LIVE=1)")
def test_live_tracked_insert_on_a_revised_copy(tmp_path):
    """Live: RESEARCH_MCP_LIVE_CITE_DOC names a .docx with Zotero citations and tracked revisions,
    RESEARCH_MCP_LIVE_CITE_ARGS a JSON {"paragraph_index", "anchor", "item_keys"}; works on a copy."""
    import shutil

    doc = os.environ.get("RESEARCH_MCP_LIVE_CITE_DOC")
    args = os.environ.get("RESEARCH_MCP_LIVE_CITE_ARGS")
    if not doc or not args:
        pytest.skip("set RESEARCH_MCP_LIVE_CITE_DOC and RESEARCH_MCP_LIVE_CITE_ARGS")
    source = tmp_path / "copy.docx"
    shutil.copyfile(doc, source)
    before = rev_elements(source)
    spec = json.loads(args)
    out = tmp_path / "out.docx"
    result = insert_citation_to_version(source, out, item_donor=None, tracked=True, author="Claude", **spec)
    new_id, inner = _new_ins(doc_xml(out), "Claude")
    assert result["revision_id"] == new_id and inner.count('w:fldCharType="begin"') == 1
    assert [r for r in rev_elements(out) if f'w:id="{new_id}"' not in r] == before
