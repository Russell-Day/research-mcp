"""Plan 864 S4 fix: the revision guard of the citation tools tolerates Word's save merging touching revisions.

The MCP's tracked edits delete a word stored in several runs as one w:del per run: adjacent sibling w:del
elements of one author and date. Word's save merges such a group into one revision (live, plan 864 S4 fix step 1),
so `compare_revision_snapshots` accepts that merge and nothing more. Synthetic docx pairs: the input as the MCP
writes it and the output as Word's save writes it.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from research_mcp.citation_ops import compare_revision_snapshots, revision_snapshot
from research_mcp.errors import VerificationFailed

from test_citation_delete import make_docx, t

AUTHOR = 'w:author="Reviewer A"'
SECONDS = "2026-11-15T21:49:42Z"
LATER_SECONDS = "2026-11-15T21:49:58Z"
MINUTE = "2026-11-15T21:49:00Z"
NEXT_MINUTE = "2026-11-15T21:50:10Z"
BOLD = "<w:rPr><w:b/></w:rPr>"


def rev(kind: str, rev_id: int, *texts: str, rpr: str = "", date: str = SECONDS, author: str = AUTHOR) -> str:
    tag = "delText" if kind == "del" else "t"
    runs = "".join(f'<w:r>{rpr}<w:{tag} xml:space="preserve">{text}</w:{tag}></w:r>' for text in texts)
    return f'<w:{kind} w:id="{rev_id}" {author} w:date="{date}">{runs}</w:{kind}>'


def runs(kind: str, rev_id: int, *pieces: tuple[str, str], date: str = MINUTE) -> str:
    """One revision holding a run per (text, rPr) piece, as Word writes a merged group."""
    tag = "delText" if kind == "del" else "t"
    body = "".join(f"<w:r>{rpr}<w:{tag}>{text}</w:{tag}></w:r>" for text, rpr in pieces)
    return f'<w:{kind} w:id="{rev_id}" {AUTHOR} w:date="{date}">{body}</w:{kind}>'


def compare(tmp_path: Path, before: str, after: str, *others: str) -> dict:
    """Compare the revisions of a paragraph `before` (plus `others`) with Word's save `after`."""
    src = make_docx(tmp_path / "in.docx", t("Lead ") + before + t(" tail"), *others, bibl=False)
    out = make_docx(tmp_path / "out.docx", t("Lead ") + after + t(" tail"), *others, bibl=False)
    return compare_revision_snapshots(revision_snapshot(src), revision_snapshot(out), "test")


SPLIT = rev("del", 10, "shar") + rev("del", 11, "ed")


@pytest.mark.parametrize("before, after, merged", [
    # two deletions of one word, one run each, equal formatting: one w:del, one run (live c01)
    (SPLIT, runs("del", 0, ("shared", "")), 2),
    # three (live c02)
    (rev("del", 10, "9") + rev("del", 11, "4") + rev("del", 12, "4"), runs("del", 3, ("944", "")), 3),
    # different formatting: one w:del, the runs kept apart (live c03)
    (rev("del", 10, "shar", rpr=BOLD) + rev("del", 11, "ed"), runs("del", 0, ("shar", BOLD), ("ed", "")), 2),
    # runs split differently, formatting per character equal
    (rev("del", 10, "shar", rpr=BOLD) + rev("del", 11, "ed", rpr=BOLD), runs("del", 0, ("sh", BOLD), ("ared", BOLD)), 2),
    # insertions (live c05)
    (rev("ins", 10, "new") + rev("ins", 11, "er"), runs("ins", 0, ("newer", "")), 2),
    # dates differing in seconds only (live c07)
    (rev("del", 10, "shar") + rev("del", 11, "ed", date=LATER_SECONDS), runs("del", 0, ("shared", "")), 2),
    # a revision of several runs merged with the next (live c16)
    (rev("del", 10, "a", "b") + rev("del", 11, "c"), runs("del", 0, ("abc", "")), 2),
    # a proofing marker between them, dropped by Word (live c12)
    (rev("del", 10, "shar") + '<w:proofErr w:type="spellStart"/>' + rev("del", 11, "ed"),
     runs("del", 0, ("shared", "")), 2),
])
def test_word_merging_a_touching_group_is_accepted(tmp_path, before, after, merged):
    result = compare(tmp_path, before, after)
    assert result["word_merged"] == 1 and result["word_merged_revisions"] == merged
    assert result["revisions_checked"] == merged and result["level"] == "serialization_only"


def test_two_groups_merged_and_a_group_left_alone(tmp_path):
    """ins,ins,del,del -> one ins + one del (live c15); a group Word did not merge still compares one to one,
    renumbered as Word renumbers it."""
    before = rev("ins", 10, "ne") + rev("ins", 11, "w") + rev("del", 12, "ol") + rev("del", 13, "d")
    after = runs("ins", 0, ("new", "")) + runs("del", 1, ("old", ""))
    kept_before = rev("del", 20, "x", date=MINUTE) + rev("del", 21, "y", date=MINUTE)
    kept_after = kept_before.replace('w:id="20"', 'w:id="2"').replace('w:id="21"', 'w:id="3"')
    result = compare(tmp_path, before + t(" mid ") + kept_before, after + t(" mid ") + kept_after)
    assert result["word_merged"] == 2 and result["word_merged_revisions"] == 4
    assert result["revisions_checked"] == 6 and result["word_renumbered_or_dates_to_minute"] == 2


def test_a_merged_group_keeps_a_nested_bookmark_and_its_partners(tmp_path):
    """A bookmark inside a member (not between members) moves into the merged revision, renumbered at both ends."""
    member = ('<w:del w:id="10" w:author="Reviewer A" w:date="2026-11-15T21:49:42Z"><w:bookmarkStart w:id="5" '
              'w:name="Kept"/><w:r><w:delText>shar</w:delText></w:r><w:bookmarkEnd w:id="5"/></w:del>')
    after = ('<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:bookmarkStart w:id="7" '
             'w:name="Kept"/><w:r><w:delText>shar</w:delText></w:r><w:bookmarkEnd w:id="7"/>'
             "<w:r><w:delText>ed</w:delText></w:r></w:del>")
    assert compare(tmp_path, member + rev("del", 11, "ed"), after)["word_merged"] == 1
    # renumbered at one end only: refused
    with pytest.raises(VerificationFailed):
        compare(tmp_path, member + rev("del", 11, "ed"), after.replace('<w:bookmarkEnd w:id="7"', '<w:bookmarkEnd w:id="5"'))


@pytest.mark.parametrize("before, after, pattern", [
    # Word merges across a bookmark (live c10), but that is not tolerated
    (rev("del", 10, "shar") + '<w:bookmarkStart w:id="3" w:name="B"/><w:bookmarkEnd w:id="3"/>' + rev("del", 11, "ed"),
     '<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r><w:delText>shar</w:delText></w:r>'
     '<w:bookmarkStart w:id="1" w:name="B"/><w:bookmarkEnd w:id="1"/><w:r><w:delText>ed</w:delText></w:r></w:del>',
     "count changed from 2 to 1"),
    # across a comment marker (live c11)
    (rev("del", 10, "shar") + '<w:commentRangeStart w:id="4"/>' + rev("del", 11, "ed") + '<w:commentRangeEnd w:id="4"/>',
     '<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r><w:delText>shar</w:delText></w:r>'
     '<w:commentRangeStart w:id="4"/><w:r><w:delText>ed</w:delText></w:r></w:del><w:commentRangeEnd w:id="4"/>',
     "count changed from 2 to 1"),
    # across an empty run (live c17) or text
    (rev("del", 10, "shar") + "<w:r><w:rPr><w:b/></w:rPr></w:r>" + rev("del", 11, "ed"),
     runs("del", 0, ("shared", "")), "count changed from 2 to 1"),
    (rev("del", 10, "shar") + t("x") + rev("del", 11, "ed"), runs("del", 0, ("shared", "")) + t("x"),
     "count changed from 2 to 1"),
    # dates in different minutes (live c08: Word keeps them apart)
    (rev("del", 10, "shar") + rev("del", 11, "ed", date=NEXT_MINUTE), runs("del", 0, ("shared", "")),
     "count changed from 2 to 1"),
    # different authors (live c09)
    (rev("del", 10, "shar") + rev("del", 11, "ed", author='w:author="Reviewer B"'), runs("del", 0, ("shared", "")),
     "count changed from 2 to 1"),
    # different types (live c14)
    (rev("del", 10, "old") + rev("ins", 11, "new"), runs("del", 0, ("oldnew", "")), "count changed from 2 to 1"),
    # the merged text changed
    (SPLIT, runs("del", 0, ("sharred", "")), "count changed from 2 to 1"),
    (SPLIT, runs("del", 0, ("shar", "")), "count changed from 2 to 1"),
    # a character's formatting changed in the merge
    (rev("del", 10, "shar", rpr=BOLD) + rev("del", 11, "ed"), runs("del", 0, ("shared", BOLD)),
     "count changed from 2 to 1"),
    # the merged revision has another author or minute
    (SPLIT, runs("del", 0, ("shared", "")).replace("Reviewer A", "Word User"), "count changed from 2 to 1"),
    (SPLIT, runs("del", 0, ("shared", ""), date="2026-11-15T21:48:00Z"), "count changed from 2 to 1"),
    # anything between the members breaks the group, whitespace included (between them, or around a proofing marker)
    (rev("del", 10, "shar") + "  " + rev("del", 11, "ed"), runs("del", 0, ("shared", "")), "count changed from 2 to 1"),
    (rev("del", 10, "shar") + '<w:proofErr w:type="spellStart"/> ' + rev("del", 11, "ed"),
     runs("del", 0, ("shared", "")), "count changed from 2 to 1"),
    (rev("del", 10, "shar") + ' <w:proofErr w:type="spellStart"/>' + rev("del", 11, "ed"),
     runs("del", 0, ("shared", "")), "count changed from 2 to 1"),
    # the merged text lost xml:space on its leading whitespace (Word drops it only from text with no outer spaces)
    (rev("del", 10, " shar") + rev("del", 11, "ed"),
     '<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r><w:delText> shared</w:delText>'
     "</w:r></w:del>", "count changed from 2 to 1"),
    # a text element attribute other than xml:space changed in the merge
    (SPLIT, '<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r><w:delText w:extra="1">'
     "shared</w:delText></w:r></w:del>", "count changed from 2 to 1"),
    # an empty formatted run inside a member dropped in the merge
    ('<w:del w:id="10" w:author="Reviewer A" w:date="2026-11-15T21:49:42Z"><w:r><w:delText>shar</w:delText></w:r>'
     "<w:r><w:rPr><w:b/></w:rPr></w:r></w:del>" + rev("del", 11, "ed"), runs("del", 0, ("shared", "")),
     "count changed from 2 to 1"),
    # a bold run holding an empty text element dropped in the merge
    ('<w:del w:id="10" w:author="Reviewer A" w:date="2026-11-15T21:49:42Z"><w:r><w:delText>shar</w:delText></w:r>'
     "<w:r><w:rPr><w:b/></w:rPr><w:delText/></w:r></w:del>" + rev("del", 11, "ed"), runs("del", 0, ("shared", "")),
     "count changed from 2 to 1"),
    # a merge of part of the group (Word merges the whole group)
    (rev("del", 10, "9") + rev("del", 11, "4") + rev("del", 12, "4"),
     runs("del", 0, ("94", "")) + runs("del", 1, ("4", "")), "count changed from 3 to 2"),
])
def test_anything_but_words_merge_still_refuses(tmp_path, before, after, pattern):
    with pytest.raises(VerificationFailed, match=pattern):
        compare(tmp_path, before, after)


def test_a_merged_revision_sharing_its_id_with_another_revision_refuses(tmp_path):
    other = rev("ins", 30, "kept", date=MINUTE)
    with pytest.raises(VerificationFailed, match="map one to one|shares its id"):
        compare(tmp_path, SPLIT + t(" and ") + other, runs("del", 30, ("shared", "")) + t(" and ") + other)


def test_an_unmerged_group_compares_one_to_one(tmp_path):
    """The input as Word itself would leave it (no merge): byte-identical, nothing merged."""
    result = compare(tmp_path, SPLIT, SPLIT)
    assert result["word_merged"] == 0 and result["level"] == "byte_identical"
    with pytest.raises(VerificationFailed, match="1 tracked revision"):
        compare(tmp_path, SPLIT, SPLIT.replace(">ed<", ">ad<"))


def test_a_lost_revision_next_to_a_merge_is_still_a_count_change(tmp_path):
    other = rev("ins", 30, "kept")
    with pytest.raises(VerificationFailed, match=r"count changed from 3 to 1 \(1 touching group"):
        compare(tmp_path, SPLIT + t(" and ") + other, runs("del", 0, ("shared", "")) + t(" and kept"))


def test_merged_text_keeping_its_outer_whitespace_is_accepted(tmp_path):
    """xml:space kept where the merged text has outer whitespace, dropped where a member's space became inner."""
    before = rev("del", 10, " shar") + rev("del", 11, "ed ") + t("x") + rev("del", 12, "a") + rev("del", 13, " b")
    after = ('<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r>'
             '<w:delText xml:space="preserve"> shared </w:delText></w:r></w:del>' + t("x")
             + runs("del", 1, ("a b", "")))
    result = compare(tmp_path, before, after)
    assert result["word_merged"] == 2 and result["word_merged_revisions"] == 4


def test_an_id_absorbed_by_two_merged_groups_refuses(tmp_path):
    """Two revisions sharing old id 11, each the later member of a group, merged into two revisions: the old id
    would map to two new ids."""
    before = SPLIT + t(" mid ") + rev("del", 12, "o") + rev("del", 11, "ld")
    after = runs("del", 0, ("shared", "")) + t(" mid ") + runs("del", 1, ("old", ""))
    with pytest.raises(VerificationFailed, match="id=11 .* into two different revisions"):
        compare(tmp_path, before, after)


def test_an_empty_formatted_run_kept_through_the_merge_is_accepted(tmp_path):
    member = ('<w:del w:id="10" w:author="Reviewer A" w:date="2026-11-15T21:49:42Z"><w:r><w:delText>shar</w:delText>'
              "</w:r><w:r><w:rPr><w:b/></w:rPr></w:r></w:del>")
    after = ('<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r><w:delText>shar</w:delText>'
             "</w:r><w:r><w:rPr><w:b/></w:rPr></w:r><w:r><w:delText>ed</w:delText></w:r></w:del>")
    assert compare(tmp_path, member + rev("del", 11, "ed"), after)["word_merged"] == 1


def test_an_empty_text_element_kept_through_the_merge_is_accepted(tmp_path):
    member = ('<w:del w:id="10" w:author="Reviewer A" w:date="2026-11-15T21:49:42Z"><w:r><w:delText>shar</w:delText>'
              "</w:r><w:r><w:rPr><w:b/></w:rPr><w:delText/></w:r></w:del>")
    after = ('<w:del w:id="0" w:author="Reviewer A" w:date="2026-11-15T21:49:00Z"><w:r><w:delText>shar</w:delText>'
             "</w:r><w:r><w:rPr><w:b/></w:rPr><w:delText/></w:r><w:r><w:delText>ed</w:delText></w:r></w:del>")
    assert compare(tmp_path, member + rev("del", 11, "ed"), after)["word_merged"] == 1
