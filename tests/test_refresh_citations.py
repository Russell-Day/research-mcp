"""Plan 602 S3: refresh_citations (refresh_ops.py) and its bibliography diff.

Unit tests run on a small synthetic docx with tracked changes and a fake refresh that renders `{Citation}`
and `{Bibliography}` the way Zotero does, from a fake library. Expectations are read back with
test_cluster_ops' own ElementTree/regex readers.
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest

from research_mcp import refresh_ops
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.refresh_ops import diff_bibliography, entry_items, first_cited_numbers, split_label

from test_cluster_ops import (
    DEL,
    INS,
    MOVE,
    NUMBERS,
    _entry,
    _esc,
    _field,
    _rewrite,
    doc_xml,
    keys_of,
    make_docx,
    read_field,
    revisions_raw,
    sha,
)

LIBRARY = {"K1": "Alpha A. Title K1. J One. 2020;1:1-2.",
           "K2": "Beta B. Title K2. J Two. 2021;2:3-4.",
           "K3": "Gamma C. Title K3. J Three. 2022;3:5-6."}
BIBL_CODE = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '


def _bibl(entries: list[str], extra: str = "") -> str:
    paragraphs = [f'<w:r><w:t>{i}.</w:t></w:r><w:r><w:tab/><w:t>{_esc(text)}</w:t></w:r>'
                  for i, text in enumerate(entries, 1)]
    head = (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{BIBL_CODE}'
            '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    body = '</w:p><w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>'.join(paragraphs)
    return f'<w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>{head}{body}{extra}{end}</w:p>'


def body(entries: list[str] | None = None, *, wrap_b: bool = False, dont_update: bool = False, bibl_extra: str = "") -> str:
    cluster_a = _field("cA", [_entry("K1"), _entry("K2")], ["(1,", "2)"], instr_split=40)
    cluster_b = _field("cB", [_entry("K3")], ["(3)"])
    if wrap_b:
        cluster_b = f'<w:ins w:id="105" w:author="Sam" w:date="2026-11-03T09:00:00Z">{cluster_b}</w:ins>'
    cluster_c = _field("cC", [_entry("K1")], ["(1)"])
    if dont_update:
        cluster_c = cluster_c.replace("&quot;noteIndex&quot;:0", "&quot;noteIndex&quot;:0,&quot;dontUpdate&quot;:true")
        cluster_c = cluster_c.replace("<w:t>(1)</w:t>", "<w:t>(1, hand)</w:t>")
    entries = entries if entries is not None else [LIBRARY[k] for k in ("K1", "K2", "K3")]
    return (
        f'<w:p><w:r><w:t xml:space="preserve">Intro text </w:t></w:r>{cluster_a}'
        f'<w:r><w:t xml:space="preserve"> more text.</w:t></w:r>{INS}</w:p>'
        f'<w:p>{DEL}<w:r><w:t xml:space="preserve">Second </w:t></w:r>{cluster_b}'
        f'<w:r><w:t xml:space="preserve"> and </w:t></w:r>{cluster_c}</w:p>'
        f'{MOVE}{_bibl(entries, bibl_extra)}<w:sectPr/>'
    )


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    return make_docx(tmp_path / "source.docx", body())


def bibliography_paragraphs(path: Path) -> list[str]:
    """Entry texts of the bibliography (w:tab as a tab), one per paragraph, read with a regex."""
    xml = doc_xml(path)
    out = []
    for chunk in xml[xml.index("ZOTERO_BIBL"):].split("</w:p>"):
        text = "".join(m.group(1) if m.group(1) is not None else "\t"
                       for m in re.finditer(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>|<w:tab/>", chunk))
        if text.strip():
            out.append(text)
        if 'fldCharType="end"' in chunk:
            break
    return out


class FakeRefresh:
    """Render `{Citation}` from keys (numbers) and `{Bibliography}` from LIBRARY in first-cited order."""

    def __init__(self, library=None):
        self.library = dict(library or LIBRARY)
        self.calls: list[tuple[Path, tuple]] = []
        self.seen_xml: str | None = None
        self.after = None

    def __call__(self, path, ids, max_wait_seconds=300.0):
        path = Path(path)
        self.calls.append((path, tuple(ids)))
        xml = self.seen_xml = doc_xml(path)
        rendered, order = {}, []
        for cid in re.findall(r'&quot;citationID&quot;:&quot;(\w+)&quot;|"citationID":"(\w+)"', xml):
            cid = cid[0] or cid[1]
            payload, text, _ = read_field(path, cid)
            order += [k for k in keys_of(payload) if k not in order]
            if text != "{Citation}":
                continue
            out = "(" + ",".join(str(NUMBERS[k]) for k in keys_of(payload)) + ")"
            rendered[cid] = out
            start = xml.index(f'"citationID":"{cid}"')
            end = xml.index("{Citation}</w:t>", start)
            seg = xml[start:end].replace('"formattedCitation":"{Citation}","plainCitation":"{Citation}"',
                                         f'"formattedCitation":"{out}","plainCitation":"{out}"')
            xml = xml[:start] + seg + out + xml[end + len("{Citation}"):]
        runs = '</w:p><w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>'.join(
            f'<w:r w:rsidR="00FF0001"><w:t>{i}.</w:t></w:r><w:r><w:tab/><w:t>{_esc(self.library[k])}</w:t></w:r>'
            for i, k in enumerate(order, 1))
        xml = re.sub(r"<w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t>\{Bibliography\}</w:t></w:r>", lambda _: runs, xml, count=1)
        if self.after:
            xml = self.after(xml)
        _rewrite(path, xml)
        return rendered, [], {"seconds": 0.1}


@pytest.fixture
def fake(monkeypatch):
    refresh = FakeRefresh()
    monkeypatch.setattr(refresh_ops, "refresh_file_and_wait", refresh)
    return refresh


def run(doc: Path, out: Path, **kw):
    return refresh_ops.refresh_to_version(doc, out, source_sha256=sha(doc), **kw)


# -- diff ------------------------------------------------------------------------------------------------


def test_split_label():
    assert split_label("18.\tDoe J. Title.") == ("18.", "Doe J. Title.")
    assert split_label("[3] Smith  J.\n") == ("[3]", "Smith J.")
    assert split_label("Smith J. 2020.") == ("", "Smith J. 2020.")


def test_diff_changed_dropped_added():
    before = ["1.\tAlpha A. Title K1. J One. 2020.", "2.\tBeta B. Title K2 hand edit. J Two. 2021.",
              "3.\tGamma C. Old gone entry. 1999.", "4.\tDelta D. Title K4. J Four. 2023."]
    after = ["1.\tAlpha A. Title K1. J One. 2020.", "2.\tBeta B. Title K2. J Two. 2021.",
             "3.\tDelta D. Title K4. J Four. 2023.", "4.\tZeta Z. Brand new. Nature. 2025."]
    diff = diff_bibliography(before, after, ["K1", "K2", "K4", "K9"])
    assert diff["changed"] == [{"item_key": "K2", "before_index": 1, "after_index": 1,
                                "before": before[1], "after": after[1]}]
    assert diff["dropped"] == [{"before_index": 2, "text": before[2]}]
    assert diff["added"] == [{"after_index": 3, "item_key": "K9", "text": after[3]}]
    assert diff["unchanged_count"] == 2 and diff["renumbered_count"] == 1


def test_diff_reordered_is_unchanged():
    before = ["Beta B. Two. 2021.", "Alpha A. One. 2020.", "Gamma C. Three. 2022."]
    after = ["Alpha A. One. 2020.", "Beta B. Two. 2021.", "Gamma C. Three. 2022."]
    diff = diff_bibliography(before, after)
    assert diff["changed"] == diff["dropped"] == diff["added"] == []
    assert diff["unchanged_count"] == 3


def test_diff_duplicate_looking_entries():
    smith_a = "Smith J. Deep learning for bladder cancer staging. Radiology. 2020a;1:1."
    smith_b = "Smith J. Deep learning for bladder cancer staging. Radiology. 2020b;2:2."
    # Both identical copies pair first-to-first; the edited one pairs with its own, not its twin.
    diff = diff_bibliography([smith_a, smith_a, smith_b], [smith_a, smith_a, smith_b])
    assert diff["unchanged_count"] == 3 and not diff["changed"]
    diff = diff_bibliography([smith_a, smith_b + " Hand note."], [smith_a, smith_b])
    assert diff["unchanged_count"] == 1
    assert [(c["before_index"], c["after_index"]) for c in diff["changed"]] == [(1, 1)]


def test_diff_hand_edit_duplicating_another_entry():
    # Reviewer case: entry 1 was hand-edited into a copy of entry 2; the refresh restores entry 1.
    # The overwritten copy (index 0) is the change, attributed to its own item.
    a = "Smith J. Study of alpha and beta. Radiology. 2020;1:1."
    b = "Smith J. Study of alpha and gamma. Radiology. 2021;2:2."
    diff = diff_bibliography(["1. " + b, "2. " + b], ["1. " + a, "2. " + b], ["KEY_A", "KEY_B"])
    assert [(c["item_key"], c["before_index"], c["after_index"]) for c in diff["changed"]] == [("KEY_A", 0, 0)]
    assert diff["unchanged_count"] == 1 and not diff["dropped"] and not diff["added"]


def test_diff_heavy_edit_same_position_is_changed():
    hand = "2.\tDoe J. LoRA ACME paper, retyped by hand. Imaging Methods 2026."
    zotero = "2.\tDoe J, Fairbanks L, Lee D. LoRA fine-tuning of a small LLM. In: Imaging Methods 2026. ACME; 2026."
    before = ["1.\tAlpha A. One.", hand, "3.\tGamma C."]
    after = ["1.\tAlpha A. One.", zotero, "3.\tGamma C."]
    diff = diff_bibliography(before, after)  # similarity 0.45: paired only because it is at the same position
    assert [c["before_index"] for c in diff["changed"]] == [1]
    assert not diff["dropped"] and not diff["added"]
    diff = diff_bibliography(before, after[:1] + after[2:] + [zotero])  # elsewhere: not paired
    assert [d["before_index"] for d in diff["dropped"]] == [1] and [a["after_index"] for a in diff["added"]] == [2]
    unrelated = ["1.\tAlpha A. One.", "2.\tSomething else entirely, nothing alike.", "3.\tGamma C."]
    diff = diff_bibliography(unrelated, after)  # below 0.3: a dropped entry and an added one
    assert not diff["changed"] and len(diff["dropped"]) == len(diff["added"]) == 1


def test_entry_items_only_when_the_order_fixes_them():
    after = ["1.\tBeta B. Title K2.", "2.\tAlpha A. Title K1.", "3.\tGamma C. Title K3."]
    data = {k: {"title": f"Title {k}"} for k in ("K1", "K2", "K3", "K4")}
    cites = ["(1)", "(2,3)"]
    assert entry_items(after, ["K2", "K1", "K3"], {}, cites, data) == ["K2", "K1", "K3"]
    assert entry_items(after, ["K2", "K1", "K3", "K4"], {"omitted": [{"uris": ["http://x/items/K4"]}]}, cites,
                       data) == ["K2", "K1", "K3"]
    assert entry_items(after, ["K2", "K1"], {}, cites, data) == [None, None, None]
    assert entry_items([e[3:] for e in after], ["K2", "K1", "K3"], {}, cites, data) == [None, None, None]
    assert entry_items(after, ["K2", "K1", "K3"], {"uncited": [{"uris": ["u"]}]}, cites, data) == [None, None, None]
    assert entry_items(after, ["K2", "K1", "K3"], {}, cites) == [None, None, None]  # nothing to confirm with
    # No title: the first author's family name confirms the key.
    untitled = {"K1": {"author": [{"family": "Alpha"}]}, "K2": {"author": [{"family": "Beta"}]},
                "K3": {"title": "Title K3"}}
    assert entry_items(after, ["K2", "K1", "K3"], {}, cites, untitled) == ["K2", "K1", "K3"]


def test_entry_items_sorted_multi_item_citation_gets_no_key():
    # Reviewer case: one citation cites Zhao then Adams in a numbered alphabetical style. Zotero
    # numbers Adams 1 and Zhao 2 and sorts the citation to "[1, 2]", which looks like first-cited
    # order. Without the text check the change to Adams's entry was attributed to ZHAO.
    data = {"ZHAO": {"title": "Graph methods for staging", "author": [{"family": "Zhao"}]},
            "ADAMS": {"title": "Deep learning study", "author": [{"family": "Adams"}]}}
    before = ["1. Adams A. Deep learning study. Original.", "2. Zhao Z. Graph methods for staging."]
    after = ["1. Adams A. Deep learning study. Revised.", "2. Zhao Z. Graph methods for staging."]
    keys = entry_items(after, ["ZHAO", "ADAMS"], {}, ["[1, 2]"], data)
    assert keys == [None, None]
    diff = diff_bibliography(before, after, keys)
    assert [(c["item_key"], c["after_index"]) for c in diff["changed"]] == [(None, 0)]
    # The same citation in citation-sequence order (Zhao listed first) keeps its keys.
    seq = ["1. Zhao Z. Graph methods for staging.", "2. Adams A. Deep learning study. Revised."]
    assert entry_items(seq, ["ZHAO", "ADAMS"], {}, ["[1, 2]"], data) == ["ZHAO", "ADAMS"]


def test_entry_items_same_title_items_get_no_key():
    # Reviewer case: Zhao and Adams are both titled "Editorial" and cited "(1,2)" in an alphabetical
    # style; each title is in both entries, so it proves nothing about which entry is whose.
    after = ["1. Adams A. Editorial. Journal. 2026.", "2. Zhao Z. Editorial. Journal. 2026."]
    data = {"ZHAO": {"title": "Editorial"}, "ADAMS": {"title": "Editorial"}}
    assert entry_items(after, ["ZHAO", "ADAMS"], {}, ["(1,2)"], data) == [None, None]
    assert entry_items(after, ["ADAMS", "ZHAO"], {}, ["(1,2)"], data) == [None, None]
    distinct = {"ADAMS": {"title": "Editorial", "author": [{"family": "Adams"}]},
                "ZHAO": {"title": "", "author": [{"family": "Zhao"}]}}
    after2 = ["1. Adams A. Editorial. Journal. 2026.", "2. Zhao Z. Commentary. Journal. 2026."]
    assert entry_items(after2, ["ADAMS", "ZHAO"], {}, ["(1,2)"], distinct) == ["ADAMS", "ZHAO"]
    # A title that also appears inside another entry is fine when no swap is possible (here "Large
    # language models in medicine" is also in a longer entry's text).
    after3 = ["1. Lee L. Large language models in medicine. Nat Med. 2023.",
              "2. Kim K. Evaluating large language models in medicine at scale. 2024."]
    data3 = {"LEE": {"title": "Large language models in medicine"},
             "KIM": {"title": "Evaluating large language models in medicine at scale"}}
    assert entry_items(after3, ["LEE", "KIM"], {}, ["(1,2)"], data3) == ["LEE", "KIM"]
    # Three items in a swap ring (each title also in the next entry) prove nothing either.
    ring = ["1. A. Alpha beta.", "2. B. Beta gamma.", "3. C. Gamma alpha."]
    ring_data = {"A": {"title": "Alpha"}, "B": {"title": "Beta"}, "C": {"title": "Gamma"}}
    assert entry_items(ring, ["A", "B", "C"], {}, ["(1-3)"], ring_data) == [None, None, None]


def test_first_cited_numbers():
    assert first_cited_numbers(["(1)", "(2,3)", "(3,4)", "(8\u201310)", "[5-7]"]) == [1, 2, 3, 4, 8, 9, 10, 5, 6, 7]
    assert first_cited_numbers(["(Smith 2020)"]) == [2020]


def test_entry_items_numbered_alphabetical_gets_no_key():
    # A numbered style sorted alphabetically: Zhao is cited first but listed second, so the text
    # numbers it [2]. Labels 1..n alone would name entry 1 (Adams) ZHAO.
    after = ["1. Adams A. Study. Revised.", "2. Zhao Z. Other study."]
    data = {"ZHAO": {"title": "Other study"}, "ADAMS": {"title": "Study"}}
    assert entry_items(after, ["ZHAO", "ADAMS"], {}, ["[2]", "[1]"], data) == [None, None]
    assert entry_items(after, ["ADAMS", "ZHAO"], {}, ["[1]", "[2]"], data) == ["ADAMS", "ZHAO"]
    # Author-date citations never number the items 1, 2, ...
    assert entry_items(after, ["ADAMS", "ZHAO"], {}, ["(Adams 2020)", "(Zhao 2021)"]) == [None, None]


def test_leading_year_is_entry_text_not_a_label():
    diff = diff_bibliography(["2020 Smith J. Important study."], ["2021 Smith J. Important study."])
    assert len(diff["changed"]) == 1 and diff["unchanged_count"] == 0
    # A whole list numbered 1..n still loses its labels: renumbering is not a change.
    diff = diff_bibliography(["1. Beta B.", "2. Alpha A."], ["1. Alpha A.", "2. Beta B."])
    assert not diff["changed"] and diff["unchanged_count"] == 2 and diff["renumbered_count"] == 2


# -- XML step and refresh ---------------------------------------------------------------------------------


def test_every_cluster_and_the_bibliography_reset_by_xml(doc, tmp_path, monkeypatch):
    seen = {}

    def stop(path, ids, max_wait_seconds=300.0):
        path = Path(path)
        seen["ids"] = tuple(ids)
        seen["fields"] = {cid: read_field(path, cid) for cid in ("cA", "cB", "cC")}
        seen["bibl"] = bibliography_paragraphs(path)
        seen["revisions"] = revisions_raw(path)
        seen["xml"] = doc_xml(path)
        raise RuntimeError("stop after the XML step")

    monkeypatch.setattr(refresh_ops, "refresh_file_and_wait", stop)
    out = tmp_path / "out.docx"
    with pytest.raises(RuntimeError, match="stop after"):
        run(doc, out)
    assert not out.exists()
    assert seen["ids"] == ("cA", "cB", "cC")
    for cid, (payload, text, _) in seen["fields"].items():
        assert text == "{Citation}", cid
        assert payload["properties"]["formattedCitation"] == payload["properties"]["plainCitation"] == "{Citation}"
    assert keys_of(seen["fields"]["cA"][0]) == ["K1", "K2"]
    assert seen["bibl"] == ["{Bibliography}"]
    assert seen["revisions"] == revisions_raw(doc)
    # The bibliography paragraph keeps its style; the code (and its custom/uncited JSON) is untouched.
    assert '<w:pStyle w:val="Bibliography"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/>' in seen["xml"]
    assert _esc(BIBL_CODE) in seen["xml"] or BIBL_CODE in seen["xml"]
    # Everything before the bibliography field other than the three clusters is byte-for-byte the same.
    strip = lambda x: re.sub(r'<w:fldChar w:fldCharType="begin"/>.*?<w:fldChar w:fldCharType="end"/>', "", x, flags=re.S)
    assert strip(seen["xml"]) == strip(doc_xml(doc))


def test_noop_refresh_reports_nothing_changed(doc, tmp_path, fake):
    out = tmp_path / "out.docx"
    original = sha(doc)
    result = run(doc, out)
    assert sha(doc) == original and out.exists()
    assert result["bibliography_unchanged_count"] == 3
    assert result["bibliography_changed"] == result["bibliography_dropped"] == result["bibliography_added"] == []
    assert result["citations_changed"] == []
    assert result["citations_refreshed"] == 3 and result["revisions_checked"] == len(revisions_raw(doc)) + 6
    assert revisions_raw(out) == revisions_raw(doc)
    assert bibliography_paragraphs(out) == [f"{i}.\t{LIBRARY[k]}" for i, k in enumerate(("K1", "K2", "K3"), 1)]


def test_hand_edited_entry_is_listed_and_restored(tmp_path, fake):
    edited = [LIBRARY["K1"], "Beta B. Title K2 (Proceedings, hand edit). J Two. 2021;2:3-4.", LIBRARY["K3"]]
    doc = make_docx(tmp_path / "edited.docx", body(edited))
    result = run(doc, tmp_path / "out.docx")
    assert result["bibliography_changed"] == [{
        "item_key": "K2", "before_index": 1, "after_index": 1,
        "before": "2.\t" + edited[1], "after": "2.\t" + LIBRARY["K2"]}]
    assert result["bibliography_unchanged_count"] == 2
    assert bibliography_paragraphs(tmp_path / "out.docx")[1] == "2.\t" + LIBRARY["K2"]


def test_changed_citation_text_reported(tmp_path, fake):
    doc = make_docx(tmp_path / "stale.docx", body().replace("<w:t>(3)</w:t>", "<w:t>(7)</w:t>"))
    result = run(doc, tmp_path / "out.docx")
    assert result["citations_changed"] == [{"citation_id": "cB", "item_keys": ["K3"], "before": "(7)", "after": "(3)"}]


def test_zotero_change_drops_and_adds(tmp_path, fake):
    doc = make_docx(tmp_path / "old.docx", body([LIBRARY["K1"], LIBRARY["K2"], "Old Q. Removed long ago. 1990."]))
    result = run(doc, tmp_path / "out.docx")
    assert [d["text"] for d in result["bibliography_dropped"]] == ["3.\tOld Q. Removed long ago. 1990."]
    assert result["bibliography_added"] == [{"after_index": 2, "item_key": "K3", "text": "3.\t" + LIBRARY["K3"]}]


def test_preview_writes_nothing(doc, tmp_path, fake):
    out = tmp_path / "out.docx"
    original = sha(doc)
    result = run(doc, out, preview=True)
    assert result["preview"] is True and result["output"] is None
    assert sha(doc) == original and not out.exists()
    assert fake.calls and fake.calls[0][0].parent != doc.parent and not fake.calls[0][0].exists()
    # In place too: a preview with output == source leaves the source as it was.
    edited = make_docx(tmp_path / "edited.docx", body([LIBRARY["K1"], "Beta B. hand edit.", LIBRARY["K3"]]))
    before = sha(edited)
    result = refresh_ops.refresh_to_version(edited, edited, source_sha256=before, preview=True)
    assert sha(edited) == before
    assert [c["before_index"] for c in result["bibliography_changed"]] == [1]


def test_cluster_inside_revision_refused(tmp_path, fake):
    doc = make_docx(tmp_path / "wrapped.docx", body(wrap_b=True))
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=r"refresh_citations: citation cB sits inside tracked revision w:ins id=105"):
        run(doc, out)
    assert not fake.calls and not out.exists()


def test_bookmark_lost_inside_a_revision_refused(tmp_path, fake):
    marker = '<w:bookmarkStart w:id="999" w:name="Example"/><w:bookmarkEnd w:id="999"/>'
    doc = make_docx(tmp_path / "bm-in-ins.docx", body().replace(INS, INS.replace("<w:r ", marker + "<w:r ", 1)))
    assert marker in doc_xml(doc)
    fake.after = lambda xml: xml.replace(marker, "")
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="tracked revision"):
        run(doc, out)
    assert not out.exists()


@pytest.mark.parametrize("extra, message", [
    ('<w:bookmarkStart w:id="9" w:name="_Ref1"/><w:bookmarkEnd w:id="9"/>', "w:bookmarkStart"),
    ('<w:ins w:id="300" w:author="R" w:date="2026-11-01T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins>',
     "tracked revision w:ins id=300"),
])
def test_bibliography_holding_a_marker_refused(tmp_path, fake, extra, message):
    doc = make_docx(tmp_path / "bm.docx", body(bibl_extra=extra))
    with pytest.raises(RefusedUnsafe, match=message):
        run(doc, tmp_path / "out.docx")
    assert not fake.calls


def test_dont_update_cluster_kept(tmp_path, fake):
    doc = make_docx(tmp_path / "hand.docx", body(dont_update=True))
    result = run(doc, tmp_path / "out.docx")
    assert fake.calls[0][1] == ("cA", "cB")
    assert read_field(tmp_path / "out.docx", "cC")[1] == "(1, hand)"
    assert any("cC" in w and "dontUpdate" in w for w in result["warnings"])


def test_refresh_that_changes_a_revision_is_refused(doc, tmp_path, fake):
    fake.after = lambda xml: xml.replace("newly added", "newly addex")
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="tracked revision"):
        run(doc, out)
    assert not out.exists()


def test_sha_mismatch_refused(doc, tmp_path, fake):
    with pytest.raises(RefusedUnsafe, match="SHA-256 mismatch"):
        refresh_ops.refresh_to_version(doc, tmp_path / "o.docx", source_sha256="0" * 64)
    assert not fake.calls


@pytest.mark.parametrize("in_place", [False, True])
@pytest.mark.parametrize("preview", [False, True])
def test_source_saved_between_check_and_staging_refuses(doc, tmp_path, monkeypatch, fake, in_place, preview):
    real = refresh_ops.audit_docx
    calls = []

    def audit(path):
        calls.append(path)
        result = real(path)
        if len(calls) == 1:  # the caller's sha check passes, then the source is saved elsewhere
            xml = doc_xml(doc).replace("Intro text ", "Intro text saved ")
            _rewrite(doc, xml)
        return result

    monkeypatch.setattr(refresh_ops, "audit_docx", audit)
    original = sha(doc)
    out = doc if in_place else tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="source changed after its SHA-256 was checked"):
        refresh_ops.refresh_to_version(doc, out, source_sha256=original, preview=preview)
    assert not fake.calls
    assert "Intro text saved " in doc_xml(doc)
    if not in_place:
        assert not out.exists()


def test_in_place_and_undo_through_server(doc, fake):
    from research_mcp import server

    original = sha(doc)
    preview = server.refresh_citations(source=str(doc), source_sha256=original, output=str(doc), preview=True)
    data = preview.model_dump() if hasattr(preview, "model_dump") else preview
    assert data["undo_token"] is None and sha(doc) == original
    result = server.refresh_citations(source=str(doc), source_sha256=original, output=str(doc))
    data = result.model_dump() if hasattr(result, "model_dump") else result
    assert data["undo_token"] and data["sha256_before"] == original and sha(doc) != original
    server.undo_document(path=str(doc))
    assert sha(doc) == original
