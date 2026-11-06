"""Plan 602 S2: edit_citation_cluster (cluster_ops.py).

Unit tests run on a small synthetic docx with tracked changes, a fake Zotero client and a fake refresh
that renders `{Citation}` the way Zotero does. Every expectation is read back with this file's own
ElementTree/regex reader, not with cluster_ops' parser.
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import cluster_ops
from research_mcp.errors import RefusedUnsafe, VerificationFailed, ZoteroUnavailable

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
URI = "http://zotero.org/users/1/items/"
REV = re.compile(r"<w:(ins|del)\b[^>]*?(?:/>|>.*?</w:\1>)", re.S)
NUMBERS = {"K1": 1, "K2": 2, "K3": 3, "K4": 4}


def _instr(cid: str, items: list[dict], formatted: str) -> str:
    payload = {"citationID": cid, "properties": {"formattedCitation": formatted, "plainCitation": formatted,
                                                 "noteIndex": 0},
               "citationItems": items, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}
    return " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, separators=(",", ":")) + " "


def _entry(key: str, **extra) -> dict:
    return {"id": NUMBERS[key], "uris": [URI + key], "itemData": {"id": NUMBERS[key], "title": f"Title {key}"}, **extra}


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _field(cid: str, items: list[dict], rendered: list[str], instr_split: int = 0, inner: str = "") -> str:
    text = _instr(cid, items, "".join(rendered))
    pieces = [text] if not instr_split else [text[:instr_split], text[instr_split:]]
    instr_runs = "".join(f'<w:r><w:instrText xml:space="preserve">{_esc(p)}</w:instrText></w:r>' for p in pieces)
    result_runs = "".join(f"<w:r><w:t>{_esc(r)}</w:t></w:r>" for r in rendered)
    return (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>{instr_runs}'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>{inner}{result_runs}'
            f'<w:r><w:fldChar w:fldCharType="end"/></w:r>')


INS = '<w:ins w:id="101" w:author="Sam" w:date="2026-11-01T10:00:00Z"><w:r w:rsidR="00AA0001"><w:t xml:space="preserve"> newly added</w:t></w:r></w:ins>'
DEL = '<w:del w:id="102" w:author="Reviewer" w:date="2026-11-02T11:00:00Z"><w:r><w:rPr><w:b/></w:rPr><w:delText>old words </w:delText></w:r></w:del>'


def _body() -> str:
    cluster_a = _field("cA", [_entry("K1", locator="12", label="page", prefix="see"), _entry("K2")],
                       ["(1,", "2)"], instr_split=40)
    cluster_b = _field("cB", [_entry("K3")], ["(3)"])
    cluster_c = _field("cC", [_entry("K3")], ["(3)"])
    cluster_d = _field("cD", [_entry("K1")], ["(1)"])
    page = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    cluster_e = _field("cE", [_entry("K2")], ["(2)"], inner=page)
    bibl = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1. Title K1</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>2. Title K2</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>3. Title K3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    return (
        f'<w:p><w:r><w:t xml:space="preserve">Intro text </w:t></w:r>{cluster_a}'
        f'<w:r><w:t xml:space="preserve"> more text.</w:t></w:r>{INS}</w:p>'
        f'<w:p>{DEL}<w:r><w:t xml:space="preserve">Second </w:t></w:r>{cluster_b}'
        f'<w:r><w:t xml:space="preserve"> and </w:t></w:r>{cluster_c}</w:p>'
        f'<w:p><w:r><w:t xml:space="preserve">Tracked </w:t></w:r>'
        f'<w:ins w:id="103" w:author="Sam" w:date="2026-11-03T09:00:00Z">{cluster_d}</w:ins></w:p>'
        f'<w:p><w:r><w:t xml:space="preserve">Nested </w:t></w:r>{cluster_e}</w:p>'
        f'{bibl}<w:sectPr/>'
    )


def make_docx(path: Path, body: str | None = None) -> Path:
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:body>{body if body is not None else _body()}</w:body></w:document>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml",
                         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                         '<Default Extension="xml" ContentType="application/xml"/></Types>')
        archive.writestr("word/document.xml", document)
        archive.writestr("word/comments.xml",
                         '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         '<w:comment w:id="0"><w:p><w:ins w:id="104" w:author="X" w:date="2026-11-01T00:00:00Z">'
                         '<w:r><w:t>c</w:t></w:r></w:ins></w:p></w:comment></w:comments>')
    return path


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def doc_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def revisions_raw(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        return [m.group(0) for part in sorted(archive.namelist()) if part.startswith("word/")
                for m in REV.finditer(archive.read(part).decode("utf-8"))]


def read_field(path: Path, cid: str) -> tuple[dict, str, list[str]]:
    """(payload, rendered text, instr pieces) of citation `cid`, read with ElementTree."""
    root = ET.fromstring(doc_xml(path))
    depth, state, instr, result = 0, None, [], []
    for el in root.iter():
        if el.tag == W + "fldChar":
            kind = el.get(W + "fldCharType")
            if kind == "begin":
                depth += 1
                if state is None or state == "done":
                    state, instr, result, mine_depth = "instr", [], [], depth
            elif kind == "separate" and depth == mine_depth and state == "instr":
                state = "result"
            elif kind == "end":
                if state == "result" and depth == mine_depth:
                    text = "".join(instr)
                    payload = json.loads(text[text.index("{"):text.rindex("}") + 1])
                    if payload.get("citationID") == cid:
                        return payload, "".join(result), instr
                    state = "done"
                depth -= 1
        elif el.tag == W + "instrText" and state == "instr":
            instr.append(el.text or "")
        elif el.tag == W + "t" and state == "result" and depth == mine_depth:
            result.append(el.text or "")
    raise AssertionError(f"citation {cid} not found")


def keys_of(payload: dict) -> list[str]:
    return [item["uris"][0].rsplit("/", 1)[-1] for item in payload["citationItems"]]


class FakeZotero:
    def __init__(self, trashed=()):
        self.trashed = set(trashed)
        self.csl_calls: list[str] = []

    def item(self, key):
        if key not in NUMBERS:
            raise ZoteroUnavailable(f"404 for {key}")
        return {"key": key, "deleted": key in self.trashed}

    def csl(self, key):
        self.csl_calls.append(key)
        if key not in NUMBERS:
            raise ZoteroUnavailable(f"404 for {key}")
        if key in self.trashed:
            raise RefusedUnsafe(f"Zotero item {key} is in the trash")
        return {"id": URI + key, "type": "article-journal", "title": f"Title {key} & more"}


@pytest.fixture
def fake_refresh(monkeypatch):
    """Render every `{Citation}` from its keys like Zotero (numbers sorted), churning rsids like Word."""
    calls = []

    def refresh(path, ids, max_wait_seconds=180.0):
        calls.append((Path(path), tuple(ids)))
        xml = doc_xml(Path(path))
        rendered = {}
        for cid in ids:
            payload, text, _ = read_field(Path(path), cid)
            assert text == "{Citation}"
            out = "(" + ",".join(str(n) for n in sorted(NUMBERS[k] for k in keys_of(payload))) + ")"
            rendered[cid] = out
            start = xml.index(f'"citationID":"{cid}"')
            end = xml.index("{Citation}</w:t>", start)
            xml = xml[:end] + out + xml[end + len("{Citation}"):]
            head = xml[:start].rsplit("ADDIN", 1)
            seg = xml[start:end].replace('"formattedCitation":"{Citation}","plainCitation":"{Citation}"',
                                         f'"formattedCitation":"{out}","plainCitation":"{out}"')
            xml = head[0] + "ADDIN" + head[1] + seg + xml[end:]
        _rewrite(Path(path), xml)
        return rendered, [], {"seconds": 0.1}

    monkeypatch.setattr(cluster_ops, "refresh_file_and_wait", refresh)
    return calls


def _rewrite(path: Path, xml: str) -> None:
    temp = path.with_suffix(".rw")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            dst.writestr(info, xml if info.filename == "word/document.xml" else src.read(info.filename))
    shutil.move(temp, path)


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    return make_docx(tmp_path / "source.docx")


def edit(doc: Path, out: Path, keys, para=1, zotero=None, **kw):
    return cluster_ops.edit_cluster_to_version(doc, out, source_sha256=sha(doc), paragraph_index=para,
                                               item_keys=keys, zotero=zotero or FakeZotero(), **kw)


# -- edits ---------------------------------------------------------------------------------------------


def test_add_key_keeps_kept_entries_and_revisions(doc, tmp_path, fake_refresh):
    before_source = sha(doc)
    out = tmp_path / "add.docx"
    result = edit(doc, out, ["K1", "K2", "K4"], citation_id="cA")
    payload, rendered, instr = read_field(out, "cA")
    assert keys_of(payload) == ["K1", "K2", "K4"]
    assert payload["citationItems"][0] == _entry("K1", locator="12", label="page", prefix="see")
    assert payload["citationItems"][2] == {"id": URI + "K4", "uris": [URI + "K4"],
                                           "itemData": {"id": URI + "K4", "type": "article-journal",
                                                        "title": "Title K4 & more"}}
    assert payload["schema"].endswith("csl-citation.json") and payload["citationID"] == "cA"
    assert rendered == "(1,2,4)" and instr[1] == ""  # second instr run emptied, JSON in the first
    assert result["item_keys_before"] == ["K1", "K2"] and result["item_keys_after"] == ["K1", "K2", "K4"]
    assert result["rendered_citation_before"] == "(1,2)" and result["rendered_citation_after"] == "(1,2,4)"
    assert result["bibliography_before"] == ["1. Title K1", "2. Title K2", "3. Title K3"]
    assert result["revisions_checked"] == 4 and result["revision_check"]["level"] == "byte_identical"
    assert revisions_raw(out) == revisions_raw(doc)
    assert sha(doc) == before_source
    assert fake_refresh == [(out.resolve(), ("cA",))]


def test_remove_and_reorder(doc, tmp_path, fake_refresh):
    out = tmp_path / "remove.docx"
    edit(doc, out, ["K2"], citation_id="cA")
    assert keys_of(read_field(out, "cA")[0]) == ["K2"]
    out2 = tmp_path / "reorder.docx"
    edit(doc, out2, ["K2", "K1"], current_text="(1,2)")
    payload, rendered, _ = read_field(out2, "cA")
    assert keys_of(payload) == ["K2", "K1"] and payload["citationItems"][1]["locator"] == "12"
    assert rendered == "(1,2)"


def test_xml_step_touches_only_the_cluster(doc, tmp_path, monkeypatch):
    """Before the refresh: JSON in the first instr run, `{Citation}` in the first result w:t, the rest
    of document.xml and every other part byte-identical."""
    seen = {}

    def refresh(path, ids, max_wait_seconds=180.0):
        seen["xml"] = doc_xml(Path(path))
        seen["revs"] = revisions_raw(Path(path))
        raise VerificationFailed("stop after the XML step")

    monkeypatch.setattr(cluster_ops, "refresh_file_and_wait", refresh)
    out = tmp_path / "x.docx"
    with pytest.raises(VerificationFailed):
        edit(doc, out, ["K1", "K3"], citation_id="cA")
    assert not out.exists()  # a failure discards the versioned output
    assert seen["revs"] == revisions_raw(doc)
    old, new = doc_xml(doc), seen["xml"]
    start = old.rindex("<w:r>", 0, old.index(" ADDIN ZOTERO_ITEM"))  # cA is the first field; its instr is split
    end = old.index('<w:r><w:fldChar w:fldCharType="end"/></w:r>', start)
    assert new[:start] == old[:start]
    assert new[len(new) - (len(old) - end):] == old[end:]
    assert '<w:t xml:space="preserve">{Citation}</w:t></w:r><w:r><w:t></w:t>' in new[start:]
    assert new.count("{Citation}") == 3  # result + formatted + plain


def test_select_ambiguous_and_not_found(doc, tmp_path, fake_refresh):
    with pytest.raises(RefusedUnsafe, match=r"2 Zotero citations .* match current_text '\(3\)'.*cB.*cC"):
        edit(doc, tmp_path / "a.docx", ["K1"], para=2, current_text="(3)")
    with pytest.raises(RefusedUnsafe, match=r"no Zotero citation in paragraph 1 .*cA '\(1,2\)' items \['K1', 'K2'\]"):
        edit(doc, tmp_path / "b.docx", ["K1"], citation_id="nope")
    with pytest.raises(RefusedUnsafe, match="exactly one of citation_id or current_text"):
        edit(doc, tmp_path / "c.docx", ["K1"], citation_id="cA", current_text="(1,2)")
    with pytest.raises(RefusedUnsafe, match="exactly one"):
        edit(doc, tmp_path / "d.docx", ["K1"])
    edit(doc, tmp_path / "e.docx", ["K1"], para=2, citation_id="cC")
    assert keys_of(read_field(tmp_path / "e.docx", "cC")[0]) == ["K1"]
    assert keys_of(read_field(tmp_path / "e.docx", "cB")[0]) == ["K3"]
    assert fake_refresh[-1][1] == ("cC",)


@pytest.mark.parametrize("keys, message", [
    ([], "item_keys is empty"),
    (["K1", "K1"], r"\['K1'\] more than once"),
    (["K1", "ZZZZ"], "'ZZZZ' is not in the local library"),
    (["K1", "K3"], "is in the trash"),
])
def test_bad_keys_refused(doc, tmp_path, fake_refresh, keys, message):
    with pytest.raises(RefusedUnsafe, match=message):
        edit(doc, tmp_path / "o.docx", keys, citation_id="cA", zotero=FakeZotero(trashed={"K3"}))
    assert not (tmp_path / "o.docx").exists() and not fake_refresh


def test_field_inside_revision_or_nested_refused(doc, tmp_path, fake_refresh):
    with pytest.raises(RefusedUnsafe, match=r"inside tracked revision w:ins id=103 by 'Sam'"):
        edit(doc, tmp_path / "a.docx", ["K2"], para=3, citation_id="cD")
    with pytest.raises(RefusedUnsafe, match="nested field"):
        edit(doc, tmp_path / "b.docx", ["K1"], para=4, citation_id="cE")
    body = _body().replace('<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(3)</w:t></w:r>',
                           '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:del w:id="105" w:author="R" '
                           'w:date="2026-11-01T00:00:00Z"><w:r><w:delText>x</w:delText></w:r></w:del>'
                           '<w:r><w:t>(3)</w:t></w:r>', 1)
    crossing = make_docx(tmp_path / "crossing.docx", body)
    with pytest.raises(RefusedUnsafe, match=r"tracked revision w:del id=105 .* lies inside citation cB"):
        edit(crossing, tmp_path / "c.docx", ["K1"], para=2, citation_id="cB")
    assert not fake_refresh


def test_existing_output_refused_and_sha_checked(doc, tmp_path, fake_refresh):
    out = tmp_path / "exists.docx"
    out.write_bytes(b"x")
    with pytest.raises(RefusedUnsafe, match="existing output"):
        edit(doc, out, ["K1"], citation_id="cA")
    with pytest.raises(RefusedUnsafe, match="SHA-256 mismatch"):
        cluster_ops.edit_cluster_to_version(doc, tmp_path / "n.docx", source_sha256="0" * 64, paragraph_index=1,
                                            item_keys=["K1"], citation_id="cA", zotero=FakeZotero())


def test_in_place_edit_and_undo(doc, fake_refresh, monkeypatch):
    from research_mcp import server

    original = sha(doc)
    monkeypatch.setattr(cluster_ops, "ZoteroClient", FakeZotero)
    result = server.edit_citation_cluster(source=str(doc), source_sha256=original, paragraph_index=1,
                                          item_keys=["K2"], output=str(doc), citation_id="cA")
    data = result.model_dump() if hasattr(result, "model_dump") else result
    assert data["undo_token"] and data["sha256_before"] == original and sha(doc) != original
    assert keys_of(read_field(doc, "cA")[0]) == ["K2"]
    server.undo_document(path=str(doc))
    assert sha(doc) == original


def test_refresh_that_changes_a_revision_is_refused(doc, tmp_path, monkeypatch, fake_refresh):
    render = cluster_ops.refresh_file_and_wait

    def refresh(path, ids, max_wait_seconds=180.0):
        out = render(path, ids, max_wait_seconds)
        _rewrite(Path(path), doc_xml(Path(path)).replace("old words ", "old wordz "))
        return out

    monkeypatch.setattr(cluster_ops, "refresh_file_and_wait", refresh)
    with pytest.raises(VerificationFailed, match=r"1 tracked revision\(s\) changed.*id=102 by 'Reviewer'"):
        edit(doc, tmp_path / "o.docx", ["K1"], citation_id="cA")
    assert not (tmp_path / "o.docx").exists()


def test_refresh_serialization_only_change_accepted(doc, tmp_path, monkeypatch, fake_refresh):
    render = cluster_ops.refresh_file_and_wait

    def refresh(path, ids, max_wait_seconds=180.0):
        out = render(path, ids, max_wait_seconds)
        xml = doc_xml(Path(path)).replace('<w:r w:rsidR="00AA0001">', '<w:r w:rsidR="00BB0002">')
        xml = xml.replace('w:id="102" w:author="Reviewer"', 'w:author="Reviewer" w:id="102"')
        _rewrite(Path(path), xml)
        return out

    monkeypatch.setattr(cluster_ops, "refresh_file_and_wait", refresh)
    result = edit(doc, tmp_path / "o.docx", ["K1"], citation_id="cA")
    check = result["revision_check"]
    assert check["level"] == "serialization_only"
    assert (check["byte_identical"], check["serialization_changed"], check["markers_or_rsids_changed"]) == (2, 1, 1)


def test_refresh_that_changes_text_outside_fields_is_refused(doc, tmp_path, monkeypatch, fake_refresh):
    render = cluster_ops.refresh_file_and_wait

    def refresh(path, ids, max_wait_seconds=180.0):
        out = render(path, ids, max_wait_seconds)
        _rewrite(Path(path), doc_xml(Path(path)).replace(" more text.", " more text!"))
        return out

    monkeypatch.setattr(cluster_ops, "refresh_file_and_wait", refresh)
    with pytest.raises(VerificationFailed, match="text outside the Zotero fields changed"):
        edit(doc, tmp_path / "o.docx", ["K1"], citation_id="cA")


MOVE = ('<w:p><w:moveFromRangeStart w:id="200" w:author="Sam" w:date="2026-11-01T00:00:00Z" w:name="move1"/>'
        '<w:moveFrom w:id="201" w:author="Sam" w:date="2026-11-01T00:00:00Z"><w:r><w:t>moved</w:t></w:r></w:moveFrom>'
        '<w:moveFromRangeEnd w:id="200"/></w:p>'
        '<w:p><w:moveToRangeStart w:id="202" w:author="Sam" w:date="2026-11-01T00:00:00Z" w:name="move1"/>'
        '<w:moveTo w:id="203" w:author="Sam" w:date="2026-11-01T00:00:00Z"><w:r><w:t>moved</w:t></w:r></w:moveTo>'
        '<w:moveToRangeEnd w:id="202"/></w:p>')


@pytest.mark.parametrize("change, message", [
    (('<w:moveFromRangeEnd w:id="200"/>', '<w:moveFromRangeEnd w:id="999"/>'), r"1 tracked revision\(s\) changed"),
    (('<w:moveToRangeEnd w:id="202"/>', ''), "tracked revision count changed from 10 to 9"),
])
def test_refresh_that_changes_a_move_range_marker_is_refused(tmp_path, monkeypatch, fake_refresh, change, message):
    """Review r1 #1: move-range markers are part of the revision snapshot (a changed id or a missing marker
    used to publish as byte_identical)."""
    source = make_docx(tmp_path / "move.docx", _body().replace("<w:sectPr/>", MOVE + "<w:sectPr/>"))
    render = cluster_ops.refresh_file_and_wait

    def refresh(path, ids, max_wait_seconds=180.0):
        out = render(path, ids, max_wait_seconds)
        _rewrite(Path(path), doc_xml(Path(path)).replace(*change))
        return out

    monkeypatch.setattr(cluster_ops, "refresh_file_and_wait", refresh)
    with pytest.raises(VerificationFailed, match=message):
        edit(source, tmp_path / "o.docx", ["K1"], citation_id="cA")
    assert not (tmp_path / "o.docx").exists()


COLOR = '<w:r><w:rPr><w:color w:val="12AB34"/></w:rPr><w:t xml:space="preserve">Intro text '


def _save_elsewhere(path: Path) -> None:
    _rewrite(path, doc_xml(path).replace('<w:r><w:t xml:space="preserve">Intro text ', COLOR, 1))


def test_source_saved_after_staging_is_not_overwritten(doc, monkeypatch, fake_refresh):
    """Review r1 #2 (race_probe): a save to the source while the edit runs refuses; the save survives."""
    plan = cluster_ops._new_payload

    def payload(*args):
        result = plan(*args)
        _save_elsewhere(doc)
        return result

    monkeypatch.setattr(cluster_ops, "_new_payload", payload)
    original = sha(doc)
    with pytest.raises(Exception, match="(?i)sha|changed"):
        cluster_ops.edit_cluster_to_version(doc, doc, source_sha256=original, paragraph_index=1,
                                            item_keys=["K1"], citation_id="cA", zotero=FakeZotero())
    assert COLOR in doc_xml(doc) and keys_of(read_field(doc, "cA")[0]) == ["K1", "K2"]


@pytest.mark.parametrize("in_place", [True, False])
def test_source_saved_between_check_and_staging_refuses(doc, tmp_path, monkeypatch, fake_refresh, in_place):
    """Review r1 #2: the staged copy must match the caller's sha, not the file's sha at staging time."""
    real_audit = cluster_ops.audit_docx
    calls = []

    def audit(path):
        result = real_audit(path)
        if not calls:
            _save_elsewhere(doc)  # the save lands right after the sha check
        calls.append(path)
        return result

    monkeypatch.setattr(cluster_ops, "audit_docx", audit)
    original = sha(doc)
    out = doc if in_place else tmp_path / "v.docx"
    with pytest.raises(RefusedUnsafe, match="source changed after its SHA-256 was checked"):
        cluster_ops.edit_cluster_to_version(doc, out, source_sha256=original, paragraph_index=1,
                                            item_keys=["K1"], citation_id="cA", zotero=FakeZotero())
    assert COLOR in doc_xml(doc) and keys_of(read_field(doc, "cA")[0]) == ["K1", "K2"]
    assert in_place or not out.exists()
    assert not fake_refresh


def test_trashed_kept_key_refused(doc, tmp_path, fake_refresh):
    """Review r1 #3: a kept item that is in the trash refuses like a new one."""
    with pytest.raises(RefusedUnsafe, match="'K1' is in the trash"):
        edit(doc, tmp_path / "o.docx", ["K1"], citation_id="cA", zotero=FakeZotero(trashed={"K1"}))
    assert not (tmp_path / "o.docx").exists()
