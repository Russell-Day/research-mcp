"""Plan 864 decision 21: Zotero document prefs (ZOTERO_PREF_n) and set_citation_style (zotero_prefs.py).

Synthetic fixtures only: a small docx (test_refresh_citations' body) plus docProps/custom.xml holding the
Zotero data in 255-character chunks, research-mcp's track-changes property and one other property; a fake
Zotero styles folder; the refresh faked as in test_refresh_citations.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from research_mcp import refresh_ops, zotero_prefs
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.zotero_prefs import (
    chunks,
    installed_styles,
    parse_data,
    read_prefs,
    resolve_style,
    set_style_to_version,
    write_prefs,
)

from test_cluster_ops import make_docx, revisions_raw, sha
from test_refresh_citations import FakeRefresh, body

URI = "http://www.zotero.org/styles/"
DATA = ('<data data-version="3" zotero-version="9.0.6"><session id="AbCd1234"/>'
        f'<style id="{URI}nature" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
        '<prefs><pref name="fieldType" value="Field"/><pref name="automaticJournalAbbreviations" value="true"/>'
        '</prefs></data>')
TRACK = ('<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="research_mcp_track_changes">'
         '<vt:lpwstr>{"on": true, "author": "Claude"}</vt:lpwstr></property>')
OTHER = ('<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="9" name="Project">'
         '<vt:lpwstr>a &amp; b</vt:lpwstr></property>')
CUSTOM_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
               '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
               'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        'officeDocument" Target="word/document.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.'
        'org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>')


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def pref_props(data: str, first_pid: int = 3) -> str:
    return "".join(f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{first_pid + i}" '
                   f'name="ZOTERO_PREF_{i + 1}"><vt:lpwstr>{_esc(chunk)}</vt:lpwstr></property>'
                   for i, chunk in enumerate(chunks(data)))


def custom_xml(data: str | None = DATA) -> str:
    return CUSTOM_HEAD + TRACK + (pref_props(data) if data is not None else "") + OTHER + "</Properties>"


def add_custom(path: Path, custom: str) -> Path:
    with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("_rels/.rels", RELS)
        archive.writestr("docProps/custom.xml", custom)
    return path


def make_doc(path: Path, data: str | None = DATA, **body_kw) -> Path:
    return add_custom(make_docx(path, body(**body_kw)), custom_xml(data))


def custom_of(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("docProps/custom.xml").decode("utf-8")


def entries(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def csl(style_id: str, title: str, cls: str = "in-text", bibliography: bool = True, parent: str | None = None) -> str:
    link = f'<link href="{parent}" rel="independent-parent"/>' if parent else ""
    tail = "" if parent else "<citation><layout/></citation>" + ("<bibliography><layout/></bibliography>"
                                                                  if bibliography else "")
    return ('<?xml version="1.0" encoding="utf-8"?><style xmlns="http://purl.org/net/xbiblio/csl" '
            f'class="{cls}" version="1.0"><info><title>{title}</title><id>{style_id}</id>{link}</info>{tail}</style>')


@pytest.fixture
def styles(tmp_path: Path, monkeypatch) -> Path:
    folder = tmp_path / "styles"
    (folder / "hidden").mkdir(parents=True)
    (folder / "american-medical-association.csl").write_text(
        csl(f"{URI}american-medical-association", "American Medical Association 11th edition"), encoding="utf-8")
    (folder / "nature.csl").write_text(csl(f"{URI}nature", "Nature"), encoding="utf-8")
    (folder / "downloaded-copy.csl").write_text(csl(f"{URI}journal-x", "Journal X"), encoding="utf-8")
    (folder / "chicago-note-bibliography.csl").write_text(
        csl(f"{URI}chicago-note-bibliography", "Chicago Manual of Style (note)", cls="note"), encoding="utf-8")
    (folder / "citations-only.csl").write_text(csl(f"{URI}citations-only", "Citations only", bibliography=False),
                                               encoding="utf-8")
    # A dependent style: class and bibliography come from its parent (here a note style).
    (folder / "hidden" / "some-journal.csl").write_text(
        csl(f"{URI}some-journal", "Some Journal", cls="in-text", parent=f"{URI}chicago-note-bibliography"),
        encoding="utf-8")
    (folder / "orphan.csl").write_text(csl(f"{URI}orphan", "Orphan", parent=f"{URI}not-installed"), encoding="utf-8")
    (folder / ".hidden-file.csl").write_text(csl(f"{URI}dotfile", "Dot"), encoding="utf-8")
    (folder / "zz-duplicate.csl").write_text(csl(f"{URI}nature", "Nature copy"), encoding="utf-8")
    monkeypatch.setenv(zotero_prefs.STYLES_DIR_ENV, str(folder))
    return folder


class Recording(FakeRefresh):
    """FakeRefresh that also records the Zotero data the refresh saw."""

    def __init__(self):
        super().__init__()
        self.prefs_seen: list[str | None] = []

    def __call__(self, path, ids, max_wait_seconds=300.0):
        self.prefs_seen.append(read_prefs(path))
        return super().__call__(path, ids, max_wait_seconds)


@pytest.fixture
def fake(monkeypatch):
    refresh = Recording()
    monkeypatch.setattr(refresh_ops, "refresh_file_and_wait", refresh)
    return refresh


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    return make_doc(tmp_path / "source.docx")


# -- Case A/B: prefs read, round trip, rechunk ---------------------------------------------------------------


def test_round_trip_is_byte_identical(doc, tmp_path):
    assert read_prefs(doc) == DATA and len(chunks(DATA)) == 2
    out = tmp_path / "same.docx"
    result = write_prefs(doc, out, read_prefs(doc))
    assert result["changed"] is False and out.read_bytes() == doc.read_bytes()
    # The XML-level rewrite with the same data is the identity too.
    text = custom_of(doc)
    assert zotero_prefs._custom_xml_with(text, DATA) == text


def test_long_data_rechunked_others_untouched(doc, tmp_path):
    long_data = DATA.replace("</prefs>", f'<pref name="pad" value="{"v" * (600 - len(DATA) - 27)}"/></prefs>')
    assert len(long_data) == 600
    out = tmp_path / "long.docx"
    assert write_prefs(doc, out, long_data) == {"changed": True, "chunks_before": 2, "chunks_after": 3}
    text = custom_of(out)
    props = re.findall(r'pid="(\d+)" name="ZOTERO_PREF_(\d+)"><vt:lpwstr>(.*?)</vt:lpwstr>', text)
    assert [(pid, n) for pid, n, _ in props] == [("3", "1"), ("4", "2"), ("10", "3")]  # new pid after max (9)
    values = [v.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&") for _, _, v in props]
    assert [len(v) for v in values] == [255, 255, 90] and "".join(values) == long_data
    assert read_prefs(out) == long_data
    # The new chunk follows the last ZOTERO_PREF; every other property is byte-identical and in place.
    assert text.index("ZOTERO_PREF_3") < text.index('name="Project"')
    assert TRACK in text and OTHER in text and text.startswith(CUSTOM_HEAD + TRACK)
    before, after = entries(doc), entries(out)
    assert list(before) == list(after)
    assert all(before[name] == after[name] for name in before if name != "docProps/custom.xml")
    # Shrinking back removes the extra properties and gives the original custom.xml back.
    back = tmp_path / "back.docx"
    write_prefs(out, back, DATA)
    assert custom_of(back) == custom_of(doc)


def test_chunks_count_utf16_units():
    assert [len(c) for c in chunks("a" * 600)] == [255, 255, 90]
    text = "a" * 254 + "\U0001F600" + "b"  # the emoji is 2 UTF-16 units: 254 + 2 crosses 255
    with pytest.raises(RefusedUnsafe, match="split a character"):
        chunks(text)
    assert chunks("") == [""]


def test_unnumbered_prefs_refused(tmp_path):
    custom = custom_xml().replace("ZOTERO_PREF_2", "ZOTERO_PREF_3")
    path = add_custom(make_docx(tmp_path / "gap.docx", body()), custom)
    with pytest.raises(RefusedUnsafe, match=r"not numbered 1..n \(\[1, 3\]\)"):
        read_prefs(path)


# -- Case C/D: installed styles ------------------------------------------------------------------------------


def test_resolve_by_id_short_name_title(styles):
    for query in (f"{URI}american-medical-association", "american-medical-association",
                  "American Medical Association 11th edition", "AMERICAN-MEDICAL-ASSOCIATION"):
        style = resolve_style(query)
        assert style.style_id == f"{URI}american-medical-association" and style.has_bibliography
        assert style.style_class == "in-text" and style.short_name == "american-medical-association"
    # The short name is the id's tail, whatever the file is called; the file stem works too.
    assert resolve_style("journal-x").file.endswith("downloaded-copy.csl")
    assert resolve_style("downloaded-copy").style_id == f"{URI}journal-x"


def test_installed_styles_as_zotero_loads_them(styles):
    found = installed_styles()
    assert f"{URI}dotfile" not in found  # names starting with "." are skipped
    assert f"{URI}orphan" not in found  # a dependent style needs its parent installed
    assert Path(found[f"{URI}nature"].file).name == "nature.csl"  # the first file wins for a duplicate id
    assert found[f"{URI}citations-only"].has_bibliography is False


def test_dependent_style_uses_parent(styles):
    style = resolve_style("some-journal")
    assert style.hidden and style.parent == f"{URI}chicago-note-bibliography"
    assert style.style_class == "note" and style.has_bibliography  # the parent's, not its own


def test_uninstalled_style_refused_naming_close_ones(styles, doc, tmp_path, fake):
    with pytest.raises(RefusedUnsafe, match=r"style 'american-medical-assoc' is not installed in Zotero .*closest "
                                             r"installed: american-medical-association"):
        resolve_style("american-medical-assoc")
    with pytest.raises(RefusedUnsafe, match=r"closest installed: american-medical-association"):  # a typo
        resolve_style("american-medical-asociation")
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="not installed"):
        set_style_to_version(doc, out, source_sha256=sha(doc), style=f"{URI}apa")
    assert not out.exists() and not fake.calls


def test_styles_dir_from_zotero_profile(tmp_path, monkeypatch):
    monkeypatch.delenv(zotero_prefs.STYLES_DIR_ENV, raising=False)
    root = tmp_path / "appdata" / "Zotero" / "Zotero"
    profile = root / "Profiles" / "abc.default"
    profile.mkdir(parents=True)
    (root / "profiles.ini").write_text("[General]\nStartWithLastProfile=1\n\n[Profile0]\nName=default\n"
                                       "IsRelative=1\nPath=Profiles/abc.default\nDefault=1\n", encoding="utf-8")
    data_dir = tmp_path / "zdata"
    escaped = str(data_dir).replace("\\", "\\\\")
    (profile / "prefs.js").write_text(f'user_pref("extensions.zotero.dataDir", "{escaped}");\n'
                                      'user_pref("extensions.zotero.useDataDir", true);\n', encoding="utf-8")
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    assert zotero_prefs.styles_dir() == data_dir / "styles"
    (profile / "prefs.js").write_text('user_pref("extensions.zotero.useDataDir", false);\n', encoding="utf-8")
    assert zotero_prefs.styles_dir() == Path.home() / "Zotero" / "styles"
    monkeypatch.setenv(zotero_prefs.STYLES_DIR_ENV, str(tmp_path / "elsewhere"))
    assert zotero_prefs.styles_dir() == tmp_path / "elsewhere"


# -- Case E/F/G: set_citation_style --------------------------------------------------------------------------


def test_applied_sets_style_then_refreshes(styles, doc, tmp_path, fake):
    out = tmp_path / "out.docx"
    revisions = revisions_raw(doc)
    result = set_style_to_version(doc, out, source_sha256=sha(doc), style="american-medical-association")
    # The refresh saw the new style, written as Zotero's Document Preferences writes it.
    assert len(fake.prefs_seen) == 1
    assert fake.prefs_seen[0] == DATA.replace(
        f'<style id="{URI}nature" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>',
        f'<style id="{URI}american-medical-association" hasBibliography="1" bibliographyStyleHasBeenSet="0"/>')
    assert read_prefs(out) == fake.prefs_seen[0]
    assert result["style_before"]["style_id"] == f"{URI}nature"
    assert result["style_after"]["style_id"] == f"{URI}american-medical-association"
    assert result["style"]["title"] == "American Medical Association 11th edition"
    assert result["session_renewed"] is False and result["preview"] is False
    assert result["source_sha256"] == sha(doc) and result["output"] == str(out.resolve())
    # Same result as refresh_citations; the bibliography layout warning is there.
    assert result["citations_refreshed"] == 3 and result["bibliography_unchanged_count"] == 3
    assert any("Bibliography paragraph style" in w for w in result["warnings"])
    # Every other custom property byte-identical; revisions unchanged.
    text = custom_of(out)
    assert TRACK in text and OTHER in text
    assert revisions_raw(out) == revisions
    assert result["revisions_checked"] >= len(revisions) and result["revision_check"]


def test_preview_writes_nothing(styles, doc, tmp_path, fake):
    original = sha(doc)
    out = tmp_path / "out.docx"
    result = set_style_to_version(doc, out, source_sha256=original, style="american-medical-association",
                                  preview=True)
    assert result["preview"] is True and result["output"] is None and result["output_sha256"] is None
    assert sha(doc) == original and not out.exists()
    assert fake.calls and not fake.calls[0][0].exists() and fake.calls[0][0].parent != doc.parent
    assert result["style_after"]["style_id"].endswith("american-medical-association")
    result = set_style_to_version(doc, doc, source_sha256=original, style="american-medical-association",
                                  preview=True)
    assert sha(doc) == original and read_prefs(doc) == DATA


def test_citation_inside_revision_refused(styles, tmp_path, fake):
    doc = make_doc(tmp_path / "wrapped.docx", wrap_b=True)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=r"set_citation_style: refresh_citations: citation cB sits inside "
                                             r"tracked revision w:ins id=105"):
        set_style_to_version(doc, out, source_sha256=sha(doc), style="american-medical-association")
    assert not fake.calls and not out.exists()


def test_citation_inside_deletion_refused(styles, tmp_path, fake):
    # A tracked deletion turns the field's instrText into delInstrText: the refresh path does not see it as a
    # citation at all, so set_citation_style checks for it itself.
    xml = body(wrap_b=True)
    start = xml.index('<w:ins w:id="105"')
    end = xml.index("</w:ins>", start) + len("</w:ins>")
    deleted = (xml[start:end].replace("<w:ins ", "<w:del ").replace("</w:ins>", "</w:del>")
               .replace("<w:instrText", "<w:delInstrText").replace("</w:instrText>", "</w:delInstrText>")
               .replace("<w:t>", "<w:delText>").replace("<w:t ", "<w:delText ").replace("</w:t>", "</w:delText>"))
    doc = add_custom(make_docx(tmp_path / "deleted.docx", xml[:start] + deleted + xml[end:]), custom_xml(DATA))
    assert zotero_prefs.zotero_fields_in_deletions(doc) == [{"field": "cB", "revision": "w:del id=105 by 'Sam'"}]
    out = tmp_path / "out.docx"
    for preview in (True, False):
        with pytest.raises(RefusedUnsafe, match=r"set_citation_style: 1 Zotero field\(s\) sit inside a tracked "
                                                 r"deletion.*cB in w:del id=105 by 'Sam'"):
            set_style_to_version(doc, out, source_sha256=sha(doc), style="american-medical-association",
                                 preview=preview)
    assert not fake.calls and not out.exists()
    assert zotero_prefs.zotero_fields_in_deletions(make_doc(tmp_path / "live.docx", wrap_b=True)) == []


def test_refresh_that_changes_another_property_refused(styles, doc, tmp_path, fake, monkeypatch):
    def dropping(path, ids, max_wait_seconds=300.0):
        result = fake(path, ids, max_wait_seconds)
        temp = Path(str(path) + ".rw")
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "docProps/custom.xml":
                    data = data.replace(b'"on": true', b'"on": false')
                dst.writestr(info, data)
        temp.replace(path)
        return result

    monkeypatch.setattr(refresh_ops, "refresh_file_and_wait", dropping)
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="a custom document property other than ZOTERO_PREF_n changed"):
        set_style_to_version(doc, out, source_sha256=sha(doc), style="american-medical-association")
    assert not out.exists()


def test_sha_mismatch_refused(styles, doc, tmp_path, fake):
    with pytest.raises(RefusedUnsafe, match="SHA-256 mismatch"):
        set_style_to_version(doc, tmp_path / "out.docx", source_sha256="0" * 64, style="nature")
    assert not fake.calls


# -- Case I: refusals ----------------------------------------------------------------------------------------


def test_class_switch_refused(styles, doc, tmp_path, fake):
    with pytest.raises(RefusedUnsafe, match=r"chicago-note-bibliography is a note style and the document's "
                                             r"citations are in-text.*Document Preferences"):
        set_style_to_version(doc, tmp_path / "o1.docx", source_sha256=sha(doc), style="chicago-note-bibliography")
    with pytest.raises(RefusedUnsafe, match="some-journal is a note style"):  # the dependent's parent decides
        set_style_to_version(doc, tmp_path / "o2.docx", source_sha256=sha(doc), style="some-journal")
    note = make_doc(tmp_path / "note.docx", DATA.replace('value="Field"/>', 'value="Field"/><pref name="noteType" '
                                                         'value="1"/>'))
    with pytest.raises(RefusedUnsafe, match="american-medical-association is an in-text style and the document's "
                                            "citations are note"):
        set_style_to_version(note, tmp_path / "o3.docx", source_sha256=sha(note), style="american-medical-association")
    assert not fake.calls


def test_no_bibliography_style_refused(styles, doc, tmp_path, fake):
    with pytest.raises(RefusedUnsafe, match="citations-only has no bibliography and the document has a Zotero "
                                            "bibliography field"):
        set_style_to_version(doc, tmp_path / "out.docx", source_sha256=sha(doc), style="citations-only")
    assert not fake.calls


def test_same_style_refused_locale_change_renews_session(styles, doc, tmp_path, fake):
    with pytest.raises(RefusedUnsafe, match=r"already uses .*nature; call refresh_citations"):
        set_style_to_version(doc, tmp_path / "o1.docx", source_sha256=sha(doc), style="nature")
    with pytest.raises(RefusedUnsafe, match="not a locale code"):
        set_style_to_version(doc, tmp_path / "o2.docx", source_sha256=sha(doc), style="nature", locale="en GB")
    out = tmp_path / "o3.docx"
    result = set_style_to_version(doc, out, source_sha256=sha(doc), style="nature", locale="en-GB")
    seen = parse_data(fake.prefs_seen[0])
    assert result["session_renewed"] is True and seen.session_id != "AbCd1234" and len(seen.session_id) == 8
    assert seen.locale == "en-GB" and seen.style_id == f"{URI}nature" and not seen.bibliography_style_has_been_set
    assert f'<style id="{URI}nature" locale="en-GB" hasBibliography="1" bibliographyStyleHasBeenSet="0"/>' in \
        fake.prefs_seen[0]
    # A style change keeps the document's locale and session.
    localized = make_doc(tmp_path / "loc.docx", DATA.replace(f'{URI}nature" ', f'{URI}nature" locale="de-DE" '))
    set_style_to_version(localized, tmp_path / "o4.docx", source_sha256=sha(localized), style="american-medical-association")
    seen = parse_data(fake.prefs_seen[1])
    assert seen.locale == "de-DE" and seen.session_id == "AbCd1234"


def test_virgin_and_legacy_documents_refused(styles, tmp_path, fake):
    virgin = make_doc(tmp_path / "virgin.docx", None)
    with pytest.raises(RefusedUnsafe, match="no Zotero preferences"):
        set_style_to_version(virgin, tmp_path / "o1.docx", source_sha256=sha(virgin), style="nature")
    legacy = make_doc(tmp_path / "legacy.docx", "AbCd1234:http://www.zotero.org/styles/nature:in-text:1:0:0")
    with pytest.raises(RefusedUnsafe, match="pre-Zotero-5 format"):
        set_style_to_version(legacy, tmp_path / "o2.docx", source_sha256=sha(legacy), style="nature")
    assert not fake.calls


def test_json_data_version_4(styles, tmp_path, fake):
    data = ('{"style":{"styleID":"http://www.zotero.org/styles/nature","locale":"en-US","hasBibliography":true,'
            '"bibliographyStyleHasBeenSet":true},"prefs":{"fieldType":"Field","noteType":0},'
            '"sessionID":"AbCd1234","zoteroVersion":"9.0.6","dataVersion":4}')
    doc = make_doc(tmp_path / "json.docx", data)
    set_style_to_version(doc, tmp_path / "out.docx", source_sha256=sha(doc), style="american-medical-association")
    assert fake.prefs_seen[0] == data.replace("styles/nature", "styles/american-medical-association").replace(
        '"bibliographyStyleHasBeenSet":true', '"bibliographyStyleHasBeenSet":false')


# -- Case H: through the server, in place, undo --------------------------------------------------------------


def test_in_place_and_undo_through_server(styles, doc, fake):
    from research_mcp import server

    original = sha(doc)
    preview = server.set_citation_style(source=str(doc), source_sha256=original, style="american-medical-association",
                                        output=str(doc), preview=True)
    data = preview.model_dump() if hasattr(preview, "model_dump") else preview
    assert data["undo_token"] is None and sha(doc) == original
    result = server.set_citation_style(source=str(doc), source_sha256=original, style="american-medical-association",
                                       output=str(doc))
    data = result.model_dump() if hasattr(result, "model_dump") else result
    assert data["undo_token"] and data["sha256_before"] == original and sha(doc) != original
    assert parse_data(read_prefs(doc)).style_id == f"{URI}american-medical-association"
    server.undo_document(path=str(doc))
    assert sha(doc) == original and read_prefs(doc) == DATA


# -- S7 review r1: fields the refresh path does not see; note styles; property attributes -------------------

from test_cluster_ops import _entry, _field  # noqa: E402

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
NOTE_DATA = DATA.replace('value="Field"/>', 'value="Field"/><pref name="noteType" value="1"/>')


def _add_part(path: Path, name: str, xml: str) -> Path:
    with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(name, xml)
    return path


def test_note_style_document_refused(styles, tmp_path, fake):
    # Note-style citations live in footnotes, which the refresh neither resets nor checks: refuse note -> note too.
    note = make_doc(tmp_path / "note.docx", NOTE_DATA)
    with pytest.raises(RefusedUnsafe, match=r"set_citation_style: the document uses a note style .*footnotes or "
                                             r"endnotes, which the refresh does not reset or check"):
        set_style_to_version(note, tmp_path / "o.docx", source_sha256=sha(note), style="chicago-note-bibliography")
    assert not fake.calls and not (tmp_path / "o.docx").exists()


def test_citation_in_footnote_refused(styles, tmp_path, fake):
    doc = _add_part(make_doc(tmp_path / "fn.docx"), "word/footnotes.xml",
                    f'<w:footnotes {W_NS}><w:footnote w:id="1"><w:p>{_field("cF", [_entry("K1")], ["(1)"])}'
                    '</w:p></w:footnote></w:footnotes>')
    assert zotero_prefs.unchecked_zotero_fields(doc) == [
        {"field": "cF", "part": "word/footnotes.xml", "reason": "part"}]
    for preview in (True, False):
        with pytest.raises(RefusedUnsafe, match=r"set_citation_style: 1 Zotero field\(s\) are outside what the "
                                                 r"refresh resets and checks.*cF in footnotes\.xml"):
            set_style_to_version(doc, tmp_path / "o.docx", source_sha256=sha(doc),
                                 style="american-medical-association", preview=preview)
    assert not fake.calls and not (tmp_path / "o.docx").exists()


def test_simple_field_citation_refused(styles, tmp_path, fake):
    simple = ('<w:p><w:fldSimple w:instr=" ADDIN ZOTERO_ITEM CSL_CITATION {&quot;citationID&quot;:&quot;cS&quot;}">'
              '<w:r><w:t>(9)</w:t></w:r></w:fldSimple></w:p>')
    xml = body().replace("<w:sectPr/>", simple + "<w:sectPr/>")
    doc = add_custom(make_docx(tmp_path / "simple.docx", xml), custom_xml(DATA))
    with pytest.raises(RefusedUnsafe, match=r"cS stored as a simple field \(w:fldSimple\) in document\.xml"):
        set_style_to_version(doc, tmp_path / "o.docx", source_sha256=sha(doc), style="american-medical-association")
    assert not fake.calls


def test_citation_in_move_from_refused(styles, tmp_path, fake):
    xml = body(wrap_b=True)
    start = xml.index('<w:ins w:id="105"')
    end = xml.index("</w:ins>", start) + len("</w:ins>")
    moved = (xml[start:end].replace("<w:ins ", '<w:moveFrom ').replace("</w:ins>", "</w:moveFrom>")
             .replace("<w:instrText", "<w:delInstrText").replace("</w:instrText>", "</w:delInstrText>")
             .replace("<w:t>", "<w:delText>").replace("<w:t ", "<w:delText ").replace("</w:t>", "</w:delText>"))
    doc = add_custom(make_docx(tmp_path / "moved.docx", xml[:start] + moved + xml[end:]), custom_xml(DATA))
    with pytest.raises(RefusedUnsafe, match=r"1 Zotero field\(s\) sit inside a tracked deletion.*cB in "
                                             r"w:moveFrom id=105 by 'Sam' \(document\.xml\)"):
        set_style_to_version(doc, tmp_path / "o.docx", source_sha256=sha(doc), style="american-medical-association")
    assert not fake.calls


def _rewriting_custom(fake, old: bytes, new: bytes):
    def refresh(path, ids, max_wait_seconds=300.0):
        result = fake(path, ids, max_wait_seconds)
        temp = Path(str(path) + ".rw")
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "docProps/custom.xml":
                    assert old in data
                    data = data.replace(old, new)
                dst.writestr(info, data)
        temp.replace(path)
        return result
    return refresh


def test_other_property_attribute_change_refused_pid_renumbering_allowed(styles, doc, tmp_path, fake, monkeypatch):
    changed = b'fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AF}" pid="9" name="Project"'
    monkeypatch.setattr(refresh_ops, "refresh_file_and_wait", _rewriting_custom(
        fake, b'fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="9" name="Project"', changed))
    with pytest.raises(VerificationFailed, match="a custom document property other than ZOTERO_PREF_n changed"):
        set_style_to_version(doc, tmp_path / "o1.docx", source_sha256=sha(doc), style="american-medical-association")
    assert not (tmp_path / "o1.docx").exists()
    # Word numbers pids itself (and renumbers after a removed chunk): a pid change alone is accepted.
    monkeypatch.setattr(refresh_ops, "refresh_file_and_wait", _rewriting_custom(fake, b'pid="9"', b'pid="20"'))
    result = set_style_to_version(doc, tmp_path / "o2.docx", source_sha256=sha(doc),
                                  style="american-medical-association")
    assert result["style_after"]["style_id"] == f"{URI}american-medical-association"
    # ... but every pid must stay present, valid and unique.
    for n, (new, problem) in enumerate([(b'pid="2"', "Project and research_mcp_track_changes share pid 2"),
                                        (b"", "Project has no pid"),
                                        (b'pid="x"', "Project has pid 'x', not an integer from 2"),
                                        ('pid="\u0662"'.encode(), "Project has pid '\u0662', not an integer"),
                                        ('pid="\u00b2"'.encode(), "Project has pid '\u00b2', not an integer"),
                                        (b'pid="2147483648"', "Project has pid '2147483648', not an integer")]):
        monkeypatch.setattr(refresh_ops, "refresh_file_and_wait", _rewriting_custom(fake, b'pid="9"', new))
        out = tmp_path / f"bad{n}.docx"
        with pytest.raises(VerificationFailed, match=r"not numbered correctly: .*" + re.escape(problem)):
            set_style_to_version(doc, out, source_sha256=sha(doc), style="american-medical-association")
        assert not out.exists()
