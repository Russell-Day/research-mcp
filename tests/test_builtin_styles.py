"""Plan 973 S3: a built-in paragraph style the document has not defined yet is added as Word itself adds it.

`builtin_styles.plan` asks a private Word, per document, on a styles-only scratch copy, and splices exactly the
`w:style` elements Word added. The non-live tests replace that Word with `fake_word`, which answers with what Word
16.0 added to the synthetic CV document (`WORD_ADDS`; `test_live_word_answers_match_the_fake` keeps it honest),
and test `extract` on the four recorded Word answers in `tests/data/builtin_styles/*.json` (styles.xml before
Word's first use of each style, and the elements Word added; Word 16.0 build 16.0.20527, 2026-11-18). The live
tests compare against Word applying the style to a copy of the document itself: an italic Normal (Word's Caption
turns italic off), a locked/semiHidden latent caption, and a "Heading 2 Char" id or name already taken.
Synthetic documents only.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from research_mcp import builtin_styles, editing, revisions
from research_mcp.audit import file_sha256
from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.sections import canonical
from test_revisions import W, document_xml, run_batch, windows_only
from test_revisions_insert_paragraph import make, para, run

DATA = Path(__file__).parent / "data" / "builtin_styles"
FIXTURES = sorted(DATA.glob("*.json"))
REVIEWER = "Reviewer B"
LIVE = os.environ.get("RESEARCH_MCP_WORD_LIVE") == "1"
live = pytest.mark.skipif(not LIVE, reason="needs RESEARCH_MCP_WORD_LIVE=1")
WD = {"Caption": -35, "Heading 2": -3}  # WdBuiltinStyle ids, also used by the live ground truth

# What Word 16.0 adds to CV_STYLES (below) on first use, as `builtin_styles.ask_word` reads it.
WORD_ADDS = {
    -35: ['<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/>'
          '<w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/><w:pPr>'
          '<w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:i/><w:iCs/>'
          '<w:color w:val="0E2841" w:themeColor="text2"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>'],
    -3: ['<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>'
         '<w:next w:val="Normal"/><w:link w:val="Heading2Char"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/>'
         '<w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/>'
         '<w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" '
         'w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>'
         '<w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="26"/><w:szCs w:val="26"/>'
         '</w:rPr></w:style>',
         '<w:style w:type="character" w:customStyle="1" w:styleId="Heading2Char"><w:name w:val="Heading 2 Char"/>'
         '<w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Heading2"/><w:uiPriority w:val="9"/><w:rPr>'
         '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" '
         'w:cstheme="majorBidi"/><w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/>'
         '<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>'],
}


def styles_of(path: Path) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/styles.xml")


def body_of(path: Path) -> ET.Element:
    return ET.fromstring(document_xml(path)).find(W + "body")


def style_elements(xml: bytes | str) -> dict[str, str]:
    text = xml.decode("utf-8") if isinstance(xml, bytes) else xml
    return {re.search(r'w:styleId="([^"]*)"', item).group(1): re.sub(r"<w:rsid\b[^>]*/>", "", item)
            for item in re.findall(r"<w:style\b[^>]*?(?:/>|>.*?</w:style>)", text, re.S)}


def plus(styles: bytes, block: str) -> bytes:
    return styles.replace(b"</w:styles>", block.encode("utf-8") + b"</w:styles>")


# -- the private Word, faked for non-live tests ---------------------------------------------------------------

@pytest.fixture(autouse=True)
def no_word_unless_live(monkeypatch):
    if not LIVE:
        def refuse(*args, **kwargs):
            raise AssertionError("a non-live test asked Word; use the fake_word fixture")
        monkeypatch.setattr(builtin_styles, "ask_word", refuse, raising=False)


def install_fake_word(monkeypatch, adds: dict[int, list[str]] | None = None) -> list[int]:
    """Replace the private Word with one that answers `adds` (default WORD_ADDS); returns the ids asked."""

    asked: list[int] = []

    def fake(scratch, wd_id, folder):
        asked.append(wd_id)
        with zipfile.ZipFile(scratch) as archive:
            body = archive.read("word/document.xml").decode("utf-8")
            baseline = archive.read("word/styles.xml")
        # the scratch package holds one plain paragraph and nothing of the real body
        assert body.count("<w:p>") == 1 and "<w:ins" not in body and "<w:del" not in body
        block = (adds or WORD_ADDS)[wd_id]
        main = re.search(r'w:styleId="([^"]*)"', block[0]).group(1)
        return builtin_styles.WordAnswer(baseline, plus(baseline, "".join(block)), main, "fake", 0.0)

    monkeypatch.setattr(builtin_styles, "ask_word", fake)
    return asked


@pytest.fixture
def fake_word(monkeypatch):
    return install_fake_word(monkeypatch)


# -- extract: Word's recorded answers ----------------------------------------------------------------------------

def _fixture_cases():
    for path in FIXTURES:
        for name in json.loads(path.read_text(encoding="utf-8"))["added"]:
            yield pytest.param(path, name, id=f"{path.stem}-{name}")


@pytest.mark.parametrize("fixture,name", list(_fixture_cases()))
def test_extract_takes_exactly_what_word_added(fixture, name):
    data = json.loads(fixture.read_text(encoding="utf-8"))
    styles = data["styles_xml"].encode("utf-8")
    word_added = {style_id: re.sub(r"<w:rsid\b[^>]*/>", "", xml) for style_id, xml in data["added"][name].items()}
    main = next(style_id for style_id, xml in word_added.items() if 'w:type="paragraph"' in xml)
    answer = builtin_styles.WordAnswer(styles, plus(styles, "".join(data["added"][name].values())), main)
    addition = builtin_styles.extract(styles, answer, name)
    assert addition.style_id == main and addition.names[0] == name
    assert {style_id: xml for style_id, _, xml in addition.definitions} == word_added


CV_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>{latent}"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph '
    'Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="SectionTitle"><w:name w:val="Section Title"/>'
    '<w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:caps/></w:rPr></w:style>'
    "</w:styles>")
CV_LATENT = '<w:latentStyles><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/></w:latentStyles>'
H2_CHAR = ('<w:style w:type="character" w:customStyle="1" w:styleId="Heading2Char"><w:name w:val="{name}"/>{link}'
           '<w:rPr><w:b/></w:rPr></w:style></w:styles>')


def cv_styles(**variant) -> str:
    """CV_STYLES, or one of the variants the review probed in Word."""

    styles = CV_STYLES.format(latent=CV_LATENT if variant.get("latent") else "")
    if variant.get("italic_normal"):
        styles = styles.replace('<w:name w:val="Normal"/><w:qFormat/>',
                                '<w:name w:val="Normal"/><w:qFormat/><w:rPr><w:i/><w:iCs/><w:sz w:val="18"/></w:rPr>')
    if variant.get("locked_latent"):
        styles = styles.replace("</w:docDefaults>", '</w:docDefaults><w:latentStyles><w:lsdException '
                                'w:name="caption" w:locked="1" w:semiHidden="1"/></w:latentStyles>')
    if "h2char" in variant:
        name, link = variant["h2char"]
        styles = styles.replace("</w:styles>", H2_CHAR.format(name=name, link=link))
    return styles


def _answer(source_styles: str, block: str, main: str, *, baseline: str | None = None,
            applied_base: str | None = None) -> builtin_styles.WordAnswer:
    base = (baseline or source_styles).encode("utf-8")
    return builtin_styles.WordAnswer(base, plus((applied_base or baseline or source_styles).encode("utf-8"), block),
                                     main)


def test_extract_refuses_when_word_changes_an_existing_style():
    """Word re-ids an unrelated style occupying the linked style's id (live: "My Character Style" in Heading2Char
    is saved as MyCharacterStyle): the occupant changes, so nothing is spliced."""

    styles = cv_styles(h2char=("My Character Style", ""))
    applied = styles.replace(H2_CHAR.format(name="My Character Style", link="").replace("</w:styles>", ""),
                             '<w:style w:type="character" w:customStyle="1" w:styleId="MyCharacterStyle">'
                             '<w:name w:val="My Character Style"/><w:rPr><w:b/></w:rPr></w:style>')
    answer = _answer(styles, "".join(WORD_ADDS[-3]), "Heading2", applied_base=applied)
    with pytest.raises(RefusedUnsafe, match="existing style.*'My Character Style'"):
        builtin_styles.extract(styles.encode("utf-8"), answer, "Heading 2")


def test_extract_names_the_style_an_addition_clashes_with():
    styles = cv_styles(h2char=("My Character Style", ""))
    with pytest.raises(RefusedUnsafe, match=r"'Heading 2 Char' \(id 'Heading2Char'\), which clashes with the "
                                            r"document's style 'My Character Style' \(id 'Heading2Char'\)"):
        builtin_styles.extract(styles.encode("utf-8"),
                               _answer(CV_STYLES.format(latent=""), "".join(WORD_ADDS[-3]), "Heading2"), "Heading 2")


def test_extract_accepts_the_free_id_word_picks():
    """A "Heading 2 Char" that is not linked to Heading 2 stays; Word adds and links "Heading 2 Char1" (live)."""

    styles = cv_styles(h2char=("Heading 2 Char", ""))
    block = "".join(WORD_ADDS[-3]).replace("Heading2Char", "Heading2Char1").replace(
        "Heading 2 Char", "Heading 2 Char1")
    addition = builtin_styles.extract(styles.encode("utf-8"), _answer(styles, block, "Heading2"), "Heading 2")
    assert addition.names == ["Heading 2", "Heading 2 Char1"]
    assert '<w:link w:val="Heading2Char1"/>' in addition.definitions[0][2]


def test_extract_refuses_a_link_that_is_not_returned():
    styles = cv_styles(h2char=("Unrelated", ""))
    # Word adds only Heading 2, linked to the document's own Heading2Char, which does not link back
    with pytest.raises(RefusedUnsafe, match="not linked back"):
        builtin_styles.extract(styles.encode("utf-8"), _answer(styles, WORD_ADDS[-3][0], "Heading2"), "Heading 2")


def test_extract_refuses_latent_and_default_changes():
    styles = cv_styles(latent=True)
    applied = styles.replace('w:qFormat="1"/>', 'w:qFormat="1" w:locked="1"/>')
    with pytest.raises(RefusedUnsafe, match="w:latentStyles"):
        builtin_styles.extract(styles.encode("utf-8"),
                               _answer(styles, WORD_ADDS[-35][0], "Caption", applied_base=applied), "Caption")
    applied = styles.replace('<w:sz w:val="22"/>', '<w:sz w:val="24"/>')
    with pytest.raises(RefusedUnsafe, match="w:docDefaults"):
        builtin_styles.extract(styles.encode("utf-8"),
                               _answer(styles, WORD_ADDS[-35][0], "Caption", applied_base=applied), "Caption")


def test_extract_maps_references_to_the_document_ids():
    """A localized default saved with id Standard: Word's scratch saves rename it Normal (live), and the added
    Caption's basedOn/next are written back as Standard."""

    styles = CV_STYLES.format(latent="").replace('w:styleId="Normal"', 'w:styleId="Standard"')
    addition = builtin_styles.extract(styles.encode("utf-8"),
                                      _answer(styles, WORD_ADDS[-35][0], "Caption",
                                              baseline=CV_STYLES.format(latent="")), "Caption")
    assert '<w:basedOn w:val="Standard"/><w:next w:val="Standard"/>' in addition.block


def test_lookup_by_name_or_style_id():
    assert builtin_styles.lookup("caption") == "Caption"
    assert builtin_styles.lookup("Heading2") == "Heading 2"
    assert builtin_styles.lookup(" heading  2 ") == "Heading 2"
    assert builtin_styles.lookup("ListParagraph") == "List Paragraph"
    assert builtin_styles.lookup("Heading 5") is None
    assert builtin_styles.lookup("TOC Heading") is None


# -- the scratch package, writing and checking ---------------------------------------------------------------

REVISED_BODY = (para("", run("Intro paragraph."))
                + para("", run("Prior work "), f'<w:ins w:id="5" w:author="{REVIEWER}" w:date="2026-10-10T17:37:00Z">'
                       + run("clearly ") + "</w:ins>", run("showed this."))
                + para("", run("Closing paragraph.")))


def test_scratch_package_holds_the_styles_and_no_content(tmp_path):
    source = make(tmp_path, "revised", REVISED_BODY, cv_styles())
    scratch = tmp_path / "scratch.docx"
    builtin_styles.write_scratch(source, scratch)
    with zipfile.ZipFile(scratch) as archive:
        names = set(archive.namelist())
        assert archive.read("word/styles.xml") == styles_of(source)
        body = archive.read("word/document.xml").decode("utf-8")
    assert "Prior work" not in body and "<w:ins" not in body and body.count("<w:p>") == 1
    assert names <= {"[Content_Types].xml", "_rels/.rels", "word/_rels/document.xml.rels", "word/document.xml",
                     "word/styles.xml", "word/theme/theme1.xml", "word/settings.xml", "word/numbering.xml",
                     "word/_rels/numbering.xml.rels"}


def test_staged_copy_gains_exactly_the_definitions(tmp_path, fake_word):
    source = make(tmp_path, "plain", para("", run("Anchor.")), cv_styles())
    addition = builtin_styles.plan(source, "Heading 2")
    assert addition.names == ["Heading 2", "Heading 2 Char"] and fake_word == [-3]
    staged = tmp_path / "staged.docx"
    builtin_styles.write_staged(source, staged, addition)
    assert styles_of(staged) == plus(styles_of(source), addition.block)
    builtin_styles.verify_added(source, staged, addition)
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(staged) as b:
        assert [i.filename for i in a.infolist()] == [i.filename for i in b.infolist()]
        assert all(a.read(name) == b.read(name) for name in a.namelist() if name != "word/styles.xml")


def test_verify_refuses_any_other_styles_change(tmp_path, fake_word):
    source = make(tmp_path, "plain", para("", run("Anchor.")), cv_styles())
    addition = builtin_styles.plan(source, "Caption")
    staged = tmp_path / "staged.docx"
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(staged, "w") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/styles.xml":
                data = plus(data, addition.block).replace(b"<w:caps/>", b"", 1)
            out.writestr(info, data)
    with pytest.raises(VerificationFailed, match="beyond the added built-in"):
        builtin_styles.verify_added(source, staged, addition)


# -- the synthetic sibling: a CV-style document with no Caption style ------------------------------------------

CV_BODY = (para('<w:pStyle w:val="SectionTitle"/>', run("Education"))
           + para("", run("Sample University, B.S. in Engineering, expected 2027."))
           + para('<w:pStyle w:val="SectionTitle"/>', run("Coursework"))
           + para("", run("Courses by term are listed below.")))
CV_HEADERS = ("Term", "Courses", "Credits", "Lab hours")
CV_ROWS = (("Fall", "4", "16", "6"), ("Spring", "5", "18", "4"))
CV_CAPTION = "Table 2. Course load by term"
TABLE_OP = {"type": "insert_table", "op_id": "t", "after_paragraph_index": 4, "headers": list(CV_HEADERS),
            "rows": [list(r) for r in CV_ROWS], "descriptive_columns": [1], "caption": CV_CAPTION,
            "tracked": True, "author": REVIEWER}


def cv_document(tmp_path: Path, latent: bool = False, **variant) -> Path:
    name = "cv-" + "-".join([str(latent)] + sorted(variant))
    path = make(tmp_path, name, CV_BODY, cv_styles(latent=latent, **variant))
    assert b'w:styleId="Caption"' not in styles_of(path) and revisions.revision_count(path) == 0
    return path


def insert_cv_table(source: Path, out: Path, **kwargs) -> dict:
    kwargs = {"tracked": True, "author": REVIEWER, **kwargs}
    return editing.insert_table_to_version(source, out, after_paragraph_index=kwargs.pop("after", 4),
                                           headers=CV_HEADERS, rows=CV_ROWS, descriptive_columns=(1,),
                                           caption=CV_CAPTION, **kwargs)


def paragraph_with(path: Path, text: str) -> ET.Element:
    found = [p for p in body_of(path).iter(W + "p") if "".join(t.text or "" for t in p.iter(W + "t")) == text]
    assert len(found) == 1, text
    return found[0]


@pytest.mark.parametrize("latent", [True, False], ids=["latent-caption", "no-latent-styles"])
def test_tracked_table_adds_word_caption_style(tmp_path, fake_word, latent):
    source = cv_document(tmp_path, latent)
    out = tmp_path / "out.docx"
    result = insert_cv_table(source, out)

    assert result["styles_added"] == ["Caption"] and result["caption_style"] == "Caption" and fake_word == [-35]
    assert result["source"] == str(source.resolve()) and result["source_sha256"] == file_sha256(source)
    # styles.xml gains exactly what Word added for this document, nothing else
    assert styles_of(out) == plus(styles_of(source), WORD_ADDS[-35][0])
    caption = paragraph_with(out, CV_CAPTION)
    assert caption.find(f"{W}pPr/{W}pStyle").get(W + "val") == "Caption"
    body = body_of(out)
    assert result["revisions_added"] and {item.get(W + "author") for item in body.iter(W + "ins")} == {REVIEWER}
    # Reject of the tracked insertion gives the input body back; the unused Caption definition stays, as in Word
    rejected = tmp_path / "rejected.docx"
    revisions.resolve_revisions(out, rejected, mode="reject", author=REVIEWER)
    assert canonical(body_of(rejected)) == canonical(body_of(source))
    assert styles_of(rejected) == styles_of(out)


def test_untracked_table_on_revised_document_adds_caption(tmp_path, fake_word):
    """Coverage (review r1 item 4): tracking off on a document with revisions takes the direct path too; the
    table is plain, the existing revision stays byte-identical and Caption is added the same way."""

    source = make(tmp_path, "revised-untracked", REVISED_BODY, cv_styles())
    out = tmp_path / "out.docx"
    result = insert_cv_table(source, out, after=3, tracked=False)
    assert result["styles_added"] == ["Caption"] and fake_word == [-35]
    assert not result["revisions_added"] and revisions.revision_count(out) == 1
    assert styles_of(out) == plus(styles_of(source), WORD_ADDS[-35][0])
    ins_before = re.search(rb"<w:ins\b.*?</w:ins>", document_xml(source).encode("utf-8"), re.S).group(0)
    assert document_xml(out).encode("utf-8").count(ins_before) == 1
    assert body_of(out).find(W + "tbl") is not None
    assert paragraph_with(out, CV_CAPTION).find(f"{W}pPr/{W}pStyle").get(W + "val") == "Caption"


def test_defined_caption_adds_nothing(tmp_path, fake_word):
    styles = CV_STYLES.format(latent="").replace(
        "</w:styles>", '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/></w:style></w:styles>')
    source = make(tmp_path, "cv-defined", CV_BODY, styles)
    out = tmp_path / "out.docx"
    result = insert_cv_table(source, out)
    assert result["styles_added"] == [] and styles_of(out) == styles_of(source) and fake_word == []


def test_failed_style_check_publishes_nothing(tmp_path, fake_word, monkeypatch):
    """Coverage (review r1 item 4): when the output's styles.xml is more than the source plus Word's addition,
    the insert refuses, the output is removed and the source is untouched."""

    source = cv_document(tmp_path)
    before = file_sha256(source)
    write = builtin_styles.write_staged

    def tampered(src, staged, addition):
        write(src, staged, addition)
        damaged = Path(str(staged) + ".tmp")
        with zipfile.ZipFile(staged) as reader, zipfile.ZipFile(damaged, "w") as writer:
            for info in reader.infolist():
                data = reader.read(info.filename)
                writer.writestr(info, data.replace(b"<w:caps/>", b"") if info.filename == "word/styles.xml" else data)
        os.replace(damaged, staged)

    monkeypatch.setattr(builtin_styles, "write_staged", tampered)
    out = tmp_path / "out.docx"
    with pytest.raises(VerificationFailed, match="beyond the added built-in"):
        insert_cv_table(source, out)
    assert not out.exists() and file_sha256(source) == before


@pytest.fixture
def in_process_batch(monkeypatch):
    from research_mcp import edit_batch

    monkeypatch.setattr(edit_batch, "_apply_operation", edit_batch._apply_operation_local)


@windows_only
def test_batch_refusal_leaves_the_source(tmp_path, fake_word, in_process_batch, monkeypatch):
    source = cv_document(tmp_path)
    before = file_sha256(source)

    def refuse(*args):
        raise VerificationFailed("styles.xml changed beyond the added built-in style definitions")

    monkeypatch.setattr(builtin_styles, "verify_added", refuse)
    result = run_batch(source, TABLE_OP)
    assert result.status != "complete" and result.counts["applied"] == 0
    assert "beyond the added built-in" in json.dumps(result.model_dump(mode="json"))
    assert file_sha256(source) == before


@windows_only
def test_batch_in_place_then_undo_restores_the_source(tmp_path, fake_word, in_process_batch):
    from research_mcp.publication import undo_document

    source = cv_document(tmp_path, latent=True)
    before, styles_before = file_sha256(source), styles_of(source)
    result = run_batch(source, TABLE_OP)
    assert result.status == "complete" and result.undo_token, result
    assert result.operations[0].details["styles_added"] == ["Caption"]
    assert styles_of(source) == plus(styles_before, WORD_ADDS[-35][0])
    undone = undo_document(str(source))
    assert undone["restored"] is True and file_sha256(source) == before


@windows_only
def test_batch_second_op_uses_the_style_the_first_added(tmp_path, fake_word, in_process_batch):
    """Coverage (review r1 item 4): op 2 sees the Caption op 1 added, adds nothing and asks Word nothing."""

    source = cv_document(tmp_path)
    result = run_batch(source, TABLE_OP, {"type": "insert_paragraph", "op_id": "p", "after_paragraph_index": 1,
                                          "text": "Figure 1. Weekly schedule", "style": "Caption",
                                          "tracked": True, "author": REVIEWER})
    assert result.status == "complete", result
    assert [op.details["styles_added"] for op in result.operations] == [["Caption"], []]
    assert fake_word == [-35]
    assert styles_of(source).count(b'w:styleId="Caption"') == 1
    assert paragraph_with(source, "Figure 1. Weekly schedule").find(f"{W}pPr/{W}pStyle").get(W + "val") == "Caption"


def test_builtin_outside_the_set_still_refuses(tmp_path, fake_word):
    styles = CV_STYLES.format(latent='<w:latentStyles><w:lsdException w:name="TOC Heading"/></w:latentStyles>')
    source = make(tmp_path, "cv-toc", CV_BODY, styles)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="has not used yet"):
        insert_cv_table(source, out, caption_style="TOC Heading")
    with pytest.raises(RefusedUnsafe, match="not a paragraph style"):
        insert_cv_table(source, out, caption_style="Table Caption")
    assert not out.exists() and fake_word == []


@pytest.mark.parametrize("tracked", [True, False])
def test_insert_paragraph_heading_2_on_revised_document(tmp_path, fake_word, tracked):
    source = make(tmp_path, f"revised-{tracked}", REVISED_BODY, cv_styles())
    assert revisions.revision_count(source) == 1
    out = tmp_path / "out.docx"
    result = editing.insert_paragraph_to_version(source, out, after_paragraph_index=2, text="Methods",
                                                 style="Heading 2", tracked=tracked, author=REVIEWER)
    assert result["styles_added"] == ["Heading 2", "Heading 2 Char"]
    assert styles_of(out) == plus(styles_of(source), "".join(WORD_ADDS[-3]))
    assert paragraph_with(out, "Methods").find(f"{W}pPr/{W}pStyle").get(W + "val") == "Heading2"
    assert result["source"] == str(source.resolve())


def test_insert_paragraph_passes_on_the_clash_refusal(tmp_path, monkeypatch):
    source = make(tmp_path, "clash", REVISED_BODY, cv_styles(h2char=("My Character Style", "")))
    install_fake_word(monkeypatch)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="existing style.*'My Character Style'"):
        editing.insert_paragraph_to_version(source, out, after_paragraph_index=2, text="Methods",
                                            style="Heading 2", tracked=True, author=REVIEWER)
    assert not out.exists()


def test_untracked_revision_free_insert_paragraph_keeps_word_path(tmp_path, fake_word, monkeypatch):
    source = make(tmp_path, "clean", CV_BODY, cv_styles())

    class WordPath(Exception):
        pass

    def word(*args, **kwargs):
        raise WordPath

    monkeypatch.setattr(editing, "IsolatedWordDocument", word)
    with pytest.raises(WordPath):  # resolves the built-in instead of refusing, then hands it to Word
        editing.insert_paragraph_to_version(source, tmp_path / "out.docx", after_paragraph_index=2, text="X",
                                            style="Heading 2")
    assert fake_word == []


# -- live: against Word itself ---------------------------------------------------------------------------------

def word_applies(source: Path, name: str) -> tuple[dict[str, str], Path]:
    """Ground truth: Word applies the built-in `name` (by id) to paragraph 1 of a copy of `source` and saves.
    Returns the saved styles by id and the copy."""

    from research_mcp.word import IsolatedWordDocument

    copy = source.with_name(source.stem + f"-word-{WD[name]}.docx")
    shutil.copyfile(source, copy)
    with IsolatedWordDocument(copy) as session:
        session.document.Paragraphs(1).Style = session.document.Styles(WD[name])
        session.save()
    return style_elements(styles_of(copy)), copy


def word_reads(path: Path, text: str) -> dict:
    from research_mcp.word import IsolatedWordDocument

    with IsolatedWordDocument(path, read_only=True) as session:
        document = session.document
        for index in range(1, int(document.Paragraphs.Count) + 1):
            paragraph = document.Paragraphs(index)
            if str(paragraph.Range.Text).rstrip("\r\x07") == text:
                return {"style": str(paragraph.Style.NameLocal), "italic": int(paragraph.Range.Font.Italic),
                        "revisions": int(document.Revisions.Count)}
    return {}


@live
def test_live_word_answers_match_the_fake(tmp_path):
    source = cv_document(tmp_path)
    for name, wd_id in WD.items():
        addition = builtin_styles.plan(source, name)
        assert [xml for _, _, xml in addition.definitions] == WORD_ADDS[wd_id], name
        assert addition.word_version.startswith("16.") and addition.word_seconds < 30


@live
@pytest.mark.parametrize("variant", ["italic_normal", "locked_latent"])
def test_live_caption_is_the_one_word_adds(tmp_path, variant):
    """Review r1 items 1-2: italic Normal (Word's Caption writes <w:i w:val="0"/> and renders upright) and a
    locked, semiHidden latent caption (Word writes <w:semiHidden/><w:locked/>): the inserted caption's style is
    Word's own, byte for byte, and Word renders the caption as it renders its own."""

    source = cv_document(tmp_path, **{variant: True})
    out = tmp_path / "out.docx"
    assert insert_cv_table(source, out)["styles_added"] == ["Caption"]
    word_styles, word_copy = word_applies(source, "Caption")
    assert style_elements(styles_of(out))["Caption"] == word_styles["Caption"]
    marker = '<w:i w:val="0"/>' if variant == "italic_normal" else "<w:locked/>"
    assert marker in word_styles["Caption"]
    ours, theirs = word_reads(out, CV_CAPTION), word_reads(word_copy, "Education")
    assert ours["style"] == theirs["style"] == "Caption"
    assert ours["italic"] == theirs["italic"] == (0 if variant == "italic_normal" else -1)


@live
def test_live_heading_2_char_id_taken_by_another_style_refuses(tmp_path):
    """Review r1 item 3: "My Character Style" holds the id Heading2Char. Word itself re-ids it on save; the MCP
    refuses, naming it, and writes nothing."""

    source = make(tmp_path, "h2-unrelated", REVISED_BODY, cv_styles(h2char=("My Character Style", "")))
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match="'My Character Style'"):
        editing.insert_paragraph_to_version(source, out, after_paragraph_index=2, text="Methods",
                                            style="Heading 2", tracked=True, author=REVIEWER)
    assert not out.exists()


@live
@pytest.mark.parametrize("link", ["", '<w:link w:val="Heading2"/>'], ids=["unlinked", "dangling-link"])
def test_live_heading_2_char_name_taken_adds_words_char1(tmp_path, link):
    """Review r1 item 3: a "Heading 2 Char" already exists (unlinked, or linking to a Heading 2 the document
    lacks). Word keeps it and adds "Heading 2 Char1"; so does the MCP, and the existing one is unchanged."""

    source = make(tmp_path, f"h2-name-{bool(link)}", REVISED_BODY, cv_styles(h2char=("Heading 2 Char", link)))
    out = tmp_path / "out.docx"
    result = editing.insert_paragraph_to_version(source, out, after_paragraph_index=2, text="Methods",
                                                 style="Heading 2", tracked=True, author=REVIEWER)
    assert result["styles_added"] == ["Heading 2", "Heading 2 Char1"]
    ours = style_elements(styles_of(out))
    word_styles, _ = word_applies(source, "Heading 2")
    assert ours["Heading2"] == word_styles["Heading2"] and ours["Heading2Char1"] == word_styles["Heading2Char1"]
    assert ours["Heading2Char"] == style_elements(styles_of(source))["Heading2Char"]
    assert word_reads(out, "Methods")["style"] == "Heading 2"


@live
def test_live_tracked_caption_style_in_word(tmp_path):
    from live_word import word_resolve

    source = cv_document(tmp_path)
    out = tmp_path / "tracked.docx"
    assert insert_cv_table(source, out)["styles_added"] == ["Caption"]
    word, word_src = word_resolve([out, source])
    assert word["opens_clean"] and not word.get("error"), word
    assert word["AcceptAll_left"] == 0 and word["RejectAll_left"] == 0, word
    assert canonical(body_of(word["reject"])) == canonical(body_of(word_src["reject"]))
    assert word_reads(out, CV_CAPTION)["style"] == "Caption"


@live
@pytest.mark.parametrize("caption_style", ["Normal", "Caption", "caption"])
def test_live_word_path_caption_style_round_trips(tmp_path, caption_style):
    source = cv_document(tmp_path, latent=True)
    out = tmp_path / "word.docx"
    result = editing.insert_table_to_version(source, out, after_paragraph_index=4, headers=CV_HEADERS,
                                             rows=CV_ROWS, descriptive_columns=(1,), caption=CV_CAPTION,
                                             caption_style=caption_style, tracked=False)
    style = "Normal" if caption_style == "Normal" else "Caption"
    assert result["caption_style"].casefold() == style.casefold()
    # Word also adds styles of its own on this path ("Normal Table", "No List"); Caption only when it is used
    assert ("Caption" in result["styles_added"]) == (style == "Caption")
    assert word_reads(out, CV_CAPTION)["style"] == style
    saved = [p for p in editing.read_paragraphs(out) if p.text == CV_CAPTION]
    assert len(saved) == 1 and saved[0].style == (None if style == "Normal" else "Caption")
