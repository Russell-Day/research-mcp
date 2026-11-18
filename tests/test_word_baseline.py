"""word_baseline: Word's save of the untouched document as the reference for format_tables and
apply_document_style (plan 973 S1). Word appends an empty paragraph after a body-final table; that paragraph alone
is tolerated, and only when it holds nothing but its properties. Fixtures are synthetic."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from research_mcp import word_baseline
from research_mcp.sections import baseline_preservation_errors
from research_mcp.word_baseline import appended_final_paragraph, choose_reference, preservation_errors, trimmed_copy

from test_tables_grid import build_cv

DOC = "word/document.xml"


def _read(path: Path, name: str = DOC) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read(name)


def _with_document(source: Path, target: Path, transform) -> Path:
    """A copy of ``source`` whose document.xml is ``transform(text)``; every other part byte-identical."""

    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as new:
        for info in old.infolist():
            data = old.read(info.filename)
            if info.filename == DOC:
                data = transform(data.decode("utf-8")).encode("utf-8")
            new.writestr(info, data)
    return target


def _append_before_sect(paragraph: str):
    def transform(xml: str) -> str:
        at = xml.rindex("<w:sectPr")
        return xml[:at] + paragraph + xml[at:]
    return transform


APPENDED = '<w:p w:rsidR="00A1B2C3"><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>'


@pytest.fixture
def cv(tmp_path) -> Path:
    return build_cv(tmp_path / "cv.docx")


@pytest.fixture
def word_form(cv, tmp_path) -> Path:
    """The CV as Word saves it: one empty paragraph after the body-final table."""

    return _with_document(cv, tmp_path / "baseline.docx", _append_before_sect(APPENDED))


def test_the_appended_paragraph_is_recognised_and_set_aside(cv, word_form, tmp_path):
    assert appended_final_paragraph(cv, word_form)
    trimmed = trimmed_copy(word_form, tmp_path / "trimmed.docx")
    assert _read(trimmed) == _read(cv)
    with zipfile.ZipFile(cv) as archive:
        for name in archive.namelist():
            assert _read(trimmed, name) == _read(cv, name)


def test_the_baseline_with_the_appended_paragraph_is_adopted(cv, word_form):
    # the plain preservation check refuses it: the reason format_tables and apply_document_style refused
    assert baseline_preservation_errors(cv, word_form)
    assert preservation_errors(cv, word_form) == []
    reference = choose_reference(cv, word_form)
    assert (reference.path, reference.compared_against, reference.problems) == (word_form, "word_baseline", ())


def test_an_identical_baseline_is_adopted(cv, tmp_path):
    copy = _with_document(cv, tmp_path / "copy.docx", lambda xml: xml)
    assert not appended_final_paragraph(cv, copy)
    assert choose_reference(cv, copy).compared_against == "word_baseline"


@pytest.mark.parametrize("paragraph", [
    '<w:p><w:r><w:t>Added</w:t></w:r></w:p>',
    '<w:p><w:pPr/><w:bookmarkStart w:id="9" w:name="x"/><w:bookmarkEnd w:id="9"/></w:p>',
])
def test_a_final_paragraph_holding_anything_is_not_tolerated(cv, tmp_path, paragraph):
    baseline = _with_document(cv, tmp_path / "busy.docx", _append_before_sect(paragraph))
    assert not appended_final_paragraph(cv, baseline)
    reference = choose_reference(cv, baseline)
    assert reference.compared_against == "source" and reference.path == cv and reference.problems


def test_a_paragraph_after_a_body_that_does_not_end_with_a_table_is_not_tolerated(tmp_path):
    from docx import Document

    document = Document()
    document.add_paragraph("Only text.")
    source = tmp_path / "text.docx"
    document.save(source)
    baseline = _with_document(source, tmp_path / "plus.docx", _append_before_sect(APPENDED))
    assert not appended_final_paragraph(source, baseline)
    assert choose_reference(source, baseline).compared_against == "source"


def test_a_changed_baseline_falls_back_to_the_source(cv, tmp_path):
    changed = _with_document(cv, tmp_path / "changed.docx",
                             lambda xml: _append_before_sect(APPENDED)(xml.replace("Lab A", "Lab Z")))
    reference = choose_reference(cv, changed)
    assert reference.compared_against == "source" and reference.problems


def test_a_missing_or_unreadable_baseline_falls_back_to_the_source(cv, tmp_path):
    assert choose_reference(cv, None).problems == ("no Word baseline was saved",)
    assert choose_reference(cv, tmp_path / "absent.docx").compared_against == "source"
    broken = tmp_path / "broken.docx"
    broken.write_bytes(b"not a zip")
    reference = choose_reference(cv, broken)
    assert reference.compared_against == "source" and "could not be read" in reference.problems[0]


@pytest.mark.parametrize("paragraph", ["<w:p/>", '<w:p w14:paraId="1A2B3C4D"><w:pPr/></w:p>'])
def test_the_trim_cuts_only_the_last_paragraph(paragraph):
    head = '<w:body><w:p><w:pPr/></w:p><w:tbl/>'
    sect = '<w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:body>'
    xml = head + paragraph + sect
    assert word_baseline._without_final_paragraph(xml) == head + sect
    assert word_baseline._without_final_paragraph(head + paragraph + "</w:body>") == head + "</w:body>"


def test_an_unfaithful_baseline_falls_back_to_the_source_in_words_body_final_form(cv, tmp_path):
    """Word's save also drops something it does not read (here, changed text stands in for it): the baseline is not
    adopted, and the source is compared with Word's appended paragraph after its final table, so that paragraph
    alone does not refuse the pass."""

    paragraph = '<w:p w14:paraId="1A2B3C4D" w:rsidR="00A1B2C3"><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>'
    folder = tmp_path / "baseline"
    folder.mkdir()
    baseline = _with_document(cv, folder / "baseline.docx",
                              lambda xml: _append_before_sect(paragraph)(xml.replace("Lab A", "Lab Z")))
    reference = choose_reference(cv, baseline)
    assert reference.compared_against == "source" and reference.problems
    assert reference.path == folder / "source_word_form.docx"
    assert appended_final_paragraph(cv, reference.path)
    source_xml, form_xml = _read(cv).decode(), _read(reference.path).decode()
    at = source_xml.rindex("<w:sectPr")
    declares_w14 = "xmlns:w14=" in source_xml[:source_xml.index(">", source_xml.index("<w:document"))]
    kept = paragraph if declares_w14 else paragraph.replace(' w14:paraId="1A2B3C4D"', "")
    assert form_xml == source_xml[:at] + kept + source_xml[at:]
    assert trimmed_copy(reference.path, tmp_path / "back.docx") and _read(tmp_path / "back.docx") == _read(cv)
    with zipfile.ZipFile(cv) as archive:
        for name in archive.namelist():
            if name != DOC:
                assert _read(reference.path, name) == _read(cv, name)


def test_an_undeclared_prefixed_element_keeps_the_plain_source(cv, tmp_path):
    folder = tmp_path / "baseline"
    folder.mkdir()
    paragraph = '<w:p><w:pPr><zz:odd xmlns:zz="urn:x"/></w:pPr></w:p>'
    baseline = _with_document(cv, folder / "baseline.docx",
                              lambda xml: _append_before_sect(paragraph)(xml.replace("Lab A", "Lab Z")))
    reference = choose_reference(cv, baseline)
    assert reference.path == cv and reference.compared_against == "source"
    assert "could not be added to the source" in reference.problems[-1]


# -- section properties (plan 973 S1 r1) ----------------------------------------------------------------------------


def _page_width(width: str):
    def transform(xml: str) -> str:
        assert xml.count('<w:pgSz w:w="12240"') == 1
        return xml.replace('<w:pgSz w:w="12240"', f'<w:pgSz w:w="{width}"')
    return transform


def test_a_baseline_with_another_page_size_is_not_adopted(cv, tmp_path):
    """Word's save changed the body-final sectPr (page width 12240 -> 1000) and appended its paragraph: the
    baseline is not adopted; the source is compared (in Word's body-final form)."""

    folder = tmp_path / "baseline"
    folder.mkdir()
    baseline = _with_document(cv, folder / "baseline.docx",
                              lambda xml: _append_before_sect(APPENDED)(_page_width("1000")(xml)))
    assert preservation_errors(cv, baseline) == ["section 1 properties, headers or footers changed"]
    reference = choose_reference(cv, baseline)
    assert reference.compared_against == "source" and reference.path == folder / "source_word_form.docx"
    assert word_baseline._sections(reference.path) == word_baseline._sections(cv)


def test_a_page_size_change_without_the_appended_paragraph_is_not_adopted(cv, tmp_path):
    baseline = _with_document(cv, tmp_path / "wide.docx", _page_width("1000"))
    assert baseline_preservation_errors(cv, baseline) == []  # the body check alone does not read it
    reference = choose_reference(cv, baseline)
    assert (reference.path, reference.compared_against) == (cv, "source")
    assert reference.problems == ("section 1 properties, headers or footers changed",)


def _with_header(path: Path, text: str) -> Path:
    from docx import Document

    document = Document(path)
    document.sections[0].header.paragraphs[0].text = text
    document.save(path)
    return path


def _with_part(source: Path, target: Path, transform) -> Path:
    """A copy of ``source`` with every part passed through ``transform(name, data) -> (name, data)``."""

    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as new:
        for info in old.infolist():
            name, data = transform(info.filename, old.read(info.filename))
            new.writestr(name, data)
    return target


def _header_part(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return next(name for name in archive.namelist() if name.startswith("word/header"))


def test_a_changed_header_is_not_adopted(cv, tmp_path):
    source = _with_header(cv, "Header text")
    header = _header_part(source)
    baseline = _with_part(source, tmp_path / "header.docx", lambda name, data: (
        name, data.replace(b"Header text", b"Other text") if name == header else data))
    assert preservation_errors(source, baseline) == ["section 1 properties, headers or footers changed"]
    assert choose_reference(source, baseline).path == source


def test_a_renamed_header_part_and_relationship_are_adopted(cv, tmp_path):
    """Word may renumber relationships and rename header parts on save; the header's content is what counts."""

    source = _with_header(cv, "Header text")
    header = _header_part(source)
    old_target, new_target = header.removeprefix("word/").encode(), b"header9.xml"
    with zipfile.ZipFile(source) as archive:
        rels = archive.read("word/_rels/document.xml.rels").decode()
    relationship = next(part for part in rels.split("<Relationship ") if old_target.decode() in part)
    old_id = relationship.split('Id="')[1].split('"')[0]

    def rename(name: str, data: bytes):
        if name == header:
            return "word/header9.xml", data
        if name == "word/_rels/document.xml.rels":
            return name, data.replace(old_target, new_target).replace(f'Id="{old_id}"'.encode(), b'Id="rId99"')
        if name == "[Content_Types].xml":
            return name, data.replace(b"/" + b"word/" + old_target, b"/word/header9.xml")
        if name == DOC:
            return name, data.replace(f'r:id="{old_id}"'.encode(), b'r:id="rId99"')
        return name, data

    baseline = _with_part(source, tmp_path / "renamed.docx", rename)
    assert _header_part(baseline) == "word/header9.xml"
    assert b'r:id="rId99"' in _read(baseline)
    assert preservation_errors(source, baseline) == []
    assert choose_reference(source, baseline).compared_against == "word_baseline"


# -- what the appended paragraph may hold (plan 973 S1 r1) ----------------------------------------------------------

SECTION_BREAK = ('<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
                 '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr></w:p>')
STRUCTURAL = [
    SECTION_BREAK,
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>',
    '<w:p><w:pPr><w:framePr w:w="2000" w:h="400"/></w:pPr></w:p>',
    '<w:p><w:pPr><w:pPrChange w:id="5" w:author="A"><w:pPr/></w:pPrChange></w:pPr></w:p>',
    '<w:p><w:pPr><w:rPr><w:ins w:id="6" w:author="A"/></w:rPr></w:pPr></w:p>',
    '<w:p><w:pPr><w:rPr><w:b/><w:rPrChange w:id="7" w:author="A"><w:rPr/></w:rPrChange></w:rPr></w:pPr></w:p>',
]


@pytest.mark.parametrize("paragraph", STRUCTURAL, ids=["sectPr", "numPr", "framePr", "pPrChange", "ins", "rPrChange"])
def test_an_appended_paragraph_with_structure_is_not_words_empty_paragraph(cv, tmp_path, paragraph):
    """Faithful otherwise: the paragraph is not set aside, so the baseline is not adopted."""

    baseline = _with_document(cv, tmp_path / "structure.docx", _append_before_sect(paragraph))
    assert not appended_final_paragraph(cv, baseline)
    assert preservation_errors(cv, baseline)
    reference = choose_reference(cv, baseline)
    assert (reference.path, reference.compared_against) == (cv, "source") and reference.problems


@pytest.mark.parametrize("paragraph", STRUCTURAL, ids=["sectPr", "numPr", "framePr", "pPrChange", "ins", "rPrChange"])
def test_an_appended_paragraph_with_structure_is_not_copied_into_the_source(cv, tmp_path, paragraph):
    """Unfaithful baseline (changed text): the source fallback stays the plain source, never the source with that
    paragraph added."""

    folder = tmp_path / "baseline"
    folder.mkdir()
    baseline = _with_document(cv, folder / "baseline.docx",
                              lambda xml: _append_before_sect(paragraph)(xml.replace("Lab A", "Lab Z")))
    reference = choose_reference(cv, baseline)
    assert (reference.path, reference.compared_against) == (cv, "source")
    assert not (folder / "source_word_form.docx").exists()
    with pytest.raises(ValueError, match="does not end with Word's appended paragraph"):
        word_baseline.source_in_word_form(cv, baseline, folder / "form.docx")


@pytest.mark.parametrize("unfaithful", [False, True], ids=["faithful", "unfaithful"])
def test_two_appended_paragraphs_are_not_tolerated(cv, tmp_path, unfaithful):
    folder = tmp_path / "baseline"
    folder.mkdir()
    baseline = _with_document(cv, folder / "baseline.docx", lambda xml: _append_before_sect(APPENDED + APPENDED)(
        xml.replace("Lab A", "Lab Z") if unfaithful else xml))
    assert not appended_final_paragraph(cv, baseline)
    assert any("paragraph/row structure changed" in problem for problem in preservation_errors(cv, baseline))
    reference = choose_reference(cv, baseline)
    assert (reference.path, reference.compared_against) == (cv, "source")
    assert not (folder / "source_word_form.docx").exists()


def _png(red: int) -> bytes:
    """A 1x1 RGB PNG of the given red value."""

    import struct
    import zlib

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(bytes([0, red, 0, 0])))
            + chunk(b"IEND", b""))


def _with_header_picture(path: Path, tmp_path: Path) -> Path:
    from docx import Document

    image = tmp_path / "logo.png"
    image.write_bytes(_png(10))
    document = Document(path)
    document.sections[0].header.paragraphs[0].add_run().add_picture(str(image))
    document.save(path)
    return path


def _media(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return next(name for name in archive.namelist() if name.startswith("word/media/"))


def test_a_changed_header_image_is_not_adopted(cv, tmp_path):
    """The header's XML is unchanged but the picture it shows differs: the baseline is not adopted."""

    source = _with_header_picture(cv, tmp_path)
    media = _media(source)
    baseline = _with_part(source, tmp_path / "image.docx", lambda name, data: (
        name, _png(200) if name == media else data))
    assert _read(baseline, _header_part(baseline)) == _read(source, _header_part(source))
    assert preservation_errors(source, baseline) == ["section 1 properties, headers or footers changed"]
    assert choose_reference(source, baseline).path == source


def test_a_renamed_header_image_and_relationship_are_adopted(cv, tmp_path):
    """The same picture under another media name and relationship id is the same header."""

    source = _with_header_picture(cv, tmp_path)
    header, media = _header_part(source), _media(source)
    rels = header.replace("word/", "word/_rels/", 1) + ".rels"
    with zipfile.ZipFile(source) as archive:
        relationships = archive.read(rels).decode()
    relationship = next(part for part in relationships.split("<Relationship ") if media.removeprefix("word/") in part)
    old_id = relationship.split('Id="')[1].split('"')[0]
    old_target, new_target = media.removeprefix("word/").encode(), b"media/image77.png"

    def rename(name: str, data: bytes):
        if name == media:
            return "word/media/image77.png", data
        if name == rels:
            return name, data.replace(old_target, new_target).replace(f'Id="{old_id}"'.encode(), b'Id="rId77"')
        if name == header:
            return name, data.replace(f'r:embed="{old_id}"'.encode(), b'r:embed="rId77"')
        return name, data

    baseline = _with_part(source, tmp_path / "renamed_image.docx", rename)
    assert b'r:embed="rId77"' in _read(baseline, header)
    assert preservation_errors(source, baseline) == []
    assert choose_reference(source, baseline).compared_against == "word_baseline"


def _as_word_saves_it(xml: bytes) -> bytes:
    """A header as Word rewrites it on save: new rsids and docPr id, wp14 drawing attributes, w:noProof on the
    picture's run, a proofing mark."""

    text = xml.decode("utf-8")
    text = text.replace("<wp:inline", '<wp:inline wp14:anchorId="1A2B3C4D" wp14:editId="5E6F7A8B"', 1)
    if "xmlns:wp14=" not in text:
        text = text.replace(" xmlns:w=", ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/'
                                         'wordprocessingDrawing" xmlns:w=', 1)
    text = re.sub(r'(<wp:docPr\b[^>]*\bid=")\d+"', r'\g<1>9"', text, count=1)
    text = text.replace("<w:p>", '<w:p w:rsidR="00C0FFEE" w:rsidRDefault="00C0FFEE">', 1)
    head, sep, tail = text.partition("<w:drawing>")
    run_start = head.rindex("<w:r>")
    head = head[:run_start] + "<w:r><w:rPr><w:noProof/></w:rPr>" + head[run_start + len("<w:r>"):]
    text = head + sep + tail
    marked = '<w:p w:rsidR="00C0FFEE" w:rsidRDefault="00C0FFEE">'
    return text.replace(marked, marked + '<w:proofErr w:type="spellStart"/>', 1).encode("utf-8")


def test_words_drawing_markup_in_a_header_is_adopted(cv, tmp_path):
    """Word rewrites a header picture's markup on save (ids, wp14 attributes, w:noProof, rsids); the picture's bytes
    and extent and the header's text are unchanged, so the baseline is adopted."""

    source = _with_header_picture(cv, tmp_path)
    header = _header_part(source)
    baseline = _with_part(source, tmp_path / "word_header.docx", lambda name, data: (
        name, _as_word_saves_it(data) if name == header else data))
    assert _read(baseline, header) != _read(source, header)
    assert b"wp14:anchorId" in _read(baseline, header) and b"<w:noProof/>" in _read(baseline, header)
    assert preservation_errors(source, baseline) == []
    assert choose_reference(source, baseline).compared_against == "word_baseline"


@pytest.mark.parametrize("change", ["extent", "bold", "field"])
def test_a_changed_header_extent_formatting_or_field_is_not_adopted(cv, tmp_path, change):
    from docx import Document

    source = _with_header_picture(cv, tmp_path)
    document = Document(source)
    document.sections[0].header.add_paragraph().add_run("Page ")
    document.save(source)
    header = _header_part(source)

    def transform(xml: bytes) -> bytes:
        text = xml.decode("utf-8")
        if change == "extent":
            text = re.sub(r'(<wp:extent cx=")(\d+)"', lambda m: f'{m.group(1)}{int(m.group(2)) * 2}"', text, count=1)
        elif change == "bold":
            text = text.replace("<w:r><w:t", "<w:r><w:rPr><w:b/></w:rPr><w:t", 1)
        else:
            text = text.replace("</w:t></w:r>", '</w:t></w:r><w:fldSimple w:instr=" PAGE "/>', 1)
        return text.encode("utf-8")

    baseline = _with_part(source, tmp_path / f"{change}.docx", lambda name, data: (
        name, transform(data) if name == header else data))
    assert _read(baseline, header) != _read(source, header)
    assert preservation_errors(source, baseline) == ["section 1 properties, headers or footers changed"]


def test_a_dropped_header_reference_is_not_adopted(cv, tmp_path):
    source = _with_header(cv, "Header text")
    baseline = _with_document(source, tmp_path / "dropped.docx",
                              lambda xml: re.sub(r"<w:headerReference\b[^>]*/>", "", xml, count=1))
    assert b"headerReference" in _read(source) and b"headerReference" not in _read(baseline)
    assert preservation_errors(source, baseline) == ["section 1 properties, headers or footers changed"]


def test_two_header_pictures_with_their_targets_swapped_are_not_adopted(cv, tmp_path):
    """Two same-size pictures in one header paragraph: swapping which picture each drawing shows is a change, even
    though the header shows the same set of pictures."""

    from docx import Document

    first, second = tmp_path / "first.png", tmp_path / "second.png"
    first.write_bytes(_png(10))
    second.write_bytes(_png(200))
    source = cv
    document = Document(source)
    paragraph = document.sections[0].header.paragraphs[0]
    paragraph.add_run().add_picture(str(first))
    paragraph.add_run().add_picture(str(second))
    document.save(source)
    header = _header_part(source)
    ids = re.findall(rb'r:embed="([^"]+)"', _read(source, header))
    assert len(ids) == 2 and ids[0] != ids[1]

    def swap(data: bytes) -> bytes:
        a, b = (b'r:embed="' + value + b'"' for value in ids)
        return data.replace(a, b"\0").replace(b, a).replace(b"\0", b)

    baseline = _with_part(source, tmp_path / "swapped.docx", lambda name, data: (
        name, swap(data) if name == header else data))
    assert re.findall(rb'r:embed="([^"]+)"', _read(baseline, header)) == ids[::-1]
    assert preservation_errors(source, baseline) == ["section 1 properties, headers or footers changed"]
    assert choose_reference(source, baseline).path == source
