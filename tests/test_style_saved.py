"""Offline regression tests for saved typography and field preservation."""

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

import pytest

from research_mcp.errors import VerificationFailed
from research_mcp.styling import (
    PROFILES,
    _body_field_instructions,
    _verify_saved_typography,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
ET.register_namespace("w", W_NS)


def _styles(font="Times New Roman", size="22", line="480", line_rule="auto"):
    styles = ET.Element(W + "styles")
    defaults = ET.SubElement(styles, W + "docDefaults")
    pdef = ET.SubElement(ET.SubElement(defaults, W + "pPrDefault"), W + "pPr")
    ET.SubElement(pdef, W + "spacing", {W + "line": line, W + "lineRule": line_rule, W + "before": "0", W + "after": "0"})
    rdef = ET.SubElement(ET.SubElement(defaults, W + "rPrDefault"), W + "rPr")
    ET.SubElement(rdef, W + "rFonts", {W + "ascii": font, W + "hAnsi": font})
    ET.SubElement(rdef, W + "sz", {W + "val": size})
    normal = ET.SubElement(styles, W + "style", {W + "type": "paragraph", W + "styleId": "Normal", W + "default": "1"})
    ET.SubElement(normal, W + "name", {W + "val": "Normal"})
    return ET.tostring(styles, encoding="utf-8", xml_declaration=True)


def _run(text, font=None, size=None):
    r = ET.Element(W + "r")
    if font or size:
        rpr = ET.SubElement(r, W + "rPr")
        if font:
            ET.SubElement(rpr, W + "rFonts", {W + "ascii": font, W + "hAnsi": font})
        if size:
            ET.SubElement(rpr, W + "sz", {W + "val": size})
    ET.SubElement(r, W + "t").text = text
    return r


def _paragraph(text, *, direct_font=None, direct_size=None, title=False):
    p = ET.Element(W + "p")
    if title:
        ppr = ET.SubElement(p, W + "pPr")
        ET.SubElement(ppr, W + "spacing", {W + "line": "240", W + "lineRule": "auto", W + "before": "240", W + "after": "240"})
    if title:
        direct_font = direct_font or "Times New Roman"
        direct_size = direct_size or "28"
    p.append(_run(text, direct_font, direct_size))
    return p


def _docx(path: Path, paragraphs, *, second_bad_margin=False, styles=None):
    document = ET.Element(W + "document")
    body = ET.SubElement(document, W + "body")
    for paragraph in paragraphs:
        body.append(paragraph)
    sect = ET.SubElement(body, W + "sectPr")
    ET.SubElement(sect, W + "pgMar", {W + side: str(2000 if second_bad_margin and side == "top" else 1440) for side in ("top", "right", "bottom", "left")})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))
        archive.writestr("word/styles.xml", styles or _styles())


def _valid_doc(path):
    return _docx(path, [_paragraph("Probe Title", title=True), _paragraph("body paragraph")])


def test_inherited_times_new_roman_double_spacing_is_accepted(tmp_path):
    path = tmp_path / "inherited.docx"
    _valid_doc(path)
    _verify_saved_typography(path, PROFILES["radiology_ai"])


def test_inherited_wrong_font_is_rejected(tmp_path):
    path = tmp_path / "wrong-font.docx"
    _docx(path, [_paragraph("Probe Title", title=True), _paragraph("body")], styles=_styles(font="Arial"))
    with pytest.raises(VerificationFailed, match="run typography"):
        _verify_saved_typography(path, PROFILES["radiology_ai"])


def test_exact_line_rule_is_rejected(tmp_path):
    path = tmp_path / "exact.docx"
    _docx(path, [_paragraph("Probe Title", title=True), _paragraph("body")], styles=_styles(line_rule="exact"))
    with pytest.raises(VerificationFailed, match="paragraph spacing"):
        _verify_saved_typography(path, PROFILES["radiology_ai"])


def test_second_section_bad_margin_is_rejected(tmp_path):
    path = tmp_path / "sections.docx"
    document = ET.Element(W + "document")
    body = ET.SubElement(document, W + "body")
    body.append(_paragraph("Probe Title", title=True))
    body.append(_paragraph("body"))
    for top in (1440, 2000):
        sect = ET.SubElement(body, W + "sectPr")
        ET.SubElement(sect, W + "pgMar", {W + side: str(top if side == "top" else 1440) for side in ("top", "right", "bottom", "left")})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))
        archive.writestr("word/styles.xml", _styles())
    with pytest.raises(VerificationFailed, match="margins"):
        _verify_saved_typography(path, PROFILES["radiology_ai"])


def test_one_incorrect_paragraph_among_compliant_ones_is_rejected(tmp_path):
    path = tmp_path / "one-bad.docx"
    _docx(path, [_paragraph("Probe Title", title=True), _paragraph("good"), _paragraph("bad", direct_font="Arial", direct_size="22"), _paragraph("good2")])
    with pytest.raises(VerificationFailed, match="run typography"):
        _verify_saved_typography(path, PROFILES["radiology_ai"])


def _field_xml(split=True, nested=False):
    p = ET.Element(W + "p")
    def field(parts):
        begin = ET.SubElement(p, W + "r"); ET.SubElement(begin, W + "fldChar", {W + "fldCharType": "begin"})
        for part in parts:
            r = ET.SubElement(p, W + "r"); ET.SubElement(r, W + "instrText").text = part
        end = ET.SubElement(p, W + "r"); ET.SubElement(end, W + "fldChar", {W + "fldCharType": "end"})
    field(["ZOTERO_", "ITEM CSL_JSON"] if split else ["ZOTERO_ITEM CSL_JSON"])
    if nested:
        field(["REF ", "anchor"])
    simple = ET.SubElement(p, W + "fldSimple", {W + "instr": " PAGE "})
    ET.SubElement(simple, W + "r").append(ET.Element(W + "t"))
    return p


def _field_doc(path, paragraph):
    document = ET.Element(W + "document"); body = ET.SubElement(document, W + "body"); body.append(paragraph); ET.SubElement(body, W + "sectPr")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))


def test_field_instructions_join_split_runs_and_include_nested_and_simple(tmp_path):
    split = tmp_path / "split.docx"; whole = tmp_path / "whole.docx"
    _field_doc(split, _field_xml(split=True, nested=True))
    _field_doc(whole, _field_xml(split=False, nested=True))
    assert _body_field_instructions(split) == _body_field_instructions(whole)
    assert _body_field_instructions(split) == ("ZOTERO_ITEM CSL_JSON", "REF anchor", "PAGE")
