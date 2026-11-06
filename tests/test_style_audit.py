from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from research_mcp.style_audit import audit_document_style

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + WNS + "}"
ET.register_namespace("w", WNS)


def _run(text, *, bold=None, italic=None):
    run = ET.Element(W + "r")
    if bold is not None or italic is not None:
        props = ET.SubElement(run, W + "rPr")
        if bold is not None:
            ET.SubElement(props, W + "b", {} if bold else {W + "val": "0"})
        if italic is not None:
            ET.SubElement(props, W + "i", {} if italic else {W + "val": "0"})
    ET.SubElement(run, W + "t").text = text
    return run


def _docx(path: Path, paragraphs, *, default_bold=False, heading=False):
    document = ET.Element(W + "document")
    body = ET.SubElement(document, W + "body")
    for paragraph in paragraphs:
        body.append(paragraph)
    ET.SubElement(body, W + "sectPr")
    styles = ET.Element(W + "styles")
    defaults = ET.SubElement(styles, W + "docDefaults")
    rdefault = ET.SubElement(ET.SubElement(defaults, W + "rPrDefault"), W + "rPr")
    if default_bold:
        ET.SubElement(rdefault, W + "b")
    if heading:
        style = ET.SubElement(styles, W + "style", {W + "type": "paragraph", W + "styleId": "Heading1"})
        ET.SubElement(style, W + "name", {W + "val": "heading 1"})
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))
        archive.writestr("word/styles.xml", ET.tostring(styles, encoding="utf-8", xml_declaration=True))


def _p(*runs, style=None):
    paragraph = ET.Element(W + "p")
    if style:
        ppr = ET.SubElement(paragraph, W + "pPr")
        ET.SubElement(ppr, W + "pStyle", {W + "val": style})
    for run in runs:
        paragraph.append(run)
    return paragraph


def test_inherited_bold_is_reported_as_whole_body_warning(tmp_path):
    path = tmp_path / "inherited.docx"
    _docx(path, [_p(_run("Title")), _p(_run("Body paragraph"))], default_bold=True)
    result = audit_document_style(path)
    assert result["paragraphs"][1]["spans"][0]["bold"] is True
    assert result["findings"][0]["code"] == "WHOLE_BODY_PARAGRAPH_BOLD"
    assert result["findings"][0]["recommendation"] == "review whole paragraph"


def test_partial_bold_recommends_label_only(tmp_path):
    path = tmp_path / "partial.docx"
    _docx(path, [_p(_run("Purpose:", bold=True), _run(" explain the study"))])
    result = audit_document_style(path)
    assert result["findings"][0]["code"] == "PARTIAL_BOLD_SPAN"
    assert result["findings"][0]["recommendation"] == "label-only"


def test_style_bold_toggle_can_cancel_inherited_default(tmp_path):
    path = tmp_path / "toggle.docx"
    document = ET.Element(W + "document")
    body = ET.SubElement(document, W + "body")
    paragraph = _p(_run("body"), style="Toggle")
    body.append(paragraph)
    ET.SubElement(body, W + "sectPr")
    styles = ET.Element(W + "styles")
    defaults = ET.SubElement(styles, W + "docDefaults")
    default_rpr = ET.SubElement(ET.SubElement(defaults, W + "rPrDefault"), W + "rPr")
    ET.SubElement(default_rpr, W + "b")
    style = ET.SubElement(styles, W + "style", {W + "type": "paragraph", W + "styleId": "Toggle"})
    ET.SubElement(style, W + "rPr").append(ET.Element(W + "b"))
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))
        archive.writestr("word/styles.xml", ET.tostring(styles, encoding="utf-8", xml_declaration=True))
    result = audit_document_style(path)
    assert result["paragraphs"][0]["spans"][0]["bold"] is False
    assert not result["findings"]


def test_heading_whole_bold_is_not_accidental_body_warning(tmp_path):
    path = tmp_path / "heading.docx"
    _docx(path, [_p(_run("Results", bold=True), style="Heading1")], heading=True)
    result = audit_document_style(path)
    assert not result["findings"]


def test_table_row_synthetic_index_matches_paragraph_reader(tmp_path):
    path = tmp_path / "table.docx"
    document = ET.Element(W + "document")
    body = ET.SubElement(document, W + "body")
    body.append(_p(_run("before")))
    table = ET.SubElement(body, W + "tbl")
    row = ET.SubElement(table, W + "tr")
    cell = ET.SubElement(row, W + "tc")
    cell.append(_p(_run("cell", bold=True)))
    ET.SubElement(body, W + "sectPr")
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))
    result = audit_document_style(path)
    assert [item["index"] for item in result["paragraphs"]] == [1, 2]
    assert result["paragraphs"][1]["scope"] == "table"
    assert result["paragraphs"][1]["spans"][0]["bold"] is True


def test_style_false_does_not_clear_inherited_true_but_direct_false_does():
    from research_mcp.style_audit import _style_layer, _direct
    rpr = ET.Element(W + "rPr")
    ET.SubElement(rpr, W + "b", {W + "val": "0"})
    assert _style_layer({"b": True, "i": False}, rpr)["b"] is True
    assert _direct({"b": True, "i": False}, rpr)["b"] is False


def test_character_style_keeps_paragraph_bold(tmp_path):
    from research_mcp.style_audit import snapshot_emphasis
    document = ET.Element(W + "document")
    body = ET.SubElement(document, W + "body")
    p = ET.SubElement(body, W + "p")
    ppr = ET.SubElement(p, W + "pPr")
    ET.SubElement(ppr, W + "pStyle", {W + "val": "StrongBody"})
    r = ET.SubElement(p, W + "r")
    ET.SubElement(ET.SubElement(r, W + "rPr"), W + "rStyle", {W + "val": "Emphasis"})
    ET.SubElement(r, W + "t").text = "X"
    styles = ET.Element(W + "styles")
    for name, kind, prop in (("StrongBody", "paragraph", "b"), ("Emphasis", "character", "i")):
        style = ET.SubElement(styles, W + "style", {W + "styleId": name, W + "type": kind})
        ET.SubElement(ET.SubElement(style, W + "rPr"), W + prop)
    path = tmp_path / "styles.docx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", ET.tostring(document))
        archive.writestr("word/styles.xml", ET.tostring(styles))
    assert snapshot_emphasis(path)[(1, 1)] == ("X", True, True)
