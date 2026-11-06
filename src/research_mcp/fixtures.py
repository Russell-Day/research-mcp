"""Deterministic synthetic Word fixtures; never derived from the manuscript."""

from __future__ import annotations

from pathlib import Path

from .constants import WD_DO_NOT_SAVE_CHANGES, word_bgr
from .errors import RefusedUnsafe


def create_table_fixture(path: str | Path, *, visible: bool = False) -> Path:
    output = Path(path).resolve()
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite fixture: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = visible
        word.DisplayAlerts = 0
        document = word.Documents.Add()
        document.Content.Text = "Synthetic table fixture.\rProtected prose.\r"
        table_range = document.Range(document.Content.End - 1, document.Content.End - 1)
        table = document.Tables.Add(table_range, 7, 3)
        rows = (
            ("Model", "Training stage", "Accuracy"),
            ("LoRA", "Early", "0.90"),
            ("LoRA", "Middle", "0.91"),
            ("LoRA", "Late", "0.92"),
            ("Full fine-tuning", "Early", "0.93"),
            ("Full fine-tuning", "Middle", "0.94"),
            ("Full fine-tuning", "Late", "0.95"),
        )
        broken_blue = word_bgr(212, 141, 84)  # OOXML 548DD4 as a Word BGR integer
        for row_index, row in enumerate(rows, start=1):
            for column_index, value in enumerate(row, start=1):
                cell = table.Cell(row_index, column_index)
                cell.Range.Text = value
                cell.Shading.BackgroundPatternColor = broken_blue

        paragraph = document.Paragraphs(1).Range
        document.Bookmarks.Add("MCPFixtureLegacy", paragraph)
        document.Bookmarks.Add("_FixtureLegitimate", document.Paragraphs(2).Range)
        document.SaveAs2(str(output))
        return output
    finally:
        if document is not None:
            try:
                document.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        pythoncom.CoUninitialize()



# --- Mixed-run résumé-shaped fixture (plan 328 S2) --------------------------
#
# Written as raw OOXML rather than through Word so offline tests can read it
# without Word and the bytes are identical on every run. It mirrors the shapes
# that broke whole-line replacement on a real résumé: entry titles made of a
# bold, a plain and an italic run; dates in their own cell; section headings
# stored in mixed case under `w:caps` ("WorK Experience" displayed as
# "WORK EXPERIENCE"); plus a body paragraph holding a QUOTE field, a content
# control and a hyperlink so boundary refusals have something to name.

_MIXED_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)


def _xml_text(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _run(text: str, *props: str) -> str:
    rpr = "".join(f"<w:{name}/>" if "=" not in name else
                  '<w:{0} w:val="{1}"/>'.format(*name.split("=", 1)) for name in props)
    return (f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}'
            f'<w:t xml:space="preserve">{_xml_text(text)}</w:t></w:r>')


def _para(*runs: str) -> str:
    return "<w:p>" + "".join(runs) + "</w:p>"


def _heading_run(text: str) -> str:
    return _run(text, "caps", "color=595959", "spacing=10")


def _cell(*paragraphs: str, width: int = 7000) -> str:
    body = "".join(paragraphs) or "<w:p/>"
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{body}</w:tc>'


def _row(*cells: str) -> str:
    return "<w:tr>" + "".join(cells) + "</w:tr>"


def _entry(bold: str, plain: str, italic: str, lines: tuple[str, ...], date: str) -> str:
    title = _para(_run(bold, "b"), _run(plain), _run(italic, "i"))
    return _row(
        _cell(title, *(_para(_run(line)) for line in lines)),
        _cell(_para(_run(date)), width=2000),
    )


MIXED_RUN_DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f"<w:document {_MIXED_NS}><w:body>"
    + _para(_run("Jordan Example", "b", "sz=36"))
    + _para(_run("Summary: ", "b"), _run("data scientist with ", "i"),
            _run("clinical", "b", "i"), _run(" research experience", "i"))
    # Table properties as Word itself writes them, so a Word round trip
    # leaves them byte-for-byte alone.
    + "<w:tbl><w:tblPr><w:tblW w:w=\"9000\" w:type=\"dxa\"/>"
    + "<w:tblLook w:val=\"0000\" w:firstRow=\"0\" w:lastRow=\"0\" w:firstColumn=\"0\" w:lastColumn=\"0\" "
    + "w:noHBand=\"0\" w:noVBand=\"0\"/></w:tblPr>"
    "<w:tblGrid><w:gridCol w:w=\"7000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
    + _row(_cell(_para(_heading_run("Education"))), _cell(_para(_heading_run("Dates")), width=2000))
    + _entry("Lakeside University", ", Honors College, ", "Springfield, ST",
             ("Major: Statistics",), "08/2022 - Present")
    + _row(_cell(_para(_run("WorK Experience", "i", "caps", "color=595959", "spacing=10"))),
           _cell(_para(), width=2000))
    + _entry("Lakeside University", ", College of Arts, Rivera Lab (Soil), ", "Teaching Assistant",
             ("Field research on soil microbes and crop yield",), "01/2024 – Present ")
    + _entry("Delta Park Museum", ", Department of Collections, ", "Outreach Volunteer",
             ("Contributed to collections review, donor records, cataloging",
              "Focused on: Visitor Programs, Place-Based Exhibition"),
             "03/2020 – Present ")
    + _row(_cell(_para(_heading_run("Projects"))), _cell(_para(), width=2000))
    + _entry("Norhaven Health", ", Survey Data Pipeline ", "(Python)", (), "07/2024")
    + "</w:tbl>"
    + _para(_run("Cited as "),
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> QUOTE "Example 2024" </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>Example 2024</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
            _run(" in the "),
            '<w:sdt><w:sdtPr><w:alias w:val="Venue"/><w:id w:val="4242"/></w:sdtPr><w:sdtEndPr/><w:sdtContent>'
            + _run("Venue Name") + "</w:sdtContent></w:sdt>",
            _run(" proceedings, see "),
            '<w:hyperlink w:anchor="Top">' + _run("the top", "u=single") + "</w:hyperlink>",
            _run(" for details."))
    + _para(_run("Plain closing paragraph."))
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>'
    "</w:body></w:document>"
)

_MIXED_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)

_MIXED_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/></Relationships>'
)

_MIXED_DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    'Target="styles.xml"/></Relationships>'
)

# Word's own defaults, so a Word save does not rewrite the fixture's table or paragraph properties.
_MIXED_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault/></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
    '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
    '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
    '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
    "</w:styles>"
)


def create_mixed_run_fixture(path: str | Path) -> Path:
    """Write the synthetic mixed-run résumé fixture; refuses to overwrite."""

    import zipfile

    output = Path(path).resolve()
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite fixture: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "x", zipfile.ZIP_DEFLATED) as archive:
        for name, data in (
            ("[Content_Types].xml", _MIXED_CONTENT_TYPES),
            ("_rels/.rels", _MIXED_RELS),
            ("word/document.xml", MIXED_RUN_DOCUMENT_XML),
            ("word/_rels/document.xml.rels", _MIXED_DOCUMENT_RELS),
            ("word/styles.xml", _MIXED_STYLES),
        ):
            info = zipfile.ZipInfo(name, date_time=(2026, 10, 30, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data)
    return output


# --------------------------------------------------------------------------
# Synthetic tracked-changes fixture (plan 501 S2). Built without Word, byte-identical
# on every run. Paragraph numbers below are layout indices.
#   1  plain paragraph with a bookmark, no revisions
#   2  another author's w:ins id=1 ("highly ") and w:del id=2 (" overall")
#   3  a ZOTERO citation field "(1)" between two plain runs
#   4  a comment (id 0) on "Commented phrase"
#   5  a run whose bold is a tracked formatting change (w:rPrChange id=3)
#   6  moved-from text (w:moveFrom id=5, range id=4)
#   7  moved-to text (w:moveTo id=7, range id=6)
#   8  a drawing inside mc:AlternateContent
#   9  a paragraph whose mark was inserted (pPr/rPr/w:ins id=8)
#  10  a tracked paragraph-property change (w:pPrChange id=9) and a deleted paragraph mark (w:del id=12)
#  11  table cell A with w:ins id=10 ("added ")
#  12  table cell B whose cell properties changed (w:tcPrChange id=11)
#  13  the row-end marker
#  14  closing paragraph
# --------------------------------------------------------------------------

TRACKED_AUTHOR = "Other Author"
TRACKED_DATE = "2026-10-10T17:37:00Z"
_TRACKED_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" '
    'mc:Ignorable="w16du"'
)


def _rev(tag: str, rev_id: int, inner: str = "") -> str:
    attrs = f'w:id="{rev_id}" w:author="{TRACKED_AUTHOR}" w:date="{TRACKED_DATE}" w16du:dateUtc="{TRACKED_DATE}"'
    return f"<w:{tag} {attrs}/>" if not inner else f"<w:{tag} {attrs}>{inner}</w:{tag}>"


def _deleted_run(text: str) -> str:
    return f'<w:r><w:delText xml:space="preserve">{_xml_text(text)}</w:delText></w:r>'


TRACKED_DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f"<w:document {_TRACKED_NS}><w:body>"
    + _para(_run("Plain "), '<w:bookmarkStart w:id="20" w:name="keep"/>', _run("opening"),
            '<w:bookmarkEnd w:id="20"/>', _run(" paragraph with no revisions."))
    + _para(_run("The model was "), _rev("ins", 1, _run("highly ")), _run("accurate"),
            _rev("del", 2, _deleted_run(" overall")), _run(" on the test set."))
    + _para(_run("Prior work "),
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>(1)</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
            _run(". The same method applies."))
    + _para('<w:commentRangeStart w:id="0"/>', _run("Commented phrase"), '<w:commentRangeEnd w:id="0"/>',
            '<w:r><w:commentReference w:id="0"/></w:r>', _run(" stays anchored."))
    + _para(f'<w:r><w:rPr><w:b/>{_rev("rPrChange", 3, "<w:rPr/>")}</w:rPr>'
            '<w:t xml:space="preserve">Bold words</w:t></w:r>', _run(" then plain."))
    + _para('<w:moveFromRangeStart w:id="4" w:author="Other Author" w:date="2026-10-10T17:37:00Z" w:name="move1"/>',
            _rev("moveFrom", 5, _run("Moved sentence. ")), '<w:moveFromRangeEnd w:id="4"/>', _run("Stay text."))
    + _para('<w:moveToRangeStart w:id="6" w:author="Other Author" w:date="2026-10-10T17:37:00Z" w:name="move1"/>',
            _rev("moveTo", 7, _run("Moved sentence. ")), '<w:moveToRangeEnd w:id="6"/>', _run("Arrival text."))
    + _para(_run("Figure here: "),
            '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" '
            'distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr w:id="30" id="30" name="Shape 1"/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>'
            '<mc:Fallback><w:pict><v:rect style="width:72pt;height:72pt"/></w:pict></mc:Fallback>'
            '</mc:AlternateContent></w:r>', _run(" caption text."))
    + f'<w:p><w:pPr><w:rPr>{_rev("ins", 8)}</w:rPr></w:pPr>{_run("Inserted paragraph mark here.")}</w:p>'
    + (f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{_rev("del", 12)}</w:rPr>{_rev("pPrChange", 9, "<w:pPr/>")}</w:pPr>'
       f'{_run("Centered by the other author.")}</w:p>')
    + "<w:tbl><w:tblPr><w:tblW w:w=\"9000\" w:type=\"dxa\"/></w:tblPr>"
    "<w:tblGrid><w:gridCol w:w=\"4500\"/><w:gridCol w:w=\"4500\"/></w:tblGrid>"
    + _row(_cell(_para(_run("Cell "), _rev("ins", 10, _run("added ")), _run("A")), width=4500),
           '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/>'
           + _rev("tcPrChange", 11, '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>')
           + "</w:tcPr>" + _para(_run("Cell B")) + "</w:tc>")
    + "</w:tbl>"
    + _para(_run("Closing paragraph."))
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>'
    "</w:body></w:document>"
)

_TRACKED_COMMENTS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    f'<w:comment w:id="0" w:author="{TRACKED_AUTHOR}" w:date="{TRACKED_DATE}" w:initials="OA">'
    '<w:p><w:r><w:t>Check this phrase.</w:t></w:r></w:p></w:comment></w:comments>'
)

_TRACKED_CONTENT_TYPES = _MIXED_CONTENT_TYPES.replace(
    "</Types>",
    '<Override PartName="/word/comments.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')

_TRACKED_DOCUMENT_RELS = _MIXED_DOCUMENT_RELS.replace(
    "</Relationships>",
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
    'Target="comments.xml"/></Relationships>')


def create_tracked_fixture(path: str | Path) -> Path:
    """Write the synthetic tracked-changes fixture (see the table above); refuses to overwrite."""

    import zipfile

    output = Path(path).resolve()
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite fixture: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "x", zipfile.ZIP_DEFLATED) as archive:
        for name, data in (
            ("[Content_Types].xml", _TRACKED_CONTENT_TYPES),
            ("_rels/.rels", _MIXED_RELS),
            ("word/document.xml", TRACKED_DOCUMENT_XML),
            ("word/_rels/document.xml.rels", _TRACKED_DOCUMENT_RELS),
            ("word/styles.xml", _MIXED_STYLES),
            ("word/comments.xml", _TRACKED_COMMENTS),
        ):
            info = zipfile.ZipInfo(name, date_time=(2026, 10, 30, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data)
    return output
