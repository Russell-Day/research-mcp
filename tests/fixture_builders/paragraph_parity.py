"""Build the synthetic paragraph-parity document and record Word's own paragraph list for it.

`tests/test_paragraphs.py` checks that `read_paragraphs` numbers paragraphs the way
`document.Paragraphs(i)` does, against a capture recorded once from live Word. This script makes both files:

    python tests/fixture_builders/paragraph_parity.py build    # tests/data/paragraph_parity.docx (offline)
    python tests/fixture_builders/paragraph_parity.py capture  # tests/data/com_paragraphs_parity.json (Windows + Word)

The document holds the structures that make Word's numbering differ from a flat `.//w:p` walk or make the
package text differ from Word's text: body tables (one row-end paragraph per row), merged cells, a cell with
two paragraphs, an empty cell, a nested table, Zotero-style citation fields and a bibliography field that
spans several paragraphs, page and line breaks, tab stops and tab characters, inline and anchored pictures,
bookmarks, a comment range and empty paragraphs. All text is made up.

`capture` copies the document under C:\\tmp, opens the copy read-only in a separate, invisible Word instance
started for this run (never an existing Word window), reads every paragraph, closes it without saving and
quits only that instance.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import struct
import subprocess
import sys
import zipfile
import zlib


ROOT = Path(__file__).resolve().parents[2]
DOCX = ROOT / "tests" / "data" / "paragraph_parity.docx"
CAPTURE = ROOT / "tests" / "data" / "com_paragraphs_parity.json"
WORK = Path(r"C:\tmp\research-mcp-parity")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def t(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'


def tab() -> str:
    return "<w:r><w:tab/></w:r>"


def p(*runs: str, style: str | None = None, ppr: str = "") -> str:
    props = (f'<w:pStyle w:val="{style}"/>' if style else "") + ppr
    return f"<w:p>{f'<w:pPr>{props}</w:pPr>' if props else ''}{''.join(runs)}</w:p>"


def field(instr: str, result: str) -> str:
    """A complex field the way Zotero writes one: begin, code, separate, result, end."""

    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {_esc(instr)} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'{t(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def cite(n: str, cid: str) -> str:
    return field(f'ADDIN ZOTERO_ITEM CSL_CITATION {{"citationID":"{cid}","citationItems":[]}}', f"({n})")


def cell(*paragraphs: str, props: str = "", width: int = 3000) -> str:
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{props}</w:tcPr>'
            f"{''.join(paragraphs) or p()}</w:tc>")


def table(rows: list[list[str]], columns: int, width: int = 3000) -> str:
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for _ in range(columns))
    body = "".join(f"<w:tr>{''.join(row)}</w:tr>" for row in rows)
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
            + "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
            + f"</w:tblBorders></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>")


def page_break() -> str:
    return p('<w:r><w:br w:type="page"/></w:r>')


def picture(n: int, anchored: bool) -> str:
    size = 457200
    graphic = (f'<wp:docPr id="{n}" name="Picture {n}"/>'
               f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
               f'<a:graphic><a:graphicData uri="{PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="dot{n}.png"/>'
               '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/>'
               f'</a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{size}" cy="{size}"/>'
               '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
    if anchored:
        shape = ('<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" '
                 'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
                 '<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>'
                 '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
                 f'<wp:extent cx="{size}" cy="{size}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
                 f'<wp:wrapSquare wrapText="bothSides"/>{graphic}</wp:anchor>')
    else:
        shape = f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{size}" cy="{size}"/>{graphic}</wp:inline>'
    return f"<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>{shape}</w:drawing></w:r>"


def _png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\x40\x80\xc0")) + chunk(b"IEND", b""))


def body() -> str:
    measured = [
        [cell(p(t("Group"))), cell(p(t("Runs"))), cell(p(t("Mean")))],
        [cell(p(t("Alpha")), props='<w:vMerge w:val="restart"/>'), cell(p(t("3"))), cell(p(t("0.91")))],
        [cell(p(), props="<w:vMerge/>"), cell(p(t("4"))), cell(p(t("0.88 ")), p(t("(rerun)")))],
        [cell(p(t("Beta"))), cell(p()), cell(p(t("0.79 "), cite("5", "cell")))],
        [cell(p(t("Gamma"))), cell(p(t("2"))), cell(p(t("0.85")))],
    ]
    groups = [
        [cell(p(t("Both groups")), props='<w:gridSpan w:val="2"/>', width=6000)],
        [cell(p(t("Early"))), cell(p(t("12")))],
        [cell(p(t("Late"))), cell(p(t("9")))],
    ]
    inner = table([[cell(p(t("a1")), width=1400), cell(p(t("b1")), width=1400)],
                   [cell(p(t("a2")), width=1400), cell(p(t("b2")), width=1400)]], 2, 1400)
    nested = [
        [cell(p(t("Outer left"))), cell(inner, p(t("Below the inner table")))],
        [cell(p(t("Outer bottom"))), cell(p())],
    ]
    results = [[cell(p(t(a))), cell(p(t(b)))] for a, b in
               (("Measure", "Value"), ("Precision", "0.90"), ("Recall", "0.86"), ("F1", "0.88"))]
    bibliography = [
        "Adams A, Baker B. A first made-up reference. J Synth Stud. 2020;1:1-10.",
        "Chen C. A second made-up reference. Synth Methods. 2021;2:11-20.",
        "Diaz D, Evans E, Fox F. A third made-up reference. Synth Rev. 2022;3:21-30.",
        "Garcia G. A fourth made-up reference. Synth Lett. 2023;4:31-40.",
    ]
    bib = []
    for n, entry in enumerate(bibliography, 1):
        runs = [t(f"{n}."), tab(), t(entry)]
        if n == 1:
            runs.insert(0, '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                           '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY </w:instrText></w:r>'
                           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        if n == len(bibliography):
            runs.append('<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        bib.append(p(*runs, style="Bibliography"))

    return "".join([
        p(t("A Synthetic Paragraph Parity Fixture"), style="Title"),
        p(),
        p(t("1. Introduction"), style="Heading1"),
        p('<w:bookmarkStart w:id="0" w:name="intro"/>', t("Short sentences stand in for prose "), cite("1", "one"),
          t(" and continue after the citation."), '<w:bookmarkEnd w:id="0"/>'),
        p(t("Two sources agree "), cite("2,3", "two"), t(", and a third disagrees "), cite("4", "three"), t(".")),
        p('<w:commentRangeStart w:id="0"/>', t("This sentence carries a reviewer comment."),
          '<w:commentRangeEnd w:id="0"/>', '<w:r><w:commentReference w:id="0"/></w:r>'),
        p(t("First line"), "<w:r><w:br/></w:r>", t("second line of the same paragraph.")),
        p(),
        p(t("1.1 Measurements"), style="Heading2"),
        p(t("Table 1 lists the measured values.")),
        table(measured, 3),
        p(t("Values are means over the runs.")),
        page_break(),
        p(t("1.2 Groups"), style="Heading2"),
        table(groups, 2),
        p(),
        p(t("1.2.1 Nested layout"), style="Heading3"),
        table(nested, 2),
        p(t("Label"), tab(), t("Value"), ppr='<w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs>'),
        p(tab(), t("Indented by a tab character.")),
        p(picture(1, anchored=False)),
        p(t("Figure 1. A single-pixel image.")),
        p(picture(2, anchored=True), t("An anchored image sits beside this text.")),
        page_break(),
        p(t("2. Results"), style="Heading1"),
        table(results, 2),
        p('<w:r><w:lastRenderedPageBreak/></w:r>', t("The results end here.")),
        p(),
        p(t("References"), style="Heading1"),
        *bib,
        p(),
    ])


def _styles() -> str:
    def style(style_id: str, name: str, level: int | None = None, based: str = "Normal") -> str:
        outline = f'<w:pPr><w:keepNext/><w:outlineLvl w:val="{level}"/></w:pPr>' if level is not None else ""
        return (f'<w:style w:type="paragraph" w:styleId="{style_id}"><w:name w:val="{name}"/>'
                f'<w:basedOn w:val="{based}"/><w:next w:val="Normal"/><w:qFormat/>{outline}</w:style>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W}">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
            + style("Title", "Title") + style("Heading1", "heading 1", 0) + style("Heading2", "heading 2", 1)
            + style("Heading3", "heading 3", 2) + style("Bibliography", "Bibliography")
            + "</w:styles>")


def _parts() -> dict[str, bytes]:
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}"><w:body>'
                + body()
                + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
                  'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    comments = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="{W}">'
                '<w:comment w:id="0" w:author="Reviewer" w:initials="R"><w:p><w:r><w:t>Check this.</w:t></w:r></w:p>'
                "</w:comment></w:comments>")
    main = "application/vnd.openxmlformats-officedocument.wordprocessingml"
    return {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
            f'<Override PartName="/word/document.xml" ContentType="{main}.document.main+xml"/>'
            f'<Override PartName="/word/styles.xml" ContentType="{main}.styles+xml"/>'
            f'<Override PartName="/word/comments.xml" ContentType="{main}.comments+xml"/></Types>').encode(),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>').encode(),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rIdStyles" Type="{REL}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rIdComments" Type="{REL}/comments" Target="comments.xml"/>'
            f'<Relationship Id="rIdImage" Type="{REL}/image" Target="media/image1.png"/></Relationships>').encode(),
        "word/document.xml": document.encode(),
        "word/styles.xml": _styles().encode(),
        "word/comments.xml": comments.encode(),
        "word/media/image1.png": _png(),
    }


def build(output: Path = DOCX) -> Path:
    """Write the package with fixed timestamps and a fixed host byte, so a rebuild is byte-identical on any OS."""

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in _parts().items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.create_system = 0  # MS-DOS, as Word writes it; the default follows the OS that builds
            archive.writestr(info, data, zipfile.ZIP_DEFLATED)
    return output


def _word_pids() -> set[int]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, check=False).stdout
    return {int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"WINWORD')}


def _com_text(text: str) -> str:
    """Range.Text minus the paragraph mark (and the cell or row-end mark)."""

    for mark in ("\r\x07", "\r", "\x07"):
        if text.endswith(mark):
            return text[: -len(mark)]
    return text


def capture(source: Path = DOCX, output: Path = CAPTURE) -> list[dict]:
    import pythoncom
    import win32com.client

    WORK.mkdir(parents=True, exist_ok=True)
    copy = WORK / source.name
    shutil.copyfile(source, copy)
    before = _word_pids()
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    try:
        started = _word_pids() - before
        if len(started) != 1 or int(word.Documents.Count):
            word = None  # not an instance this run started: never close or quit it
            raise SystemExit(f"refusing: expected one new, empty Word instance (new pids: {sorted(started)})")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        document = word.Documents.Open(FileName=str(copy), ReadOnly=True, AddToRecentFiles=False, Visible=False,
                                       ConfirmConversions=False, OpenAndRepair=False, NoEncodingDialog=True)
        try:
            rows = []
            for index in range(1, int(document.Paragraphs.Count) + 1):
                paragraph = document.Paragraphs(index)
                rows.append({
                    "i": index,
                    "text": _com_text(str(paragraph.Range.Text)),
                    "style": str(paragraph.Style.NameLocal),
                    "in_table": bool(paragraph.Range.Information(12)),  # wdWithInTable
                })
        finally:
            document.Close(0)  # wdDoNotSaveChanges
    finally:
        if word is not None:
            word.Quit(0)
        pythoncom.CoUninitialize()
    copy.unlink(missing_ok=True)
    output.write_text(json.dumps(rows, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    return rows


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("action", choices=("build", "capture"))
    args = parser.parse_args(argv)
    if args.action == "build":
        print(build())
    else:
        rows = capture()
        print(f"{CAPTURE}: {len(rows)} paragraphs")
    return 0


if __name__ == "__main__":
    sys.exit(main())
