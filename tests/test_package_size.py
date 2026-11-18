"""package_size.size_growth: say why a write made the .docx much larger (chk_18a037b8933e45d5)."""

from __future__ import annotations

import os
import zipfile
from pathlib import Path

from research_mcp.package_size import size_growth

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' \
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
RELS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'


def _docx(path: Path, *, fonts: dict[str, bytes], media: dict[str, bytes], subset: bool = False) -> Path:
    """fonts: font name -> embedded regular bytes (stored uncompressed so sizes are predictable)."""
    table, rels = [], []
    for n, (name, data) in enumerate(fonts.items(), start=1):
        table.append(f'<w:font w:name="{name}"><w:embedRegular r:id="rId{n}" w:fontKey="{{0}}"/></w:font>')
        rels.append(f'<Relationship Id="rId{n}" Type="font" Target="fonts/font{n}.odttf"/>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", f"<w:document {W}><w:body/></w:document>")
        z.writestr("word/settings.xml", f"<w:settings {W}><w:embedTrueTypeFonts/>"
                   + ("<w:saveSubsetFonts/>" if subset else "") + "</w:settings>")
        z.writestr("word/fontTable.xml", f'<w:fonts {W}><w:font w:name="Arial"/>{"".join(table)}</w:fonts>')
        z.writestr("word/_rels/fontTable.xml.rels", f"<Relationships {RELS}>{''.join(rels)}</Relationships>")
        for n, data in enumerate(fonts.values(), start=1):
            z.writestr(f"word/fonts/font{n}.odttf", data)
        for name, data in media.items():
            z.writestr(f"word/media/{name}", data)
    return path


def _noise(n: int, seed: int) -> bytes:
    return os.urandom(n) if seed < 0 else bytes((i * 7919 + seed) % 251 for i in range(n))


def test_small_growth_is_not_reported(tmp_path):
    before = _docx(tmp_path / "a.docx", fonts={}, media={"image1.png": _noise(100_000, 1)})
    after = _docx(tmp_path / "b.docx", fonts={}, media={"image1.png": _noise(100_000, 1)})
    assert size_growth(before, after) is None


def test_newly_embedded_fonts_are_named_and_renamed_media_is_unchanged(tmp_path):
    a, b = _noise(300_000, -1), _noise(200_000, -1)
    noto = _noise(50_000, -1)
    before = _docx(tmp_path / "a.docx", fonts={"Noto Sans Symbols": noto},
                   media={"image1.png": a, "image2.png": b})
    # Word renumbers media on save; the same bytes under swapped names are not a change.
    after = _docx(tmp_path / "b.docx", fonts={"Noto Sans Symbols": noto, "Calibri": _noise(800_000, -1)},
                  media={"image1.png": b, "image2.png": a})
    report = size_growth(before, after)
    assert report is not None
    assert [(f["font"], f["style"]) for f in report["fonts_embedded"]] == [("Calibri", "regular")]
    assert report["fonts_embedded"][0]["part"] == "word/fonts/font2.odttf"
    assert report["media_changed"] is False
    assert report["parts"]["fonts"]["bytes_after"] - report["parts"]["fonts"]["bytes_before"] > 790_000
    assert report["parts"]["media"]["bytes_after"] == report["parts"]["media"]["bytes_before"]
    assert "Calibri regular" in report["note"]
    assert "'Embed only the characters used in the document'" in report["note"]
    assert "Image bytes are unchanged." in report["note"]


def test_reencoded_media_is_reported(tmp_path):
    before = _docx(tmp_path / "a.docx", fonts={}, media={"image1.png": _noise(100_000, -1)})
    after = _docx(tmp_path / "b.docx", fonts={}, media={"image1.png": _noise(900_000, -1)})
    report = size_growth(before, after)
    assert report["media_changed"] is True and report["fonts_embedded"] == []
    assert "re-encoded or replaced" in report["note"]


def test_unreadable_package_yields_no_report(tmp_path):
    before = tmp_path / "a.docx"
    before.write_bytes(b"x")
    after = tmp_path / "b.docx"
    after.write_bytes(b"not a zip" * 100_000)
    assert size_growth(before, after) is None


def test_publish_inplace_reports_growth(tmp_path):
    from research_mcp import publication
    from research_mcp.checkpoints import file_sha256

    target = _docx(tmp_path / "doc.docx", fonts={}, media={"image1.png": _noise(100_000, -1)})
    staged = _docx(tmp_path / "staged.docx", fonts={"Calibri": _noise(900_000, -1)},
                   media={"image1.png": _noise(100_000, -1)})
    with publication.capture(tool="test") as records:
        publication.publish_inplace(staged, target, expected_sha256=file_sha256(target))
    fields = publication.write_fields(records)
    assert fields["size_growth"]["fonts_embedded"][0]["font"] == "Calibri"
