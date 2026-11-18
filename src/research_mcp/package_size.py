"""Report when a write makes the .docx package much larger, and say which parts grew.

Word's save can grow a document that the edit itself barely touched. The case that prompted this
(BIOINF 527 PS2, 2026-11-08: 1.26 MB -> 2.75 MB after eight text-only paragraph inserts) was not the
images: their bytes were identical. The document, a Google Docs export, has "Embed fonts in the file" on
(``w:embedTrueTypeFonts``) without "Embed only the characters used" (``w:saveSubsetFonts``), so Word's save
embedded its theme fonts Calibri and Cambria in full (about 3.2 MB raw, 1.6 MB compressed). Any Word save of
that document does the same (a plain save of the untouched handout: 2.74 MB; with 'Embed only the characters
used' ticked: 0.89 MB), so the result names the cause and the remedy instead of leaving the growth unexplained.

Read only: nothing here changes the package. A failure to read either package yields no report.
"""

from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Report only growth that matters for upload limits and mail, not the few KB every save moves.
MIN_GROWTH_BYTES = 256 * 1024
MIN_GROWTH_RATIO = 0.20
_STYLES = {"embedRegular": "regular", "embedBold": "bold", "embedItalic": "italic",
           "embedBoldItalic": "bold italic"}


def _mb(n: int) -> str:
    return f"{n / 1e6:.2f} MB"  # decimal MB, as the item reported 1.26 MB


def _fonts(z: zipfile.ZipFile) -> dict[tuple[str, str], str]:
    """(font name, style) -> embedded part name, from fontTable.xml and its relationships."""
    try:
        table = ET.fromstring(z.read("word/fontTable.xml"))
    except KeyError:
        return {}
    targets: dict[str, str] = {}
    try:
        rels = ET.fromstring(z.read("word/_rels/fontTable.xml.rels"))
        for rel in rels.iter(f"{{{PKG_REL}}}Relationship"):
            target = rel.get("Target") or ""
            targets[rel.get("Id") or ""] = (target.lstrip("/") if target.startswith("/")
                                            else posixpath.normpath(posixpath.join("word", target)))
    except KeyError:
        pass
    found: dict[tuple[str, str], str] = {}
    for font in table.iter(f"{{{W}}}font"):
        name = font.get(f"{{{W}}}name") or "?"
        for tag, style in _STYLES.items():
            embed = font.find(f"{{{W}}}{tag}")
            if embed is not None:
                found[(name, style)] = targets.get(embed.get(f"{{{R}}}id") or "", "")
    return found


def _settings(z: zipfile.ZipFile) -> set[str]:
    try:
        root = ET.fromstring(z.read("word/settings.xml"))
    except KeyError:
        return set()
    names = set()
    for child in root:
        local = child.tag.rsplit("}", 1)[-1]
        if local in ("embedTrueTypeFonts", "saveSubsetFonts"):
            if (child.get(f"{{{W}}}val") or "true").lower() not in ("0", "false", "off"):
                names.add(local)
    return names


def _kind(name: str) -> str:
    if name.startswith("word/fonts/"):
        return "fonts"
    if name.startswith("word/media/"):
        return "media"
    return "other"


def size_growth(before: str | Path, after: str | Path, *, min_bytes: int = MIN_GROWTH_BYTES,
                min_ratio: float = MIN_GROWTH_RATIO) -> dict | None:
    """None unless ``after`` is at least ``min_bytes`` and ``min_ratio`` larger than ``before``.

    Otherwise: sizes, stored (compressed) bytes by kind of part (fonts, media, other), fonts embedded now that
    were not before, whether any media part's bytes changed, and a one-line ``note`` with the cause and the
    remedy when it is known.
    """
    try:
        size_before, size_after = Path(before).stat().st_size, Path(after).stat().st_size
        grew = size_after - size_before
        if grew < min_bytes or grew < min_ratio * max(size_before, 1):
            return None
        with zipfile.ZipFile(before) as zb, zipfile.ZipFile(after) as za:
            stored_b = {i.filename: i.compress_size for i in zb.infolist()}
            stored_a = {i.filename: i.compress_size for i in za.infolist()}
            fonts_b, fonts_a = _fonts(zb), _fonts(za)
            settings = _settings(za)
            # (CRC-32, size) from the directory: no image is decompressed or held in memory.
            media_b = sorted((i.CRC, i.file_size) for i in zb.infolist() if _kind(i.filename) == "media")
            media_a = sorted((i.CRC, i.file_size) for i in za.infolist() if _kind(i.filename) == "media")
    except Exception:  # read only and advisory: a report must never fail or block a publish
        return None

    # By kind, not by part name: Word renumbers media and font parts on save (image1 <-> image4).
    parts = {kind: {"bytes_before": sum(v for n, v in stored_b.items() if _kind(n) == kind),
                    "bytes_after": sum(v for n, v in stored_a.items() if _kind(n) == kind)}
             for kind in ("fonts", "media", "other")}
    added = [{"font": name, "style": style, "part": part, "bytes": stored_a.get(part, 0)}
             for (name, style), part in sorted(fonts_a.items()) if (name, style) not in fonts_b]
    # Word renames and reorders media on save, so compare the multiset of images, not names.
    media_changed = media_b != media_a
    font_bytes = sum(stored_a.get(part, 0) for part in {item["part"] for item in added})  # a shared part once

    notes = [f"The file grew from {_mb(size_before)} to {_mb(size_after)}."]
    if added:
        listed = ", ".join(f"{item['font']} {item['style']}" for item in added)
        stored = f"; {_mb(font_bytes)} stored" if font_bytes else ""
        if "embedTrueTypeFonts" in settings:
            cause = " because the document has 'Embed fonts in the file' on"
            if "saveSubsetFonts" not in settings:
                cause += " without 'Embed only the characters used in the document'"
        else:  # settings missing or saying off: do not claim a cause the package does not show
            cause = ""
        notes.append(
            f"Word's save embedded fonts the document did not carry before ({listed}{stored}){cause}. "
            "Word embeds every font the document uses or names (theme fonts included) on any save with these "
            "settings, so this is usually not the edited text. To shrink it, in Word: File > Options > Save, "
            "tick 'Embed only the characters used in the document' (or untick 'Embed fonts in the file'), "
            "then save.")
    if media_changed:
        notes.append("Image bytes changed: Word re-encoded or replaced at least one picture.")
    elif media_b:
        notes.append("Image bytes are unchanged.")
    return {"bytes_before": size_before, "bytes_after": size_after, "grew_by": grew, "parts": parts,
            "fonts_embedded": added, "media_changed": media_changed, "note": " ".join(notes)}
