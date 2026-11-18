"""Word's built-in paragraph styles, added to a document the way Word adds them on first use (plan 973 S3).

A built-in style (Caption, Heading 1, ...) has no definition in ``word/styles.xml`` until a document uses it;
Word then writes one. What Word writes depends on the document (the properties the style would inherit anyway
from Normal and the document defaults are left out, a toggle Normal turns on is written as off, a latent entry's
locked/semiHidden flags are carried over, a linked "... Char" style gets a free id), so the definition is not
modelled here: Word is asked, per document. ``plan`` builds a scratch package holding only the document's
styles, theme, settings that bear on styles (compatibility mode, theme font language, colour mapping) and
numbering, and a body of one plain paragraph (no revisions, comments or other content); a private Word
(``IsolatedWordDocument``; the real document is never opened) saves it once as is and once after applying the
built-in by its ``WdBuiltinStyle`` id (so a localized Word picks the same style), and the ``w:style`` elements the
second save has and the first lacks are what Word adds. Any other difference between the two saves (an existing
style Word changed, such as an unrelated "Heading 2 Char" it would link, ``w:latentStyles`` or ``w:docDefaults``)
refuses. References to existing styles (``w:basedOn``, ``w:next``, ``w:link``) are mapped back to the document's
own ids by style name, since Word may renumber ids on save (a localized default ``Standard`` saves as ``Normal``).
The direct-XML tools then splice exactly those elements into the real package's styles.xml (``write_staged``) and
check that nothing else in styles.xml changed (``verify_added``). Asking Word costs one Word start and two saves
of a one-paragraph document (about 3-6 seconds), only when a style is missing.

A built-in outside ``BUILTINS`` is not added; the tools keep refusing it.
"""

from __future__ import annotations

import re
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

from .errors import RefusedUnsafe

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W = "{" + W_NS + "}"
STYLES_PART = "word/styles.xml"
WORD_SECONDS = 60.0

# WdBuiltinStyle ids (Word object model); applied by id, never by the localized name.
BUILTINS: dict[str, int] = {
    "Caption": -35,
    "Heading 1": -2,
    "Heading 2": -3,
    "Heading 3": -4,
    "Heading 4": -5,
    "Title": -63,
    "List Paragraph": -180,
    "Quote": -181,
}

_STYLE_RE = re.compile(r"<w:style\b[^>]*?(?:/>|>.*?</w:style>)", re.S)
_STYLE_ID_RE = re.compile(r'w:styleId="([^"]*)"')
_NAME_RE = re.compile(r'<w:name w:val="([^"]*)"\s*/>')
_RSID_RE = re.compile(r"<w:rsid\b[^>]*/>")
_REF_RE = re.compile(r'<w:(basedOn|next|link) w:val="([^"]*)"\s*/>')
_PREFIX_RE = re.compile(r'(?:</?|\s)([A-Za-z_][\w.-]*):[\w.-]+')
_XMLNS_RE = re.compile(r'xmlns:([\w.-]+)="([^"]*)"')
# settings.xml children that bear on how Word writes a style definition; the rest (protection, mail merge,
# attached template, tracking, footnote separators, ...) stays out of the scratch package.
_SETTINGS_KEPT = ("defaultTabStop", "characterSpacingControl", "compat", "themeFontLang", "clrSchemeMapping",
                  "decimalSymbol", "listSeparator")
_REL_TYPES = {"theme": "theme/theme1.xml", "settings": "settings.xml", "numbering": "numbering.xml"}
_CONTENT_TYPES = {
    "document": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "styles": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    "settings": "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    "numbering": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
}


def _key(value: str) -> str:
    return "".join(value.split()).casefold()


def lookup(style: str) -> str | None:
    """The BUILTINS name `style` names (by name or styleId, case- and space-insensitive), else None."""

    wanted = _key(style or "")
    return next((name for name in BUILTINS if _key(name) == wanted), None)


@dataclass
class Addition:
    """What Word adds for the built-in `name` in one document: its w:style elements (styleId, w:name, xml), the
    paragraph style first, ready to splice into that document's styles.xml."""

    name: str
    style_id: str
    style_name: str
    definitions: list[tuple[str, str, str]]
    word_version: str = ""
    word_seconds: float = 0.0

    @property
    def names(self) -> list[str]:
        return [self.name if index == 0 else style_name
                for index, (_, style_name, _) in enumerate(self.definitions)]

    @property
    def block(self) -> str:
        return "".join(xml for _, _, xml in self.definitions)


@dataclass
class WordAnswer:
    """Word's two saves of the scratch package: as is, and with the built-in applied to its paragraph."""

    baseline_styles: bytes
    applied_styles: bytes
    applied_style_id: str | None  # the w:pStyle Word gave the paragraph
    version: str = ""
    seconds: float = 0.0


# --- the scratch package --------------------------------------------------------------------------------------


def _rels(archive: zipfile.ZipFile, part: str) -> list[ET.Element]:
    path = PurePosixPath(part)
    rels = str(path.parent / "_rels" / (path.name + ".rels"))
    if rels not in archive.namelist():
        return []
    return list(ET.fromstring(archive.read(rels)).iter(f"{{{PKG_REL_NS}}}Relationship"))


def _target(base: str, target: str) -> str:
    parts: list[str] = []
    for piece in (target.lstrip("/") if target.startswith("/") else f"{PurePosixPath(base).parent}/{target}").split("/"):
        if piece == "..":
            if parts:
                parts.pop()
        elif piece not in ("", "."):
            parts.append(piece)
    return "/".join(parts)


def _settings(xml: str) -> str:
    root = re.search(r"<w:settings\b[^>]*>", xml)
    if root is None:
        raise RefusedUnsafe("settings.xml does not bind the w: prefix on its root as Word writes it")
    kept = []
    for name in _SETTINGS_KEPT:
        for match in re.finditer(rf"<w:{name}\b(?:[^>]*/>|.*?</w:{name}>)", xml, re.S):
            kept.append((match.start(), match.group(0)))
    head = xml[:root.end()]
    return head + "".join(item for _, item in sorted(kept)) + "</w:settings>"


def write_scratch(source: str | Path, scratch: str | Path) -> None:
    """A package Word can open holding only `source`'s styles, theme, numbering and style-related settings,
    with one plain paragraph as its body."""

    parts: dict[str, bytes] = {}
    overrides: dict[str, str] = {"word/document.xml": _CONTENT_TYPES["document"],
                                 STYLES_PART: _CONTENT_TYPES["styles"]}
    relationships = [("styles", "styles.xml")]
    with zipfile.ZipFile(source) as archive:
        parts[STYLES_PART] = archive.read(STYLES_PART)
        names = set(archive.namelist())
        defaults = re.findall(r"<Default\b[^>]*/>", archive.read("[Content_Types].xml").decode("utf-8"))
        for rel in _rels(archive, "word/document.xml"):
            kind = rel.get("Type", "").rsplit("/", 1)[-1]
            if kind not in _REL_TYPES or rel.get("TargetMode") == "External":
                continue
            part = _target("word/document.xml", rel.get("Target", ""))
            if part not in names or any(target == _REL_TYPES[kind] for _, target in relationships):
                continue
            data = archive.read(part)
            name = "word/" + _REL_TYPES[kind]
            if kind == "settings":
                data = _settings(data.decode("utf-8")).encode("utf-8")
            elif kind == "numbering":
                # picture bullets: carry the numbering part's own relationships and their media along
                for inner in _rels(archive, part):
                    inner_part = _target(part, inner.get("Target", ""))
                    if inner.get("TargetMode") != "External" and inner_part in names:
                        parts[_target(name, inner.get("Target", ""))] = archive.read(inner_part)
                rels_name = str(PurePosixPath(part).parent / "_rels" / (PurePosixPath(part).name + ".rels"))
                if rels_name in names:
                    parts["word/_rels/numbering.xml.rels"] = archive.read(rels_name)
            parts[name] = data
            overrides[name] = _CONTENT_TYPES[kind]
            relationships.append((kind, _REL_TYPES[kind]))
    parts["word/document.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>Base paragraph.</w:t></w:r></w:p></w:body>'
        "</w:document>").encode("utf-8")
    parts["word/_rels/document.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        + "".join(f'<Relationship Id="rId{index}" Type="{R_NS}/{kind}" Target="{target}"/>'
                  for index, (kind, target) in enumerate(relationships, start=1))
        + "</Relationships>").encode("utf-8")
    parts["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" Type="{R_NS}/officeDocument" '
        'Target="word/document.xml"/></Relationships>').encode("utf-8")
    parts["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + "".join(defaults)
        + "".join(f'<Override PartName="/{name}" ContentType="{kind}"/>' for name, kind in overrides.items())
        + "</Types>").encode("utf-8")
    with zipfile.ZipFile(scratch, "w", zipfile.ZIP_DEFLATED) as writer:
        writer.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            writer.writestr(name, data)


def ask_word(scratch: Path, wd_id: int, folder: Path) -> WordAnswer:
    """Have a private Word save `scratch` as is, then with the built-in `wd_id` applied to its paragraph."""

    from .word import IsolatedWordDocument

    baseline, applied = folder / "baseline.docx", folder / "applied.docx"
    started = time.monotonic()
    with IsolatedWordDocument(scratch, timeout_seconds=WORD_SECONDS) as session:
        document = session.document
        version = f"{session.word.Version} build {session.word.Build}"
        # Both saves follow an edit, so their difference is the style alone, not Word's save noise.
        document.Paragraphs(1).Range.InsertAfter(" ")
        session.save_as(baseline)
        document.Paragraphs(1).Style = document.Styles(wd_id)
        document.Paragraphs(1).Range.InsertAfter(" ")
        session.save_as(applied)
    seconds = time.monotonic() - started
    with zipfile.ZipFile(applied) as archive:
        body = archive.read("word/document.xml").decode("utf-8")
        applied_styles = archive.read(STYLES_PART)
    with zipfile.ZipFile(baseline) as archive:
        baseline_styles = archive.read(STYLES_PART)
    style = re.search(r'<w:pStyle w:val="([^"]*)"', body)
    return WordAnswer(baseline_styles, applied_styles, style.group(1) if style else None, version, seconds)


# --- what Word added --------------------------------------------------------------------------------------------


def _styles(xml: str) -> dict[str, str]:
    return {_STYLE_ID_RE.search(item).group(1): _RSID_RE.sub("", item) for item in _STYLE_RE.findall(xml)}


def _style_name(xml: str) -> str:
    match = _NAME_RE.search(xml)
    return match.group(1) if match else _STYLE_ID_RE.search(xml).group(1)


def _namespaces(xml: str) -> dict[str, str]:
    root = re.search(r"<w:styles\b[^>]*>", xml)
    return dict(_XMLNS_RE.findall(root.group(0))) if root else {}


def _outside_styles(xml: str) -> dict[str, str]:
    """styles.xml outside its w:style elements, by top-level piece (latentStyles, docDefaults, the rest)."""

    rest = _STYLE_RE.sub("", xml)
    found = {}
    for name in ("latentStyles", "docDefaults"):
        match = re.search(rf"<w:{name}\b.*?</w:{name}>", rest, re.S)
        found[name] = match.group(0) if match else ""
        rest = rest.replace(found[name], "")
    found["the rest of styles.xml"] = rest
    return found


def extract(source_styles: bytes, answer: WordAnswer, name: str) -> Addition:
    """The w:style elements Word added for `name`, with references mapped to `source_styles`' own ids.
    Refuses when Word changed anything else, or when what it added cannot be spliced into this styles.xml."""

    source_xml = source_styles.decode("utf-8")
    base_xml, applied_xml = answer.baseline_styles.decode("utf-8"), answer.applied_styles.decode("utf-8")
    base, applied = _styles(base_xml), _styles(applied_xml)
    source = _styles(source_xml)
    source_by_name = {_key(_style_name(xml)): style_id for style_id, xml in source.items()}

    def refuse(why: str) -> RefusedUnsafe:
        return RefusedUnsafe(f"style {name!r} is a Word built-in this document has not defined, and Word {why}; "
                             f"apply {name!r} once in Word, then insert again")

    changed = [_style_name(xml) for style_id, xml in base.items() if applied.get(style_id, xml) != xml]
    gone = [_style_name(xml) for style_id, xml in base.items() if style_id not in applied]
    if changed or gone:
        raise refuse("would also change the document's existing style(s) "
                     + ", ".join(repr(item) for item in changed + gone)
                     + " (for a linked style: a style of that name that is not linked to it)")
    outside_before, outside_after = _outside_styles(base_xml), _outside_styles(applied_xml)
    differs = [piece for piece in outside_before if outside_before[piece] != outside_after[piece]]
    if differs:
        raise refuse("would also change " + ", ".join(f"w:{piece}" if " " not in piece else piece
                                                       for piece in differs))
    added = [style_id for style_id in applied if style_id not in base]
    main = answer.applied_style_id
    if main not in added:
        raise refuse(f"gave the paragraph the style {main!r} without adding a definition")
    added.remove(main)
    added.insert(0, main)

    base_names = {style_id: _style_name(xml) for style_id, xml in base.items()}
    definitions = []
    for style_id in added:
        xml = applied[style_id]
        style_name = _style_name(xml)
        occupant = source.get(style_id) or source.get(source_by_name.get(_key(style_name), "\0"))
        if occupant is not None:
            raise refuse(f"would add the style {style_name!r} (id {style_id!r}), which clashes with the document's "
                         f"style {_style_name(occupant)!r} (id {_STYLE_ID_RE.search(occupant).group(1)!r})")
        if "<w:numPr" in xml:
            raise refuse(f"would number {style_name!r} with the document's lists")

        def remap(match: re.Match) -> str:
            kind, target = match.group(1), match.group(2)
            if target in added:
                return match.group(0)
            mapped = source_by_name.get(_key(base_names.get(target, "\0")))
            if mapped is None:
                raise refuse(f"would point {style_name!r} ({kind}) at {target!r}, which this document lacks")
            return f'<w:{kind} w:val="{mapped}"/>'

        xml = _REF_RE.sub(remap, xml)
        definitions.append((style_id, style_name, xml))

    by_id = {style_id: xml for style_id, _, xml in definitions}
    for style_id, style_name, xml in definitions:
        link = re.search(r'<w:link w:val="([^"]*)"', xml)
        if link is None:
            continue
        partner = by_id.get(link.group(1), source.get(link.group(1)))
        back = re.search(r'<w:link w:val="([^"]*)"', partner or "")
        if back is None or back.group(1) != style_id:
            raise refuse(f"would link {style_name!r} to the document's style "
                         f"{_style_name(partner) if partner else link.group(1)!r}, which is not linked back to it")

    wanted, have = _namespaces(applied_xml), _namespaces(source_xml)
    for prefix in sorted(set(_PREFIX_RE.findall("".join(by_id.values()))) - {"xml"}):
        if wanted.get(prefix) is None or have.get(prefix) != wanted[prefix]:
            raise refuse(f"writes the prefix {prefix!r}, which this styles.xml does not bind the same way")
    return Addition(name, main, definitions[0][1], definitions, answer.version, answer.seconds)


# --- writing and checking a document -------------------------------------------------------------------------


def _read(path: str | Path, part: str) -> bytes | None:
    with zipfile.ZipFile(path) as archive:
        return archive.read(part) if part in archive.namelist() else None


def _splice(styles: bytes, block: str) -> bytes:
    text = styles.decode("utf-8")
    root_tag = re.search(r"<w:styles\b[^>]*>", text)
    if root_tag is None or f'xmlns:w="{W_NS}"' not in root_tag.group(0):
        raise RefusedUnsafe("styles.xml does not bind the w: prefix on its root as Word writes it; apply the style "
                            "once in Word, then insert again")
    end = text.rindex("</w:styles>")
    return (text[:end] + block + text[end:]).encode("utf-8")


def plan(source: str | Path, style: str) -> Addition | None:
    """What Word adds to `source` for the built-in `style` names, asked of a private Word on a scratch package;
    None when `style` is not one of BUILTINS."""

    name = lookup(style)
    if name is None:
        return None
    styles = _read(source, STYLES_PART)
    if styles is None:
        raise RefusedUnsafe(f"this document has no styles part to add {name!r} to; apply it once in Word")
    _splice(styles, "")  # refuses a styles.xml the definitions could not be spliced into, before Word starts
    with tempfile.TemporaryDirectory(prefix="research-mcp-builtin-style-") as folder:
        scratch = Path(folder) / "scratch.docx"
        write_scratch(source, scratch)
        answer = ask_word(scratch, BUILTINS[name], Path(folder))
    return extract(styles, answer, name)


def write_staged(source: str | Path, staged: str | Path, addition: Addition) -> None:
    """Copy `source` to `staged` with the addition at the end of styles.xml; every other entry byte-identical."""

    with zipfile.ZipFile(source) as reader, zipfile.ZipFile(staged, "w") as writer:
        for info in reader.infolist():
            data = reader.read(info.filename)
            if info.filename == STYLES_PART:
                data = _splice(data, addition.block)
            writer.writestr(info, data, compress_type=info.compress_type)


def verify_added(source: str | Path, output: str | Path, addition: Addition) -> None:
    """Refuse unless the output's styles.xml is the source's plus exactly the added definitions."""

    from .errors import VerificationFailed

    before, after = _read(source, STYLES_PART), _read(output, STYLES_PART)
    if after != _splice(before, addition.block):
        raise VerificationFailed("styles.xml changed beyond the added built-in style definitions",
                                 expected=addition.names, actual="other changes in word/styles.xml")
    ids = [item.get(W + "styleId") for item in ET.fromstring(after).findall(W + "style")]
    if any(ids.count(style_id) != 1 for style_id, _, _ in addition.definitions):
        raise VerificationFailed("an added style id is not unique in styles.xml",
                                 expected=[item[0] for item in addition.definitions], actual=ids)


def run_with_style(source: str | Path, addition: Addition | None, operation):
    """Run `operation(staged_source)` on a copy of `source` that defines the built-in, check the output, and
    report the original source and `styles_added`. With no addition it runs on `source` itself."""

    from .audit import file_sha256

    if addition is None:
        result = operation(Path(source))
        result["styles_added"] = []
        return result
    with tempfile.TemporaryDirectory(prefix="research-mcp-styles-") as folder:
        staged = Path(folder) / "staged.docx"
        write_staged(source, staged, addition)
        result = operation(staged)
    try:
        verify_added(source, result["output"], addition)
    except Exception:
        Path(result["output"]).unlink(missing_ok=True)
        raise
    result["source"] = str(Path(source).resolve())
    result["source_sha256"] = file_sha256(source)
    result["styles_added"] = addition.names
    return result
