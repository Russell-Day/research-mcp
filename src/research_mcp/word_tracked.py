"""Plan 916 S4: apply_document_style and format_tables recorded with Word's own Track Changes.

While the research-mcp tracking switch is on (or the tool is called with ``tracked=True``) both tools make their
existing Word pass with Word's ``TrackRevisions`` on, in the isolated document only, and restore the document's
own setting before the save. After the save, in the parent process:

1. ``relabel`` sets ``w:author`` on every revision Word recorded to the switch author (the source is revision-free,
   so every revision is new), and puts ``word/people.xml`` back as the source had it, since Word adds the account
   of the user it runs as there. ``Application.UserName`` is never changed.
2. ``untracked_notes`` lists what the pass changed without a revision: cell merges (``format_tables`` makes them
   with tracking off, because Word records a tracked merge as deleted and retyped text), page margins of a section
   with no ``w:sectPrChange``, header/footer text changed in a part with no revision.
3. ``add_untracked_comments`` adds one comment per note by the switch author (``comment_ops.add_comment``, direct
   XML), anchored on the whole text of the first paragraph with text at that place.
4. ``run_word_check`` (worker kind ``tracked_check``) has Word save the input as is (baseline), Accept All on the
   output before the comments, Reject All on the final output, and open the final output. Reject All must give
   the baseline's formatting (``formatting_snapshot``: per paragraph its properties and per character its run
   properties, table/row/cell properties and grid, section properties, header/footer text) apart from the places
   the comments name (paragraph properties by effective value: an explicit zero indent or spacing, or an off
   keepNext-type toggle, that the paragraph inherits anyway counts as absent, since Word drops those when it
   records a change and Reject All does not write them back); Accept All must pass the tool's own verification (today's untracked result); no revision
   may survive either. Anything else refuses and nothing is published.
"""

from __future__ import annotations

import copy
import os
import re
import shutil
import tempfile
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Sequence
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr

from .errors import RefusedUnsafe, VerificationFailed
from .schema_models import FormatTablesResult

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PEOPLE_PART = "word/people.xml"
REVISION_TAGS = ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "tblPrChange", "trPrChange",
                 "tcPrChange", "sectPrChange", "tblGridChange", "tblPrExChange", "numberingChange", "cellIns",
                 "cellDel", "cellMerge")
# Move range markers carry an author too (Word's formatting passes write none; relabelled for completeness).
AUTHORED_TAGS = REVISION_TAGS + ("moveFromRangeStart", "moveToRangeStart")
# Start tags as Word writes them (prefix w, attributes in any order, either quote); relabel then re-reads every
# part with the XML parser and refuses if any revision still has another author, so a miss can't pass silently.
_REV_START = re.compile(r"<w:(?:" + "|".join(AUTHORED_TAGS) + r")(?=[\s/>])[^>]*>")
_AUTHOR_ATTR = re.compile(r'\bw:author\s*=\s*("[^"]*"|\'[^\']*\')')
TWIPS_PER_INCH = 1440
SIDES = ("top", "bottom", "left", "right")


class TrackedFormatTablesResult(FormatTablesResult):
    """format_tables' result with the tracking fields (plan 916 S4)."""

    tracking: str = "off"
    tracking_author: str | None = None
    revisions_added: int = 0
    untracked: list[dict] = []
    tracked_check: dict | None = None
    compared_against: str | None = None


# -- which mode ----------------------------------------------------------------------------------------------

def tracking_for(source: str | Path, tracked: bool | None, author: str | None = None) -> tuple[bool, str]:
    """(on, author): ``tracked`` wins over the document's switch; the author is the call's, else the switch's,
    else "Claude"."""

    from .track_setting import DEFAULT_AUTHOR, read_setting

    setting = read_setting(source)
    on = setting.on if tracked is None else bool(tracked)
    chosen = (author or "").strip() or (setting.author if setting.is_set else "") or DEFAULT_AUTHOR
    return on, chosen


REVISED_HINT = (
    "; for formatting on a document with tracked changes use style_text (bold, italic) or format_paragraph (font, "
    "size, indents, spacing, numbering), which work as direct XML and record tracked formatting changes"
)


def set_word_tracking(document: Any, on: bool) -> bool:
    """Set Word's TrackRevisions on the isolated document; returns the previous value to restore before the save."""

    previous = bool(document.TrackRevisions)
    document.TrackRevisions = bool(on)
    return previous


# -- relabel -------------------------------------------------------------------------------------------------

def _read_parts(path: Path) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        return infos, {info.filename: archive.read(info.filename) for info in infos}


def _write_parts(path: Path, infos: Sequence[zipfile.ZipInfo], parts: dict[str, bytes]) -> None:
    temp = path.with_name(f".{path.stem}.relabel-{os.getpid()}.tmp")
    with zipfile.ZipFile(temp, "w") as archive:
        for info in infos:
            if info.filename in parts:
                archive.writestr(info, parts[info.filename], compress_type=info.compress_type)
    os.replace(temp, path)


def _relabel_xml(xml: str, author: str) -> tuple[str, int]:
    value = quoteattr(author)
    count = 0

    def tag(match: re.Match) -> str:
        nonlocal count
        start = match.group(0)
        if not _AUTHOR_ATTR.search(start):
            return start
        count += 1
        return _AUTHOR_ATTR.sub(lambda _m: f"w:author={value}", start, count=1)

    return _REV_START.sub(tag, xml), count


def _drop_people_references(parts: dict[str, bytes]) -> None:
    """Remove people.xml's relationship and content-type override (text edits; every other byte kept)."""

    rels = "word/_rels/document.xml.rels"
    if rels in parts:
        parts[rels] = re.sub(rb"<Relationship\b[^>]*\bTarget=\"(?:/word/)?people\.xml\"[^>]*/>", b"", parts[rels])
    types = "[Content_Types].xml"
    if types in parts:
        parts[types] = re.sub(rb"<Override\b[^>]*\bPartName=\"/word/people\.xml\"[^>]*/>", b"", parts[types])


def relabel(path: str | Path, author: str, source: str | Path) -> int:
    """Every revision's ``w:author`` becomes ``author``; ``word/people.xml`` goes back to the source's state.

    The caller guarantees ``source`` held no revision, so every revision in ``path`` is one Word recorded for
    this pass. Returns the number of revision elements in ``path`` (those with and without an author)."""

    from .revisions import revision_count

    path, source = Path(path), Path(source)
    if revision_count(source):
        raise RefusedUnsafe("relabel needs a revision-free source; refusing to relabel existing revisions")
    before = revision_count(path)
    infos, parts = _read_parts(path)
    with zipfile.ZipFile(source) as archive:
        source_people = archive.read(PEOPLE_PART) if PEOPLE_PART in archive.namelist() else None
    for name in list(parts):
        if name.startswith("word/") and name.endswith(".xml") and name != PEOPLE_PART:
            text = parts[name].decode("utf-8")
            new, count = _relabel_xml(text, author)
            if count:
                parts[name] = new.encode("utf-8")
    if PEOPLE_PART in parts:
        if source_people is not None:
            parts[PEOPLE_PART] = source_people
        else:
            del parts[PEOPLE_PART]
            _drop_people_references(parts)
            infos = [info for info in infos if info.filename != PEOPLE_PART]
    _write_parts(path, infos, parts)
    if revision_count(path) != before:
        raise VerificationFailed("relabelling Word's revisions changed their count", expected=before,
                                 actual=revision_count(path))
    others = sorted(revision_authors(path) - {author})
    if others:
        raise VerificationFailed("revisions by another author remain after relabelling", expected=[author],
                                 actual=others)
    return before


def revision_authors(path: str | Path) -> set[str]:
    authors: set[str] = set()
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                root = ET.fromstring(archive.read(name))
                for element in root.iter():
                    if element.tag.startswith(W) and element.tag[len(W):] in AUTHORED_TAGS:
                        if element.get(W + "author") is not None:
                            authors.add(element.get(W + "author"))
    return authors


def revision_kinds(path: str | Path, *, by_part: bool = False) -> dict[str, int]:
    """Revision elements by kind (``"kind in part"`` with ``by_part``)."""

    kinds: dict[str, int] = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                root = ET.fromstring(archive.read(name))
                for element in root.iter():
                    if element.tag.startswith(W) and element.tag[len(W):] in REVISION_TAGS:
                        kind = element.tag[len(W):] + (f" in {name}" if by_part else "")
                        kinds[kind] = kinds.get(kind, 0) + 1
    return dict(sorted(kinds.items()))


# -- formatting snapshot -------------------------------------------------------------------------------------

_DROP_ATTR = re.compile(r"\}(rsid\w*|paraId|textId)$")
_DROP_CHILD = {W + tag for tag in ("rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange",
                                   "sectPrChange", "tblGridChange", "tblPrExChange", "numberingChange", "ins", "del",
                                   "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge")}


def canonical(element: ET.Element | None, *, drop: Iterable[str] = ()) -> str:
    """Order-insensitive text of a properties element without revision records and rsids ("" for none/empty)."""

    if element is None:
        return ""
    dropped = _DROP_CHILD | {W + tag for tag in drop}

    def walk(node: ET.Element) -> str:
        attrs = sorted((key, value) for key, value in node.attrib.items() if not _DROP_ATTR.search(key))
        children = sorted(walk(child) for child in node if child.tag not in dropped)
        tag = node.tag.replace(W, "w:")
        return f"<{tag}{''.join(f' {k}={v!r}' for k, v in attrs)}>{''.join(children)}</{tag}>"

    children = sorted(walk(child) for child in element if child.tag not in dropped)
    return "".join(child for child in children if not re.fullmatch(r"<w:rPr></w:rPr>", child))


def _paragraph(node: ET.Element, view: str = "accepted",
               effective: "_View | None" = None) -> tuple[str, tuple[tuple[str, str], ...]]:
    """(canonical pPr, ((char, canonical rPr), ...)). ``view="rejected"`` reads the paragraph as Reject All would
    leave it: deleted text in, inserted text out, each run's and the paragraph's old properties where a tracked
    formatting change holds them. ``effective`` (optional) reads the chosen pPr and each rPr by effective value
    (``_View``)."""

    rejected = view == "rejected"
    ppr = node.find(W + "pPr")
    if rejected and ppr is not None and ppr.find(f"{W}pPrChange/{W}pPr") is not None:
        ppr = ppr.find(f"{W}pPrChange/{W}pPr")
    chars: list[tuple[str, str]] = []
    skipped = (W + "ins", W + "moveTo") if rejected else (W + "del", W + "moveFrom")

    def runs(parent: ET.Element) -> None:
        for child in parent:
            tag = child.tag
            if tag in skipped:
                continue
            if tag == W + "r":
                rpr_node = child.find(W + "rPr")
                if rejected and rpr_node is not None and rpr_node.find(f"{W}rPrChange/{W}rPr") is not None:
                    rpr_node = rpr_node.find(f"{W}rPrChange/{W}rPr")
                rpr = canonical(effective.rpr(ppr, rpr_node) if effective is not None else rpr_node)
                for item in child:
                    if item.tag == W + "t" or (rejected and item.tag == W + "delText"):
                        chars.extend((ch, rpr) for ch in item.text or "")
                    elif item.tag == W + "tab":
                        chars.append(("\t", rpr))
                    elif item.tag in (W + "br", W + "cr"):
                        chars.append(("\n", rpr))
            elif tag not in (W + "pPr",):
                runs(child)

    runs(node)
    if effective is not None:
        ppr = effective.ppr(ppr)
    return canonical(ppr, drop=("sectPr",)), tuple(chars)


# Explicit values that change nothing on the page: Word drops them when it records a formatting change and does
# not write them back on Reject All (measured, plan 916 S6: a table whose caption and cells carried w:ind 0,
# w:spacing before/after 0 with autospacing 0, w:b w:val="0" and the document's default font in w:rFonts). The
# Reject All check reads past them with ``_View``; every other property is compared as written.
_ZERO_ATTRS = {"ind": ("left", "start", "right", "end", "firstLine", "hanging"),
               "spacing": ("before", "after", "beforeAutospacing", "afterAutospacing")}
_PARA_TOGGLES = ("keepNext", "keepLines", "pageBreakBefore", "contextualSpacing")
_RUN_TOGGLES = ("b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "vanish")
# Run properties compared by attribute against the last layer that sets the same attribute.
_RUN_ATTRS = {"rFonts": ("ascii", "hAnsi", "cs", "eastAsia"), "sz": ("val",), "szCs": ("val",)}
_OFF = {"0", "false", "off"}


def _twips(value: str | None) -> int:
    try:
        return int(float(value)) if value is not None else 0
    except ValueError:
        return 0


def _toggle_on(node: ET.Element) -> bool:
    return (node.get(W + "val") or "1").lower() not in _OFF


def _effective_values(layers: Sequence[ET.Element | None]) -> dict[str, Any]:
    """What the zero-default paragraph properties come to over ``layers`` (inherited first, direct last)."""

    values: dict[str, Any] = {"left": 0, "right": 0, "first": 0, "before": 0, "after": 0,
                              "beforeAutospacing": False, "afterAutospacing": False}
    values.update(dict.fromkeys(_PARA_TOGGLES, False))
    for layer in layers:
        if layer is None:
            continue
        ind = layer.find(W + "ind")
        if ind is not None:
            for key, names in (("left", ("start", "left")), ("right", ("end", "right"))):
                found = next((ind.get(W + name) for name in names if ind.get(W + name) is not None), None)
                if found is not None:
                    values[key] = _twips(found)
            if ind.get(W + "hanging") is not None:
                values["first"] = -_twips(ind.get(W + "hanging"))
            elif ind.get(W + "firstLine") is not None:
                values["first"] = _twips(ind.get(W + "firstLine"))
        spacing = layer.find(W + "spacing")
        if spacing is not None:
            for key in ("before", "after"):
                if spacing.get(W + key) is not None:
                    values[key] = _twips(spacing.get(W + key))
            for key in ("beforeAutospacing", "afterAutospacing"):
                if spacing.get(W + key) is not None:
                    values[key] = spacing.get(W + key).lower() not in _OFF
        for toggle in _PARA_TOGGLES:
            node = layer.find(W + toggle)
            if node is not None:
                values[toggle] = _toggle_on(node)
    return values


def _normalize_toggles(element: ET.Element, toggles: Sequence[str]) -> None:
    """``<w:b w:val="1"/>`` / ``"true"`` / ``"on"`` -> ``<w:b/>`` (the same property, as Word rewrites it)."""

    for toggle in toggles:
        node = element.find(W + toggle)
        if node is not None and _toggle_on(node) and node.get(W + "val") is not None:
            del node.attrib[W + "val"]


def effective_ppr(ppr: ET.Element | None, inherited: list[ET.Element | None]) -> ET.Element | None:
    """A copy of ``ppr`` without the explicit zero / off values (``_ZERO_ATTRS``, ``_PARA_TOGGLES``) that leave the
    paragraph's effective values as they are; every other property is kept as written."""

    if ppr is None:
        return None
    result = copy.deepcopy(ppr)
    _normalize_toggles(result, _PARA_TOGGLES)
    target = _effective_values(inherited + [result])
    for tag, attrs in _ZERO_ATTRS.items():
        node = result.find(W + tag)
        if node is None:
            continue
        for attr in attrs:
            value = node.get(W + attr)
            if value is None:
                continue
            del node.attrib[W + attr]
            if _effective_values(inherited + [result]) != target:
                node.set(W + attr, value)
        if not node.attrib and len(node) == 0:
            result.remove(node)
    for toggle in _PARA_TOGGLES:
        node = result.find(W + toggle)
        if node is None:
            continue
        index = list(result).index(node)
        result.remove(node)
        if _effective_values(inherited + [result]) != target:
            result.insert(index, node)
    return result


def effective_rpr(rpr: ET.Element | None, inherited: list[ET.Element | None]) -> ET.Element | None:
    """A copy of ``rpr`` without what the run inherits anyway: a toggle (``_RUN_TOGGLES``) set off where no
    inherited layer sets it, and an ``_RUN_ATTRS`` attribute equal to the last inherited layer that sets it (fonts
    only when no layer, the run's included, names a theme font), and a ``w:color`` whose attributes all equal the
    last inherited color's. Every other property is kept as written."""

    if rpr is None:
        return None
    result = copy.deepcopy(rpr)
    _normalize_toggles(result, _RUN_TOGGLES)
    layers = [layer for layer in inherited if layer is not None]
    for toggle in _RUN_TOGGLES:
        node = result.find(W + toggle)
        if node is not None and not _toggle_on(node) and all(layer.find(W + toggle) is None for layer in layers):
            result.remove(node)
    themed = any(any(key.endswith("Theme") for key in node.attrib)
                 for layer in layers + [result] for node in layer.findall(W + "rFonts"))
    for tag, attrs in _RUN_ATTRS.items():
        node = result.find(W + tag)
        if node is None or (tag == "rFonts" and themed):
            continue
        for attr in attrs:
            value = node.get(W + attr)
            if value is None:
                continue
            inherited_value = None
            for layer in layers:
                found = layer.find(W + tag)
                if found is not None and found.get(W + attr) is not None:
                    inherited_value = found.get(W + attr)
            if inherited_value is not None and inherited_value.lower() == value.lower():
                del node.attrib[W + attr]
        if not node.attrib and len(node) == 0:
            result.remove(node)
    # A color is inherited only as a whole: w:val is a theme color's cached RGB when w:themeColor (tint, shade) is
    # set, so a direct RGB equal to that cache is still a different color (plan 916 S4 review r3).
    color = result.find(W + "color")
    inherited_color = next((layer.find(W + "color") for layer in reversed(layers)
                            if layer.find(W + "color") is not None), None)
    if color is not None and inherited_color is not None and len(color) == 0 and _attrs(color) == _attrs(
            inherited_color):
        result.remove(color)
    return result


def _attrs(node: ET.Element) -> dict[str, str]:
    return {key: value.lower() for key, value in node.attrib.items()}


class _Inherited:
    """The property layers a paragraph or run inherits (docDefaults, its table's style, its paragraph style chain,
    its list level; for a run also its character style chain), read from one package's styles and numbering."""

    def __init__(self, parts: dict[str, bytes]) -> None:
        from .paragraph_format import Styles

        self.styles = Styles(parts)

    def view(self, table_style: str | None = None, in_table: bool = False) -> "_View | None":
        """The view for a body paragraph, or for a cell paragraph of a table with ``table_style``; None (compare as
        written) when a conditional format (w:tblStylePr) of that style sets paragraph or run properties, since
        which of them applies to a cell is not modelled here."""

        chain = self.styles.chain(table_style) if in_table else []
        if any(style.find(f"{W}tblStylePr/{W}pPr") is not None or style.find(f"{W}tblStylePr/{W}rPr") is not None
               for style in chain):
            return None
        return _View(self, [style.find(W + "pPr") for style in reversed(chain)],
                     [style.find(W + "rPr") for style in reversed(chain)])


@dataclass
class _View:
    inherited: _Inherited
    table_ppr: list[ET.Element | None]
    table_rpr: list[ET.Element | None]

    def ppr(self, ppr: ET.Element | None) -> ET.Element | None:
        styles = self.inherited.styles
        level = styles._level(styles.numbering(ppr))
        layers = ([styles.ppr_default] + self.table_ppr
                  + [style.find(W + "pPr") for style in reversed(styles.chain(styles._para_style(ppr)))]
                  + [level.find(W + "pPr") if level is not None else None])
        return effective_ppr(ppr, layers)

    def rpr(self, ppr: ET.Element | None, rpr: ET.Element | None) -> ET.Element | None:
        styles = self.inherited.styles
        layers = ([styles.rpr_default] + self.table_rpr
                  + [style.find(W + "rPr") for style in reversed(styles.chain(styles._para_style(ppr)))])
        rstyle = rpr.find(W + "rStyle") if rpr is not None else None
        if rstyle is not None:
            layers += [style.find(W + "rPr") for style in reversed(styles.chain(rstyle.get(W + "val")))]
        return effective_rpr(rpr, layers)


def _section(sect: ET.Element, targets: dict[str, str], parts: dict[str, bytes], number: int) -> dict:
    out: dict = {}
    pg_mar = sect.find(W + "pgMar")
    out[("sect", number, "pgMar")] = canonical(pg_mar) if pg_mar is None else tuple(
        sorted((key.replace(W, "w:"), value) for key, value in pg_mar.attrib.items()))
    out[("sect", number, "props")] = canonical(sect, drop=("headerReference", "footerReference", "pgMar"))
    for kind in ("header", "footer"):
        for ref in sect.findall(W + kind + "Reference"):
            target = targets.get(ref.get("{" + R_NS + "}id") or "")
            text = part_text(parts[target]) if target in parts else None
            out[("sect", number, kind, ref.get(W + "type") or "default")] = text
    return out


def part_text(data: bytes) -> str:
    """Visible text of a header/footer part: paragraphs joined by newlines, deleted text left out."""

    root = ET.fromstring(data)
    lines = []
    for node in root.iter(W + "p"):
        lines.append("".join(ch for ch, _ in _paragraph(node)[1]))
    return "\n".join(lines).strip()


def _rel_targets(parts: dict[str, bytes]) -> dict[str, str]:
    rels = parts.get("word/_rels/document.xml.rels")
    targets: dict[str, str] = {}
    if rels:
        for rel in ET.fromstring(rels):
            target = rel.get("Target") or ""
            if rel.get("TargetMode") == "External":
                continue
            target = target.lstrip("/") if target.startswith("/") else "word/" + target
            targets[rel.get("Id") or ""] = target
    return targets


def formatting_snapshot(path: str | Path, view: str = "accepted", *, effective: bool = False) -> dict:
    """Formatting of every paragraph, table and section, keyed by place (Word renumbers nothing here).

    Keys: ``("p", i)`` body paragraph, ``("t", path, "cell"|"row"|"tbl"|"grid")`` table parts, ``("t", path, "p",
    k)`` a paragraph of a cell (path = table number, then row/cell numbers, nested tables appended), ``("sect", n,
    "pgMar"|"props"|"header"/"footer", type)``. A paragraph value is (canonical pPr, ((char, canonical rPr), ...)),
    read as ``view`` ("accepted", or "rejected": see ``_paragraph``); table and section values are as saved.
    ``effective=True`` reads each paragraph's properties through ``effective_ppr`` (explicit zero / off values the
    paragraph would inherit anyway are left out)."""

    _infos, parts = _read_parts(Path(path))
    root = ET.fromstring(parts["word/document.xml"])
    body = root.find(W + "body")
    targets = _rel_targets(parts)
    snapshot: dict = {}
    sections = 0

    inherited = _Inherited(parts) if effective else None

    def table(node: ET.Element, place: tuple) -> None:
        style = node.find(f"{W}tblPr/{W}tblStyle")
        cell_view = (inherited.view(style.get(W + "val") if style is not None else None, in_table=True)
                     if inherited is not None else None)
        snapshot[("t", place, "tbl")] = canonical(node.find(W + "tblPr"))
        snapshot[("t", place, "grid")] = canonical(node.find(W + "tblGrid"))
        for r, row in enumerate(node.findall(W + "tr"), start=1):
            snapshot[("t", place + (r,), "row")] = canonical(row.find(W + "trPr"))
            snapshot[("t", place + (r,), "tblPrEx")] = canonical(row.find(W + "tblPrEx"))
            for c, cell in enumerate(row.findall(W + "tc"), start=1):
                here = place + (r, c)
                snapshot[("t", here, "cell")] = canonical(cell.find(W + "tcPr"))
                k = 0
                for item in cell:
                    if item.tag == W + "p":
                        k += 1
                        snapshot[("t", here, "p", k)] = _paragraph(item, view, cell_view)
                    elif item.tag == W + "tbl":
                        k += 1
                        table(item, here + (k,))

    body_view = inherited.view() if inherited is not None else None
    t = i = 0
    for child in body:
        if child.tag == W + "p":
            i += 1
            snapshot[("p", i)] = _paragraph(child, view, body_view)
            sect = child.find(f"{W}pPr/{W}sectPr")
            if sect is not None:
                sections += 1
                snapshot.update(_section(sect, targets, parts, sections))
        elif child.tag == W + "tbl":
            t += 1
            table(child, (t,))
        elif child.tag == W + "sectPr":
            sections += 1
            snapshot.update(_section(child, targets, parts, sections))
    return snapshot


def differences(expected: dict, actual: dict, excluded: Callable[[tuple], bool],
                normalize: Callable[[tuple, Any], Any] | None = None) -> list[tuple]:
    """Keys whose values differ, ``excluded`` keys aside; ``normalize(key, value)`` (optional) maps both sides
    before they are compared."""

    same = (lambda key, value: value) if normalize is None else normalize
    return [key for key in sorted(set(expected) | set(actual), key=repr)
            if not excluded(key) and same(key, expected.get(key)) != same(key, actual.get(key))]


def describe_difference(key: tuple, expected: dict, actual: dict) -> str:
    """Where and what differs, without the document's text (characters are counted, not quoted)."""

    old, new = expected.get(key), actual.get(key)
    if key[0] == "p" or (key[0] == "t" and len(key) == 4 and key[2] == "p"):
        where = f"body paragraph {key[1]}" if key[0] == "p" else f"table {key[1]} paragraph {key[3]}"
        if old is None or new is None:
            return f"{where} missing on one side"
        if old[0] != new[0]:
            return f"{where} paragraph properties {old[0][:160]!r} -> {new[0][:160]!r}"
        if "".join(ch for ch, _ in old[1]) != "".join(ch for ch, _ in new[1]):
            return f"{where} text differs ({len(old[1])} -> {len(new[1])} characters)"
        for n, ((_, a), (_, b)) in enumerate(zip(old[1], new[1]), start=1):
            if a != b:
                return f"{where} character {n} run properties {a[:160]!r} -> {b[:160]!r}"
    if key[0] == "sect" and key[2] in ("header", "footer"):
        return f"section {key[1]} {key[2]} ({key[3]}) text differs"
    return f"{key!r}: {str(old)[:160]!r} -> {str(new)[:160]!r}"


# -- untracked notes -----------------------------------------------------------------------------------------

@dataclass
class Note:
    kind: str  # "margins" | "merge" | "header" | "footer"
    text: str
    section: int | None = None
    cell: tuple[int, int, int] | None = None  # (table, row, column), 1-based, the first merged cell
    rows: tuple[int, int] | None = None
    excluded: list[tuple] = field(default_factory=list)  # snapshot key prefixes the Reject check skips
    # snapshot keys compared without the merge itself (w:vMerge): the merged cells' properties
    merge_keys: list[tuple] = field(default_factory=list)

    def covers(self, key: tuple) -> bool:
        return any(key[:len(prefix)] == prefix for prefix in self.excluded)

    def normalize(self, key: tuple, value: Any) -> Any:
        if key in self.merge_keys and isinstance(value, str):
            return _VMERGE.sub("", value)
        return value


_VMERGE = re.compile(r"<w:vMerge\b[^>]*></w:vMerge>")


def _inches(twips: str | None) -> str:
    try:
        return f'{int(twips or 0) / TWIPS_PER_INCH:g}"'
    except ValueError:
        return f"{twips} twips"


def _sections(xml: bytes) -> list[ET.Element]:
    body = ET.fromstring(xml).find(W + "body")
    found = []
    for child in body:
        if child.tag == W + "p":
            sect = child.find(f"{W}pPr/{W}sectPr")
            if sect is not None:
                found.append(sect)
        elif child.tag == W + "sectPr":
            found.append(child)
    return found


def margin_note(number: int, old: ET.Element | None, new: ET.Element | None) -> Note | None:
    before = {side: (old.get(W + side) if old is not None else None) for side in SIDES}
    after = {side: (new.get(W + side) if new is not None else None) for side in SIDES}
    if before == after:
        return None
    if len(set(before.values())) == 1 and len(set(after.values())) == 1:
        change = f"from {_inches(before['top'])} to {_inches(after['top'])}"
    else:
        change = "(" + ", ".join(f"{side} {_inches(before[side])} to {_inches(after[side])}"
                                 for side in SIDES if before[side] != after[side]) + ")"
    return Note("margins", f"Page margins changed {change} (not tracked: Word can't track page setup)",
                section=number, excluded=[("sect", number, "pgMar")])


def untracked_notes(source: str | Path, output: str | Path, merges: Sequence[tuple[int, Any]] = ()) -> list[Note]:
    """What the pass changed without a revision: merges (from the specs), margins of a section with no
    sectPrChange, header/footer text in a part with no revision."""

    notes: list[Note] = []
    for table_index, group in merges:
        # Only the merge itself is untracked: the merged cells' w:vMerge, and the label the cells below the first
        # no longer hold (their paragraphs). The first cell's paragraphs (text, run and paragraph properties) and
        # every other cell property stay in the Reject check; merged_cell_errors checks the first cell's text and
        # character formatting and that the cells below hold no text.
        rows = (group.start_row, group.end_row)
        notes.append(Note(
            "merge", f"Cells in rows {rows[0]}-{rows[1]} of column {group.column} of table {table_index} were merged "
            "(not tracked: Word can't track cell merges)",
            cell=(table_index, rows[0], group.column), rows=rows,
            excluded=[("t", (table_index, r, group.column), "p") for r in range(rows[0] + 1, rows[1] + 1)],
            merge_keys=[("t", (table_index, r, group.column), "cell") for r in range(rows[0], rows[1] + 1)]))
    _i, src = _read_parts(Path(source))
    _o, out = _read_parts(Path(output))
    old_sections, new_sections = _sections(src["word/document.xml"]), _sections(out["word/document.xml"])
    for number, (old, new) in enumerate(zip(old_sections, new_sections), start=1):
        if new.find(W + "sectPrChange") is None:
            note = margin_note(number, old.find(W + "pgMar"), new.find(W + "pgMar"))
            if note:
                notes.append(note)
    old_targets, new_targets = _rel_targets(src), _rel_targets(out)
    for number, (old, new) in enumerate(zip(old_sections, new_sections), start=1):
        for kind in ("header", "footer"):
            changed = []
            for ref in new.findall(W + kind + "Reference"):
                kind_type = ref.get(W + "type") or "default"
                target = new_targets.get(ref.get("{" + R_NS + "}id") or "")
                old_ref = next((item for item in old.findall(W + kind + "Reference")
                                if (item.get(W + "type") or "default") == kind_type), None)
                old_target = old_targets.get(old_ref.get("{" + R_NS + "}id") or "") if old_ref is not None else None
                old_text = part_text(src[old_target]) if old_target in src else ""
                new_text = part_text(out[target]) if target in out else ""
                if old_text != new_text and target in out and not _has_revision(out[target]):
                    changed.append(kind_type)
            if changed:
                notes.append(Note(kind, f"{kind.capitalize()} text changed (not tracked)", section=number,
                                  excluded=[("sect", number, kind)]))
    return notes


def _has_revision(data: bytes) -> bool:
    root = ET.fromstring(data)
    return any(element.tag.startswith(W) and element.tag[len(W):] in REVISION_TAGS for element in root.iter())


def merged_cell_errors(expected: dict, actual: dict, notes: Sequence[Note]) -> list[str]:
    """For each merge note: the first merged cell must hold the same paragraphs with the same characters and the
    same run properties per character as ``expected`` (the input), and the cells below it no text. Both snapshots
    are ``formatting_snapshot``s; the document's text is never quoted."""

    errors: list[str] = []
    for note in notes:
        if note.kind != "merge" or note.cell is None or note.rows is None:
            continue
        table, first_row, column = note.cell
        place = (table, first_row, column)
        where = f"the first merged cell (table {table}, row {first_row}, column {column})"
        old_keys = sorted(key for key in expected if key[:3] == ("t", place, "p"))
        new_keys = sorted(key for key in actual if key[:3] == ("t", place, "p"))
        if old_keys != new_keys:
            errors.append(f"{where} has {len(new_keys)} paragraph(s) instead of {len(old_keys)}")
            continue
        for key in old_keys:
            old_chars, new_chars = expected[key][1], actual[key][1]
            if "".join(ch for ch, _ in old_chars) != "".join(ch for ch, _ in new_chars):
                errors.append(f"{where} paragraph {key[3]} text differs ({len(old_chars)} -> {len(new_chars)} "
                              "characters)")
                continue
            changed = [n for n, ((_, a), (_, b)) in enumerate(zip(old_chars, new_chars), start=1) if a != b]
            if changed:
                errors.append(f"{where} paragraph {key[3]}: the character formatting of {len(changed)} of "
                              f"{len(old_chars)} characters differs (from character {changed[0]})")
        for row in range(first_row + 1, note.rows[1] + 1):
            below = (table, row, column)
            if any(actual[key][1] for key in actual if key[:3] == ("t", below, "p")):
                errors.append(f"the merged cell at table {table}, row {row}, column {column} still holds text")
    return errors


# -- cell changes a merge drops ------------------------------------------------------------------------------

_TCPR_ORDER = ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
               "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge",
               "tcPrChange")


def _top_tables(body: ET.Element) -> list[ET.Element]:
    """Tables not inside another table, in document order (Word's ``Document.Tables``)."""

    found: list[ET.Element] = []

    def walk(node: ET.Element) -> None:
        for child in node:
            if child.tag == W + "tbl":
                found.append(child)
            else:
                walk(child)

    walk(body)
    return found


def _serialize_w(element: ET.Element, what: str = "cell properties") -> str:
    """``element`` written with the ``w:`` prefix; refuses anything outside the main namespace."""

    if not element.tag.startswith(W):
        raise RefusedUnsafe(f"{what} hold {element.tag}, which the tracked-change repair can't copy; "
                            "run with tracked=false or make this change in Word")
    attrs = []
    for key, value in element.attrib.items():
        if not key.startswith(W):
            raise RefusedUnsafe(f"{what} hold the attribute {key}, which the tracked-change repair can't copy; "
                                "run with tracked=false or make this change in Word")
        attrs.append(f" w:{key[len(W):]}={quoteattr(value)}")
    name = "w:" + element.tag[len(W):]
    children = "".join(_serialize_w(child, what) for child in element)
    return f"<{name}{''.join(attrs)}>{children}</{name}>" if children else f"<{name}{''.join(attrs)}/>"


def _start_tag_end(data: bytes, index: int) -> tuple[int, bool]:
    """(end, empty) of the start tag beginning at ``index``: ``end`` is just past its ``>``, ``empty`` whether it
    ends ``/>`` (an empty-element tag). Quoted attribute values are skipped, since XML allows ``>`` in them.
    Expat reports an empty element's end event only after the whole tag, so the start tag alone tells."""

    quote = None
    position = index + 1
    while position < len(data):
        char = data[position:position + 1]
        if quote:
            if char == quote:
                quote = None
        elif char in (b'"', b"'"):
            quote = char
        elif char == b">":
            return position + 1, data[position - 1:position] == b"/"
        position += 1
    raise VerificationFailed("unterminated start tag in Word's XML", expected=">", actual=index)


def _cell_spans(data: bytes) -> list[list[list[dict]]]:
    """Byte spans of every top-level table's cells: tables -> rows -> cells -> {"tc_open_end", "tcPr": (start,
    end, empty) | None}. Rows and cells are the direct ``w:tr``/``w:tc`` children, as ``_top_tables`` reads them."""

    import xml.parsers.expat

    parser = xml.parsers.expat.ParserCreate(namespace_separator="}")
    stack: list[tuple[str, int]] = []
    tables: list[list[list[dict]]] = []
    tbl_depth: list[int] = []

    def start(name: str, _attrs: dict) -> None:
        tag = "{" + name if "}" in name else name
        index = parser.CurrentByteIndex
        depth = len(stack)
        parent = stack[-1][0] if stack else ""
        if tag == W + "tbl":
            if not tbl_depth:
                tables.append([])
            tbl_depth.append(depth)
        elif len(tbl_depth) == 1 and tag == W + "tr" and parent == W + "tbl" and depth == tbl_depth[0] + 1:
            tables[-1].append([])
        elif len(tbl_depth) == 1 and tag == W + "tc" and parent == W + "tr" and depth == tbl_depth[0] + 2:
            end, empty = _start_tag_end(data, index)
            if empty:
                raise VerificationFailed("Word wrote an empty table cell element (<w:tc/>), which the tracked-change "
                                         "repair can't edit; run with tracked=false or make this change in Word")
            tables[-1][-1].append({"tc_open_end": end, "tcPr": None})
        elif len(tbl_depth) == 1 and tag == W + "tcPr" and parent == W + "tc" and depth == tbl_depth[0] + 3:
            end, empty = _start_tag_end(data, index)
            tables[-1][-1][-1]["tcPr"] = [index, end, True] if empty else [index, None, False]
        stack.append((tag, index))

    def end(name: str) -> None:
        tag, opened = stack.pop()
        index = parser.CurrentByteIndex
        if tag == W + "tbl":
            tbl_depth.pop()
        elif tag == W + "tcPr" and len(tbl_depth) == 1 and len(stack) == tbl_depth[0] + 3:
            span = tables[-1][-1][-1]["tcPr"]
            if span is not None and span[0] == opened and span[1] is None:
                span[1] = _start_tag_end(data, index)[0]  # the end tag </w:tcPr> starts at ``index``

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(data, True)
    return tables


def _next_revision_id(parts: dict[str, bytes]) -> int:
    ids = [int(value) for name, data in parts.items() if name.endswith(".xml")
           for value in re.findall(rb"\bw:id\s*=\s*[\"'](-?\d+)[\"']", data)]
    return max(ids, default=0) + 1


def _revision_date(xml: bytes) -> str:
    match = re.search(rb"<w:(?:tcPrChange|tblPrChange|pPrChange|rPrChange)\b[^>]*\bw:date=\"([^\"]+)\"", xml)
    if match:
        return match.group(1).decode("ascii")
    from datetime import datetime

    return datetime.now().replace(microsecond=0).isoformat() + "Z"


def restore_merge_row_changes(stage: str | Path, source: str | Path, merges: Sequence[tuple[int, Any]],
                              author: str) -> int:
    """Word drops the tracked cell-property changes (``w:tcPrChange``) of every cell in the rows a merge touches,
    though the merge itself is made with tracking off. Each cell of those rows whose properties now differ from
    the source's (apart from ``w:vMerge``, the merge itself, named by its comment) and has no ``w:tcPrChange`` gets
    one again: the old properties are the source's plus the cell's ``w:vMerge``, so Reject restores the borders and
    keeps the merge. Direct text insertion into ``word/document.xml``; every other byte is kept. Returns the count."""

    stage = Path(stage)
    infos, parts = _read_parts(stage)
    _i, src = _read_parts(Path(source))
    data = parts["word/document.xml"]
    old_tables = _top_tables(ET.fromstring(src["word/document.xml"]).find(W + "body"))
    new_tables = _top_tables(ET.fromstring(data).find(W + "body"))
    spans = _cell_spans(data)
    if len(spans) != len(new_tables) or len(old_tables) != len(new_tables):
        raise VerificationFailed("table count differs while restoring the merged rows' cell changes",
                                 expected=len(old_tables), actual=len(spans))
    next_id = _next_revision_id(parts)
    date = _revision_date(data)
    inserts: list[tuple[int, int, bytes]] = []  # (start, end, replacement)
    rows_done: set[tuple[int, int]] = set()
    for table_index, group in merges:
        old_rows = old_tables[table_index - 1].findall(W + "tr")
        new_rows = new_tables[table_index - 1].findall(W + "tr")
        for row in range(group.start_row, group.end_row + 1):
            if (table_index, row) in rows_done:
                continue
            rows_done.add((table_index, row))
            old_cells = old_rows[row - 1].findall(W + "tc")
            new_cells = new_rows[row - 1].findall(W + "tc")
            cell_spans = spans[table_index - 1][row - 1]
            if not (len(old_cells) == len(new_cells) == len(cell_spans)):
                raise VerificationFailed(f"table {table_index} row {row} cell count differs while restoring the "
                                         "merged rows' cell changes", expected=len(old_cells), actual=len(new_cells))
            for old_cell, new_cell, span in zip(old_cells, new_cells, cell_spans):
                old_pr, new_pr = old_cell.find(W + "tcPr"), new_cell.find(W + "tcPr")
                if new_pr is not None and new_pr.find(W + "tcPrChange") is not None:
                    continue
                if canonical(old_pr, drop=("vMerge",)) == canonical(new_pr, drop=("vMerge",)):
                    continue
                previous = ET.Element(W + "tcPr")
                for child in (old_pr if old_pr is not None else ()):
                    if child.tag != W + "vMerge":
                        previous.append(child)
                merge = new_pr.find(W + "vMerge") if new_pr is not None else None
                if merge is not None:
                    at = next((n for n, child in enumerate(previous)
                               if child.tag[len(W):] in _TCPR_ORDER
                               and _TCPR_ORDER.index(child.tag[len(W):]) > _TCPR_ORDER.index("vMerge")),
                              len(previous))
                    previous.insert(at, merge)
                change = (f"<w:tcPrChange w:id=\"{next_id}\" w:author={quoteattr(author)} w:date=\"{date}\">"
                          f"{_serialize_w(previous)}</w:tcPrChange>").encode("utf-8")
                next_id += 1
                tc_pr = span["tcPr"]
                if tc_pr is None:
                    inserts.append((span["tc_open_end"], span["tc_open_end"], b"<w:tcPr>" + change + b"</w:tcPr>"))
                elif tc_pr[2]:
                    inserts.append((tc_pr[0], tc_pr[1], b"<w:tcPr>" + change + b"</w:tcPr>"))
                else:
                    close = data.rindex(b"</", tc_pr[0], tc_pr[1])
                    inserts.append((close, close, change))
    for start, end, text in sorted(inserts, reverse=True):
        data = data[:start] + text + data[end:]
    if inserts:
        parts["word/document.xml"] = data
        _write_parts(stage, infos, parts)
    return len(inserts)



# -- paragraph changes Word records incompletely ----------------------------------------------------------------

_PPR_OUTSIDE = ("rPr", "sectPr", "pPrChange")  # not part of a w:pPrChange's old properties


def _paragraph_spans(data: bytes) -> list[dict]:
    """Byte spans of every ``w:p`` in document order (``ET.iter`` order): {"open": [start, end, empty] (the start
    tag), "open_end", "pPr": [start, end, empty] | None, "old": [start, end, empty] | None (the ``w:pPr`` inside the
    paragraph's own ``w:pPrChange``)}. ``empty`` marks an empty-element tag (``<w:pPr/>``)."""

    import xml.parsers.expat

    parser = xml.parsers.expat.ParserCreate(namespace_separator="}")
    stack: list[tuple[str, int]] = []
    found: list[dict] = []
    open_paragraphs: list[dict] = []

    def start(name: str, _attrs: dict) -> None:
        tag = "{" + name if "}" in name else name
        index = parser.CurrentByteIndex
        tags = [t for t, _ in stack]
        if tag == W + "p":
            end, empty = _start_tag_end(data, index)
            item = {"open": [index, end, empty], "open_end": end, "pPr": None, "old": None}
            found.append(item)
            open_paragraphs.append(item)
        elif tag == W + "pPr" and open_paragraphs and tags[-1:] == [W + "p"]:
            end, empty = _start_tag_end(data, index)
            open_paragraphs[-1]["pPr"] = [index, end, True] if empty else [index, None, False]
        elif tag == W + "pPr" and open_paragraphs and tags[-3:] == [W + "p", W + "pPr", W + "pPrChange"]:
            end, empty = _start_tag_end(data, index)
            open_paragraphs[-1]["old"] = [index, end, True] if empty else [index, None, False]
        stack.append((tag, index))

    def end(name: str) -> None:
        tag, opened = stack.pop()
        index = parser.CurrentByteIndex
        if tag == W + "p":
            open_paragraphs.pop()
        elif tag == W + "pPr" and open_paragraphs:
            for key in ("pPr", "old"):
                span = open_paragraphs[-1][key]
                if span is not None and span[0] == opened and span[1] is None:
                    span[1] = _start_tag_end(data, index)[0]  # the end tag </w:pPr> starts at ``index``

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(data, True)
    return found


def restore_paragraph_changes(stage: str | Path, source: str | Path, author: str) -> int:
    """Word's ``w:pPrChange`` does not always hold the paragraph's old properties as they were: it drops
    ``w:beforeAutospacing``/``w:afterAutospacing`` from ``w:spacing`` (measured, Word 16), so Reject All would lose
    them. The source is revision-free and the pass adds no paragraph, so the old properties are exactly the
    source's: every paragraph of ``word/document.xml`` whose ``w:pPrChange`` differs from the source's ``w:pPr``
    (its mark's ``w:rPr`` and ``w:sectPr`` aside) gets the source's, and one whose properties changed with no
    ``w:pPrChange`` gets one by ``author``. Direct text edits; every other byte is kept. Returns the count."""

    stage = Path(stage)
    infos, parts = _read_parts(stage)
    _i, src = _read_parts(Path(source))
    data = parts["word/document.xml"]
    old_paragraphs = list(ET.fromstring(src["word/document.xml"]).iter(W + "p"))
    new_paragraphs = list(ET.fromstring(data).iter(W + "p"))
    spans = _paragraph_spans(data)
    from .word_baseline import appended_final_paragraph

    if len(new_paragraphs) == len(old_paragraphs) + 1 == len(spans) and appended_final_paragraph(source, stage):
        # plan 973 S1: the empty paragraph Word appends after a body-final table has no source paragraph; Reject
        # All of its changes is checked against Word's own save of the input, which holds it too
        new_paragraphs, spans = new_paragraphs[:-1], spans[:-1]
    if not (len(old_paragraphs) == len(new_paragraphs) == len(spans)):
        raise VerificationFailed("paragraph count differs while checking Word's paragraph changes",
                                 expected=len(old_paragraphs), actual=len(new_paragraphs))
    next_id = _next_revision_id(parts)
    date = _revision_date(data)
    edits: list[tuple[int, int, bytes]] = []
    for old_p, new_p, span in zip(old_paragraphs, new_paragraphs, spans):
        old_pr, new_pr = old_p.find(W + "pPr"), new_p.find(W + "pPr")
        want = canonical(old_pr, drop=("rPr", "sectPr"))
        change = new_pr.find(W + "pPrChange") if new_pr is not None else None
        if change is None and canonical(new_pr, drop=("rPr", "sectPr")) == want:
            continue
        if change is not None and canonical(change.find(W + "pPr"), drop=("rPr", "sectPr")) == want:
            continue
        previous = ET.Element(W + "pPr")
        for child in (old_pr if old_pr is not None else ()):
            if child.tag[len(W):] not in _PPR_OUTSIDE:
                previous.append(child)
        old_xml = _serialize_w(previous, "paragraph properties").encode("utf-8")
        if change is not None:
            if span["old"] is None:
                raise VerificationFailed("Word wrote a paragraph change with no old properties; run with "
                                         "tracked=false or make this change in Word")
            edits.append((span["old"][0], span["old"][1], old_xml))
            continue
        record = (f"<w:pPrChange w:id=\"{next_id}\" w:author={quoteattr(author)} w:date=\"{date}\">"
                  f"{old_xml.decode('utf-8')}</w:pPrChange>").encode("utf-8")
        next_id += 1
        pr = span["pPr"]
        if pr is None and span["open"][2]:  # <w:p/>: open it to hold the properties
            start, end = span["open"][0], span["open"][1]
            opened = data[start:end - 2].rstrip() + b">"
            name = re.match(rb"<([^\s/>]+)", opened).group(1)
            edits.append((start, end, opened + b"<w:pPr>" + record + b"</w:pPr></" + name + b">"))
        elif pr is None:
            edits.append((span["open_end"], span["open_end"], b"<w:pPr>" + record + b"</w:pPr>"))
        elif pr[2]:
            edits.append((pr[0], pr[1], b"<w:pPr>" + record + b"</w:pPr>"))
        else:
            close = data.rindex(b"</", pr[0], pr[1])
            edits.append((close, close, record))
    for start, end, text in sorted(edits, reverse=True):
        data = data[:start] + text + data[end:]
    if edits:
        parts["word/document.xml"] = data
        _write_parts(stage, infos, parts)
    return len(edits)

# -- comments ------------------------------------------------------------------------------------------------

def _anchor_candidates(path: Path, note: Note) -> list[tuple[int, str]]:
    """(paragraph_index, whole visible text) of the paragraphs the note's comment may sit on, best first."""

    from .sections import read_layout

    layout = read_layout(path)
    paragraphs = [p for p in layout.paragraphs if p.node is not None and not p.nested]
    if note.cell is not None:
        table, row, column = note.cell
        return [(p.index, p.text) for p in paragraphs
                if p.table == table - 1 and p.row == row - 1 and p.cell == column - 1 and p.text.strip()]
    section, number = [], 1
    for p in paragraphs:
        if number == (note.section or 1) and p.text.strip():
            section.append((p.index, p.text))
        if p.table is None and p.node.find(f"{W}pPr/{W}sectPr") is not None:
            number += 1
    return section


def add_untracked_comments(path: str | Path, notes: Sequence[Note], author: str) -> list[dict]:
    """One comment per note by ``author`` on the whole text of the first paragraph that takes one."""

    from .comment_ops import add_comment

    path = Path(path)
    added: list[dict] = []
    if not notes:
        return added
    with tempfile.TemporaryDirectory(prefix="research-mcp-untracked-") as folder:
        current = path
        for n, note in enumerate(notes):
            reasons = []
            for index, anchor in _anchor_candidates(current, note)[:20]:
                target = Path(folder) / f"step{n}.docx"
                try:
                    result = add_comment(current, target, paragraph_index=index, anchor=anchor, text=note.text,
                                         author=author)
                except RefusedUnsafe as exc:
                    reasons.append(f"paragraph {index}: {exc}")
                    continue
                current = target
                added.append({"kind": note.kind, "text": note.text, "comment_id": result["comment_id"],
                              "paragraph_index": index})
                break
            else:
                raise RefusedUnsafe(
                    f"could not add the comment {note.text!r} for a change Word can't track: "
                    + ("; ".join(reasons[:3]) or "no paragraph with text at that place")
                    + ". Run the tool with tracked=false, or make this change in Word")
        shutil.copyfile(current, path)
    return added


# -- the Word check ------------------------------------------------------------------------------------------

def _story_revisions(document: Any) -> int:
    total = 0
    for story in document.StoryRanges:
        current = story
        while current is not None:
            total += int(current.Revisions.Count)
            current = current.NextStoryRange
    return total


def _remaining(document: Any) -> list[Any]:
    found = []
    for story in document.StoryRanges:
        current = story
        while current is not None:
            found.extend(current.Revisions)
            current = current.NextStoryRange
    return found


def word_check_local(jobs: list[dict]) -> dict:
    """Worker entry (kind ``tracked_check``). Each job: ``{"action": "baseline"|"accept"|"reject"|"open",
    "source", "target"}``; baseline/accept/reject copy ``source`` to ``target`` and save Word's result there."""

    from .word import IsolatedWordDocument

    results = []
    for job in jobs:
        action = job["action"]
        if action == "open":
            with IsolatedWordDocument(job["source"], read_only=True, timeout_seconds=120) as session:
                results.append({"action": action, "revisions": _story_revisions(session.document),
                                "comments": int(session.document.Comments.Count)})
            continue
        if action == "baseline":
            # Document.Save on an unmodified document writes nothing, so the target would keep the input's bytes;
            # SaveAs2 always writes Word's own form (plan 973 S1: an empty paragraph after a body-final table).
            with IsolatedWordDocument(job["source"], read_only=True, timeout_seconds=120) as session:
                count = _story_revisions(session.document)
                session.save_as(job["target"])
            results.append({"action": action, "command": None, "revisions_before": count,
                            "revisions_after": count, "resolved_one_by_one": 0})
            continue
        shutil.copyfile(job["source"], job["target"])
        with IsolatedWordDocument(job["target"], timeout_seconds=120) as session:
            document = session.document
            before = _story_revisions(document)
            singles = 0
            command = None
            if action in ("accept", "reject"):
                # The Review tab's Accept/Reject All Changes button runs Word's built-in command AcceptAllChangesInDoc
                # / RejectAllChangesInDoc, which resolves every revision. Document.AcceptAllRevisions (the object
                # model) skips a formatting change on the document's last paragraph mark (measured, Word 16), so it
                # is only the fallback; anything left is then resolved one revision at a time and counted.
                command = "AcceptAllChangesInDoc" if action == "accept" else "RejectAllChangesInDoc"
                try:
                    document.Activate()
                    getattr(session.word.WordBasic, command)  # a WordBasic command runs when it is read
                except Exception:
                    command = "AcceptAllRevisions" if action == "accept" else "RejectAllRevisions"
                    getattr(document, command)()
                for _ in range(3):
                    left = _remaining(document)
                    if not left:
                        break
                    for revision in left:
                        getattr(revision, "Accept" if action == "accept" else "Reject")()
                        singles += 1
            elif action != "baseline":
                raise RefusedUnsafe(f"unknown tracked_check action {action!r}")
            left_count = _story_revisions(document)
            session.save()
        results.append({"action": action, "command": command, "revisions_before": before,
                        "revisions_after": left_count, "resolved_one_by_one": singles})
    return {"jobs": results}


def run_word_check(input_copy: Path, pre_comments: Path, final: Path, folder: Path) -> dict[str, Any]:
    from .operation_process import run_operation

    paths = {"baseline": folder / "baseline.docx", "accept": folder / "accepted.docx",
             "reject": folder / "rejected.docx"}
    jobs = [
        {"action": "baseline", "source": str(input_copy), "target": str(paths["baseline"])},
        {"action": "accept", "source": str(pre_comments), "target": str(paths["accept"])},
        {"action": "reject", "source": str(final), "target": str(paths["reject"])},
        {"action": "open", "source": str(final), "target": None},
    ]
    result = run_operation("tracked_check", {"jobs": jobs}, timeout_seconds=600)
    return {"paths": paths, "jobs": result["jobs"]}


# -- the whole tail ------------------------------------------------------------------------------------------

@contextmanager
def input_copy(stage: str | Path, on: bool) -> Iterator[Path | None]:
    """A copy of the staged input taken before the Word pass (None when tracking is off), removed afterwards."""

    if not on:
        yield None
        return
    with tempfile.TemporaryDirectory(prefix="research-mcp-input-") as folder:
        copy = Path(folder) / "input.docx"
        shutil.copyfile(stage, copy)
        yield copy


def off_fields() -> dict[str, Any]:
    return {"tracking": "off", "tracking_author": None, "revisions_added": 0, "untracked": [], "tracked_check": None}


def finish(stage: str | Path, input_copy: str | Path, *, author: str, verify_accept: Callable[[Path, Path], None],
           merges: Sequence[tuple[int, Any]] = ()) -> dict[str, Any]:
    """Relabel, comment the untracked changes and check the result in Word; raises (nothing published) when Reject
    All does not give the input's formatting outside the commented places, or Accept All fails the tool's own
    verification, or a revision survives either.

    Reject All is compared with, and ``verify_accept(accepted, reference)`` runs the tool's own checks on Word's
    Accept All against, ``reference``: Word's save of the untouched input when it preserves the input
    (``word_baseline.choose_reference``; plan 973 S1), else the input copy; ``compared_against`` in the result says
    which."""

    from .revisions import revision_count

    stage, input_copy = Path(stage), Path(input_copy)
    restored = restore_merge_row_changes(stage, input_copy, merges, author) if merges else 0
    paragraphs_repaired = restore_paragraph_changes(stage, input_copy, author)
    added = relabel(stage, author, input_copy)
    kinds = revision_kinds(stage)
    notes = untracked_notes(input_copy, stage, merges)
    merged = merged_cell_errors(formatting_snapshot(input_copy, "rejected", effective=True),
                                formatting_snapshot(stage, "rejected", effective=True), notes)
    if merged:
        raise VerificationFailed("the cell merge changed what Reject All can't give back: " + "; ".join(merged[:5])
                                 + ". Run with tracked=false or make this change in Word")
    with tempfile.TemporaryDirectory(prefix="research-mcp-tracked-") as name:
        folder = Path(name)
        pre_comments = folder / "pre_comments.docx"
        shutil.copyfile(stage, pre_comments)
        untracked = add_untracked_comments(stage, notes, author)
        if revision_count(stage) != added:
            raise VerificationFailed("adding the untracked-change comments changed the revisions",
                                     expected=added, actual=revision_count(stage))
        check = run_word_check(input_copy, pre_comments, stage, folder)
        paths = check["paths"]
        for side in ("accept", "reject"):
            left = revision_kinds(paths[side], by_part=True)
            if left:
                raise VerificationFailed(
                    f"Word {'Accept' if side == 'accept' else 'Reject'} All left revisions Word cannot resolve: "
                    + ", ".join(f"{count} {kind}" for kind, count in left.items())
                    + "; run with tracked=false or make this change in Word")
        from .word_baseline import choose_reference

        reference = choose_reference(input_copy, paths["baseline"])
        baseline = formatting_snapshot(reference.path, effective=True)
        rejected = formatting_snapshot(paths["reject"], effective=True)
        merged = merged_cell_errors(baseline, rejected, notes)
        if merged:
            raise VerificationFailed("after Word Reject All, " + "; ".join(merged[:5]) + ". Run with tracked=false "
                                     "or make this change in Word")

        def normalize(key: tuple, value: Any) -> Any:
            for note in notes:
                value = note.normalize(key, value)
            return value

        diff = differences(baseline, rejected, lambda key: any(note.covers(key) for note in notes), normalize)
        if diff:
            raise VerificationFailed(
                f"Word Reject All does not give back the input's formatting in {len(diff)} place(s), so the change "
                "would not be fully recorded as tracked changes: "
                + "; ".join(describe_difference(key, baseline, rejected) for key in diff[:5])
                + ". Run with tracked=false or make this change in Word")
        verify_accept(paths["accept"], reference.path)
        jobs = {job["action"]: job for job in check["jobs"]}
    return {
        "tracking": "on",
        "tracking_author": author,
        "revisions_added": added,
        "untracked": untracked,
        "tracked_check": {
            "revision_kinds": kinds,
            "word_revisions": jobs["open"]["revisions"],
            "word_commands": {side: jobs[side]["command"] for side in ("accept", "reject")},
            "resolved_one_by_one": {side: jobs[side]["resolved_one_by_one"] for side in ("accept", "reject")},
            "reject_all_matches_input": True,
            "accept_all_verified": True,
            "places_excluded": len(notes),
            "merged_row_cell_changes_restored": restored,
            "paragraph_changes_repaired": paragraphs_repaired,
        },
        "compared_against": reference.compared_against,
    }
