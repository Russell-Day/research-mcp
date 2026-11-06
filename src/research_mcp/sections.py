"""Offline planning and verification for minimal-diff text edits.

Whole-line replacement used to be refused whenever the line spanned runs with
different formatting -- which is every résumé entry title ("**Delta Park
Museum**, Department of Collections, *Outreach Volunteer*"). The fix is to
never replace the line at all: diff the paragraph's current text against the
requested text and touch only the characters that differ.

This module does everything that needs no Word:

* reads each paragraph's stored text with a per-character formatting
  fingerprint and a caps flag (`w:caps`/`w:smallCaps` make stored
  "WorK Experience" display as "WORK EXPERIENCE");
* plans the character edits (`plan_text_diff`) and, for rows and sections,
  the paragraph and row inserts and deletes, as an ordered list of steps that
  `editing._execute_plan` performs bottom-up in Word;
* records what the result must look like, so `verify_plan` can prove it on
  the saved candidate: requested text, unchanged characters' formatting,
  inserted characters' formatting, every unit outside the edit unchanged,
  cell counts, no unrequested empty cell, date cells still dates.

Formatting rule (D8): unchanged characters keep their runs. A replacement
takes the formatting of the first character it replaces. A pure insertion
takes the formatting of the character before it -- at the start of a
paragraph, the one after -- except that text which prefixes the next word
("Seasonal " before "Outreach") takes the next word's formatting, which is
what typing it there in Word produces. A neighbour inside a field, hyperlink
or content control is passed over when the other neighbour is outside one.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
import difflib
from io import BytesIO
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import NS, W
from .paragraphs import read_snapshot


W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_RSID_ATTRS = {"rsidR", "rsidRPr", "rsidRDefault", "rsidP", "rsidDel", "rsidTr", "rsidSect", "rsidRDefault"}
# cnfStyle caches which table-style conditions (first column, banding) apply; it
# formats nothing and Word adds or drops it on save (editing one cell of the JBJS
# Table 1 made Word drop it from row 3's cells and paragraphs but not row 4's).
_DROP_TAGS = {W + "proofErr", W + "lastRenderedPageBreak", W + "cnfStyle"}
# Empty property containers carry nothing; Word adds an empty sdtEndPr on its first save of a document.
_EMPTY_DROP_TAGS = {W + "rPr", W + "pPr", W + "sdtEndPr"}
_OFF = {"0", "false", "off"}


# --------------------------------------------------------------------------
# Canonical XML: what "unchanged" means after Word re-saves a package.
# --------------------------------------------------------------------------

def _canon(element: ET.Element) -> None:
    for key in list(element.attrib):
        local = key.rsplit("}", 1)[-1]
        if (key.startswith(W) and local in _RSID_ATTRS) or key in (W14 + "paraId", W14 + "textId") or (
            key == XML_SPACE and element.tag == W + "t"
        ):
            del element.attrib[key]
    attributes = sorted(element.attrib.items())
    element.attrib.clear()
    element.attrib.update(attributes)
    children = [child for child in element if child.tag not in _DROP_TAGS]
    for child in children:
        _canon(child)
    children = [
        child for child in children
        if not (child.tag in _EMPTY_DROP_TAGS and len(child) == 0 and not child.attrib)
    ]
    if element.tag.endswith("Pr"):
        children.sort(key=lambda child: child.tag)
    merged: list[ET.Element] = []
    for child in children:
        if child.tag == W + "r" and merged and merged[-1].tag == W + "r" and _mergeable(child) and _mergeable(merged[-1]) \
                and _rpr_key(child) == _rpr_key(merged[-1]):
            previous = merged[-1]
            for part in child:
                if part.tag != W + "rPr":
                    previous.append(part)
            _join_text(previous)
        else:
            merged.append(child)
    element[:] = merged


def _mergeable(run: ET.Element) -> bool:
    return all(child.tag in (W + "rPr", W + "t", W + "tab") for child in run)


def _rpr_key(run: ET.Element) -> str:
    rpr = run.find("w:rPr", NS)
    return ET.tostring(rpr, encoding="unicode") if rpr is not None else ""


def _join_text(run: ET.Element) -> None:
    parts: list[ET.Element] = []
    for child in list(run):
        if child.tag == W + "t" and parts and parts[-1].tag == W + "t":
            parts[-1].text = (parts[-1].text or "") + (child.text or "")
        else:
            parts.append(child)
    run[:] = parts


def canonical(element: ET.Element | None) -> str:
    if element is None:
        return ""
    clone = copy.deepcopy(element)
    _canon(clone)
    return ET.tostring(clone, encoding="unicode")


# --------------------------------------------------------------------------
# Document model.
# --------------------------------------------------------------------------

@dataclass
class Para:
    index: int
    node: ET.Element | None
    text: str = ""
    fps: tuple[str, ...] = ()
    caps: tuple[bool, ...] = ()
    table: int | None = None
    row: int | None = None
    cell: int | None = None
    line: int | None = None
    nested: bool = False
    exact: bool = True
    has_fields: bool = False
    style: str | None = None
    outline: int | None = None

    @property
    def row_end(self) -> bool:
        return self.node is None

    @property
    def displayed(self) -> str:
        return display(self.text, self.caps)

    def ppr(self, *, mark: bool = True) -> str:
        """Canonical paragraph properties; `mark=False` ignores the mark's own run formatting."""

        if self.node is None:
            return ""
        ppr = self.node.find("w:pPr", NS)
        if ppr is not None and not mark:
            ppr = copy.deepcopy(ppr)
            for rpr in ppr.findall("w:rPr", NS):
                ppr.remove(rpr)
        return canonical(ppr)


@dataclass
class Row:
    table: int
    row: int
    node: ET.Element
    cells: list[list[int]]
    end_index: int
    unit: int
    complex: str | None = None

    @property
    def first_index(self) -> int:
        return self.cells[0][0] if self.cells and self.cells[0] else self.end_index


@dataclass
class Unit:
    kind: str  # "p", "tbl", "row", "other"
    node: ET.Element
    first: int | None = None
    last: int | None = None
    _key: str | None = None

    def key(self) -> str:
        if self._key is None:
            if self.kind == "tbl":
                self._key = canonical(self.node.find("w:tblPr", NS)) + canonical(self.node.find("w:tblGrid", NS))
            else:
                self._key = canonical(self.node)
        return self._key


@dataclass
class Layout:
    path: Path
    paragraphs: list[Para]
    units: list[Unit]
    tables: list[list[Row]]
    styles_outline: dict[str, int]
    body: ET.Element | None = None
    comment_keys: dict[str, tuple] = field(default_factory=dict)  # comment id -> (author, initials, text)

    def row_at_unit(self, position: int) -> Row | None:
        for rows in self.tables:
            for row in rows:
                if row.unit == position:
                    return row
        return None

    def para(self, index: int) -> Para:
        if not 1 <= index <= len(self.paragraphs):
            raise RefusedUnsafe(f"paragraph_index {index} is outside 1..{len(self.paragraphs)}")
        return self.paragraphs[index - 1]

    def row_of(self, index: int) -> Row | None:
        item = self.para(index)
        if item.table is None or item.nested:
            return None
        return self.tables[item.table][item.row]

    def unit_of(self, index: int) -> int:
        row = self.row_of(index)
        if row is not None:
            return row.unit
        item = self.para(index)
        for position, unit in enumerate(self.units):
            if unit.first is not None and unit.first <= index <= unit.last:
                return position
        raise RefusedUnsafe(f"paragraph {index} is not addressable for this edit (nested table)")


def _on(element: ET.Element | None) -> bool | None:
    if element is None:
        return None
    return element.get(W + "val", "true").lower() not in _OFF


def _rpr_caps(rpr: ET.Element | None) -> bool | None:
    if rpr is None:
        return None
    flags = [_on(rpr.find(f"w:{tag}", NS)) for tag in ("caps", "smallCaps")]
    present = [flag for flag in flags if flag is not None]
    return any(present) if present else None


def _styles(archive: zipfile.ZipFile) -> tuple[dict[str, bool], dict[str, int]]:
    try:
        root = ET.fromstring(archive.read("word/styles.xml"))
    except KeyError:
        return {}, {}
    own_caps: dict[str, bool | None] = {}
    own_level: dict[str, int | None] = {}
    based: dict[str, str] = {}
    for style in root.findall("w:style", NS):
        sid = style.get(W + "styleId")
        if not sid:
            continue
        own_caps[sid] = _rpr_caps(style.find("w:rPr", NS))
        level = style.find("w:pPr/w:outlineLvl", NS)
        own_level[sid] = int(level.get(W + "val")) if level is not None and level.get(W + "val", "").isdigit() else None
        parent = style.find("w:basedOn", NS)
        if parent is not None and parent.get(W + "val"):
            based[sid] = parent.get(W + "val")

    def resolve(table: dict, sid: str | None, depth: int = 0):
        if sid is None or depth > 10 or sid not in table:
            return None
        if table[sid] is not None:
            return table[sid]
        return resolve(table, based.get(sid), depth + 1)

    caps = {sid: bool(resolve(own_caps, sid)) for sid in own_caps}
    levels = {sid: value for sid in own_level if (value := resolve(own_level, sid)) is not None and value < 9}
    return caps, levels


def _walk_paragraph(paragraph: ET.Element, style_caps: dict[str, bool], para_style: str | None):
    """Characters in `_paragraph_text` order with formatting context.

    Mirrors `paragraphs._paragraph_text` exactly (w:t, w:tab, w:br, w:cr at
    any depth outside pPr) so offsets agree with every other tool.
    """

    text: list[str] = []
    fps: list[str] = []
    caps: list[bool] = []
    field_depth = [0]
    para_caps = style_caps.get(para_style or "", False)

    def emit(node: ET.Element, rpr_key: str, run_caps: bool, context: str) -> None:
        for child in node:
            tag = child.tag
            if tag == W + "pPr":
                continue
            if tag == W + "r":
                rpr = child.find("w:rPr", NS)
                direct = _rpr_caps(rpr)
                if direct is None:
                    style = rpr.find("w:rStyle", NS) if rpr is not None else None
                    direct = style_caps.get(style.get(W + "val"), False) if style is not None else para_caps
                emit(child, canonical(rpr), bool(direct), context)
                continue
            if tag == W + "rPr":
                continue
            if tag == W + "fldChar":
                kind = child.get(W + "fldCharType")
                if kind == "begin":
                    field_depth[0] += 1
                elif kind == "end":
                    field_depth[0] = max(0, field_depth[0] - 1)
                continue
            if tag in (W + "t", W + "tab", W + "br", W + "cr"):
                if tag == W + "t":
                    chunk = child.text or ""
                elif tag == W + "tab":
                    chunk = "\t"
                else:
                    chunk = "\x0c" if child.get(W + "type") == "page" else "\v"
                flags = context + ("f" if field_depth[0] else "")
                for character in chunk:
                    text.append(character)
                    fps.append(flags + "|" + rpr_key)
                    caps.append(run_caps)
                continue
            extra = {W + "hyperlink": "h", W + "sdt": "s", W + "fldSimple": "f"}.get(tag, "")
            emit(child, rpr_key, run_caps, context + extra)

    emit(paragraph, "", para_caps, "")
    return "".join(text), tuple(fps), tuple(caps)


def read_layout(path: str | Path) -> Layout:
    snapshot = read_snapshot(path)
    with zipfile.ZipFile(BytesIO(snapshot.package)) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        style_caps, style_levels = _styles(archive)
        comment_keys = {}
        if "word/comments.xml" in archive.namelist():
            for comment in ET.fromstring(archive.read("word/comments.xml")).findall("w:comment", NS):
                text = "".join(node.text or "" for node in comment.iter(W + "t"))
                comment_keys[comment.get(W + "id")] = (comment.get(W + "author"), comment.get(W + "initials"), text)
    body = root.find("w:body", NS)
    paragraphs: list[Para] = []
    units: list[Unit] = []
    tables: list[list[Row]] = []

    def add_paragraph(node: ET.Element | None, **where) -> int:
        index = len(paragraphs) + 1
        if node is None:
            paragraphs.append(Para(index=index, node=None, **where))
            return index
        style_node = node.find("w:pPr/w:pStyle", NS)
        style = style_node.get(W + "val") if style_node is not None else None
        text, fps, caps = _walk_paragraph(node, style_caps, style)
        level_node = node.find("w:pPr/w:outlineLvl", NS)
        if level_node is not None and level_node.get(W + "val", "").isdigit():
            outline = int(level_node.get(W + "val"))
            outline = outline if outline < 9 else None
        else:
            outline = style_levels.get(style or "")
        paragraphs.append(Para(
            index=index, node=node, text=text, fps=fps, caps=caps, style=style, outline=outline,
            exact=not any(node.find(f".//w:{tag}", NS) is not None for tag in ("drawing", "pict", "object")),
            has_fields=any(node.find(f".//w:{tag}", NS) is not None for tag in ("fldChar", "instrText", "fldSimple")),
            **where,
        ))
        return index

    def walk_nested(table: ET.Element, top: int, row: int, cell: int) -> None:
        for nested_row in table.findall("w:tr", NS):
            for nested_cell in nested_row.findall("w:tc", NS):
                for child in nested_cell:
                    if child.tag == W + "p":
                        add_paragraph(child, table=top, row=row, cell=cell, nested=True)
                    elif child.tag == W + "tbl":
                        walk_nested(child, top, row, cell)
            add_paragraph(None, table=top, row=row, cell=cell, nested=True)

    if body is not None:
        for child in body:
            if child.tag == W + "p":
                index = add_paragraph(child)
                units.append(Unit("p", child, index, index))
            elif child.tag == W + "tbl":
                table_number = len(tables)
                rows: list[Row] = []
                units.append(Unit("tbl", child))
                for row_number, row_node in enumerate(child.findall("w:tr", NS)):
                    first = len(paragraphs) + 1
                    cells: list[list[int]] = []
                    complex_reason = None
                    for cell_number, cell_node in enumerate(row_node.findall("w:tc", NS)):
                        lines: list[int] = []
                        for item in cell_node:
                            if item.tag == W + "p":
                                lines.append(add_paragraph(
                                    item, table=table_number, row=row_number, cell=cell_number, line=len(lines)))
                            elif item.tag == W + "tbl":
                                complex_reason = complex_reason or f"cell {cell_number + 1} holds a nested table"
                                walk_nested(item, table_number, row_number, cell_number)
                            elif item.tag != W + "tcPr":
                                complex_reason = complex_reason or (
                                    f"cell {cell_number + 1} holds a {item.tag.rsplit('}', 1)[-1]} element")
                        cells.append(lines)
                    end = add_paragraph(None, table=table_number, row=row_number)
                    units.append(Unit("row", row_node, first, end))
                    rows.append(Row(table_number, row_number, row_node, cells, end, len(units) - 1, complex_reason))
                tables.append(rows)
            else:
                units.append(Unit("other", child))
    return Layout(Path(path), paragraphs, units, tables, style_levels, body, comment_keys)


# --------------------------------------------------------------------------
# Caps-aware text helpers (D12).
# --------------------------------------------------------------------------

def display(text: str, caps) -> str:
    return "".join(ch.upper() if flag else ch for ch, flag in zip(text, caps))


def _char_eq(stored: str, is_caps: bool, wanted: str) -> bool:
    return stored == wanted or (is_caps and stored.upper() == wanted.upper())


def caps_equal(stored: str, caps, wanted: str) -> bool:
    return len(stored) == len(wanted) and all(_char_eq(a, c, b) for a, c, b in zip(stored, caps, wanted))


def caps_find(stored: str, caps, wanted: str) -> list[int]:
    """Start offsets where `wanted` matches; case is relaxed only on caps characters."""

    hits = []
    width = len(wanted)
    for start in range(0, len(stored) - width + 1):
        if all(_char_eq(stored[start + k], caps[start + k], wanted[k]) for k in range(width)):
            hits.append(start)
    return hits


def _strip_span(text: str) -> tuple[int, int]:
    start, end = 0, len(text)
    while start < end and text[start].isspace():
        start += 1
    while end > start and text[end - 1].isspace():
        end -= 1
    return start, end


def caps_equal_stripped(item: Para, wanted: str) -> bool:
    start, end = _strip_span(item.text)
    return caps_equal(item.text[start:end], item.caps[start:end], wanted.strip())


# --------------------------------------------------------------------------
# Character diff (D8).
# --------------------------------------------------------------------------

@dataclass
class Edit:
    start: int
    end: int
    text: str
    mode: str  # "replace", "delete", "after", "before", "empty"
    source: int | None  # offset of the character whose formatting the new text takes

    def describe(self, old: str) -> str:
        removed = old[self.start:self.end]
        if self.mode == "delete":
            return f"delete {removed!r} at characters {self.start}-{self.end}"
        if self.mode == "replace":
            return f"replace {removed!r} with {self.text!r} at characters {self.start}-{self.end}"
        return f"insert {self.text!r} at character {self.start}"


_TOKEN = re.compile(r"\w+|\s+|[^\w\s]", re.UNICODE)


def _tokens(text: str) -> list[tuple[int, int]]:
    return [(match.start(), match.end()) for match in _TOKEN.finditer(text)]


def _in_object(fp: str) -> bool:
    return bool(fp.split("|", 1)[0])


def plan_text_diff(old: str, new: str, caps, fps=None) -> list[Edit]:
    """Minimal edits turning stored `old` into requested `new`, left to right.

    Tokens are matched on displayed text, so a caps run's stored case
    ("WorK") survives a request that spells it as displayed ("WORK"). A
    changed token span whose characters carry more than one formatting is
    refined character by character, so characters it shares with the
    request keep their own runs: `AB**CD**EF` -> `AXCDYF` replaces B and E
    and leaves the bold CD alone.
    """

    fps = fps or ("|",) * len(old)
    old_spans, new_spans = _tokens(old), _tokens(new)
    shown = display(old, caps)
    old_keys = [shown[a:b] for a, b in old_spans]
    new_keys = [new[a:b] for a, b in new_spans]
    matcher = difflib.SequenceMatcher(None, old_keys, new_keys, autojunk=False)
    spans: list[tuple[int, int, str]] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        a = old_spans[i1][0] if i1 < len(old_spans) else len(old)
        b = old_spans[i2 - 1][1] if i2 > i1 else a
        c = new_spans[j1][0] if j1 < len(new_spans) else len(new)
        d = new_spans[j2 - 1][1] if j2 > j1 else c
        while a < b and c < d and _char_eq(old[a], caps[a], new[c]):
            a, c = a + 1, c + 1
        while a < b and c < d and _char_eq(old[b - 1], caps[b - 1], new[d - 1]):
            b, d = b - 1, d - 1
        if a == b and c == d:
            continue
        if a < b and c < d and len({(fps[k], caps[k]) for k in range(a, b)}) > 1:
            spans.extend(_refine(shown, new, a, b, c, d))
        else:
            spans.append((a, b, new[c:d]))
    edits = []
    for number, (start, end, text) in enumerate(spans):
        # An insertion may slide right over identical characters, but never
        # onto a character the next edit changes.
        limit = spans[number + 1][0] - 1 if number + 1 < len(spans) else len(old)
        edits.append(_classify(old, start, end, text, fps, limit))
    return edits


def _refine(shown: str, new: str, a: int, b: int, c: int, d: int) -> list[tuple[int, int, str]]:
    """Character-level edits inside one changed span, keeping its equal characters."""

    matcher = difflib.SequenceMatcher(None, shown[a:b], new[c:d], autojunk=False)
    return [
        (a + i1, a + i2, new[c + j1:c + j2])
        for tag, i1, i2, j1, j2 in matcher.get_opcodes() if tag != "equal"
    ]


def _classify(old: str, start: int, end: int, text: str, fps, limit: int | None = None) -> Edit:
    if start < end:
        return Edit(start, end, text, "replace" if text else "delete", start if text else None)
    # Slide a pure insertion right past identical characters so that
    # " Seasonal" after ", " becomes "Seasonal " before "Outreach".
    limit = len(old) if limit is None else limit
    while start < min(limit, len(old)) and text and old[start] == text[0]:
        text = text[1:] + old[start]
        start += 1
    if not old:
        return Edit(0, 0, text, "empty", None)
    before = start - 1 if start > 0 else None
    after = start if start < len(old) else None
    prefixes_next_word = (
        after is not None
        and text[-1].isspace()
        and not text[0].isspace()
        and (old[after].isalnum())
        and (before is None or old[before].isspace())
    )
    choice = after if (before is None or prefixes_next_word) else before
    other = before if choice == after else after
    if other is not None and _in_object(fps[choice]) and not _in_object(fps[other]):
        choice = other
    return Edit(start, start, text, "before" if choice == after else "after", choice)


def apply_edits(old: str, fps, caps, edits: list[Edit]):
    """Expected stored text, per-character fingerprints and caps flags."""

    text: list[str] = []
    out_fps: list[str | None] = []
    out_caps: list[bool] = []
    cursor = 0
    for edit in edits:
        text.extend(old[cursor:edit.start])
        out_fps.extend(fps[cursor:edit.start])
        out_caps.extend(caps[cursor:edit.start])
        source_fp = fps[edit.source] if edit.source is not None else None
        source_caps = caps[edit.source] if edit.source is not None else False
        text.extend(edit.text)
        out_fps.extend([source_fp] * len(edit.text))
        out_caps.extend([source_caps] * len(edit.text))
        cursor = edit.end
    text.extend(old[cursor:])
    out_fps.extend(fps[cursor:])
    out_caps.extend(caps[cursor:])
    return "".join(text), out_fps, out_caps


# --------------------------------------------------------------------------
# Dates (D11).
# --------------------------------------------------------------------------

_MONTH = (r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|"
          r"sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?")
_ONE_DATE = (
    rf"(?:(?:expected|exp\.?|est\.?)\s+)?(?:\d{{1,2}}/\d{{1,2}}/\d{{2,4}}|\d{{1,2}}/\d{{4}}|\d{{4}}-\d{{1,2}}(?:-\d{{1,2}})?|"
    rf"{_MONTH}\s+\d{{1,2}},?\s+\d{{4}}|{_MONTH},?\s+\d{{4}}|(?:spring|summer|fall|autumn|winter)\s+\d{{4}}|"
    rf"(?:19|20)\d{{2}}|present|current|now|ongoing)"
)
_DATE_RANGE = rf"{_ONE_DATE}(?:\s*(?:-|–|—|to|through|until)\s*{_ONE_DATE})?"
_DATE_LINE = re.compile(rf"^\s*{_DATE_RANGE}(?:\s*[,;]\s*{_DATE_RANGE})*\s*$", re.IGNORECASE)


def is_date_like(text: str) -> bool:
    lines = [line for line in re.split(r"[\r\n\v]", text) if line.strip()]
    return bool(lines) and all(_DATE_LINE.match(line) for line in lines)


# --------------------------------------------------------------------------
# Plans.
# --------------------------------------------------------------------------

@dataclass
class ExpLine:
    text: str
    caps: list[bool]
    fps: list[str | None]
    ppr: str | None
    requested: str | None = None  # None: the line must be byte-for-byte (canonically) unchanged
    canon: str | None = None
    mark: bool = True  # compare the paragraph mark's run formatting too
    # The paragraph this line was edited from: (stored text, per-character formatting).
    # Verification aligns it with the result on its own, without the plan.
    origin: tuple[str, tuple[str, ...]] | None = None


@dataclass
class ExpUnit:
    kind: str  # "row" or "p"
    lines: list[list[ExpLine]]  # per cell (a body paragraph is one cell of one line)
    tcprs: list[str] = field(default_factory=list)
    trpr: str = ""
    dates: list[bool] = field(default_factory=list)  # per cell: must still read as a date
    label: str = ""


@dataclass
class TextStep:
    index: int
    text: str
    caps: list[bool]
    edits: list[Edit]


@dataclass
class InsertLinesStep:
    index: int
    text: str
    caps: list[bool]
    where: str  # "before" / "after"
    lines: list[str]


@dataclass
class DeleteParasStep:
    first: int
    last: int
    trailing: bool  # last line of its cell: merge into the line before
    previous: int | None
    fix_format: bool
    texts: list[str] = field(default_factory=list)
    caps: list[tuple[bool, ...]] = field(default_factory=list)


@dataclass
class DuplicateRowStep:
    template: int  # first paragraph index of the template row
    before: int | None  # first paragraph index of the row the copy goes before
    append_after: int | None  # first paragraph index of the table's last row


@dataclass
class DeleteRowStep:
    index: int  # any paragraph of the row


@dataclass
class EditPlan:
    op: str
    steps: list = field(default_factory=list)
    unit_lo: int = 0
    unit_hi: int = 0
    expected: list[ExpUnit] = field(default_factory=list)
    paragraph_delta: int = 0
    first_index: int = 1
    last_index: int = 1
    report: dict = field(default_factory=dict)


def _line_expectation(item: Para) -> ExpLine:
    return ExpLine(item.text, list(item.caps), list(item.fps), item.ppr(), None, canonical(item.node))


def _check_editable(item: Para, op: str) -> None:
    if item.row_end:
        raise RefusedUnsafe(f"{op}: paragraph {item.index} is a table row-end marker and holds no text")
    if not item.exact:
        raise RefusedUnsafe(
            f"{op}: paragraph {item.index} contains a drawing, so its text and offsets cannot be matched "
            "against Word reliably")


def _check_line(text: str, op: str) -> None:
    for bad, name in (("\r", "carriage return"), ("\n", "line feed"), ("\x07", "cell marker")):
        if bad in text:
            raise RefusedUnsafe(f"{op}: paragraph text may not contain a {name}; give one entry per paragraph")


def _plan_cell(
    plan: EditPlan, op: str, lines: list[Para], new_lines: list[str], *, base_shift: int = 0,
) -> list[ExpLine]:
    """Steps (appended bottom-up) and expectations for one cell's paragraphs.

    `lines` describe the cell as it stands when its steps run; `base_shift`
    moves their indices (a freshly duplicated row sits below its template).
    """

    for text in new_lines:
        _check_line(text, op)
    old_keys = [item.displayed for item in lines]
    matcher = difflib.SequenceMatcher(None, old_keys, list(new_lines), autojunk=False)
    expected: list[ExpLine | None] = [None] * len(new_lines)
    origin: dict[int, int] = {}  # old line position -> new line position
    relaxed: set[int] = set()  # old lines whose paragraph mark is replaced by a deleted line's
    codes = matcher.get_opcodes()
    for tag, i1, i2, j1, j2 in codes:
        if tag == "equal":
            for offset in range(i2 - i1):
                expected[j1 + offset] = _line_expectation(lines[i1 + offset])
                origin[i1 + offset] = j1 + offset
    for tag, i1, i2, j1, j2 in reversed(codes):
        if tag == "equal":
            continue
        paired = min(i2 - i1, j2 - j1)
        if j2 - j1 > paired:
            added = list(new_lines[j1 + paired:j2])
            following = i1 + paired
            if following < len(lines):
                anchor = lines[following]
                where = "before"
                source = 0 if anchor.text else None
            else:
                anchor = lines[following - 1]
                where = "after"
                source = len(anchor.text) - 1 if anchor.text else None
            _check_editable(anchor, op)
            plan.steps.append(InsertLinesStep(anchor.index + base_shift, anchor.text, list(anchor.caps), where, added))
            for offset, text in enumerate(added):
                fp = anchor.fps[source] if source is not None else None
                is_caps = anchor.caps[source] if source is not None else False
                # A paragraph made by splitting the anchor keeps its paragraph
                # properties; Word gives the new mark the typed text's formatting.
                expected[j1 + paired + offset] = ExpLine(
                    text, [is_caps] * len(text), [fp] * len(text), anchor.ppr(mark=False), text, mark=False)
        if i2 - i1 > paired:
            doomed = lines[i1 + paired:i2]
            for item in doomed:
                _refuse_objects(item.node, op, f"delete paragraph {item.index}")
            trailing = i2 == len(lines)
            previous = lines[i1 + paired - 1] if trailing else None
            if trailing and previous is None:
                raise RefusedUnsafe(f"{op}: a cell must keep at least one paragraph")
            fix = bool(trailing and previous is not None and previous.ppr() != doomed[-1].ppr())
            if trailing:
                # The cell's last paragraph mark survives the delete, so the line
                # before inherits it; Word's ParagraphFormat restores the rest.
                relaxed.add(i1 + paired - 1)
            plan.steps.append(DeleteParasStep(
                doomed[0].index + base_shift, doomed[-1].index + base_shift, trailing,
                previous.index + base_shift if previous is not None else None, fix,
                [item.text for item in doomed], [item.caps for item in doomed]))
        for offset in reversed(range(paired)):
            item = lines[i1 + offset]
            wanted = new_lines[j1 + offset]
            _check_editable(item, op)
            edits = plan_text_diff(item.text, wanted, item.caps, item.fps)
            if edits:
                plan.steps.append(TextStep(item.index + base_shift, item.text, list(item.caps), edits))
            text, fps, caps = apply_edits(item.text, item.fps, item.caps, edits)
            expected[j1 + offset] = ExpLine(text, caps, fps, item.ppr(), wanted, origin=(item.text, item.fps))
            origin[i1 + offset] = j1 + offset
    for position in relaxed:
        if position in origin:
            item, line = lines[position], expected[origin[position]]
            expected[origin[position]] = ExpLine(
                line.text, line.caps, line.fps, item.ppr(mark=False),
                line.requested if line.requested is not None else display(line.text, line.caps), mark=False,
                origin=(item.text, item.fps))
    return [line for line in expected if line is not None]


def _split_cell(text: str) -> list[str]:
    return text.replace("\r\n", "\n").split("\n")


def _cell_text(layout: Layout, indices: list[int]) -> str:
    return "\n".join(layout.para(index).displayed for index in indices)


def _row_texts(layout: Layout, row: Row) -> list[str]:
    return [_cell_text(layout, cell) for cell in row.cells]


def _tcprs(row: Row) -> list[str]:
    return [canonical(cell.find("w:tcPr", NS)) for cell in row.node.findall("w:tc", NS)]


def _trpr(row: Row) -> str:
    return canonical(row.node.find("w:trPr", NS))


def _plan_row(
    plan: EditPlan, layout: Layout, op: str, row: Row, new_cells: list[str], *, base_shift: int = 0,
    label: str = "",
) -> ExpUnit:
    if row.complex:
        raise RefusedUnsafe(f"{op}: table {row.table + 1} row {row.row + 1}: {row.complex}; edit it paragraph by paragraph")
    if len(new_cells) != len(row.cells):
        raise RefusedUnsafe(
            f"{op}: table {row.table + 1} row {row.row + 1} has {len(row.cells)} cells; "
            f"{len(new_cells)} cell texts were given (the cell count must stay the same)")
    dates = []
    unit = ExpUnit("row", [], _tcprs(row), _trpr(row), [], label)
    for number in reversed(range(len(row.cells))):
        old = _cell_text(layout, row.cells[number])
        wanted = new_cells[number]
        must_date = is_date_like(old)
        if must_date and not is_date_like(wanted):
            raise RefusedUnsafe(
                f"{op}: table {row.table + 1} row {row.row + 1} cell {number + 1} holds a date ({old!r}) "
                f"but the requested text {wanted!r} does not read as a date or date range")
        dates.append(must_date)
        lines = [layout.para(index) for index in row.cells[number]]
        unit.lines.append(_plan_cell(plan, op, lines, _split_cell(wanted), base_shift=base_shift))
    unit.lines.reverse()
    dates.reverse()
    unit.dates = dates
    return unit


def _unchanged_row(layout: Layout, row: Row) -> ExpUnit:
    lines = [[_line_expectation(layout.para(index)) for index in cell] for cell in row.cells]
    return ExpUnit("row", lines, _tcprs(row), _trpr(row), [False] * len(row.cells))


def _row_size(row: Row) -> int:
    return row.end_index - row.first_index + 1


def _record_text(report: dict, prefix: str, layout: Layout, index: int) -> None:
    item = layout.para(index)
    report[prefix + "_stored"] = item.text
    report[prefix + "_displayed"] = item.displayed


def plan_paragraph(layout: Layout, op: str, index: int, new_text: str, *, find: str | None = None) -> EditPlan:
    item = layout.para(index)
    _check_editable(item, op)
    _check_line(new_text, op)
    plan = EditPlan(op=op, first_index=index, last_index=index)
    edits = plan_text_diff(item.text, new_text, item.caps, item.fps)
    if edits:
        plan.steps.append(TextStep(index, item.text, list(item.caps), edits))
    text, fps, caps = apply_edits(item.text, item.fps, item.caps, edits)
    line = ExpLine(text, caps, fps, item.ppr(), new_text, origin=(item.text, item.fps))
    row = layout.row_of(index)
    if row is not None and not item.nested:
        unit = _unchanged_row(layout, row)
        unit.lines[item.cell][item.line] = line
        old_cell = _cell_text(layout, row.cells[item.cell])
        new_cell = "\n".join(
            new_text if member == index else layout.para(member).displayed for member in row.cells[item.cell])
        if is_date_like(old_cell) and not is_date_like(new_cell):
            raise RefusedUnsafe(
                f"{op}: paragraph {index} is in a date cell ({old_cell!r}); the result {new_cell!r} "
                "would no longer read as a date or date range")
        unit.dates[item.cell] = is_date_like(old_cell)
        plan.unit_lo = row.unit
    elif item.nested:
        raise RefusedUnsafe(f"{op}: paragraph {index} is inside a nested table, which this edit does not support")
    else:
        unit = ExpUnit("p", [[line]])
        plan.unit_lo = layout.unit_of(index)
    plan.unit_hi = plan.unit_lo + 1
    plan.expected = [unit]
    plan.report = {
        "paragraph_index": index,
        "text_before": item.text,
        "text_before_displayed": item.displayed,
        "text_after": text,
        "text_after_displayed": display(text, caps),
        "diff": [edit.describe(item.text) for edit in edits],
    }
    if find is not None:
        plan.report["find"] = find
    return plan


def plan_replace_text(layout: Layout, op: str, index: int, find: str, replace: str) -> EditPlan:
    item = layout.para(index)
    _check_editable(item, op)
    hits = caps_find(item.text, item.caps, find)
    if len(hits) != 1:
        raise RefusedUnsafe(
            f"find must occur exactly once in paragraph {index}; it occurs {len(hits)} times"
            + (" (case is relaxed only inside caps/small-caps runs)" if any(item.caps) else ""))
    start = hits[0]
    # Stored text outside the match, so its stored case is kept (D12).
    requested = item.text[:start] + replace + item.text[start + len(find):]
    return plan_paragraph(layout, op, index, requested, find=find)


def find_row(layout: Layout, match: str, *, table_index: int | None = None) -> Row:
    wanted = match.strip()
    hits: list[Row] = []
    for number, rows in enumerate(layout.tables):
        if table_index is not None and number != table_index - 1:
            continue
        for row in rows:
            found = False
            for cell in row.cells:
                items = [layout.para(index) for index in cell]
                joined = "\n".join(item.text for item in items)
                joined_caps = [flag for item in items for flag in (*item.caps, False)][:len(joined)]
                start, end = _strip_span(joined)
                if caps_equal(joined[start:end], joined_caps[start:end], wanted) or any(
                    caps_equal_stripped(item, wanted) for item in items if item.text.strip()
                ):
                    found = True
                    break
            if found:
                hits.append(row)
    if len(hits) != 1:
        where = f" in table {table_index}" if table_index else ""
        raise RefusedUnsafe(
            f"match {match!r} must equal one cell's text (or one paragraph of a cell) in exactly one table row{where}; "
            f"it matches {len(hits)} rows")
    return hits[0]


def plan_table_row(layout: Layout, op: str, match: str, cells: list[str], *, table_index: int | None = None) -> EditPlan:
    row = find_row(layout, match, table_index=table_index)
    plan = EditPlan(op=op, first_index=row.first_index, last_index=row.end_index)
    unit = _plan_row(plan, layout, op, row, list(cells), label=f"table {row.table + 1} row {row.row + 1}")
    plan.unit_lo, plan.unit_hi = row.unit, row.unit + 1
    plan.expected = [unit]
    plan.paragraph_delta = sum(len(cell) for cell in unit.lines) - sum(len(cell) for cell in row.cells)
    plan.report = {
        "table": row.table + 1, "row": row.row + 1,
        "cells_before": _row_texts(layout, row),
        "cells_before_stored": ["\n".join(layout.para(i).text for i in cell) for cell in row.cells],
        "cells_after": list(cells),
    }
    return plan


# --------------------------------------------------------------------------
# Sections.
# --------------------------------------------------------------------------

def _is_caps_heading(item: Para) -> bool:
    letters = [flag for ch, flag in zip(item.text, item.caps) if ch.isalpha()]
    return bool(letters) and all(letters) and len(item.text.strip()) <= 120


def _style_heading_level(layout: Layout, item: Para) -> int | None:
    if item.style and (item.style in layout.styles_outline):
        return layout.styles_outline[item.style]
    if item.style and item.style.lower().startswith(("heading", "title")):
        digits = "".join(ch for ch in item.style if ch.isdigit())
        return int(digits) - 1 if digits else 0
    return None


def _heading_kind(layout: Layout, item: Para):
    """("caps",) for a caps heading, ("style", level) for an outline heading, else None."""

    if item.row_end or not item.text.strip():
        return None
    if _is_caps_heading(item):
        return ("caps",)
    level = _style_heading_level(layout, item)
    return ("style", level) if level is not None else None


def _ends_section(start_kind, kind) -> bool:
    if kind is None:
        return False
    if start_kind == ("caps",):
        return kind == ("caps",)
    return kind[0] == "style" and kind[1] <= start_kind[1]


def find_heading(layout: Layout, heading: str, *, after: int = 0) -> Para:
    hits = [
        item for item in layout.paragraphs[after:]
        if not item.row_end and item.text.strip() and caps_equal_stripped(item, heading)
        and _heading_kind(layout, item) is not None
    ]
    if not hits:
        loose = [item.index for item in layout.paragraphs[after:] if not item.row_end and caps_equal_stripped(item, heading)]
        detail = f"; paragraphs {loose} have that text but are not headings" if loose else ""
        raise RefusedUnsafe(
            f"heading {heading!r} was not found (headings are caps/small-caps paragraphs or outline-level "
            f"styles; case is relaxed only inside caps runs){detail}")
    if len(hits) > 1:
        raise RefusedUnsafe(f"heading {heading!r} matches {len(hits)} headings (paragraphs {[i.index for i in hits]})")
    return hits[0]


@dataclass
class Section:
    heading: Para
    kind: str  # "rows" or "paragraphs"
    rows: list[Row] = field(default_factory=list)
    paragraphs: list[Para] = field(default_factory=list)
    end_heading: Para | None = None
    table: int | None = None


def locate_section(layout: Layout, heading: str, *, until_heading: str | None = None) -> Section:
    head = find_heading(layout, heading)
    start_kind = _heading_kind(layout, head)
    stop = find_heading(layout, until_heading, after=head.index) if until_heading else None
    if head.table is not None:
        if head.nested:
            raise RefusedUnsafe("replace_section: the heading is inside a nested table")
        rows = layout.tables[head.table]
        section = Section(head, "rows", table=head.table)
        for row in rows[head.row + 1:]:
            members = [layout.para(index) for cell in row.cells for index in cell]
            if stop is not None and any(item.index == stop.index for item in members):
                section.end_heading = stop
                break
            if stop is None:
                ender = next((item for item in members if _ends_section(start_kind, _heading_kind(layout, item))), None)
                if ender is not None:
                    section.end_heading = ender
                    break
            section.rows.append(row)
        if stop is not None and section.end_heading is None:
            raise RefusedUnsafe(f"until_heading {until_heading!r} is not in the same table as heading {heading!r}")
        return section
    section = Section(head, "paragraphs")
    for item in layout.paragraphs[head.index:]:
        if stop is not None and item.index == stop.index:
            section.end_heading = stop
            break
        if stop is None and _ends_section(start_kind, _heading_kind(layout, item)):
            section.end_heading = item
            break
        if item.table is not None:
            if stop is None and not section.paragraphs:
                # A heading paragraph directly above a table: the table's rows are the section.
                rows = layout.tables[item.table]
                section = Section(head, "rows", rows=list(rows), table=item.table)
                for position, row in enumerate(rows):
                    members = [layout.para(i) for cell in row.cells for i in cell]
                    ender = next((m for m in members if _ends_section(start_kind, _heading_kind(layout, m))), None)
                    if ender is not None:
                        section.rows = list(rows[:position])
                        section.end_heading = ender
                        break
                return section
            raise RefusedUnsafe(
                f"replace_section: section {heading!r} mixes body paragraphs and a table; "
                "use replace_table_row for the table, or until_heading to stop before it")
        section.paragraphs.append(item)
    return section


def plan_section(
    layout: Layout, op: str, heading: str, *, paragraphs: list[str] | None = None,
    rows: list[list[str]] | None = None, until_heading: str | None = None,
) -> EditPlan:
    section = locate_section(layout, heading, until_heading=until_heading)
    report = {
        "heading_index": section.heading.index,
        "heading_stored": section.heading.text,
        "heading_displayed": section.heading.displayed,
        "ends_at_heading": section.end_heading.displayed if section.end_heading else None,
    }
    if section.kind == "rows":
        if rows is None:
            raise RefusedUnsafe(
                f"{op}: section {heading!r} is table rows (table {section.table + 1}); give `rows` (cell texts per row)")
        plan = _plan_section_rows(layout, op, section, [list(row) for row in rows])
    else:
        if paragraphs is None:
            raise RefusedUnsafe(f"{op}: section {heading!r} is body paragraphs; give `paragraphs`")
        plan = _plan_section_paragraphs(layout, op, section, list(paragraphs))
    plan.report = {**report, **plan.report}
    return plan


def _plan_section_rows(layout: Layout, op: str, section: Section, new_rows: list[list[str]]) -> EditPlan:
    old_rows = section.rows
    heading_row = layout.row_of(section.heading.index) if section.heading.table is not None else None
    table_rows = layout.tables[section.table]
    plan = EditPlan(op=op)
    if not old_rows:
        raise RefusedUnsafe(f"{op}: section {section.heading.displayed!r} has no rows to use as a formatting template")
    for row in old_rows:
        if row.complex:
            raise RefusedUnsafe(f"{op}: table {row.table + 1} row {row.row + 1}: {row.complex}")
    old_keys = ["\x1f".join(_row_texts(layout, row)) for row in old_rows]
    new_keys = ["\x1f".join(row) for row in new_rows]
    matcher = difflib.SequenceMatcher(None, old_keys, new_keys, autojunk=False)
    expected: list[ExpUnit | None] = [None] * len(new_rows)
    codes = matcher.get_opcodes()
    for tag, i1, i2, j1, j2 in codes:
        if tag == "equal":
            for offset in range(i2 - i1):
                expected[j1 + offset] = _unchanged_row(layout, old_rows[i1 + offset])
    delta = 0
    for tag, i1, i2, j1, j2 in reversed(codes):
        if tag == "equal":
            continue
        paired = min(i2 - i1, j2 - j1)
        if j2 - j1 > paired:
            position = i1 + paired
            template = old_rows[position - 1] if position > 0 else old_rows[position]
            following = old_rows[position] if position < len(old_rows) else (
                table_rows[old_rows[-1].row + 1] if old_rows[-1].row + 1 < len(table_rows) else None)
            _refuse_template_objects(layout, op, template)
            appended = False
            for offset in reversed(range(j2 - j1 - paired)):
                if following is not None:
                    insert_at = following.first_index
                    plan.steps.append(DuplicateRowStep(template.first_index, following.first_index, None))
                elif not appended:
                    insert_at = table_rows[-1].end_index + 1
                    plan.steps.append(DuplicateRowStep(template.first_index, None, table_rows[-1].first_index))
                    appended = True
                else:
                    # Later (earlier-numbered) copies go above the first appended one.
                    plan.steps.append(DuplicateRowStep(template.first_index, insert_at, None))
                shift = insert_at - template.first_index
                expected[j1 + paired + offset] = _plan_row(
                    plan, layout, op, template, new_rows[j1 + paired + offset], base_shift=shift,
                    label=f"new row {j1 + paired + offset + 1} of the section (copied from table row {template.row + 1})")
                delta += _row_size(template) + _added_lines(expected[j1 + paired + offset], template)
        if i2 - i1 > paired:
            for row in reversed(old_rows[i1 + paired:i2]):
                _refuse_template_objects(layout, op, row, verb="delete")
                plan.steps.append(DeleteRowStep(row.first_index))
                delta -= _row_size(row)
        for offset in reversed(range(paired)):
            row = old_rows[i1 + offset]
            expected[j1 + offset] = _plan_row(
                plan, layout, op, row, new_rows[j1 + offset], label=f"table {row.table + 1} row {row.row + 1}")
            delta += _added_lines(expected[j1 + offset], row)
    plan.expected = [unit for unit in expected if unit is not None]
    plan.unit_lo, plan.unit_hi = old_rows[0].unit, old_rows[-1].unit + 1
    plan.first_index = old_rows[0].first_index
    plan.last_index = old_rows[-1].end_index
    plan.paragraph_delta = delta
    plan.report = {
        "section_kind": "rows",
        "rows_before": [_row_texts(layout, row) for row in old_rows],
        "rows_after": new_rows,
        "heading_row": heading_row.row + 1 if heading_row else None,
    }
    return plan


def _added_lines(unit: ExpUnit, row: Row) -> int:
    return sum(len(cell) for cell in unit.lines) - sum(len(cell) for cell in row.cells)


_OBJECTS = (("fldChar", "field"), ("fldSimple", "field"), ("sdt", "content control"),
            ("hyperlink", "hyperlink"), ("drawing", "drawing"), ("pict", "drawing"), ("object", "embedded object"),
            ("bookmarkStart", "bookmark"), ("commentRangeStart", "comment"), ("commentReference", "comment"),
            ("ins", "tracked revision"), ("del", "tracked revision"), ("footnoteReference", "footnote"),
            ("endnoteReference", "endnote"))


def _object_in(node: ET.Element | None) -> str | None:
    if node is None:
        return None
    for tag, name in _OBJECTS:
        found = node.find(f".//w:{tag}", NS)
        if found is not None:
            if tag == "bookmarkStart":
                return f"bookmark {found.get(W + 'name')!r}"
            return name
    return None


def _refuse_objects(node: ET.Element | None, op: str, action: str) -> None:
    what = _object_in(node)
    if what is not None:
        raise RefusedUnsafe(f"{op}: would {action}, which holds a {what}; edit around it instead")


def _refuse_template_objects(layout: Layout, op: str, row: Row, *, verb: str = "copy") -> None:
    _refuse_objects(row.node, op, f"{verb} table {row.table + 1} row {row.row + 1}")


def _plan_section_paragraphs(layout: Layout, op: str, section: Section, new_lines: list[str]) -> EditPlan:
    old = section.paragraphs
    plan = EditPlan(op=op)
    if not old:
        raise RefusedUnsafe(f"{op}: section {section.heading.displayed!r} has no paragraphs to use as a formatting template")
    if not new_lines:
        raise RefusedUnsafe(f"{op}: give at least one paragraph (an empty section would lose its formatting)")
    lines = _plan_cell(plan, op, old, new_lines)
    for step in plan.steps:
        if isinstance(step, DeleteParasStep) and step.trailing:
            # The paragraph after the section is the next heading, so nothing is "last in a cell" here.
            step.trailing = False
    plan.expected = [ExpUnit("p", [[line]]) for line in lines]
    plan.unit_lo = layout.unit_of(old[0].index)
    plan.unit_hi = layout.unit_of(old[-1].index) + 1
    plan.first_index, plan.last_index = old[0].index, old[-1].index
    plan.paragraph_delta = len(lines) - len(old)
    plan.report = {
        "section_kind": "paragraphs",
        "paragraphs_before": [item.displayed for item in old],
        "paragraphs_after": new_lines,
    }
    return plan


# --------------------------------------------------------------------------
# Word-save baseline preservation.
# --------------------------------------------------------------------------
# Edits are verified against Word's own save of the untouched document, so
# that save must first be shown to preserve the source. Tolerated, and only
# these: bookmark/comment/content-control/drawing ids renumbered, runs split
# or merged without changing any character's formatting, rsids, paraIds,
# proofing marks, empty property containers, and Word's own `_GoBack`
# bookmark appearing or going.

_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_SKIP_SUBTREES = {W + "pPr", W + "rPr", W + "sdtPr", W + "tblPr", W + "trPr", W + "tcPr", W + "sectPr"}
_WORD_OWN_BOOKMARKS = {"_GoBack"}


def _instr(text: str) -> str:
    return " ".join(text.split())


def _sdt_identity(node: ET.Element) -> str:
    props = node.find("w:sdtPr", NS)
    if props is None:
        return ""
    props = copy.deepcopy(props)
    for child in props.findall("w:id", NS):
        props.remove(child)
    return canonical(props)


def anchor_inventory(path: str | Path) -> dict[str, list]:
    """Bookmarks, comments, fields, content controls and hyperlinks with their text extents.

    Extents are (paragraph ordinal, character offset) pairs in document
    order, so they survive id renumbering and run splitting but not a
    bookmark collapsing or a comment moving.
    """

    snapshot = read_snapshot(path)
    with zipfile.ZipFile(BytesIO(snapshot.package)) as archive:
        names = set(archive.namelist())
        root = ET.fromstring(archive.read("word/document.xml"))
        comments = ET.fromstring(archive.read("word/comments.xml")) if "word/comments.xml" in names else None
        rels_name = "word/_rels/document.xml.rels"
        rels = ET.fromstring(archive.read(rels_name)) if rels_name in names else None
    targets = {}
    if rels is not None:
        for rel in rels:
            targets[rel.get("Id")] = (rel.get("Target"), rel.get("TargetMode"))
    comment_keys = {}
    if comments is not None:
        for comment in comments.findall("w:comment", NS):
            text = "".join(node.text or "" for node in comment.iter(W + "t"))
            comment_keys[comment.get(W + "id")] = (comment.get(W + "author"), comment.get(W + "initials"), text)
    state = {"para": 0, "offset": 0}
    bookmarks: dict[str, list] = {}
    open_bookmarks: dict[str, tuple] = {}
    comment_starts: dict[str, tuple] = {}
    comment_ranges: list = []
    fields: list = []
    field_stack: list[dict] = []
    controls: list = []
    links: list = []

    def here() -> tuple[int, int]:
        return state["para"], state["offset"]

    def visit(node: ET.Element) -> None:
        tag = node.tag
        if tag in _SKIP_SUBTREES:
            return
        if tag == W + "p":
            state["para"] += 1
            state["offset"] = 0
        elif tag == W + "t":
            state["offset"] += len(node.text or "")
        elif tag in (W + "tab", W + "br", W + "cr"):
            state["offset"] += 1
        elif tag == W + "bookmarkStart":
            open_bookmarks[node.get(W + "id")] = (
                node.get(W + "name"), here(), node.get(W + "colFirst"), node.get(W + "colLast"))
        elif tag == W + "bookmarkEnd":
            started = open_bookmarks.pop(node.get(W + "id"), None)
            if started is not None:
                name, start, first, last = started
                bookmarks.setdefault(name, []).append((start, here(), first, last))
        elif tag == W + "commentRangeStart":
            comment_starts[node.get(W + "id")] = here()
        elif tag == W + "commentRangeEnd":
            key = comment_keys.get(node.get(W + "id"), ("?", None, ""))
            comment_ranges.append(("range", key, comment_starts.pop(node.get(W + "id"), None), here()))
        elif tag == W + "commentReference":
            key = comment_keys.get(node.get(W + "id"), ("?", None, ""))
            comment_ranges.append(("reference", key, here(), None))
        elif tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                field_stack.append({"begin": here(), "instr": []})
            elif kind == "end" and field_stack:
                item = field_stack.pop()
                fields.append(("complex", _instr("".join(item["instr"])), item["begin"], here()))
        elif tag == W + "instrText":
            if field_stack:
                field_stack[-1]["instr"].append(node.text or "")
        if tag in (W + "fldSimple", W + "sdt", W + "hyperlink"):
            start = here()
            for child in node:
                visit(child)
            if tag == W + "fldSimple":
                fields.append(("simple", _instr(node.get(W + "instr", "")), start, here()))
            elif tag == W + "sdt":
                controls.append((_sdt_identity(node), start, here()))
            else:
                links.append((targets.get(node.get(_R + "id")), node.get(W + "anchor"), start, here()))
            return
        for child in node:
            visit(child)

    body = root.find("w:body", NS)
    if body is not None:
        visit(body)
    for name, start, _, _ in open_bookmarks.values():
        bookmarks.setdefault(name, []).append((start, None, None, None))
    for name in _WORD_OWN_BOOKMARKS:
        bookmarks.pop(name, None)
    return {
        "bookmarks": sorted((name, extents) for name, extents in bookmarks.items()),
        "comments": sorted(comment_ranges, key=repr),
        "fields": fields,
        "content controls": controls,
        "hyperlinks": links,
    }


def _first_difference(old: list, new: list) -> str:
    for a, b in zip(old, new):
        if a != b:
            return f"{a!r} became {b!r}"
    if len(old) != len(new):
        extra = old[len(new)] if len(old) > len(new) else new[len(old)]
        return f"{len(old)} -> {len(new)} ({'lost' if len(old) > len(new) else 'gained'} {extra!r})"
    return "differs"


def baseline_preservation_errors(source: str | Path, baseline: str | Path) -> list[str]:
    """Differences between a document and Word's save of it beyond the tolerated normalization."""

    before, after = read_layout(source), read_layout(baseline)
    problems: list[str] = []
    if len(before.paragraphs) != len(after.paragraphs) or len(before.units) != len(after.units):
        return [f"paragraph/row structure changed ({len(before.paragraphs)} -> {len(after.paragraphs)} paragraphs, "
                f"{len(before.units)} -> {len(after.units)} body elements)"]
    for old, new in zip(before.paragraphs, after.paragraphs):
        if old.row_end != new.row_end or old.text != new.text:
            problems.append(f"paragraph {old.index} text changed")
            continue
        if old.fps != new.fps or old.caps != new.caps:
            offset = next(k for k in range(len(old.text)) if old.fps[k] != new.fps[k] or old.caps[k] != new.caps[k])
            problems.append(
                f"paragraph {old.index} character {offset} ({old.text[offset]!r}) formatting "
                f"{_short(old.fps[offset])} became {_short(new.fps[offset])}")
        if old.ppr() != new.ppr():
            problems.append(f"paragraph {old.index} ({old.text[:40]!r}) paragraph properties changed")
    for position, (old, new) in enumerate(zip(before.units, after.units)):
        if old.kind != new.kind:
            problems.append(f"unit {position + 1} changed kind ({old.kind} -> {new.kind})")
        elif old.kind == "tbl" and old.key() != new.key():
            problems.append(f"unit {position + 1}: table properties or grid changed")
    for old_rows, new_rows in zip(before.tables, after.tables):
        for old, new in zip(old_rows, new_rows):
            if _trpr(old) != _trpr(new) or _tcprs(old) != _tcprs(new):
                problems.append(f"table {old.table + 1} row {old.row + 1}: row or cell properties changed")
    old_anchors, new_anchors = anchor_inventory(source), anchor_inventory(baseline)
    for kind in old_anchors:
        if old_anchors[kind] != new_anchors[kind]:
            problems.append(f"{kind}: {_first_difference(old_anchors[kind], new_anchors[kind])}")
    return problems


# --------------------------------------------------------------------------
# Bookmark and comment extents (D11). Checked on every staged candidate.
# --------------------------------------------------------------------------
# Positions are offsets into the document's text stream: every layout
# paragraph's text followed by one mark (a row-end marker is an empty
# paragraph), the characters `_walk_paragraph` reads. An edit may move an
# anchor only as Word's own typing would (measured in Word 2026-10-30,
# scratch/resume_fix_328/r3/probe_insert*.py):
#
# * text inserted at a position lands inside a bookmark that starts there,
#   outside a comment that starts there, and outside any bookmark or comment
#   that ends there; a collapsed bookmark stays before it;
# * text that replaces characters belongs to the anchors that held them, so
#   an anchor starting at or ending after the replaced characters covers the
#   new text; a boundary strictly inside the replaced characters may land
#   anywhere in the new text;
# * an anchor whose characters are all deleted disappears in Word; that is refused
#   (a coauthor's comment or a cross-reference bookmark is never dropped silently);
# * inserted table rows may fall on either side of a boundary they touch.
#
# Unchanged anchored characters therefore always stay anchored.

_ANCHOR_MARKS = {W + "bookmarkStart", W + "bookmarkEnd", W + "commentRangeStart", W + "commentRangeEnd",
                 W + "commentReference"}


def _paragraph_starts(lengths: list[int]) -> list[int]:
    starts, position = [], 0
    for length in lengths:
        starts.append(position)
        position += length + 1
    return starts


def stream_text(layout: Layout) -> str:
    return "".join(item.text + "\r" for item in layout.paragraphs)


def anchor_extents(layout: Layout) -> dict:
    """Bookmarks (name -> (start, end)) and comments ([(key, start, end)]) as stream offsets."""

    starts = _paragraph_starts([len(item.text) for item in layout.paragraphs])
    own = {id(item.node): item.index for item in layout.paragraphs if item.node is not None}
    cursor = [0]
    marks: list[tuple[str, str | None, str | None, int]] = []

    def visit(node: ET.Element) -> None:
        tag = node.tag
        if tag in _SKIP_SUBTREES:
            return
        index = own.get(id(node)) if tag == W + "p" else None
        if index is not None:
            cursor[0] = starts[index - 1]
        elif tag == W + "t":
            cursor[0] += len(node.text or "")
        elif tag in (W + "tab", W + "br", W + "cr"):
            cursor[0] += 1
        elif tag in _ANCHOR_MARKS:
            marks.append((tag, node.get(W + "id"), node.get(W + "name"), cursor[0]))
        for child in node:
            visit(child)
        if index is not None:
            width = len(layout.para(index).text)
            if cursor[0] - starts[index - 1] != width:
                raise VerificationFailed(
                    f"cannot measure bookmark and comment positions in paragraph {index} "
                    f"({cursor[0] - starts[index - 1]} characters counted, {width} expected)")
            cursor[0] = starts[index - 1] + width + 1
        elif tag == W + "tr":
            cursor[0] += 1  # the row-end marker

    if layout.body is not None:
        visit(layout.body)
    bookmarks: dict[str, tuple[int, int | None]] = {}
    open_bookmarks: dict[str | None, tuple[str, int]] = {}
    comments: dict[str | None, dict[str, int]] = {}
    for tag, mark_id, name, position in marks:
        if tag == W + "bookmarkStart":
            open_bookmarks[mark_id] = (name or "", position)
        elif tag == W + "bookmarkEnd":
            started = open_bookmarks.pop(mark_id, None)
            if started is not None and started[0] not in _WORD_OWN_BOOKMARKS:
                bookmarks[started[0]] = (started[1], position)
        else:
            role = {W + "commentRangeStart": "start", W + "commentRangeEnd": "end"}.get(tag, "reference")
            comments.setdefault(mark_id, {}).setdefault(role, position)
    for name, start in open_bookmarks.values():
        if name not in _WORD_OWN_BOOKMARKS:
            bookmarks[name] = (start, None)
    ranges = []
    for mark_id, seen in comments.items():
        start = seen.get("start", seen.get("reference", seen.get("end")))
        end = seen.get("end", seen.get("reference", start))
        ranges.append((layout.comment_keys.get(mark_id, ("?", None, "")), start, end))
    ranges.sort(key=lambda item: (item[1], item[2]))
    return {"bookmarks": bookmarks, "comments": ranges}


@dataclass
class StreamEdit:
    """`removed` characters at `position` replaced by `added` new ones."""

    position: int
    removed: int
    added: int
    either_side: bool = False  # an inserted table row


def map_anchor_point(point: int, role: str, edit: StreamEdit) -> tuple[int, int]:
    """Where an anchor boundary may be after `edit`, as an inclusive (low, high) range.

    `role` says where the boundary ends up when text is inserted exactly at
    it: "before" the new text (a bookmark's start, so the bookmark takes it;
    every end, so the anchor does not; both ends of a collapsed anchor) or
    "after" it (a comment's start, so the comment does not take it).
    """

    at, removed, added = edit.position, edit.removed, edit.added
    if point < at:
        return point, point
    if point > at + removed:
        return point - removed + added, point - removed + added
    if not removed:
        if edit.either_side:
            return at, at + added
        return (at, at) if role == "before" else (at + added, at + added)
    if point == at:
        return at, at
    if point == at + removed:
        return at + added, at + added
    return at, at + added


def _row_members(rows: list, index: int) -> tuple[int, int]:
    key = rows[index - 1]
    if key is None:
        raise VerificationFailed(f"paragraph {index} is not in a table row")
    low = high = index
    while low > 1 and rows[low - 2] == key:
        low -= 1
    while high < len(rows) and rows[high] == key:
        high += 1
    return low, high


def plan_stream_edits(layout: Layout, plan: EditPlan) -> tuple[list[StreamEdit], list[int]]:
    """The plan's steps as stream edits, in the order Word performs them, and the final paragraph lengths."""

    lengths = [len(item.text) for item in layout.paragraphs]
    rows: list = [(item.table, item.row) if item.table is not None else None for item in layout.paragraphs]
    edits: list[StreamEdit] = []
    copies = 0

    def start(index: int) -> int:
        return sum(lengths[:index - 1]) + index - 1

    for step in plan.steps:
        if isinstance(step, TextStep):
            if lengths[step.index - 1] != len(step.text):
                raise VerificationFailed(f"{plan.op}: cannot follow paragraph {step.index} through the edit plan")
            base = start(step.index)
            for edit in reversed(step.edits):
                edits.append(StreamEdit(base + edit.start, edit.end - edit.start, len(edit.text)))
            lengths[step.index - 1] += sum(len(edit.text) - (edit.end - edit.start) for edit in step.edits)
        elif isinstance(step, InsertLinesStep):
            at = start(step.index) + (len(step.text) if step.where == "after" else 0)
            edits.append(StreamEdit(at, 0, sum(len(line) + 1 for line in step.lines)))
            spot = step.index - 1 if step.where == "before" else step.index
            lengths[spot:spot] = [len(line) for line in step.lines]
            rows[spot:spot] = [rows[step.index - 1]] * len(step.lines)
        elif isinstance(step, DeleteParasStep):
            removed = sum(lengths[step.first - 1:step.last]) + step.last - step.first + 1
            at = start(step.previous) + lengths[step.previous - 1] if step.trailing else start(step.first)
            edits.append(StreamEdit(at, removed, 0))
            del lengths[step.first - 1:step.last], rows[step.first - 1:step.last]
        elif isinstance(step, DuplicateRowStep):
            low, high = _row_members(rows, step.template)
            copied = lengths[low - 1:high]
            spot = step.before if step.before is not None else _row_members(rows, step.append_after)[1] + 1
            edits.append(StreamEdit(start(spot), 0, sum(length + 1 for length in copied), either_side=True))
            copies += 1
            lengths[spot - 1:spot - 1] = copied
            rows[spot - 1:spot - 1] = [("copy", copies)] * len(copied)
        elif isinstance(step, DeleteRowStep):
            low, high = _row_members(rows, step.index)
            edits.append(StreamEdit(start(low), sum(lengths[low - 1:high]) + high - low + 1, 0))
            del lengths[low - 1:high], rows[low - 1:high]
        else:  # pragma: no cover - planner and checker are versioned together
            raise VerificationFailed(f"{plan.op}: unknown plan step {type(step).__name__}")
    return edits, lengths


@dataclass
class _Expected:
    start: tuple[int, int]
    end: tuple[int, int]
    may_vanish: bool = False


def _expect(start: int, end: int | None, start_role: str, edits: list[StreamEdit]) -> _Expected:
    end = start if end is None else end
    if start == end:
        start_role = "before"  # a collapsed anchor stays before typed text
    low_start = high_start = start
    low_end = high_end = end
    vanish = False
    for edit in edits:
        if edit.removed and edit.position <= low_start and high_end <= edit.position + edit.removed and not (
                low_start == high_start == edit.position and low_end == high_end == edit.position + edit.removed
                and edit.added):
            vanish = True  # every character of the anchor is deleted (a whole replaced span keeps it)
        low_start = map_anchor_point(low_start, start_role, edit)[0]
        high_start = map_anchor_point(high_start, start_role, edit)[1]
        low_end = map_anchor_point(low_end, "before", edit)[0]
        high_end = map_anchor_point(high_end, "before", edit)[1]
    return _Expected((low_start, high_start), (low_end, high_end), vanish)


def _span_text(text: str, start: int, end: int | None) -> str:
    return text[start:end if end is not None else start].replace("\r", "¶").replace("\x07", "")


def anchor_extent_errors(before: Layout, after: Layout, plan: EditPlan) -> list[str]:
    """Bookmarks and comments whose extent in `after` is not their `before` extent mapped through the plan."""

    edits, lengths = plan_stream_edits(before, plan)
    return anchor_errors_for_edits(before, after, edits, lengths)


def anchor_errors_for_edits(before: Layout, after: Layout, edits: list[StreamEdit], lengths: list[int]) -> list[str]:
    """Bookmarks and comments whose extent in `after` is not their `before` extent mapped through `edits`.

    `edits` are applied in order, each in the coordinates left by the ones
    before it; `lengths` are the paragraph lengths they should leave. Used by
    the plan checker above and by the direct-XML revision engine.
    """

    if lengths != [len(item.text) for item in after.paragraphs]:
        return ["the document's paragraphs do not line up with the edit plan, so bookmark and comment "
                "positions cannot be checked"]
    old, new = anchor_extents(before), anchor_extents(after)
    text, old_text = stream_text(after), stream_text(before)
    problems: list[str] = []

    def check(label: str, original: tuple, wanted: _Expected, actual: tuple | None) -> None:
        was = _span_text(old_text, original[0], original[1])
        if actual is None:
            if wanted.may_vanish:
                problems.append(f"{label} on {was!r} would be deleted because all of its text is deleted; "
                                "remove or move it first, or keep some of its text")
            else:
                problems.append(f"{label} on {was!r} was lost")
            return
        start, end = actual[0], actual[1] if actual[1] is not None else actual[0]
        if wanted.start[0] <= start <= wanted.start[1] and wanted.end[0] <= end <= wanted.end[1]:
            return
        if wanted.start[0] == wanted.start[1] and wanted.end[0] == wanted.end[1]:
            should = repr(_span_text(text, wanted.start[0], wanted.end[0]))
        else:
            should = f"characters {wanted.start} to {wanted.end} of the edited text"
        problems.append(f"{label} would cover {_span_text(text, start, end)!r} instead of {should} (was {was!r})")

    for name, extent in sorted(old["bookmarks"].items()):
        wanted = _expect(extent[0], extent[1], "before", edits)
        check(f"bookmark {name!r}", extent, wanted, new["bookmarks"].get(name))
    for name in sorted(set(new["bookmarks"]) - set(old["bookmarks"])):
        problems.append(f"bookmark {name!r} appeared on {_span_text(text, *new['bookmarks'][name])!r}")
    pending: dict[tuple, list] = {}
    for key, start, end in new["comments"]:
        pending.setdefault(key, []).append((start, end))
    for key, start, end in old["comments"]:
        wanted = _expect(start, end, "after", edits)
        found = pending.get(key) or []
        actual = None
        for candidate in found:
            if wanted.start[0] <= candidate[0] <= wanted.start[1] and wanted.end[0] <= candidate[1] <= wanted.end[1]:
                actual = candidate
                break
        if actual is None and found and not wanted.may_vanish:
            actual = found[0]
        if actual is not None:
            found.remove(actual)
        check(f"comment ({key[0]}: {key[2][:40]!r})", (start, end), wanted, actual)
    for key, found in pending.items():
        for start, end in found:
            problems.append(f"comment ({key[0]}: {key[2][:40]!r}) appeared on {_span_text(text, start, end)!r}")
    return problems


# --------------------------------------------------------------------------
# Verification (D11). Called on every staged candidate; no parameter skips it.
# --------------------------------------------------------------------------

def _fail(message: str, expected=None, actual=None):
    raise VerificationFailed(message, expected=expected, actual=actual)


def _unit_paragraphs(layout: Layout, position: int) -> list[list[Para]]:
    unit = layout.units[position]
    if unit.kind == "p":
        return [[layout.para(unit.first)]]
    row = layout.row_at_unit(position)
    return [[layout.para(index) for index in cell] for cell in row.cells]


def verify_plan(source: str | Path, candidate: str | Path, plan: EditPlan) -> dict:
    before = read_layout(source)
    after = read_layout(candidate)
    lo, hi = plan.unit_lo, plan.unit_hi
    count = len(plan.expected)
    old_units, new_units = before.units, after.units
    if len(new_units) != len(old_units) - (hi - lo) + count:
        _fail(f"{plan.op}: the document's paragraph/row structure changed outside the edited range",
              expected=len(old_units) - (hi - lo) + count, actual=len(new_units))
    for position in range(lo):
        if old_units[position].key() != new_units[position].key():
            _fail(f"{plan.op}: content before the edit changed (unit {position + 1}, {old_units[position].kind})")
    for offset in range(len(old_units) - hi):
        old_unit, new_unit = old_units[hi + offset], new_units[lo + count + offset]
        if old_unit.key() != new_unit.key():
            _fail(f"{plan.op}: content after the edit changed (unit {hi + offset + 1}, {old_unit.kind})")
    touched = 0
    for number, wanted in enumerate(plan.expected):
        unit = new_units[lo + number]
        if unit.kind != ("row" if wanted.kind == "row" else "p"):
            _fail(f"{plan.op}: expected a {wanted.kind} at edited position {number + 1}, found {unit.kind}")
        cells = _unit_paragraphs(after, lo + number)
        label = wanted.label or f"edited {wanted.kind} {number + 1}"
        if wanted.kind == "row":
            if len(cells) != len(wanted.lines):
                _fail(f"{plan.op}: {label} has {len(cells)} cells; expected {len(wanted.lines)}")
            row = after.row_at_unit(lo + number)
            if _tcprs(row) != wanted.tcprs:
                _fail(f"{plan.op}: {label}: cell properties changed")
            if _trpr(row) != wanted.trpr:
                _fail(f"{plan.op}: {label}: row properties changed")
        for cell_number, (paras, lines) in enumerate(zip(cells, wanted.lines)):
            where = f"{label}, cell {cell_number + 1}" if wanted.kind == "row" else label
            if len(paras) != len(lines):
                _fail(f"{plan.op}: {where} has {len(paras)} paragraphs; expected {len(lines)}",
                      expected=[line.text for line in lines], actual=[p.text for p in paras])
            for para, line in zip(paras, lines):
                if line.requested is None and line.canon is not None:
                    if canonical(para.node) != line.canon:
                        _fail(f"{plan.op}: {where}: unrequested change to paragraph {para.text!r}")
                    continue
                touched += 1
                if para.text != line.text:
                    _fail(f"{plan.op}: {where}: stored text differs from the plan", expected=line.text, actual=para.text)
                if line.requested is not None and not caps_equal(para.text, para.caps, line.requested):
                    _fail(f"{plan.op}: {where}: result text differs from the requested text",
                          expected=line.requested, actual=para.displayed)
                for offset, (want, got) in enumerate(zip(line.fps, para.fps)):
                    if want is not None and want != got:
                        _fail(f"{plan.op}: {where}: character {offset} ({para.text[offset]!r}) has formatting "
                              f"{_short(got)} instead of {_short(want)}")
                if line.origin is not None:
                    _check_kept_formatting(plan.op, where, line.origin, para)
                if line.ppr is not None and para.ppr(mark=line.mark) != line.ppr:
                    _fail(f"{plan.op}: {where}: paragraph properties changed for {para.text!r}",
                          expected=line.ppr, actual=para.ppr(mark=line.mark))
            if wanted.kind == "row":
                old_empty = all(not line.text.strip() for line in lines)
                if all(not para.text.strip() for para in paras) and not old_empty:
                    _fail(f"{plan.op}: {where} went empty without being asked")
                requested = [line.requested if line.requested is not None else line.text for line in lines]
                if any(line.strip() for line in requested) and all(not para.text.strip() for para in paras):
                    _fail(f"{plan.op}: {where} went empty, which was not requested")
                if wanted.dates and wanted.dates[cell_number] and not is_date_like("\n".join(p.displayed for p in paras)):
                    _fail(f"{plan.op}: {where} held a date and no longer reads as one: "
                          f"{' / '.join(p.displayed for p in paras)!r}")
    anchors = anchor_extent_errors(before, after, plan)
    if anchors:
        _fail(f"{plan.op}: the edit would move a bookmark or comment: " + "; ".join(anchors[:5]), actual=anchors[:20])
    marked = anchor_extents(before)
    return {"consistency_check": "passed", "touched_paragraphs": touched,
            "units_verified_unchanged": len(old_units) - (hi - lo),
            "anchors_verified": len(marked["bookmarks"]) + len(marked["comments"])}


_LCS_CELLS = 1_000_000


def _lcs(a: list, b: list) -> int:
    previous = [0] * (len(b) + 1)
    for item in a:
        current = [0]
        for position, other in enumerate(b):
            current.append(previous[position] + 1 if item == other else max(previous[position + 1], current[-1]))
        previous = current
    return previous[-1]


def _kept_pairs(old: str, new: str) -> list[tuple[int, int]]:
    """Which source characters an edit leaves unchanged, from the two texts alone.

    Unchanged words (a word-level diff), plus, inside a word rewritten into
    one word, the characters a character-level diff keeps. Coincidental
    letters shared by unrelated words of a rewritten phrase do not count.
    """

    old_spans, new_spans = _tokens(old), _tokens(new)
    matcher = difflib.SequenceMatcher(
        None, [old[a:b] for a, b in old_spans], [new[a:b] for a, b in new_spans], autojunk=False)
    pairs: list[tuple[int, int]] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for (a, b), (c, _) in zip(old_spans[i1:i2], new_spans[j1:j2]):
                pairs.extend((a + k, c + k) for k in range(b - a))
        elif tag == "replace" and i2 - i1 == j2 - j1:
            for (a, b), (c, d) in zip(old_spans[i1:i2], new_spans[j1:j2]):
                if old[a].isalnum() and new[c].isalnum():
                    words = difflib.SequenceMatcher(None, old[a:b], new[c:d], autojunk=False)
                    for block in words.get_matching_blocks():
                        pairs.extend((a + block.a + k, c + block.b + k) for k in range(block.size))
    return pairs


def _check_kept_formatting(op: str, where: str, origin: tuple[str, tuple[str, ...]], para: Para) -> None:
    """Every source character the edit leaves unchanged keeps its full formatting.

    Independent of the plan: `_kept_pairs` decides from the two texts which
    characters are unchanged, and the result must be able to share at least
    that many characters with the source *with* their formatting (run
    properties plus field/hyperlink/content-control context). The sharing is
    measured as a longest common subsequence, so it does not depend on which
    of several equally good alignments the edit used. A character that
    survived the edit but lost its bold fails here even when the plan
    expected it to.
    """

    old_text, old_fps = origin
    old = list(zip(old_text, old_fps))
    new = list(zip(para.text, para.fps))
    kept = _kept_pairs(old_text, para.text)
    head = 0
    while head < min(len(old), len(new)) and old[head] == new[head]:
        head += 1
    tail = 0
    while tail < min(len(old), len(new)) - head and old[-1 - tail] == new[-1 - tail]:
        tail += 1
    old_mid, new_mid = old[head:len(old) - tail], new[head:len(new) - tail]
    if not old_mid or not new_mid:
        shared = head + tail
    elif len(old_mid) * len(new_mid) <= _LCS_CELLS:
        shared = head + tail + _lcs(old_mid, new_mid)
    else:  # very long rewrites: a greedy alignment, which can only undercount (fail safe)
        shared = head + tail + sum(block.size for block in difflib.SequenceMatcher(
            None, old_mid, new_mid, autojunk=False).get_matching_blocks())
    if shared >= len(kept):
        return
    detail = ""
    for i, j in kept:
        if old_fps[i] != para.fps[j]:
            detail = (f"; e.g. character {i} ({old_text[i]!r}) was {_short(old_fps[i])} "
                      f"and is now {_short(para.fps[j])}")
            break
    _fail(f"{op}: {where}: {len(kept) - shared} unchanged character(s) lost their formatting{detail}",
          expected=old_text, actual=para.text)


def _short(fp: str | None) -> str:
    if fp is None:
        return "(any)"
    context, _, rpr = fp.partition("|")
    tags = []
    if rpr:
        for node in ET.fromstring(rpr):
            tag = node.tag.rsplit("}", 1)[-1]
            value = node.get(W + "val")
            tags.append(f"{tag}={value}" if value not in (None, "true", "1") and tag not in ("rFonts",) else tag)
    where = {"h": "hyperlink", "s": "content control", "f": "field"}
    inside = ",".join(where[ch] for ch in context if ch in where)
    return f"[{' '.join(tags) or 'plain'}{' in ' + inside if inside else ''}]"


def consistency_errors(source: str | Path, candidate: str | Path) -> list[str]:
    """Batch-level D11 check over every table whose shape is unchanged.

    Every touched row keeps its cell count and every date cell still reads as
    a date. Tables are paired by position among tables with the same grid; a
    table whose row count changed was rewritten by a section op, which
    verified itself.
    """

    before, after = read_layout(source), read_layout(candidate)
    problems: list[str] = []
    old_sig = [len(rows) for rows in before.tables]
    new_sig = [len(rows) for rows in after.tables]
    matcher = difflib.SequenceMatcher(None, old_sig, new_sig, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            continue
        for offset in range(i2 - i1):
            old_rows, new_rows = before.tables[i1 + offset], after.tables[j1 + offset]
            for old_row, new_row in zip(old_rows, new_rows):
                where = f"table {j1 + offset + 1} row {new_row.row + 1}"
                if len(old_row.cells) != len(new_row.cells):
                    problems.append(f"{where}: cell count {len(old_row.cells)} -> {len(new_row.cells)}")
                    continue
                for number, (old_cell, new_cell) in enumerate(zip(old_row.cells, new_row.cells)):
                    old_text = _cell_text(before, old_cell)
                    new_text = _cell_text(after, new_cell)
                    if old_text == new_text:
                        continue
                    if is_date_like(old_text) and not is_date_like(new_text):
                        problems.append(f"{where} cell {number + 1}: date {old_text!r} became {new_text!r}")
    return problems
