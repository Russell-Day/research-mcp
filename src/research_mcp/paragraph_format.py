"""Paragraph formatting as direct OOXML, tracked or plain (plan 916 S3, decision 5).

``format_paragraphs`` sets font, size, indents, spacing and list numbering on one paragraph or a range of
paragraphs of one container. Word never opens: only the touched ``w:pPr`` and ``w:rPr`` elements change,
every other byte of ``word/document.xml`` and every other zip entry is copied unchanged (``revisions.rezip``).

What is formatted: the paragraph's own properties (``w:numPr``, ``w:ind``, ``w:spacing``), its mark's run
properties and every run of visible text (inside a ``w:ins`` too). Runs inside a ``w:del`` and the runs of a
field (code and result: citations, cross-references) keep their properties. A property is written only where
the paragraph's or run's effective value (direct formatting over the numbering level, the paragraph and
character styles and the document defaults) differs from the requested one.

``like_paragraph`` copies that paragraph's effective indents, spacing and numbering, and the font and size
that most of its visible characters show; explicit fields win over it.

Tracked: a ``w:pPrChange`` per paragraph whose properties change (holding its old ``w:pPr``), a
``w:rPrChange`` on the mark and on each changed run (holding its old ``w:rPr``), by ``author``, ids ascending
in document order (mark, paragraph, runs). As in ``tracked_format`` (plan 864 S6): a run inside the author's
own pending ``w:ins`` takes the formatting in place (L1); a run, mark or paragraph that already carries the
author's own formatting change keeps it, since its old properties are still the state before the revision
(L3); another author's formatting change is refused, naming it. Plain: the new formatting only; text that
carries a tracked formatting change is refused (it would be undone by rejecting that change).

Verified before anything is written: bytes outside the formatted paragraphs identical; every existing
revision byte-identical except a ``w:ins`` holding a formatted run (same start tag, same content apart from
run properties); accepted and original text unchanged; Reject of the new revisions gives the input paragraph
(canonical XML); Accept gives the requested effective values and leaves every other paragraph and run
property unchanged; new ids unique in the package.
"""

from __future__ import annotations

import copy
import datetime as dt
import re
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from xml.sax.saxutils import escape

from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import W
from .revisions import (
    DEFAULT_AUTHOR, REV_TAGS, Node, ParagraphModel, _et_final_chars, _fragment, _result,
    _verify_written, _Writer, _WRAPPERS, attr_value, describe_wrapper, et_resolve, load, next_revision_id,
    parse, revision_stamp, rezip,
)
from .tracked_format import ppr_change, ppr_rank, rpr_change, rpr_rank

OP = "format_paragraph"
MAX_INDENT_PT = 1584  # Word's limit (22 inches)
MAX_SIZE_PT = 1638
_KEYS = ("left", "right", "first", "before", "after")
# Attributes a requested value replaces (the alternatives Word reads for the same property go with it).
_IND_ATTRS = {"left": ("w:left", "w:start", "w:leftChars", "w:startChars"),
              "right": ("w:right", "w:end", "w:rightChars", "w:endChars"),
              "first": ("w:firstLine", "w:hanging", "w:firstLineChars", "w:hangingChars")}
_SPACING_ATTRS = {"before": ("w:before", "w:beforeLines", "w:beforeAutospacing"),
                  "after": ("w:after", "w:afterLines", "w:afterAutospacing")}
_FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:cs", "w:asciiTheme", "w:hAnsiTheme", "w:cstheme")
_MOVES = ("w:moveFrom", "w:moveTo")
_ATTR = re.compile(r'([\w:.-]+)\s*=\s*(["\'])(.*?)\2', re.S)
_ON = {"1", "true", "on"}


# --------------------------------------------------------------------------
# Raw element helpers.
# --------------------------------------------------------------------------

def _children(raw: str, tag: str) -> tuple[str, list[tuple[str, str]]]:
    """Open start tag and (name, raw) children of the element `raw` ("" -> a new empty `tag`)."""

    if not raw:
        return f"<{tag}>", []
    tree = parse(raw).children[0]
    head = re.sub(r"\s*/>\Z", ">", tree.head(raw)) if tree.closed else tree.head(raw)
    return head, [(item.name, raw[item.start:item.end]) for item in tree.children]


def _close(head: str, children: list[tuple[str, str]], tag: str) -> str:
    return f"{head}{''.join(raw for _, raw in children)}</{tag}>"


def _attrs_of(raw: str) -> list[tuple[str, str]]:
    if not raw:
        return []
    tree = parse(raw).children[0]
    head = tree.head(raw)
    rest = head[re.match(r"<\s*[\w:.-]+", head).end():]
    return [(name, value if quote == '"' else value.replace('"', "&quot;")) for name, quote, value in
            _ATTR.findall(rest)]


def _updated(raw: str, tag: str, updates: dict[str, str | None]) -> str:
    """The empty element `tag` with `raw`'s attributes, `updates` applied (None removes one)."""

    attrs = [(name, value) for name, value in _attrs_of(raw) if name not in updates]
    attrs += [(name, value) for name, value in updates.items() if value is not None]
    return f"<{tag}" + "".join(f' {name}="{value}"' for name, value in attrs) + "/>"


def _put(children: list[tuple[str, str]], name: str, raw: str | None, rank) -> list[tuple[str, str]]:
    """`children` with `name` replaced by `raw` (None removes it) at its schema position."""

    children = [item for item in children if item[0] != name]
    if raw is not None:
        at = next((number for number, item in enumerate(children) if rank(item[0]) > rank(name)), len(children))
        children.insert(at, (name, raw))
    return children


def _child_raw(children: list[tuple[str, str]], name: str) -> str:
    return next((raw for item, raw in children if item == name), "")


def _attr_text(value: str) -> str:
    return escape(value, {'"': "&quot;"})


def _tw(value: float) -> int:
    return int(round(value * 20))


def _pt(twips_or_half: int, unit: int) -> str:
    value = twips_or_half / unit
    return f"{value:g}"


# --------------------------------------------------------------------------
# Effective properties: document defaults, styles, numbering, theme fonts.
# --------------------------------------------------------------------------

def _int(element: ET.Element | None, name: str) -> int | None:
    if element is None:
        return None
    value = element.get(W + name)
    if value is None:
        return None
    try:
        return int(round(float(value)))
    except ValueError:
        return None


class Styles:
    """What a paragraph or run inherits: docDefaults, the style chain (w:basedOn), list levels, theme fonts."""

    def __init__(self, parts: dict[str, bytes]) -> None:
        self.styles: dict[str, ET.Element] = {}
        self.default_para: str | None = None
        self.ppr_default = self.rpr_default = None
        if "word/styles.xml" in parts:
            root = ET.fromstring(parts["word/styles.xml"])
            defaults = root.find(W + "docDefaults")
            if defaults is not None:
                self.ppr_default = defaults.find(f"{W}pPrDefault/{W}pPr")
                self.rpr_default = defaults.find(f"{W}rPrDefault/{W}rPr")
            for style in root.findall(W + "style"):
                style_id = style.get(W + "styleId")
                if style_id is None:
                    continue
                self.styles[style_id] = style
                if style.get(W + "type") == "paragraph" and (style.get(W + "default") or "").lower() in _ON:
                    self.default_para = style_id
        self.nums: dict[str, ET.Element] = {}
        self.abstracts: dict[str, ET.Element] = {}
        if "word/numbering.xml" in parts:
            root = ET.fromstring(parts["word/numbering.xml"])
            for item in root.findall(W + "abstractNum"):
                self.abstracts[item.get(W + "abstractNumId") or ""] = item
            for item in root.findall(W + "num"):
                self.nums[item.get(W + "numId") or ""] = item
        self.theme: dict[str, str] = {}
        themes = sorted(name for name in parts if re.fullmatch(r"word/theme/theme\d*\.xml", name))
        if themes:
            drawing = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            root = ET.fromstring(parts[themes[0]])
            for kind in ("major", "minor"):
                fonts = root.find(f".//{drawing}{kind}Font")
                if fonts is None:
                    continue
                for slot, tag in (("latin", "latin"), ("cs", "cs"), ("ea", "ea")):
                    node = fonts.find(drawing + tag)
                    if node is not None and node.get("typeface"):
                        self.theme[f"{kind}:{slot}"] = node.get("typeface")

    def chain(self, style_id: str | None) -> list[ET.Element]:
        """The style and its w:basedOn ancestors, the style itself first."""

        out, seen = [], set()
        while style_id and style_id in self.styles and style_id not in seen:
            seen.add(style_id)
            style = self.styles[style_id]
            out.append(style)
            based = style.find(W + "basedOn")
            style_id = based.get(W + "val") if based is not None else None
        return out

    def _para_style(self, ppr: ET.Element | None) -> str | None:
        node = ppr.find(W + "pStyle") if ppr is not None else None
        if node is not None and node.get(W + "val"):
            return node.get(W + "val")
        return self.default_para

    # -- numbering -----------------------------------------------------------

    def numbering(self, ppr: ET.Element | None) -> tuple[str, int] | None:
        """(numId, ilvl) the paragraph is numbered with, None when it is not numbered."""

        num_id = level = None
        direct = ppr.find(W + "numPr") if ppr is not None else None
        if direct is not None:
            node = direct.find(W + "numId")
            num_id = node.get(W + "val") if node is not None else None
            node = direct.find(W + "ilvl")
            level = _int(node, "val")
        if num_id is None:
            for style in self.chain(self._para_style(ppr)):
                numpr = style.find(f"{W}pPr/{W}numPr")
                if numpr is None:
                    continue
                node = numpr.find(W + "numId")
                if node is None:
                    continue
                num_id = node.get(W + "val")
                if level is None:
                    level = _int(numpr.find(W + "ilvl"), "val")
                break
        if num_id is None or num_id.strip() == "0" or num_id not in self.nums:
            return None
        return num_id, level or 0

    def _level(self, numbering: tuple[str, int] | None) -> ET.Element | None:
        if numbering is None:
            return None
        num_id, level = numbering
        num = self.nums[num_id]
        for override in num.findall(W + "lvlOverride"):
            if _int(override, "ilvl") == level:
                lvl = override.find(W + "lvl")
                if lvl is not None:
                    return lvl
        abstract_ref = num.find(W + "abstractNumId")
        abstract = self.abstracts.get(abstract_ref.get(W + "val") if abstract_ref is not None else "")
        if abstract is None:
            return None
        return next((lvl for lvl in abstract.findall(W + "lvl") if _int(lvl, "ilvl") == level), None)

    # -- paragraph properties ----------------------------------------------

    def para(self, ppr: ET.Element | None) -> dict:
        """Effective left/right/first (twips, first negative = hanging), before/after (twips), numbering."""

        numbering = self.numbering(ppr)
        level = self._level(numbering)
        layers = [self.ppr_default]
        layers += [style.find(W + "pPr") for style in reversed(self.chain(self._para_style(ppr)))]
        layers += [level.find(W + "pPr") if level is not None else None, ppr]
        values = dict.fromkeys(_KEYS, 0)
        for layer in layers:
            if layer is None:
                continue
            ind = layer.find(W + "ind")
            if ind is not None:
                for key, names in (("left", ("start", "left")), ("right", ("end", "right"))):
                    found = next((_int(ind, name) for name in names if _int(ind, name) is not None), None)
                    if found is not None:
                        values[key] = found
                hanging, first = _int(ind, "hanging"), _int(ind, "firstLine")
                if hanging is not None:
                    values["first"] = -hanging
                elif first is not None:
                    values["first"] = first
            spacing = layer.find(W + "spacing")
            if spacing is not None:
                for key in ("before", "after"):
                    found = _int(spacing, key)
                    if found is not None:
                        values[key] = found
        values["numbering"] = numbering
        return values

    # -- run properties ------------------------------------------------------

    def _run_layers(self, ppr: ET.Element | None, rpr: ET.Element | None) -> list[ET.Element | None]:
        layers = [self.rpr_default]
        layers += [style.find(W + "rPr") for style in reversed(self.chain(self._para_style(ppr)))]
        rstyle = rpr.find(W + "rStyle") if rpr is not None else None
        if rstyle is not None:
            layers += [style.find(W + "rPr") for style in reversed(self.chain(rstyle.get(W + "val")))]
        layers.append(rpr)
        return layers

    def cs_size(self, ppr: ET.Element | None, rpr: ET.Element | None) -> int | None:
        """The complex-script size (w:szCs, half-points) the run inherits or sets, None when nothing sets it."""

        size = None
        for layer in self._run_layers(ppr, rpr):
            found = _int(layer.find(W + "szCs"), "val") if layer is not None else None
            if found is not None:
                size = found
        return size

    def bold(self, ppr: ET.Element | None, rpr: ET.Element | None) -> bool:
        """Whether the run shows bold (the last layer that sets w:b decides)."""

        bold = False
        for layer in self._run_layers(ppr, rpr):
            node = layer.find(W + "b") if layer is not None else None
            if node is not None:
                bold = (node.get(W + "val") or "true").lower() not in ("0", "false", "off")
        return bold

    def centered(self, ppr: ET.Element | None) -> bool:
        """Whether the paragraph is centered (direct w:jc, else its style chain's)."""

        layers = [self.ppr_default] + [style.find(W + "pPr") for style in reversed(self.chain(self._para_style(ppr)))]
        jc = None
        for layer in layers + [ppr]:
            node = layer.find(W + "jc") if layer is not None else None
            if node is not None:
                jc = node.get(W + "val")
        return jc == "center"

    def run(self, ppr: ET.Element | None, rpr: ET.Element | None) -> tuple[str | None, str | None, int]:
        """Effective (ascii font, hAnsi font, size in half-points) of a run (or a paragraph mark: its rPr)."""

        layers = self._run_layers(ppr, rpr)
        fonts: dict[str, str | None] = {"ascii": None, "hAnsi": None}
        size = 20
        for layer in layers:
            if layer is None:
                continue
            node = layer.find(W + "rFonts")
            if node is not None:
                for slot in fonts:
                    theme = node.get(W + slot + "Theme")
                    if theme is not None:
                        fonts[slot] = self._theme_font(theme)
                    elif node.get(W + slot) is not None:
                        fonts[slot] = node.get(W + slot)
            found = _int(layer.find(W + "sz"), "val")
            if found is not None:
                size = found
        return fonts["ascii"], fonts["hAnsi"], size

    def _theme_font(self, theme: str) -> str | None:
        kind = "major" if theme.startswith("major") else "minor"
        slot = "cs" if theme.endswith("Bidi") else ("ea" if theme.endswith("EastAsia") else "latin")
        return self.theme.get(f"{kind}:{slot}")


# --------------------------------------------------------------------------
# Request.
# --------------------------------------------------------------------------

def _request(*, font, size, left_indent, right_indent, first_line_indent, space_before, space_after,
             remove_numbering, op: str) -> dict:
    want: dict = {"font": None, "size": None, **dict.fromkeys(_KEYS), "numbering": None}
    if font is not None:
        font = font.strip()
        if not font or len(font) > 64 or any(ord(ch) < 32 for ch in font):
            raise RefusedUnsafe(f"{op}: font must be a font name (1-64 characters)")
        want["font"] = font
    if size is not None:
        half = size * 2
        if not 1 <= size <= MAX_SIZE_PT or abs(half - round(half)) > 1e-6:
            raise RefusedUnsafe(f"{op}: size must be 1-{MAX_SIZE_PT} pt in steps of 0.5 pt (got {size:g})")
        want["size"] = int(round(half))
    for key, value, low in (("left", left_indent, -MAX_INDENT_PT), ("right", right_indent, -MAX_INDENT_PT),
                            ("first", first_line_indent, -MAX_INDENT_PT), ("before", space_before, 0),
                            ("after", space_after, 0)):
        if value is None:
            continue
        if not low <= value <= MAX_INDENT_PT:
            raise RefusedUnsafe(f"{op}: {key} must be {low}..{MAX_INDENT_PT} pt (got {value:g})")
        want[key] = _tw(value)
    if remove_numbering:
        want["numbering"] = "remove"
    return want


# --------------------------------------------------------------------------
# Plan and write one paragraph.
# --------------------------------------------------------------------------

class _Run:
    def __init__(self, run: Node, text: str) -> None:
        self.run, self.text = run, text
        self.rpr = ""
        self.kind = "new"  # "new" (rPrChange added), "own_insertion" (L1), "own_format_change" (L3), "plain"
        self.font = self.size = None  # what is written (None: that property is left alone)
        self.in_place: Node | None = None


class _Paragraph:
    def __init__(self, formatter: "Formatter", index: int) -> None:
        self.f, self.index = formatter, index
        doc, xml = formatter.doc, formatter.xml
        self.node = doc.paragraph(index, formatter.op)
        if doc.layout.para(index).nested:
            raise self.refuse("it is inside a nested table; format it in Word")
        self.model = ParagraphModel(xml, self.node)
        self.ppr_node = self.node.child("w:pPr")
        self.ppr = xml[self.ppr_node.start:self.ppr_node.end] if self.ppr_node is not None else ""
        self.ppr_et = _fragment(xml, self.ppr) if self.ppr else None
        mark = self.ppr_node.child("w:rPr") if self.ppr_node is not None else None
        self.mark_rpr = xml[mark.start:mark.end] if mark is not None else ""
        self.new_ppr: str | None = None
        self.ppr_changes: dict = {}
        self.mark_change: dict = {}
        self.mark_kind = "new"
        self.ppr_kind = "new"
        self.runs: list[_Run] = []
        self.field_runs: set[int] = set()
        self.eligible: set[int] = set()
        self.own_changes: set[str | None] = set()  # ids of the author's own formatting changes kept (L3)

    def refuse(self, what: str) -> RefusedUnsafe:
        return RefusedUnsafe(f"{self.f.op}: paragraph {self.index}: {what}")

    # -- checks ---------------------------------------------------------------

    def _change_head(self, raw: str, tag: str) -> str | None:
        match = re.search(rf"<{tag}\b[^>]*>", raw)
        return match.group(0) if match else None

    def _check_existing(self, raw: str, tag: str, what: str) -> str:
        """'new' (no change of that kind yet) or 'own' (the author's own, kept: L3); refuses otherwise."""

        head = self._change_head(raw, tag)
        if head is None:
            return "new"
        described = _describe_change(head, tag)
        if not self.f.tracked:
            raise self.refuse(f"{what} carries the tracked formatting change {described}; untracked "
                              "formatting is refused there (rejecting that change would undo it); format it "
                              "tracked, or accept or reject that change first")
        if attr_value(head, "w:author") != self.f.author:
            raise self.refuse(f"{what} carries another author's tracked formatting change {described}; accept or "
                              "reject it first")
        self.own_changes.add(attr_value(head, "w:id"))
        return "own"

    def _check_moves(self) -> None:
        for item in self.model.items:
            for node in item.wrappers:
                if node.name in _MOVES:
                    raise self.refuse(f"it holds moved text ({describe_wrapper(self.f.xml, node)}); accept or "
                                      "reject the move first")
        if re.search(r"<w:(moveFrom|moveTo)\b", self.mark_rpr):
            raise self.refuse("its paragraph mark is part of a tracked move; accept or reject the move first")

    # -- planning -------------------------------------------------------------

    def plan(self, want: dict) -> bool:
        """Work out what changes; True when anything does."""

        self._check_moves()
        f = self.f
        head, children = _children(self.ppr, "w:pPr")
        original = list(children)
        current = f.styles.para(self.ppr_et)
        # numbering first: removing a list also removes the list level's indents
        numbering = want["numbering"]
        if numbering == "remove" and current["numbering"] is not None:
            children = _put(children, "w:numPr", None, ppr_rank)
            if f.styles.numbering(self._et(head, children)) is not None:  # the style numbers it: switch it off
                children = _put(children, "w:numPr", '<w:numPr><w:numId w:val="0"/></w:numPr>', ppr_rank)
            self.ppr_changes["numbering"] = None
        elif isinstance(numbering, tuple) and current["numbering"] != numbering:
            num_id, level = numbering
            children = _put(children, "w:numPr",
                            f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{_attr_text(num_id)}"/></w:numPr>',
                            ppr_rank)
            self.ppr_changes["numbering"] = {"num_id": num_id, "level": level}
        after = f.styles.para(self._et(head, children))
        ind_updates: dict[str, str | None] = {}
        spacing_updates: dict[str, str | None] = {}
        old_ind = _child_raw(children, "w:ind")
        start_names = any(name in ("w:start", "w:end") for name, _ in _attrs_of(old_ind))
        for key in _KEYS:
            value = want[key]
            if value is None or after[key] == value:
                continue
            self.ppr_changes[key] = value
            if key in _IND_ATTRS:
                for name in _IND_ATTRS[key]:
                    ind_updates[name] = None
                if key == "first":
                    ind_updates["w:hanging" if value < 0 else "w:firstLine"] = str(abs(value))
                else:
                    name = {"left": ("w:left", "w:start"), "right": ("w:right", "w:end")}[key][start_names]
                    ind_updates[name] = str(value)
            else:
                for name in _SPACING_ATTRS[key]:
                    spacing_updates[name] = None
                spacing_updates[f"w:{key}"] = str(value)
        if ind_updates:
            children = _put(children, "w:ind", _updated(old_ind, "w:ind", ind_updates), ppr_rank)
        if spacing_updates:
            old = _child_raw(children, "w:spacing")
            children = _put(children, "w:spacing", _updated(old, "w:spacing", spacing_updates), ppr_rank)
        if self.ppr_changes:
            self.ppr_kind = self._check_existing(self.ppr, "w:pPrChange", "its paragraph properties")
        self.planned_children = children
        self.original_children = original
        self.ppr_head = head
        # mark
        mark_et = _fragment(f.xml, self.mark_rpr) if self.mark_rpr else None
        self.mark_change = self._run_change(mark_et, want)
        if self.mark_change:
            self.mark_kind = self._check_existing(self.mark_rpr, "w:rPrChange", "its paragraph mark")
        # runs
        for item in self.model.items:
            run = item.run
            if run is not None and (getattr(item, "field", False) or getattr(item, "kind", "") == "field"):
                self.field_runs.add(id(run))
        seen: dict[int, _Run] = {}
        for char in self.model.final:
            if id(char.run) in self.field_runs:
                continue
            self.eligible.add(id(char.run))
            entry = seen.get(id(char.run))
            if entry is None:
                entry = seen[id(char.run)] = _Run(char.run, "")
                self.runs.append(entry)
            entry.text += char.ch
        planned = []
        for entry in self.runs:
            rpr_node = entry.run.child("w:rPr")
            entry.rpr = f.xml[rpr_node.start:rpr_node.end] if rpr_node is not None else ""
            change = self._run_change(_fragment(f.xml, entry.rpr) if entry.rpr else None, want)
            if not change:
                continue
            entry.font, entry.size = change.get("font"), change.get("size")
            state = self._check_existing(entry.rpr, "w:rPrChange", f"the text {entry.text[:40]!r}")
            if state == "own":
                entry.kind = "own_format_change"
            elif not f.tracked:
                entry.kind = "plain"
            else:
                char = next(char for char in self.model.final if char.run is entry.run)
                wrappers = [node for node in char.wrappers if node.name in _WRAPPERS]
                if len(wrappers) == 1 and wrappers[0].name == "w:ins" and \
                        wrappers[0].attr(f.xml, "w:author") == f.author:
                    entry.kind, entry.in_place = "own_insertion", wrappers[0]
            planned.append(entry)
        self.runs = planned
        return bool(self.ppr_changes or self.mark_change or self.runs)

    def _et(self, head: str, children: list[tuple[str, str]]) -> ET.Element:
        return _fragment(self.f.xml, _close(head, children, "w:pPr"))

    def _run_change(self, rpr_et: ET.Element | None, want: dict) -> dict:
        ascii_font, hansi_font, size = self.f.styles.run(self.ppr_et, rpr_et)
        change = {}
        if want["font"] is not None and (ascii_font != want["font"] or hansi_font != want["font"]):
            change["font"] = want["font"]
        if want["size"] is not None and (size != want["size"] or
                                         self.f.styles.cs_size(self.ppr_et, rpr_et) not in (None, want["size"])):
            change["size"] = want["size"]  # Word's size box sets w:sz and w:szCs together
        return change

    # -- writing ----------------------------------------------------------------

    def write(self, writer: _Writer) -> str:
        f, xml = self.f, self.f.xml
        edits: list[tuple[int, int, str]] = []
        if self.ppr_changes or self.mark_change:
            children = list(self.planned_children)
            if self.mark_change:
                change = rpr_change(writer, self.mark_rpr, "") if f.tracked and self.mark_kind == "new" else ""
                new_mark = _run_props(self.mark_rpr, self.mark_change.get("font"), self.mark_change.get("size"),
                                      change)
                children = _put(children, "w:rPr", new_mark, ppr_rank)
            if self.ppr_changes and f.tracked and self.ppr_kind == "new":
                children = _put(children, "w:pPrChange", ppr_change(writer, self.ppr), ppr_rank)
            self.new_ppr = _close(self.ppr_head, children, "w:pPr")
            if self.ppr_node is not None:
                edits.append((self.ppr_node.start, self.ppr_node.end, self.new_ppr))
            elif self.node.closed:
                head = re.sub(r"\s*/>\Z", ">", self.node.head(xml))
                edits.append((self.node.start, self.node.end, head + self.new_ppr + "</w:p>"))
            else:
                at = self.node.start + len(self.node.head(xml))
                edits.append((at, at, self.new_ppr))
        for entry in self.runs:
            change = rpr_change(writer, entry.rpr, entry.text) if entry.kind == "new" and f.tracked else ""
            new_rpr = _run_props(entry.rpr, entry.font, entry.size, change)
            rpr_node = entry.run.child("w:rPr")
            if rpr_node is not None:
                edits.append((rpr_node.start, rpr_node.end, new_rpr))
            else:
                at = entry.run.start + len(entry.run.head(xml))
                edits.append((at, at, new_rpr))
        out, cursor = [], self.node.start
        for start, end, text in sorted(edits, key=lambda item: (item[0], item[1])):
            out.append(xml[cursor:start])
            out.append(text)
            cursor = end
        out.append(xml[cursor:self.node.end])
        return "".join(out)

    def report(self) -> dict:
        item: dict = {"paragraph_index": self.index}
        changes = {}
        for key, value in self.ppr_changes.items():
            changes[{"left": "left_indent", "right": "right_indent", "first": "first_line_indent",
                     "before": "space_before", "after": "space_after"}.get(key, key)] = \
                value if key == "numbering" else float(_pt(value, 20))
        if changes:
            item["paragraph"] = changes
        if self.mark_change:
            item["mark"] = _run_report(self.mark_change)
        item["runs_formatted"] = len(self.runs)
        in_place = [{"kind": entry.kind, "text": entry.text[:80]} for entry in self.runs
                    if entry.kind in ("own_insertion", "own_format_change")]
        if in_place:
            item["formatted_in_place"] = in_place
        return item


def _describe_change(head: str, tag: str) -> str:
    return (f"{tag} id={attr_value(head, 'w:id')} by {attr_value(head, 'w:author') or ''!r} "
            f"({attr_value(head, 'w:date')})")


def _run_report(change: dict) -> dict:
    out = {}
    if "font" in change:
        out["font"] = change["font"]
    if "size" in change:
        out["size"] = change["size"] / 2
    return out


def _run_props(rpr: str, font: str | None, size: int | None, change: str) -> str:
    """`rpr` with the font (ascii/hAnsi/cs, theme attributes dropped) and size (w:sz, w:szCs) set, at their
    schema positions, and `change` (a new w:rPrChange, or "") appended last."""

    head, children = _children(rpr, "w:rPr")
    if font is not None:
        value = _attr_text(font)
        updates = {name: None for name in _FONT_ATTRS}
        updates.update({"w:ascii": value, "w:hAnsi": value, "w:cs": value})
        children = _put(children, "w:rFonts", _updated(_child_raw(children, "w:rFonts"), "w:rFonts", updates),
                        rpr_rank)
    if size is not None:
        children = _put(children, "w:sz", f'<w:sz w:val="{size}"/>', rpr_rank)
        children = _put(children, "w:szCs", f'<w:szCs w:val="{size}"/>', rpr_rank)
    if change:
        children.append(("w:rPrChange", change))
    return _close(head, children, "w:rPr")


# --------------------------------------------------------------------------
# Formatter.
# --------------------------------------------------------------------------

class Formatter:
    def __init__(self, parts, doc, *, paragraph_index: int, through_paragraph_index: int | None, want: dict,
                 like_paragraph: int | None, tracked: bool, author: str, op: str = OP) -> None:
        self.parts, self.doc, self.xml, self.op = parts, doc, doc.xml, op
        self.tracked, self.author = tracked, author
        self.styles = Styles(parts)
        last = through_paragraph_index if through_paragraph_index is not None else paragraph_index
        if last < paragraph_index:
            raise RefusedUnsafe(f"{op}: through_paragraph_index {last} is before paragraph_index {paragraph_index}")
        total = len(doc.layout.paragraphs)
        for index in (paragraph_index, last):
            if not 1 <= index <= total:
                raise RefusedUnsafe(f"{op}: paragraph_index {index} is outside 1..{total}")
        self.first, self.last = paragraph_index, last
        self._check_container()
        self.like = None
        if like_paragraph is not None:
            if not 1 <= like_paragraph <= total:
                raise RefusedUnsafe(f"{op}: like_paragraph {like_paragraph} is outside 1..{total}")
            if paragraph_index <= like_paragraph <= last:
                raise RefusedUnsafe(f"{op}: like_paragraph {like_paragraph} is one of the paragraphs to format; "
                                    "give another paragraph to copy from")
            self.like = self.like_values(like_paragraph)
            want = _merge_like(want, self.like)
        if not any(want[key] is not None for key in ("font", "size", "numbering", *_KEYS)):
            raise RefusedUnsafe(f"{op}: give at least one of font, size, left_indent, right_indent, "
                                "first_line_indent, space_before, space_after, remove_numbering, like_paragraph")
        self.want = want
        self.paragraphs = [_Paragraph(self, index) for index in range(self.first, self.last + 1)]
        self.changed = [item for item in self.paragraphs if item.plan(want)]
        if not self.changed:
            span = f"paragraph {self.first}" if self.first == self.last else f"paragraphs {self.first}..{self.last}"
            raise RefusedUnsafe(f"{op}: {span} is already formatted as requested; nothing to change")

    def _check_container(self) -> None:
        layout = self.doc.layout

        def where(index: int) -> str:
            item = layout.para(index)
            if item.table is None:
                return "the body"
            return f"table {item.table + 1} row {item.row + 1} cell {item.cell + 1}"

        for index in range(self.first, self.last + 1):
            item = layout.para(index)
            if item.row_end:
                raise RefusedUnsafe(f"{self.op}: paragraph {index} is a table row-end marker, not a paragraph; a "
                                    "range must lie in the body or in one table cell (give one op per cell)")
            if item.nested:
                raise RefusedUnsafe(f"{self.op}: paragraph {index} is inside a nested table; format it in Word")
        home = (layout.para(self.first).table, layout.para(self.first).row, layout.para(self.first).cell)
        for index in range(self.first + 1, self.last + 1):
            item = layout.para(index)
            if (item.table, item.row, item.cell) != home:
                raise RefusedUnsafe(
                    f"{self.op}: paragraphs {self.first}..{self.last} do not lie in one container: paragraph "
                    f"{index - 1} is in {where(index - 1)} and paragraph {index} in {where(index)}; give one op "
                    "per cell (or per run of body paragraphs)")

    def like_values(self, index: int) -> dict:
        """Paragraph `index`'s effective indents, spacing and numbering, and its dominant font and size."""

        node = self.doc.paragraph(index, self.op)
        model = ParagraphModel(self.xml, node)
        ppr_node = node.child("w:pPr")
        ppr = _fragment(self.xml, self.xml[ppr_node.start:ppr_node.end]) if ppr_node is not None else None
        values = self.styles.para(ppr)
        fields = {id(item.run) for item in model.items
                  if item.run is not None and (getattr(item, "field", False) or getattr(item, "kind", "") == "field")}
        counts: Counter = Counter()
        cache: dict[str, tuple] = {}
        for char in model.final:
            if id(char.run) in fields or not char.ch.strip():
                continue
            if char.rpr not in cache:
                cache[char.rpr] = self.styles.run(ppr, _fragment(self.xml, char.rpr) if char.rpr else None)
            counts[cache[char.rpr]] += 1
        if counts:
            ascii_font, hansi_font, size = counts.most_common(1)[0][0]
        else:
            mark = ppr_node.child("w:rPr") if ppr_node is not None else None
            mark_et = _fragment(self.xml, self.xml[mark.start:mark.end]) if mark is not None else None
            ascii_font, hansi_font, size = self.styles.run(ppr, mark_et)
        values["font"] = ascii_font if ascii_font and ascii_font == hansi_font else (ascii_font or None)
        values["size"] = size
        values["index"] = index
        return values

    def apply(self, *, first_id: int, stamp: str) -> tuple[str, list[tuple[_Paragraph, str]], list[dict]]:
        writer = _Writer(first_id, self.author, stamp, True)
        out, cursor, written = [], 0, []
        for item in self.changed:
            new = item.write(writer)
            out.append(self.xml[cursor:item.node.start])
            out.append(new)
            cursor = item.node.end
            written.append((item, new))
        out.append(self.xml[cursor:])
        return "".join(out), written, writer.added


def _merge_like(want: dict, like: dict) -> dict:
    merged = dict(want)
    for key in ("font", "size", *_KEYS):
        if merged[key] is None:
            merged[key] = like[key]
    if merged["numbering"] is None:
        merged["numbering"] = like["numbering"] if like["numbering"] is not None else "remove"
    return merged


# --------------------------------------------------------------------------
# Verification.
# --------------------------------------------------------------------------

def _strip_run_props(element: ET.Element) -> ET.Element:
    clone = copy.deepcopy(element)
    for run in clone.iter(W + "r"):
        for rpr in run.findall(W + "rPr"):
            run.remove(rpr)
    return clone


def _props_split(rpr: ET.Element | None, drop: set[str]) -> str:
    """Canonical rPr/pPr without the children named in `drop` (local names) and revision markers."""

    from .sections import canonical

    if rpr is None:
        return ""
    clone = copy.deepcopy(rpr)
    for child in list(clone):
        local = child.tag.rsplit("}", 1)[-1]
        if local in drop or local in REV_TAGS:
            clone.remove(child)
    return canonical(clone) if len(clone) else ""


def _attrs_minus(element: ET.Element | None, drop: set[str]) -> dict:
    if element is None:
        return {}
    return {key: value for key, value in element.attrib.items() if key.rsplit("}", 1)[-1] not in drop}


def _verify(formatter: Formatter, new_xml: str, written: list[tuple[_Paragraph, str]], added: list[dict]) -> None:
    from .sections import canonical

    op, xml, author, tracked = formatter.op, formatter.xml, formatter.author, formatter.tracked
    styles, want = formatter.styles, formatter.want
    cursor = new_cursor = 0
    for item, new in written:
        gap = xml[cursor:item.node.start]
        if new_xml[new_cursor:new_cursor + len(gap)] != gap:
            raise VerificationFailed(f"{op}: bytes outside the formatted paragraphs changed")
        new_cursor += len(gap) + len(new)
        cursor = item.node.end
    if new_xml[new_cursor:] != xml[cursor:]:
        raise VerificationFailed(f"{op}: bytes outside the formatted paragraphs changed")
    if added and not tracked:
        raise VerificationFailed(f"{op}: an untracked format added revisions")
    new_ids = {entry["id"] for entry in added}
    order: list[int] = []
    for item, new in written:
        old_raw = xml[item.node.start:item.node.end]
        old_tree, new_tree = parse(old_raw), parse(new)

        def revisions(tree: Node) -> list[Node]:
            return [node for node in tree.iter() if node.local in REV_TAGS and node.name.startswith("w:")]

        holders = {wrapper.head(xml) for entry in item.runs for char in item.model.final if char.run is entry.run
                   for wrapper in char.wrappers if wrapper.name == "w:ins"}
        old_revs, kept = revisions(old_tree), []
        for node in revisions(new_tree):
            number = node.attr(new, "w:id")
            if node.name in ("w:rPrChange", "w:pPrChange") and number in new_ids and \
                    node.attr(new, "w:author") == author:
                order.append(int(number))
                if not re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$", node.attr(new, "w:date") or ""):
                    raise VerificationFailed(f"{op}: new revision has the wrong date format")
            else:
                kept.append(node)
        if len(kept) != len(old_revs):
            raise VerificationFailed(f"{op}: a pre-existing revision in paragraph {item.index} was added or lost")
        for old, new_node in zip(old_revs, kept):
            before, after = old_raw[old.start:old.end], new[new_node.start:new_node.end]
            if before == after:
                continue
            if old.name != "w:ins" or new_node.name != "w:ins" or old.head(old_raw) != new_node.head(new) or \
                    old.head(old_raw) not in holders or \
                    canonical(_strip_run_props(_fragment(xml, before))) != \
                    canonical(_strip_run_props(_fragment(xml, after))):
                raise VerificationFailed(f"{op}: a pre-existing revision in paragraph {item.index} changed")
        new_model = ParagraphModel(new, new_tree.children[0])
        if new_model.text != item.model.text or new_model.original != item.model.original:
            raise VerificationFailed(f"{op}: the text of paragraph {item.index} changed")
        old_p, new_p = _fragment(xml, old_raw), _fragment(xml, new)
        edited_ins = {entry.in_place.attr(xml, "w:id") for entry in item.runs if entry.kind == "own_insertion"}
        is_new = lambda el: el.tag in (W + "rPrChange", W + "pPrChange") and el.get(W + "id") in new_ids and \
            el.get(W + "author") == author  # noqa: E731
        own_changes = item.own_changes
        is_edited = lambda el: el.get(W + "author") == author and (  # noqa: E731
            (el.tag == W + "ins" and el.get(W + "id") in edited_ins)
            or (el.tag in (W + "rPrChange", W + "pPrChange") and el.get(W + "id") in own_changes))
        if tracked:
            rejected = copy.deepcopy(new_p)
            et_resolve(rejected, lambda el: is_new(el) or is_edited(el), "reject")
            baseline = copy.deepcopy(old_p)
            et_resolve(baseline, is_edited, "reject")
            if canonical(rejected) != canonical(baseline):
                raise VerificationFailed(
                    f"{op}: rejecting the new formatting revisions does not give back paragraph {item.index}")
        accepted = copy.deepcopy(new_p)
        et_resolve(accepted, is_new, "accept")
        _verify_accepted(formatter, item, old_p, accepted)
    if order != sorted(order) or len(set(order)) != len(order) or len(order) != len(added):
        raise VerificationFailed(f"{op}: new revision ids do not ascend left to right")


def _verify_accepted(formatter: Formatter, item: _Paragraph, old_p: ET.Element, accepted: ET.Element) -> None:
    op, styles, want = formatter.op, formatter.styles, formatter.want
    label = f"{op}: paragraph {item.index}"
    old_ppr, new_ppr = old_p.find(W + "pPr"), accepted.find(W + "pPr")
    # paragraph properties: requested values in effect, everything else as it was
    before, after = styles.para(old_ppr), styles.para(new_ppr)
    numbering_changed = before["numbering"] != after["numbering"]
    if want["numbering"] == "remove" and after["numbering"] is not None:
        raise VerificationFailed(f"{label} is still numbered")
    if isinstance(want["numbering"], tuple) and after["numbering"] != want["numbering"]:
        raise VerificationFailed(f"{label} does not have the requested numbering")
    if want["numbering"] is None and numbering_changed:
        raise VerificationFailed(f"{label}: its numbering changed")
    for key in _KEYS:
        if want[key] is not None:
            if after[key] != want[key]:
                raise VerificationFailed(f"{label}: {key} is {after[key]}, not {want[key]} twips")
        elif not numbering_changed and after[key] != before[key]:
            raise VerificationFailed(f"{label}: {key} changed though it was not requested")
    drop = {"numPr", "ind", "spacing", "rPr", "pPrChange"}
    if _props_split(old_ppr, drop) != _props_split(new_ppr, drop):
        raise VerificationFailed(f"{label}: a paragraph property that was not requested changed")
    touched = {name.split(":")[1] for key in _IND_ATTRS if want[key] is not None for name in _IND_ATTRS[key]}
    if _attrs_minus(_find(old_ppr, "ind"), touched) != _attrs_minus(_find(new_ppr, "ind"), touched):
        raise VerificationFailed(f"{label}: an indent that was not requested changed")
    touched = {name.split(":")[1] for key in _SPACING_ATTRS if want[key] is not None for name in _SPACING_ATTRS[key]}
    if _attrs_minus(_find(old_ppr, "spacing"), touched) != _attrs_minus(_find(new_ppr, "spacing"), touched):
        raise VerificationFailed(f"{label}: spacing that was not requested changed")
    # the mark
    _check_run(formatter, old_ppr, _find(old_ppr, "rPr"), _find(new_ppr, "rPr"), f"{label}'s paragraph mark")
    # every character: eligible runs show the requested font/size, other properties unchanged
    got, was = _et_final_chars(accepted), _et_final_chars(old_p)
    if "".join(ch for ch, _ in got) != item.model.text or "".join(ch for ch, _ in was) != item.model.text:
        raise VerificationFailed(f"{label}: accepting the formatting changes changes its text")
    final = item.model.final
    if len(final) != len(got):
        raise VerificationFailed(f"{label}: characters could not be aligned for the check")
    cache: dict[tuple[str, str], None] = {}
    for char, (ch, key), (_, old_key) in zip(final, got, was):
        if id(char.run) not in item.eligible:
            if key != old_key:
                raise VerificationFailed(f"{label}: {ch!r} (deleted or field text) changed formatting")
            continue
        if (key, old_key) in cache:
            continue
        cache[(key, old_key)] = None
        _check_run(formatter, old_ppr, ET.fromstring(old_key) if old_key else None,
                   ET.fromstring(key) if key else None, f"{label}: character {ch!r}")


def _find(element: ET.Element | None, local: str) -> ET.Element | None:
    return element.find(W + local) if element is not None else None


def _check_run(formatter: Formatter, ppr: ET.Element | None, old: ET.Element | None, new: ET.Element | None,
               what: str) -> None:
    want, styles = formatter.want, formatter.styles
    drop = set()
    if want["font"] is not None:
        drop.add("rFonts")
    if want["size"] is not None:
        drop |= {"sz", "szCs"}
    drop.add("rPrChange")
    if _props_split(old, drop) != _props_split(new, drop):
        raise VerificationFailed(f"{what}: a run property that was not requested changed")
    ascii_font, hansi_font, size = styles.run(ppr, new)
    if want["font"] is not None:
        if ascii_font != want["font"] or hansi_font != want["font"]:
            raise VerificationFailed(f"{what} is not in {want['font']}")
        fonts = lambda el: _attrs_minus(_find(el, "rFonts"), {name.split(":")[1] for name in _FONT_ATTRS})  # noqa
        if fonts(old) != fonts(new):
            raise VerificationFailed(f"{what}: a font slot that was not requested changed")
    if want["size"] is not None and (size != want["size"] or styles.cs_size(ppr, new) not in (None, want["size"])):
        raise VerificationFailed(f"{what} is not {want['size'] / 2:g} pt")


# --------------------------------------------------------------------------
# Operation.
# --------------------------------------------------------------------------

def format_paragraphs(source, output, *, paragraph_index: int, through_paragraph_index: int | None = None,
                      font: str | None = None, size: float | None = None, left_indent: float | None = None,
                      right_indent: float | None = None, first_line_indent: float | None = None,
                      space_before: float | None = None, space_after: float | None = None,
                      remove_numbering: bool = False, like_paragraph: int | None = None, tracked: bool,
                      author: str | None = DEFAULT_AUTHOR, now: dt.datetime | None = None, op: str = OP) -> dict:
    """Format paragraphs `paragraph_index`..`through_paragraph_index` (one container), tracked as
    w:pPrChange/w:rPrChange or plain; verified, then written to the new file `output`."""

    author = (author or DEFAULT_AUTHOR).strip()
    if tracked and not author:
        raise RefusedUnsafe(f"{op}: author must not be empty")
    want = _request(font=font, size=size, left_indent=left_indent, right_indent=right_indent,
                    first_line_indent=first_line_indent, space_before=space_before, space_after=space_after,
                    remove_numbering=remove_numbering, op=op)
    parts, doc = load(source)
    formatter = Formatter(parts, doc, paragraph_index=paragraph_index, through_paragraph_index=through_paragraph_index,
                          want=want, like_paragraph=like_paragraph, tracked=tracked, author=author, op=op)
    first_id = next_revision_id(parts)
    stamp = revision_stamp(doc.xml, now)
    new_xml, written, added = formatter.apply(first_id=first_id, stamp=stamp)
    _verify(formatter, new_xml, written, added)
    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, op)
    except Exception:
        Path(output).unlink(missing_ok=True)
        raise
    applied = {key: value for key, value in formatter.want.items() if value is not None}
    requested = {}
    for key, value in applied.items():
        if key == "font":
            requested["font"] = value
        elif key == "size":
            requested["size"] = value / 2
        elif key == "numbering":
            requested["numbering"] = "none" if value == "remove" else {"num_id": value[0], "level": value[1]}
        else:
            requested[{"left": "left_indent", "right": "right_indent", "first": "first_line_indent",
                       "before": "space_before", "after": "space_after"}[key]] = float(_pt(value, 20))
    return _result(
        source, output, op, doc,
        paragraph_index=paragraph_index,
        through_paragraph_index=formatter.last,
        like_paragraph=like_paragraph,
        tracked=tracked,
        author=author if tracked else None,
        requested=requested,
        paragraphs=[item.report() for item in formatter.changed],
        unchanged_paragraphs=[item.index for item in formatter.paragraphs if item not in formatter.changed],
        revisions_added=added,
    )


def apply_operation(source, output, operation) -> dict:
    """edit_batch entry for the format_paragraph op (tracked/author resolved by edit_batch from the switch)."""

    return format_paragraphs(
        source, output, paragraph_index=operation.paragraph_index,
        through_paragraph_index=operation.through_paragraph_index, font=operation.font, size=operation.size,
        left_indent=operation.left_indent, right_indent=operation.right_indent,
        first_line_indent=operation.first_line_indent, space_before=operation.space_before,
        space_after=operation.space_after, remove_numbering=operation.remove_numbering,
        like_paragraph=operation.like_paragraph, tracked=bool(operation.tracked),
        author=operation.author or DEFAULT_AUTHOR)


# --------------------------------------------------------------------------
# Blank-answer warning (decision 6).
# --------------------------------------------------------------------------

def blank_fill_note(path, index: int) -> str | None:
    """When paragraph `index` is empty (accepted text) and numbered, or its mark is >= 2 pt off the nearest
    paragraph of body text, a one-line hint; it names format_paragraph(like_paragraph=M) when a body paragraph
    M exists (headings, titles and title-like lines are never M), else None."""

    parts, doc = load(path)
    if not 1 <= index <= len(doc.layout.paragraphs):
        return None
    item = doc.layout.para(index)
    if item.row_end or item.text.strip():
        return None
    styles = Styles(parts)
    node = doc.paragraph(index, OP)
    ppr_node = node.child("w:pPr")
    ppr = _fragment(doc.xml, doc.xml[ppr_node.start:ppr_node.end]) if ppr_node is not None else None
    mark = ppr_node.child("w:rPr") if ppr_node is not None else None
    mark_et = _fragment(doc.xml, doc.xml[mark.start:mark.end]) if mark is not None else None
    numbered = styles.numbering(ppr) is not None
    size = styles.run(ppr, mark_et)[2]
    profiles = _Profiles(doc, styles)
    body = profiles.nearest_body(index)
    what = ["numbered"] if numbered else []
    if body is None:
        common = profiles.body_size
        if common is not None and abs(size - common) >= 4:
            what.append(f"{_pt(size, 2)} pt (most text is {_pt(common, 2)} pt)")
        if not what:
            return None
        return (f"Paragraph {index} is {' and '.join(what)}; no paragraph reads as body text to match it to, so "
                "give format_paragraph the font, size and indents it needs.")
    probe = Formatter.__new__(Formatter)
    probe.parts, probe.doc, probe.xml, probe.op, probe.styles = parts, doc, doc.xml, OP, styles
    like = probe.like_values(body)
    if abs(size - like["size"]) >= 4:
        what.append(f"{_pt(size, 2)} pt (paragraph {body} is {_pt(like['size'], 2)} pt)")
    if not what:
        return None
    return (f"Paragraph {index} is {' and '.join(what)}; format_paragraph(paragraph_index={index}, "
            f"like_paragraph={body}) matches it to paragraph {body}.")


class _Profiles:
    """Per paragraph with text: whether it is a heading by style or outline level, numbered, centered, mostly
    bold, and its dominant size; and the document's body size (the dominant size over non-heading text)."""

    TITLE_STEP = 6  # half-points: a line 3 pt or more above the body size reads as a title
    BOLD_TITLE_CHARS = 80  # a mostly bold line reads as a title only when short and not indented

    def __init__(self, doc, styles: Styles) -> None:
        self.doc, self.styles = doc, styles
        self.info: dict[int, dict] = {}
        cache: dict[tuple[str, str], tuple[int, bool]] = {}
        sizes: Counter = Counter()
        for number in range(1, len(doc.layout.paragraphs) + 1):
            item = doc.layout.para(number)
            if item.row_end or not item.text.strip():
                continue
            node = doc.paragraph(number, OP)
            ppr_node = node.child("w:pPr")
            ppr_raw = doc.xml[ppr_node.start:ppr_node.end] if ppr_node is not None else ""
            ppr = _fragment(doc.xml, ppr_raw) if ppr_raw else None
            counts: Counter = Counter()
            bold = total = 0
            for char in ParagraphModel(doc.xml, node).final:
                if not char.ch.strip():
                    continue
                key = (ppr_raw, char.rpr or "")
                if key not in cache:
                    rpr = _fragment(doc.xml, char.rpr) if char.rpr else None
                    cache[key] = (styles.run(ppr, rpr)[2], styles.bold(ppr, rpr))
                char_size, char_bold = cache[key]
                counts[char_size] += 1
                bold += char_bold
                total += 1
            if not total:
                continue
            style = (item.style or "").lower()
            heading = item.outline is not None or style.startswith(("heading", "title", "subtitle"))
            self.info[number] = {"heading": heading, "numbered": styles.numbering(ppr) is not None,
                                 "centered": styles.centered(ppr), "bold": bold * 2 > total,
                                 "indented": styles.para(ppr)["left"] > 0, "chars": total,
                                 "size": counts.most_common(1)[0][0]}
            if not heading:
                sizes.update(counts)
        self.body_size = sizes.most_common(1)[0][0] if sizes else None

    def title_like(self, number: int) -> bool:
        info = self.info[number]
        if info["heading"] or info["centered"]:
            return True
        if info["bold"] and not info["indented"] and info["chars"] <= self.BOLD_TITLE_CHARS:
            return True  # a bold answer (indented, or a full sentence or more) stays body text
        return self.body_size is not None and info["size"] - self.body_size >= self.TITLE_STEP

    def nearest_body(self, index: int) -> int | None:
        """The nearest paragraph that reads as body text (not a heading, title or title-like line: centered,
        short, unindented and mostly bold, or 3 pt or more above the body size), preferring one that is not numbered (the previous
        one on a tie); None when there is none."""

        candidates = []
        for distance in range(1, len(self.doc.layout.paragraphs)):
            for other in (index - distance, index + distance):
                if other in self.info and not self.title_like(other):
                    candidates.append(other)
        return next((other for other in candidates if not self.info[other]["numbered"]),
                    candidates[0] if candidates else None)
