"""Plan 864: research-mcp's tracked-changes switch, kept in the document.

The switch is a custom document property (docProps/custom.xml), ``research_mcp_track_changes``, holding the
JSON ``{"on": true, "author": "Claude"}`` as a vt:lpwstr. It is research-mcp's own setting, not Word's
``w:trackRevisions``: it travels with the file (a copy under another name keeps it) and Word keeps custom
properties when it saves. It is written as a direct zip edit, never through Word; every other property
(ZOTERO_PREF_n among them) and every other part stay byte-identical.

While it is on, edit_document's text ops (replace_text, replace_paragraph, insert_text) are made as tracked
changes by the switch's author (an op's own ``author`` overrides it), delete_citation follows it, and ops
that cannot be tracked yet refuse with NOT_TRACKABLE_MESSAGE unless the op says ``tracked: false``.

How an op asks: ``read_setting(path)`` -> ``TrackSetting(on, author, is_set)``; ``route_operation`` turns a
text op into its tracked form (edit_batch calls it for every op)."""

from __future__ import annotations

from dataclasses import dataclass
import html
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

from .errors import RefusedUnsafe, VerificationFailed

PROPERTY_NAME = "research_mcp_track_changes"
DEFAULT_AUTHOR = "Claude"
SET_OP = "set_track_changes"
NOT_TRACKABLE_MESSAGE = (
    "{op} can't be recorded as a tracked change yet; turn tracking off for this op (`tracked: false`) or make "
    "it in Word with Track Changes on."
)
OFF_WITH_REVISIONS_NOTE = (
    "This document has tracked changes and tracking is off for your edits; set_track_changes(on) to track them."
)
# edit_document ops that cannot be made as tracked changes yet; tools are named by the server.
NOT_TRACKABLE_OPS: tuple[str, ...] = ()  # plan 916 S5: insert_table is tracked by direct XML
NOT_TRACKABLE_TOOLS = ("clean_bookmarks",)  # apply_document_style, format_tables: Word tracks them (plan 916 S4)
# Text ops that become tracked edits while the switch is on.
ROUTED_OPS = ("replace_text", "replace_paragraph", "insert_text")
# Ops that keep their own type and take `tracked`/`author` themselves (plan 864 S2: insert_paragraph writes the
# new paragraphs as a tracked insertion by direct XML; S3a: replace_table_row and replace_section record their
# line edits, added and removed lines or paragraphs as tracked changes by direct XML, and refuse adding or
# removing a whole row; plan 916 S5: insert_table writes the table and caption as a tracked insertion).
TRACKED_IN_PLACE_OPS = ("insert_paragraph", "replace_table_row", "replace_section", "insert_table")

CUSTOM_PART = "docProps/custom.xml"
CUSTOM_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
_CONTENT_TYPES = "[Content_Types].xml"
_PACKAGE_RELS = "_rels/.rels"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass(frozen=True)
class TrackSetting:
    on: bool
    author: str
    is_set: bool  # the property exists (on or off); unset documents behave as off

    def as_dict(self) -> dict:
        """What inspect/audit show: ``{"on", "author"}`` (author None when the switch was never set)."""
        return {"on": self.on, "author": self.author if self.is_set else None}


OFF = TrackSetting(on=False, author=DEFAULT_AUTHOR, is_set=False)


def not_trackable(op: str) -> RefusedUnsafe:
    return RefusedUnsafe(NOT_TRACKABLE_MESSAGE.format(op=op))


# -- reading -----------------------------------------------------------------------------------------------

def _custom_part_name(parts_names, rels_xml: bytes | None) -> str | None:
    """The package's custom-properties part, by its relationship (docProps/custom.xml in practice)."""

    if rels_xml:
        try:
            root = ET.fromstring(rels_xml)
        except ET.ParseError as exc:
            raise RefusedUnsafe(f"cannot read {_PACKAGE_RELS}: {exc}") from exc
        for rel in root.iter(f"{{{_REL_NS}}}Relationship"):
            if rel.get("Type") == CUSTOM_REL_TYPE and rel.get("TargetMode") != "External":
                target = (rel.get("Target") or "").lstrip("/")
                if target in parts_names:
                    return target
    return CUSTOM_PART if CUSTOM_PART in parts_names else None


def _parse_value(text: str) -> TrackSetting:
    try:
        value = json.loads(text)
    except json.JSONDecodeError as exc:
        raise RefusedUnsafe(f"the {PROPERTY_NAME} property is not valid JSON ({text[:80]!r}); "
                            "call set_track_changes to rewrite it") from exc
    if not isinstance(value, dict) or not isinstance(value.get("on"), bool):
        raise RefusedUnsafe(f"the {PROPERTY_NAME} property must be JSON like {{\"on\": true, \"author\": "
                            f"\"Claude\"}}, found {text[:80]!r}; call set_track_changes to rewrite it")
    author = value.get("author")
    if not isinstance(author, str) or not author.strip():
        author = DEFAULT_AUTHOR
    return TrackSetting(on=value["on"], author=author, is_set=True)


def _setting_from_xml(xml: bytes) -> TrackSetting:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        raise RefusedUnsafe(f"cannot read {CUSTOM_PART}: {exc}") from exc
    for prop in root:
        if prop.tag.rsplit("}", 1)[-1] == "property" and prop.get("name") == PROPERTY_NAME:
            values = list(prop)
            if len(values) != 1 or values[0].tag != f"{{{VT_NS}}}lpwstr":
                raise RefusedUnsafe(f"the {PROPERTY_NAME} property is not a text (vt:lpwstr) value; "
                                    "call set_track_changes to rewrite it")
            return _parse_value(values[0].text or "")
    return OFF


def read_setting(path: str | Path) -> TrackSetting:
    """The document's switch: off (unset) when the property is missing."""

    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        rels = archive.read(_PACKAGE_RELS) if _PACKAGE_RELS in names else None
        part = _custom_part_name(names, rels)
        if part is None:
            return OFF
        return _setting_from_xml(archive.read(part))


def describe(path: str | Path) -> dict:
    """inspect/audit's ``track_changes`` field; never raises (a broken property is reported, not fatal)."""

    try:
        return read_setting(path).as_dict()
    except (RefusedUnsafe, OSError, zipfile.BadZipFile, KeyError) as exc:
        return {"on": False, "author": None, "error": str(exc)}


# -- writing -----------------------------------------------------------------------------------------------

_PROPERTY_ELEMENT = re.compile(
    r"<(?P<prefix>[\w.-]+:)?property\b(?P<attrs>[^>]*)>.*?</(?P=prefix)?property\s*>|"
    r"<(?:[\w.-]+:)?property\b[^>]*/>", re.S)
_NAME_ATTR = re.compile(r"""\bname\s*=\s*(["'])(?P<name>.*?)\1""", re.S)
_PID_ATTR = re.compile(r"""\bpid\s*=\s*(["'])\s*(?P<pid>\d+)\s*\1""")


def _value_xml(vt_prefix: str | None, on: bool, author: str) -> str:
    payload = html.escape(json.dumps({"on": on, "author": author}, ensure_ascii=False), quote=False)
    if vt_prefix:
        return f"<{vt_prefix}:lpwstr>{payload}</{vt_prefix}:lpwstr>"
    return f'<vt:lpwstr xmlns:vt="{VT_NS}">{payload}</vt:lpwstr>'


def _new_custom_xml(on: bool, author: str) -> bytes:
    prop = (f'<property fmtid="{FMTID}" pid="2" name="{PROPERTY_NAME}">'
            f"{_value_xml('vt', on, author)}</property>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<Properties xmlns="{CUSTOM_NS}" xmlns:vt="{VT_NS}">{prop}</Properties>').encode("utf-8")


def _updated_custom_xml(xml: bytes, on: bool, author: str) -> bytes:
    """Replace the switch's own <property> element, or append one; every other byte is kept."""

    try:
        text = xml.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise RefusedUnsafe(f"{CUSTOM_PART} is not UTF-8; set the switch after saving the document in Word") from exc
    root_match = re.search(r"<(?P<prefix>[\w.-]+:)?Properties\b(?P<attrs>[^>]*?)(?P<close>/?)>", text)
    if root_match is None:
        raise RefusedUnsafe(f"{CUSTOM_PART} has no Properties root")
    prefix = root_match.group("prefix") or ""
    vt_match = re.search(rf"""xmlns:([\w.-]+)\s*=\s*(["']){re.escape(VT_NS)}\2""", root_match.group("attrs"))
    vt_prefix = vt_match.group(1) if vt_match else None
    value = _value_xml(vt_prefix, on, author)
    pids, own = [], None
    for match in _PROPERTY_ELEMENT.finditer(text, root_match.end()):
        attrs = match.group(0)[: match.group(0).find(">")]
        name = _NAME_ATTR.search(attrs)
        pid = _PID_ATTR.search(attrs)
        if pid:
            pids.append(int(pid.group("pid")))
        if name and html.unescape(name.group("name")) == PROPERTY_NAME:
            if own is not None:
                raise RefusedUnsafe(f"{CUSTOM_PART} holds {PROPERTY_NAME} twice; remove one in Word "
                                    "(File > Info > Properties > Advanced Properties > Custom)")
            own = match
    if own is not None:
        start_tag = own.group(0)[: own.group(0).find(">") + 1]
        if start_tag.endswith("/>"):
            start_tag = start_tag[:-2].rstrip() + ">"
        element = f"{start_tag}{value}</{prefix}property>"
        return (text[: own.start()] + element + text[own.end():]).encode("utf-8")
    pid = max(pids, default=1) + 1
    element = f'<{prefix}property fmtid="{FMTID}" pid="{pid}" name="{PROPERTY_NAME}">{value}</{prefix}property>'
    if root_match.group("close"):
        head = text[: root_match.start()] + root_match.group(0)[:-2].rstrip() + ">"
        return (head + element + f"</{prefix}Properties>" + text[root_match.end():]).encode("utf-8")
    close = text.rindex(f"</{prefix}Properties")
    return (text[:close] + element + text[close:]).encode("utf-8")


def _added_to_package(content_types: bytes, rels: bytes) -> tuple[bytes, bytes]:
    types = content_types.decode("utf-8")
    if f'PartName="/{CUSTOM_PART}"' not in types:
        close = types.rindex("</Types>")
        types = (types[:close] + f'<Override PartName="/{CUSTOM_PART}" ContentType="{CUSTOM_CONTENT_TYPE}"/>'
                 + types[close:])
    rel_text = rels.decode("utf-8")
    used = set(re.findall(r"""\bId\s*=\s*["']([^"']+)["']""", rel_text))
    number = max([int(m) for m in re.findall(r"rId(\d+)", " ".join(used))] + [0]) + 1
    while f"rId{number}" in used:
        number += 1
    close = rel_text.rindex("</Relationships>")
    rel_text = (rel_text[:close] + f'<Relationship Id="rId{number}" Type="{CUSTOM_REL_TYPE}" '
                f'Target="{CUSTOM_PART}"/>' + rel_text[close:])
    return types.encode("utf-8"), rel_text.encode("utf-8")


def _other_properties(xml: bytes) -> list[tuple]:
    root = ET.fromstring(xml)
    return [(prop.get("fmtid"), prop.get("pid"), prop.get("name"), ET.tostring(prop))
            for prop in root if prop.get("name") != PROPERTY_NAME]


def write_setting(source: str | Path, output: str | Path, *, on: bool, author: str | None = None) -> dict:
    """Write the switch into a new package at ``output`` (direct zip edit; ``output`` must not exist).

    ``author`` None keeps the current author (else "Claude"). An unchanged setting copies the file and
    reports ``read_only``. Checked before returning: the new value reads back, every other entry is
    byte-identical except the custom part (and, when it was created, the content types and package rels),
    and every other custom property is unchanged."""

    source, output = Path(source), Path(output)
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    try:
        previous = read_setting(source)
        shown_before = previous.as_dict()
    except RefusedUnsafe as exc:  # an unreadable value is what set_track_changes is told to rewrite
        previous = None
        shown_before = {"on": False, "author": None, "error": str(exc)}
    author = (author or "").strip() or (previous.author if previous else "") or DEFAULT_AUTHOR
    wanted = TrackSetting(on=on, author=author, is_set=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    base = {"track_changes": wanted.as_dict(), "previous": shown_before, "index_delta": 0}
    if previous == wanted:
        output.write_bytes(source.read_bytes())
        return {**base, "changed": False, "read_only": True,
                "summary": f"tracking already {'on' if on else 'off'} (author {author!r}); nothing written"}
    with zipfile.ZipFile(source) as archive:
        infos = archive.infolist()
        names = {info.filename for info in infos}
        rels = archive.read(_PACKAGE_RELS) if _PACKAGE_RELS in names else None
        part = _custom_part_name(names, rels)
        changed: dict[str, bytes] = {}
        created = part is None
        if created:
            if rels is None or _CONTENT_TYPES not in names:
                raise RefusedUnsafe("the package has no _rels/.rels or [Content_Types].xml; save it in Word first")
            part = CUSTOM_PART
            changed[part] = _new_custom_xml(on, author)
            changed[_CONTENT_TYPES], changed[_PACKAGE_RELS] = _added_to_package(archive.read(_CONTENT_TYPES), rels)
        else:
            old_custom = archive.read(part)
            changed[part] = _updated_custom_xml(old_custom, on, author)
        with zipfile.ZipFile(output, "x") as target:
            for info in infos:
                data = changed.get(info.filename)
                target.writestr(info, archive.read(info.filename) if data is None else data,
                                compress_type=info.compress_type)
            if created:
                target.writestr(part, changed[part], compress_type=zipfile.ZIP_DEFLATED)
    _verify(source, output, part, created, wanted)
    return {**base, "changed": True, "created_part": created,
            "summary": f"tracking {'on' if on else 'off'} (author {author!r})"}


def _verify(source: Path, output: Path, part: str, created: bool, wanted: TrackSetting) -> None:
    try:
        written = read_setting(output)
    except RefusedUnsafe as exc:
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{SET_OP}: the written property does not read back: {exc}") from exc
    allowed = {part} | ({_CONTENT_TYPES, _PACKAGE_RELS} if created else set())
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
        old = {name: before.read(name) for name in before.namelist()}
        new = {name: after.read(name) for name in after.namelist()}
    problems = []
    if written != wanted:
        problems.append(f"read back {written} instead of {wanted}")
    if set(new) - set(old) - ({part} if created else set()) or set(old) - set(new):
        problems.append("the package's entries changed")
    problems += [f"{name} changed" for name in old if name not in allowed and old[name] != new.get(name)]
    if not created and _other_properties(old[part]) != _other_properties(new[part]):
        problems.append("another custom property changed")
    if problems:
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{SET_OP}: " + "; ".join(problems))


def apply_operation(source, output, operation) -> dict:
    """edit_batch entry for the set_track_changes op."""

    from .audit import audit_docx, file_sha256

    details = write_setting(source, output, on=operation.on, author=operation.author)
    details["output_sha256"] = file_sha256(output)
    details["paragraph_count"] = audit_docx(output).paragraph_count
    return details


# -- routing -----------------------------------------------------------------------------------------------

def route_operation(original, operation, setting: TrackSetting):
    """Decide how one edit_document op runs under the switch.

    ``original`` is the op as the caller sent it (its ``model_fields_set`` says which fields were given);
    ``operation`` is the same op after index mapping and text cleaning. Returns ``(op_to_run, info)``:
    ``info`` is None for ops the switch does not touch, else ``{"tracking", "tracking_author"?, "routed_as"?}``.
    Raises NOT_TRACKABLE_MESSAGE for an op that cannot be tracked while tracking applies to it."""

    from .edit_models import TrackedReplaceParagraphOperation, TrackedReplaceTextOperation

    kind = operation.type
    given = getattr(original, "model_fields_set", set())
    if kind == "delete_paragraph":  # plan 864 S3c: as delete_citation, but the author falls back to the switch's
        tracked = operation.tracked if operation.tracked is not None else (setting.on if setting.is_set else True)
        author = operation.author or setting.author or DEFAULT_AUTHOR
        routed = operation.model_copy(update={"tracked": tracked, "author": author})
        return routed, _info(tracked, author)
    if kind == "style_text":  # plan 864 S6: tracked formatting follows the switch (never set: plain)
        tracked = operation.tracked if operation.tracked is not None else setting.on
        author = operation.author or setting.author or DEFAULT_AUTHOR
        routed = operation.model_copy(update={"tracked": tracked, "author": author})
        return routed, _info(tracked, author)
    if kind == "delete_citation":
        tracked = operation.tracked if ("tracked" in given or not setting.is_set) else setting.on
        author = operation.author if ("author" in given or not setting.is_set) else setting.author
        routed = operation.model_copy(update={"tracked": tracked, "author": author})
        return routed, _info(tracked, author)
    if kind == "sweep_characters":
        author = operation.author if ("author" in given or not setting.on) else setting.author
        return operation.model_copy(update={"author": author}), _info(True, author)
    if kind not in ROUTED_OPS + NOT_TRACKABLE_OPS + TRACKED_IN_PLACE_OPS:
        return operation, None
    tracked = operation.tracked if operation.tracked is not None else setting.on
    author = operation.author or setting.author or DEFAULT_AUTHOR
    if not tracked:
        return operation, _info(False, None)
    if kind in NOT_TRACKABLE_OPS:
        raise not_trackable(kind)
    if kind in TRACKED_IN_PLACE_OPS:
        info = _info(True, author)
        info["routed_as"] = kind
        return operation.model_copy(update={"tracked": True, "author": author}), info
    base = {"op_id": operation.op_id, "depends_on": list(operation.depends_on), "group": operation.group,
            "keep": list(operation.keep), "paragraph_index": operation.paragraph_index, "author": author}
    if kind == "replace_text":
        if not operation.match_case or operation.occurrence != 1 or operation.expected_count != 1:
            raise RefusedUnsafe("Only one unique, case-sensitive match is supported; inspect and use a unique anchor.")
        routed = TrackedReplaceTextOperation(type="tracked_replace_text", find=operation.find,
                                             replace=operation.replace, **base)
    elif kind == "replace_paragraph":
        routed = TrackedReplaceParagraphOperation(type="tracked_replace_paragraph", text=operation.text, **base)
    else:  # insert_text: replace the anchor with itself plus the text, as one tracked insertion
        anchor, text = operation.anchor, operation.text
        replace = anchor + text if operation.position == "after" else text + anchor
        routed = TrackedReplaceTextOperation(type="tracked_replace_text", find=anchor, replace=replace, **base)
    info = _info(True, author)
    info["routed_as"] = routed.type
    return routed, info


def _info(tracked: bool, author: str | None) -> dict:
    info: dict = {"tracking": "on" if tracked else "off"}
    if tracked and author:
        info["tracking_author"] = author
    return info
