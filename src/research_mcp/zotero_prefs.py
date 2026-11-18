"""Plan 864 decision 21: the document's Zotero preferences and `set_citation_style`.

Zotero's Word plugin keeps one string per document, its document data, in custom document properties
``ZOTERO_PREF_1..n`` (docProps/custom.xml), each holding the next 255 characters (UTF-16 code units, as Word
stores them). Zotero 5+ writes XML (data-version 3)::

    <data data-version="3" zotero-version="9.0.6"><session id="qySbiIQu"/>
    <style id="http://www.zotero.org/styles/apa" locale="en-US" hasBibliography="1"
    bibliographyStyleHasBeenSet="1"/><prefs><pref name="fieldType" value="Field"/></prefs></data>

and data-version 4 is the same as JSON. ``set_citation_style`` changes the style the way Zotero's Document
Preferences dialog does (``Session.setDocPrefs`` in Zotero's integration.js): the new style id, the locale
(kept unless given), ``hasBibliography`` from the installed style, ``bibliographyStyleHasBeenSet`` reset so
Zotero lays out Word's Bibliography paragraph style for the new style on the next refresh. The session id is
renewed when only the locale changes, because a Zotero session cached under the old id keeps its style and
locale while the style id is unchanged (``Session.setData``). The prefs are written by direct zip edit (every
other part and custom property byte-identical), then the document goes through the refresh_citations path
(`refresh_ops._refresh_staged`), with its guards and bibliography diff.

Styles are the ones Zotero loads: ``<data dir>/styles/*.csl`` and ``styles/hidden/*.csl`` (names starting with
"." skipped; the first file wins for a duplicate id). A dependent style takes its class and bibliography from
its installed parent. The data dir is ``RESEARCH_MCP_ZOTERO_STYLES_DIR`` (the styles folder itself) when set,
else the Zotero profile's ``extensions.zotero.dataDir`` (when ``useDataDir`` is true), else ``~/Zotero``.

Not done, refused instead: switching between an in-text and a note style (Zotero then converts every
citation field to or from footnotes, which only its Document Preferences dialog in Word does), a style with
no bibliography for a document with a bibliography field, and documents with no Zotero data yet.
"""

from __future__ import annotations

from dataclasses import dataclass
import difflib
import html
import json
import os
from pathlib import Path
import re
import secrets
import shutil
import string
import tempfile
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

from .errors import RefusedUnsafe, VerificationFailed

OP = "set_citation_style"
PREF_PREFIX = "ZOTERO_PREF_"
CHUNK_UNITS = 255  # Word plugin: each ZOTERO_PREF_n holds at most 255 UTF-16 code units
STYLE_URI = "http://www.zotero.org/styles/"
CSL_NS = "http://purl.org/net/xbiblio/csl"
STYLES_DIR_ENV = "RESEARCH_MCP_ZOTERO_STYLES_DIR"
BIBLIOGRAPHY_LAYOUT_WARNING = (
    "Zotero lays out Word's Bibliography paragraph style (indents, line and entry spacing) for the new style on "
    "this refresh, as its Document Preferences dialog does; re-run apply_document_style if a journal profile "
    "sets those.")

_CUSTOM_PART = "docProps/custom.xml"
_PACKAGE_RELS = "_rels/.rels"
_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
_VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


# -- installed styles --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class CitationStyle:
    style_id: str
    title: str
    style_class: str  # "in-text" or "note" (a dependent style's parent decides)
    has_bibliography: bool
    file: str
    parent: str | None = None
    hidden: bool = False

    @property
    def short_name(self) -> str:
        return self.style_id[len(STYLE_URI):] if self.style_id.startswith(STYLE_URI) else Path(self.file).stem

    def as_dict(self) -> dict[str, Any]:
        return {"style_id": self.style_id, "short_name": self.short_name, "title": self.title,
                "style_class": self.style_class, "has_bibliography": self.has_bibliography,
                "file": self.file, "parent": self.parent}


def _profile_data_dir() -> Path | None:
    """Zotero's custom data directory from its profile prefs.js, when the profile uses one."""

    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    root = Path(appdata) / "Zotero" / "Zotero"
    candidates: list[Path] = []
    ini = root / "profiles.ini"
    if ini.is_file():
        sections = re.split(r"^\[", ini.read_text(encoding="utf-8", errors="replace"), flags=re.M)
        for section in sorted(sections, key=lambda s: "Default=1" not in s):
            match = re.search(r"^Path=(.+)$", section, re.M)
            if match:
                rel = "IsRelative=0" not in section
                path = match.group(1).strip()
                candidates.append(root / path if rel else Path(path))
    candidates += sorted((root / "Profiles").glob("*")) if (root / "Profiles").is_dir() else []
    for profile in candidates:
        prefs = profile / "prefs.js"
        if not prefs.is_file():
            continue
        text = prefs.read_text(encoding="utf-8", errors="replace")
        use = re.search(r'user_pref\("extensions\.zotero\.useDataDir",\s*(true|false)\)', text)
        data = re.search(r'user_pref\("extensions\.zotero\.dataDir",\s*"((?:[^"\\]|\\.)*)"\)', text)
        if use and use.group(1) == "true" and data:
            return Path(json.loads(f'"{data.group(1)}"'))
        return None
    return None


def styles_dir() -> Path:
    """The folder Zotero loads styles from (see the module docstring)."""

    override = os.environ.get(STYLES_DIR_ENV)
    if override:
        return Path(override)
    data_dir = _profile_data_dir() or Path.home() / "Zotero"
    return data_dir / "styles"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _read_csl(path: Path) -> dict[str, Any] | None:
    try:
        root = ET.parse(path).getroot()
    except (ET.ParseError, OSError):
        return None
    if _local(root.tag) != "style":
        return None
    info = next((child for child in root if _local(child.tag) == "info"), None)
    if info is None:
        return None
    style_id = title = parent = None
    for child in info:
        name = _local(child.tag)
        if name == "id" and style_id is None:
            style_id = (child.text or "").strip()
        elif name == "title" and title is None:
            title = (child.text or "").strip()
        elif name == "link" and parent is None and child.get("rel") in ("source", "independent-parent"):
            parent = child.get("href")
    if not style_id:
        return None
    return {"id": style_id, "title": title or style_id, "class": root.get("class") or "",
            "bibliography": any(_local(el.tag) == "bibliography" for el in root.iter()), "parent": parent}


def installed_styles(directory: str | Path | None = None) -> dict[str, CitationStyle]:
    """Every style Zotero would load from ``directory`` (default: `styles_dir()`), by id."""

    directory = Path(directory) if directory is not None else styles_dir()
    raw: dict[str, tuple[dict[str, Any], Path, bool]] = {}
    for folder, hidden in ((directory, False), (directory / "hidden", True)):
        if not folder.is_dir():
            continue
        for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
            if path.name.startswith(".") or path.suffix.lower() != ".csl" or not path.is_file():
                continue
            data = _read_csl(path)
            if data and data["id"] not in raw and data["parent"] != data["id"]:
                raw[data["id"]] = (data, path, hidden)
    styles: dict[str, CitationStyle] = {}
    for style_id, (data, path, hidden) in raw.items():
        source = data
        if data["parent"]:
            parent = raw.get(data["parent"])
            if parent is None:
                continue  # Zotero cannot use a dependent style whose parent is not installed
            source = parent[0]
        styles[style_id] = CitationStyle(style_id=style_id, title=data["title"], style_class=source["class"],
                                         has_bibliography=source["bibliography"], file=str(path),
                                         parent=data["parent"], hidden=hidden)
    return styles


def resolve_style(query: str, directory: str | Path | None = None) -> CitationStyle:
    """The installed style ``query`` names: its id, its short name (id after .../styles/ or file stem), or its
    title (case-insensitive). Refused, naming the closest installed styles, when none matches."""

    directory = Path(directory) if directory is not None else styles_dir()
    styles = installed_styles(directory)
    wanted = (query or "").strip()
    if not wanted:
        raise RefusedUnsafe(f"{OP}: give a style id or short name, e.g. american-medical-association")
    if wanted in styles:
        return styles[wanted]
    folded = wanted.casefold()
    for matcher in (lambda s: s.style_id.casefold() == folded,
                    lambda s: s.short_name.casefold() == folded,
                    lambda s: Path(s.file).stem.casefold() == folded,
                    lambda s: s.title.casefold() == folded):
        hits = [s for s in styles.values() if matcher(s)]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            raise RefusedUnsafe(f"{OP}: style {query!r} matches several installed styles: "
                                + ", ".join(s.style_id for s in hits) + "; give the full id")
    short = folded[len(STYLE_URI):] if folded.startswith(STYLE_URI) else folded
    names = {s.short_name.casefold(): s.short_name for s in styles.values()}
    close = difflib.get_close_matches(short, list(names), n=5, cutoff=0.5)
    close += [names[n] for n in sorted(names) if short and short in n and names[n] not in
              [names.get(c) for c in close]][: max(0, 5 - len(close))]
    shown = ", ".join(names.get(c, c) for c in close) or "none similar"
    if not styles:
        raise RefusedUnsafe(f"{OP}: no Zotero styles found in {directory}; is Zotero installed there? "
                            f"({STYLES_DIR_ENV} overrides the folder)")
    raise RefusedUnsafe(f"{OP}: style {query!r} is not installed in Zotero ({directory}); closest installed: "
                        f"{shown}. Install it in Zotero (Settings > Cite > Styles) or pick an installed one.")


# -- the document data ---------------------------------------------------------------------------------------


@dataclass(frozen=True)
class DocumentData:
    text: str
    data_version: int
    style_id: str
    locale: str | None
    has_bibliography: bool
    bibliography_style_has_been_set: bool
    session_id: str
    note_type: int  # 0 in-text, 1 footnotes, 2 endnotes

    def summary(self) -> dict[str, Any]:
        return {"style_id": self.style_id, "locale": self.locale, "has_bibliography": self.has_bibliography,
                "bibliography_style_has_been_set": self.bibliography_style_has_been_set,
                "note_type": self.note_type, "data_version": self.data_version}


_STYLE_ELEMENT = re.compile(r"<style\b[^>]*?/?>(?:</style>)?")
_SESSION_ID = re.compile(r"""(<session\b[^>]*?\bid\s*=\s*)(["'])(.*?)\2""")


def parse_data(text: str) -> DocumentData:
    """Read Zotero's document data (XML data-version 3 or JSON data-version 4)."""

    if text.startswith("{"):
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise RefusedUnsafe(f"{OP}: the document's Zotero data is not readable JSON: {exc}") from exc
        style = data.get("style") if isinstance(data, dict) else None
        if not isinstance(style, dict) or not style.get("styleID"):
            raise RefusedUnsafe(f"{OP}: the document's Zotero data has no style")
        prefs = data.get("prefs") or {}
        return DocumentData(text=text, data_version=4, style_id=style["styleID"], locale=style.get("locale") or None,
                            has_bibliography=bool(style.get("hasBibliography")),
                            bibliography_style_has_been_set=bool(style.get("bibliographyStyleHasBeenSet")),
                            session_id=str(data.get("sessionID") or ""), note_type=_note_type(prefs.get("noteType")))
    if not text.startswith("<"):
        raise RefusedUnsafe(f"{OP}: the document's Zotero data is in the pre-Zotero-5 format; open the document "
                            "in Word and use Zotero's Document Preferences once to upgrade it")
    try:
        root = ET.fromstring(text)
    except ET.ParseError as exc:
        raise RefusedUnsafe(f"{OP}: the document's Zotero data is not readable XML: {exc}") from exc
    style = root.find("style")
    if root.tag != "data" or style is None or not style.get("id"):
        raise RefusedUnsafe(f"{OP}: the document's Zotero data has no <style id=...>")
    if len(_STYLE_ELEMENT.findall(text)) != 1:
        raise RefusedUnsafe(f"{OP}: the document's Zotero data holds more than one <style>; reset it in Word")
    session = root.find("session")
    prefs = {p.get("name"): p.get("value") for p in root.iterfind("prefs/pref")}
    return DocumentData(text=text, data_version=int(root.get("data-version") or 2), style_id=style.get("id"),
                        locale=style.get("locale") or None, has_bibliography=style.get("hasBibliography") == "1",
                        bibliography_style_has_been_set=style.get("bibliographyStyleHasBeenSet") == "1",
                        session_id=session.get("id", "") if session is not None else "",
                        note_type=_note_type(prefs.get("noteType")))


def _note_type(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _new_session_id() -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(8))


def with_style(data: DocumentData, style: CitationStyle, locale: str | None, *, renew_session: bool) -> str:
    """The document data with the new style, as Zotero's Document Preferences writes it; every other byte
    of the string kept (session and prefs), the session id renewed when asked."""

    if data.data_version == 4:
        payload = json.loads(data.text)
        payload["style"] = {**payload["style"], "styleID": style.style_id, "locale": locale,
                            "hasBibliography": style.has_bibliography, "bibliographyStyleHasBeenSet": False}
        if renew_session:
            payload["sessionID"] = _new_session_id()
        return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    element = (f'<style id="{html.escape(style.style_id)}" ' + (f'locale="{html.escape(locale)}" ' if locale else "")
               + f'hasBibliography="{1 if style.has_bibliography else 0}" bibliographyStyleHasBeenSet="0"/>')
    text = _STYLE_ELEMENT.sub(lambda _: element, data.text, count=1)
    if renew_session:
        text, count = _SESSION_ID.subn(lambda m: f"{m.group(1)}{m.group(2)}{_new_session_id()}{m.group(2)}", text,
                                       count=1)
        if not count:
            raise RefusedUnsafe(f"{OP}: the document's Zotero data has no session id to renew")
    return text


# -- ZOTERO_PREF_n in docProps/custom.xml --------------------------------------------------------------------

_PROPERTY_ELEMENT = re.compile(
    r"<(?P<prefix>[\w.-]+:)?property\b(?P<attrs>[^>]*)>.*?</(?P=prefix)?property\s*>|"
    r"<(?:[\w.-]+:)?property\b[^>]*/>", re.S)
_NAME_ATTR = re.compile(r"""\bname\s*=\s*(["'])(?P<name>.*?)\1""", re.S)
_PID_ATTR = re.compile(r"""\bpid\s*=\s*(["'])\s*(?P<pid>\d+)\s*\1""")
_LPWSTR = re.compile(r"<(?P<p>[\w.-]+:)?lpwstr\b[^>]*?(?:/>|>(?P<text>.*?)</(?P=p)?lpwstr\s*>)", re.S)


@dataclass
class _Pref:
    number: int
    start: int
    end: int
    raw: str
    value: str | None  # None: not a single vt:lpwstr


def _custom_part(archive: zipfile.ZipFile) -> str | None:
    from .track_setting import _custom_part_name

    names = set(archive.namelist())
    rels = archive.read(_PACKAGE_RELS) if _PACKAGE_RELS in names else None
    return _custom_part_name(names, rels)


def _prefs_of(text: str) -> tuple[list[_Pref], list[int]]:
    prefs, pids = [], []
    for match in _PROPERTY_ELEMENT.finditer(text):
        raw = match.group(0)
        head = raw[: raw.find(">")]
        pid = _PID_ATTR.search(head)
        if pid:
            pids.append(int(pid.group("pid")))
        name = _NAME_ATTR.search(head)
        name = html.unescape(name.group("name")) if name else ""
        if not name.startswith(PREF_PREFIX) or not name[len(PREF_PREFIX):].isdigit():
            continue
        values = list(_LPWSTR.finditer(raw))
        value = None
        if len(values) == 1:
            value = html.unescape(values[0].group("text") or "")
        prefs.append(_Pref(int(name[len(PREF_PREFIX):]), match.start(), match.end(), raw, value))
    return prefs, pids


def _joined(prefs: list[_Pref]) -> str:
    by_number: dict[int, _Pref] = {}
    for pref in prefs:
        if pref.number in by_number:
            raise RefusedUnsafe(f"{OP}: {PREF_PREFIX}{pref.number} appears twice in {_CUSTOM_PART}")
        if pref.value is None:
            raise RefusedUnsafe(f"{OP}: {PREF_PREFIX}{pref.number} is not a text (vt:lpwstr) property")
        by_number[pref.number] = pref
    numbers = sorted(by_number)
    if numbers != list(range(1, len(numbers) + 1)):
        raise RefusedUnsafe(f"{OP}: the Zotero preference properties are not numbered 1..n ({numbers}); "
                            "reset them with Zotero's Document Preferences in Word")
    return "".join(by_number[n].value or "" for n in numbers)


def read_prefs(path: str | Path) -> str | None:
    """The document's Zotero data (ZOTERO_PREF_1..n joined), or None when it has none."""

    with zipfile.ZipFile(path) as archive:
        part = _custom_part(archive)
        if part is None:
            return None
        text = archive.read(part).decode("utf-8")
    prefs, _ = _prefs_of(text)
    return _joined(prefs) if prefs else None


def chunks(data: str) -> list[str]:
    """``data`` split as the Word plugin stores it: 255 UTF-16 code units per property."""

    units = data.encode("utf-16-le")
    size = CHUNK_UNITS * 2
    try:
        return [units[i:i + size].decode("utf-16-le") for i in range(0, len(units), size)] or [""]
    except UnicodeDecodeError as exc:
        raise RefusedUnsafe(f"{OP}: the Zotero data would split a character across two properties") from exc


def _without_prefs(text: str) -> str:
    prefs, _ = _prefs_of(text)
    for pref in reversed(prefs):
        text = text[: pref.start] + text[pref.end:]
    return text


def _custom_xml_with(text: str, data: str) -> str:
    prefs, pids = _prefs_of(text)
    by_number = {p.number: p for p in prefs}
    wanted = chunks(data)
    last = max(prefs, key=lambda p: p.start)
    prefix = re.match(r"<([\w.-]+:)?", last.raw).group(1) or ""
    vt = _LPWSTR.search(last.raw)
    vt_prefix = (vt.group("p") or "") if vt else "vt:"
    new_pid = max(pids, default=1)
    edits: list[tuple[int, int, str]] = []
    added = ""
    for number, chunk in enumerate(wanted, 1):
        escaped = html.escape(chunk, quote=False)
        pref = by_number.get(number)
        if pref is not None:
            if pref.value == chunk:
                continue
            value = _LPWSTR.search(pref.raw)
            raw = pref.raw[: value.start()] + f"<{vt_prefix}lpwstr>{escaped}</{vt_prefix}lpwstr>" + pref.raw[value.end():]
            edits.append((pref.start, pref.end, raw))
        else:
            new_pid += 1
            added += (f'<{prefix}property fmtid="{_FMTID}" pid="{new_pid}" name="{PREF_PREFIX}{number}">'
                      f"<{vt_prefix}lpwstr>{escaped}</{vt_prefix}lpwstr></{prefix}property>")
    edits += [(p.start, p.end, "") for p in prefs if p.number > len(wanted)]
    if added:
        edits.append((last.end, last.end, added))
    for start, end, new in sorted(edits, key=lambda e: (e[0], e[1]), reverse=True):
        text = text[:start] + new + text[end:]
    return text


def write_prefs(source: str | Path, output: str | Path, data: str) -> dict[str, Any]:
    """Write ``data`` as ZOTERO_PREF_1..n into a new package at ``output`` (must not exist): existing
    properties keep their place and pid (only a changed value is rewritten), new ones follow the last one with
    the next pid, extra ones are removed. Unchanged data copies the file byte for byte. Checked before
    returning: the data reads back, every other zip entry and every other byte of custom.xml are unchanged."""

    source, output = Path(source), Path(output)
    if output.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    with zipfile.ZipFile(source) as archive:
        part = _custom_part(archive)
        if part is None:
            raise RefusedUnsafe(f"{OP}: the document has no {_CUSTOM_PART}, so no Zotero preferences")
        old_text = archive.read(part).decode("utf-8")
        old_prefs, _ = _prefs_of(old_text)
        if not old_prefs:
            raise RefusedUnsafe(f"{OP}: the document has no {PREF_PREFIX}n properties")
        old_data = _joined(old_prefs)
        before_chunks = len(old_prefs)
        if old_data == data:
            shutil.copyfile(source, output)
            return {"changed": False, "chunks_before": before_chunks, "chunks_after": before_chunks}
        new_text = _custom_xml_with(old_text, data)
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "x") as target:
            for info in archive.infolist():
                payload = new_text.encode("utf-8") if info.filename == part else archive.read(info.filename)
                target.writestr(info, payload, compress_type=info.compress_type)
    problems = []
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
        if before.namelist() != after.namelist():
            problems.append("the package's entries changed")
        else:
            problems += [f"{name} changed" for name in before.namelist()
                         if name != part and before.read(name) != after.read(name)]
        written = after.read(part).decode("utf-8")
    if _without_prefs(written) != _without_prefs(old_text):
        problems.append(f"{_CUSTOM_PART} changed outside the {PREF_PREFIX}n properties")
    written_prefs, _ = _prefs_of(written)
    if _joined(written_prefs) != data:
        problems.append("the Zotero data does not read back")
    if any(len(p.value.encode("utf-16-le")) > CHUNK_UNITS * 2 for p in written_prefs):
        problems.append(f"a {PREF_PREFIX}n property is longer than {CHUNK_UNITS} characters")
    if problems:
        output.unlink(missing_ok=True)
        raise VerificationFailed(f"{OP}: " + "; ".join(problems))
    return {"changed": True, "chunks_before": before_chunks, "chunks_after": len(written_prefs)}


def other_custom_properties(path: str | Path) -> list[tuple[str, tuple[tuple[str, str], ...], str]]:
    """Every custom property except ZOTERO_PREF_n, in order: (name, attributes but `pid`, value XML).

    Compared after Word saves. `pid` is left out on purpose: Word numbers custom properties itself and renumbers
    the ones after a removed ZOTERO_PREF_n chunk; every other attribute (fmtid, linkTarget, ...) must stay."""

    with zipfile.ZipFile(path) as archive:
        part = _custom_part(archive)
        if part is None:
            return []
        root = ET.fromstring(archive.read(part))
    out = []
    for prop in root:
        name = prop.get("name") or ""
        if name.startswith(PREF_PREFIX) and name[len(PREF_PREFIX):].isdigit():
            continue
        attributes = tuple(sorted((key, value) for key, value in prop.attrib.items() if key != "pid"))
        out.append((name, attributes, "".join(ET.tostring(child, encoding="unicode") for child in prop)))
    return out


def custom_pid_problems(path: str | Path) -> list[str]:
    """Problems with the custom properties' `pid`s: present, ASCII digits, 2..2**31-1 (xsd:int), unique."""

    with zipfile.ZipFile(path) as archive:
        part = _custom_part(archive)
        if part is None:
            return []
        root = ET.fromstring(archive.read(part))
    problems: list[str] = []
    seen: dict[int, str] = {}
    for prop in root:
        name = prop.get("name") or "?"
        pid = prop.get("pid")
        if pid is None:
            problems.append(f"{name} has no pid")
        elif not re.fullmatch(r"[0-9]{1,10}", pid) or not 2 <= int(pid) <= 2**31 - 1:
            problems.append(f"{name} has pid {pid!r}, not an integer from 2 to {2**31 - 1}")
        elif int(pid) in seen:
            problems.append(f"{name} and {seen[int(pid)]} share pid {pid}")
        else:
            seen[int(pid)] = name
    return problems


# -- the tool ------------------------------------------------------------------------------------------------

_LOCALE = re.compile(r"^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$")
_WNS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_DELETING = (_WNS + "del", _WNS + "moveFrom")
_CITATION_ID = re.compile(r'"citationID"\s*:\s*"([^"]+)"')
_BODY_PART = "word/document.xml"


def _field_name(instruction: str) -> str:
    match = _CITATION_ID.search(instruction)
    return match.group(1) if match else (instruction.split(" ", 2)[1:2] or ["ZOTERO"])[0]


def unchecked_zotero_fields(path: str | Path) -> list[dict[str, str]]:
    """Zotero fields the refresh path would neither reset nor check, though Zotero re-renders them in the new style.

    The refresh reads complex fields with a `w:instrText` instruction in the main body only, so these are missed:
    a field whose instruction is deleted text (`w:delInstrText`, in a `w:del` or `w:moveFrom`; reason "deletion"),
    a field in any other part: footnotes, endnotes, headers, footers, comments (reason "part"), and a simple
    field `w:fldSimple w:instr="ADDIN ZOTERO_..."` anywhere (reason "simple"). Each entry names the field, its
    part and, for a deletion, the revision."""

    found: list[dict[str, str]] = []
    with zipfile.ZipFile(Path(path)) as archive:
        parts = sorted(name for name in archive.namelist()
                       if name.startswith("word/") and name.endswith(".xml") and "/" not in name[len("word/"):])
        roots = [(name, ET.fromstring(archive.read(name))) for name in parts]
    for part, root in roots:
        stack: list[dict[str, Any]] = []

        def report(instruction: str, reason: str, revision: ET.Element | None = None) -> None:
            entry = {"field": _field_name(instruction), "part": part, "reason": reason}
            if revision is not None:
                entry["revision"] = (f"w:{revision.tag[len(_WNS):]} id={revision.get(_WNS + 'id')} by "
                                     f"{revision.get(_WNS + 'author')!r}")
            found.append(entry)

        def visit(node: ET.Element, revision: ET.Element | None) -> None:
            tag = node.tag
            inner = node if tag in _DELETING else revision
            if tag == _WNS + "fldSimple":
                instruction = (node.get(_WNS + "instr") or "").strip()
                if instruction.startswith("ADDIN ZOTERO_"):
                    report(instruction, "simple", inner)
            elif tag == _WNS + "fldChar":
                kind = node.get(_WNS + "fldCharType", "")
                if kind == "begin":
                    stack.append({"instr": [], "deleted": None, "done": False})
                elif kind in ("separate", "end") and stack:
                    field = stack[-1]
                    if not field["done"]:
                        field["done"] = True
                        instruction = "".join(field["instr"]).strip()
                        if instruction.startswith("ADDIN ZOTERO_"):
                            if field["deleted"] is not None:
                                report(instruction, "deletion", field["deleted"])
                            elif part != _BODY_PART:
                                report(instruction, "part")
                    if kind == "end":
                        stack.pop()
            elif tag in (_WNS + "instrText", _WNS + "delInstrText") and stack and not stack[-1]["done"]:
                stack[-1]["instr"].append(node.text or "")
                if tag == _WNS + "delInstrText" and stack[-1]["deleted"] is None:
                    stack[-1]["deleted"] = revision if revision is not None else node
            for child in node:
                visit(child, inner)

        visit(root, None)
    return found


def zotero_fields_in_deletions(path: str | Path) -> list[dict[str, str]]:
    """The "deletion" entries of `unchecked_zotero_fields` in the main body: field and revision."""

    return [{"field": f["field"], "revision": f["revision"]} for f in unchecked_zotero_fields(path)
            if f["reason"] == "deletion" and f["part"] == _BODY_PART]


def _describe_unchecked(entry: dict[str, str]) -> str:
    where = entry["part"][len("word/"):]
    if entry["reason"] == "deletion":
        return f"{entry['field']} in {entry['revision']} ({where})"
    if entry["reason"] == "simple":
        return f"{entry['field']} stored as a simple field (w:fldSimple) in {where}"
    return f"{entry['field']} in {where}"


def plan_style_change(path: Path, style: str, locale: str | None,
                      styles_directory: str | Path | None = None) -> dict[str, Any]:
    """Everything decided before a byte is written: the target style, the new data, or a refusal."""

    from .refresh_ops import plan_refresh_xml
    from .revisions import load

    text = read_prefs(path)
    if text is None:
        raise RefusedUnsafe(f"{OP}: the document has no Zotero preferences (no {PREF_PREFIX}n properties); "
                            "insert one citation with Zotero in Word first")
    data = parse_data(text)
    target = resolve_style(style, styles_directory)
    if locale is not None:
        locale = locale.strip()
        if not _LOCALE.match(locale):
            raise RefusedUnsafe(f"{OP}: locale {locale!r} is not a locale code like en-US or de-DE")
    new_locale = data.locale if locale is None else locale
    if target.style_id == data.style_id and new_locale == data.locale:
        raise RefusedUnsafe(f"{OP}: the document already uses {target.style_id}"
                            + (f" ({new_locale})" if new_locale else "")
                            + "; call refresh_citations to re-render it")
    doc_class = "note" if data.note_type else "in-text"
    if target.style_class != doc_class:
        raise RefusedUnsafe(
            f"{OP}: {target.short_name} is a{'n' if target.style_class == 'in-text' else ''} {target.style_class} "
            f"style and the document's citations are {doc_class}; switching moves every citation "
            f"{'into' if target.style_class == 'note' else 'out of'} footnotes, which only Zotero's Document "
            "Preferences in Word does (Zotero tab > Document Preferences). Pick a "
            f"{doc_class} style here.")
    if doc_class == "note":
        raise RefusedUnsafe(
            f"{OP}: the document uses a note style ({data.style_id}); its citations live in footnotes or endnotes, "
            "which the refresh does not reset or check yet. Change a note style in Word (Zotero tab > Document "
            "Preferences).")
    unchecked = unchecked_zotero_fields(path)
    deleted = [f for f in unchecked if f["reason"] == "deletion"]
    if deleted:
        raise RefusedUnsafe(
            f"{OP}: {len(deleted)} Zotero field(s) sit inside a tracked deletion, which Zotero would re-render in "
            "the new style inside the revision: "
            + "; ".join(_describe_unchecked(d) for d in deleted)
            + ". Accept or reject the deletion first (list_revisions shows it).")
    if unchecked:
        raise RefusedUnsafe(
            f"{OP}: {len(unchecked)} Zotero field(s) are outside what the refresh resets and checks (the main "
            "body's complex fields), so Zotero would re-render them unchecked: "
            + "; ".join(_describe_unchecked(f) for f in unchecked)
            + ". Change the style in Word (Zotero tab > Document Preferences) for this document.")
    _, doc = load(path)
    try:
        _, reset, _kept, bibliographies = plan_refresh_xml(doc.xml, doc.body)
    except RefusedUnsafe as exc:
        raise RefusedUnsafe(f"{OP}: {exc}") from exc
    if not reset and not bibliographies:
        raise RefusedUnsafe(f"{OP}: the document has no Zotero citation or bibliography to re-render")
    if bibliographies and not target.has_bibliography:
        raise RefusedUnsafe(f"{OP}: {target.short_name} has no bibliography and the document has a Zotero "
                            "bibliography field; remove the bibliography in Word first or pick another style")
    renew = target.style_id == data.style_id  # only the locale changes: a cached Zotero session would keep it
    new_text = with_style(data, target, new_locale, renew_session=renew)
    return {"data": data, "target": target, "locale": new_locale, "new_text": new_text, "session_renewed": renew}


def _apply_staged(staged: Path, expected: str, plan: dict[str, Any], max_wait_seconds: float) -> dict[str, Any]:
    from . import refresh_ops
    from .audit import audit_docx

    actual = audit_docx(staged).sha256.upper()
    if actual != expected:
        raise RefusedUnsafe(f"{OP}: the source changed after its SHA-256 was checked (expected {expected}, "
                            f"staged {actual}). Reinspect the document with inspect_document.")
    others_before = other_custom_properties(staged)
    temp = staged.with_name(f".~research-mcp-prefs-{secrets.token_hex(4)}.docx")
    try:
        written = write_prefs(staged, temp, plan["new_text"])
        os.replace(temp, staged)
    finally:
        temp.unlink(missing_ok=True)
    try:
        result = refresh_ops._refresh_staged(staged, audit_docx(staged).sha256.upper(), max_wait_seconds)
    except RefusedUnsafe as exc:
        raise RefusedUnsafe(f"{OP}: {exc}") from exc
    after_text = read_prefs(staged)
    after = parse_data(after_text) if after_text else None
    if after is None or after.style_id != plan["target"].style_id:
        raise VerificationFailed(f"{OP}: after the refresh the document's Zotero style is "
                                 f"{after.style_id if after else None}, not {plan['target'].style_id}")
    if plan["locale"] and after.locale != plan["locale"]:
        raise VerificationFailed(f"{OP}: after the refresh the document's Zotero locale is {after.locale}, "
                                 f"not {plan['locale']}")
    if other_custom_properties(staged) != others_before:
        raise VerificationFailed(f"{OP}: a custom document property other than {PREF_PREFIX}n changed")
    pid_problems = custom_pid_problems(staged)
    if pid_problems:
        raise VerificationFailed(f"{OP}: after the refresh the custom document properties are not numbered "
                                 "correctly: " + "; ".join(pid_problems))
    warnings = list(result.get("warnings") or [])
    if plan["target"].has_bibliography and result.get("bibliography_after"):
        warnings.append(BIBLIOGRAPHY_LAYOUT_WARNING)
    return {**result, "warnings": warnings, "style_after": after.summary(),
            "prefs_chunks_before": written["chunks_before"], "prefs_chunks_after": len(chunks(after_text))}


def set_style_to_version(
    source: str | Path,
    output: str | Path,
    *,
    source_sha256: str,
    style: str,
    locale: str | None = None,
    preview: bool = False,
    max_wait_seconds: float = 300.0,
    styles_directory: str | Path | None = None,
) -> dict[str, Any]:
    """Change the document's Zotero citation style and re-render it into ``output`` (or in place)."""

    from .audit import audit_docx
    from .word import staged_document

    src = Path(source).resolve(strict=True)
    out = Path(output).resolve()
    if not preview and out != src and out.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    expected = source_sha256.upper()
    actual = audit_docx(src).sha256.upper()
    if expected != actual:
        raise RefusedUnsafe(f"source document SHA-256 mismatch: expected {expected}, actual {actual}. "
                            "Reinspect the document with inspect_document.")
    plan = plan_style_change(src, style, locale, styles_directory)
    head = {"style": plan["target"].as_dict(), "style_before": plan["data"].summary(),
            "session_renewed": plan["session_renewed"]}
    if preview:
        with tempfile.TemporaryDirectory(prefix="research-mcp-style-preview-") as temp:
            staged = Path(temp) / src.name
            shutil.copyfile(src, staged)
            result = _apply_staged(staged, expected, plan, max_wait_seconds)
        return {**result, "success": True, "preview": True, "source": str(src), "source_sha256": actual,
                "output": None, "output_sha256": None, **head}
    with staged_document(src, out) as staged:
        result = _apply_staged(Path(staged), expected, plan, max_wait_seconds)
    return {**result, "success": True, "preview": False, "source": str(src), "source_sha256": actual,
            "output": str(out), "output_sha256": audit_docx(out).sha256, **head}
