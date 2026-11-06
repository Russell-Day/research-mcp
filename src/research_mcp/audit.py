"""Structural auditing of docx packages without starting Word.

Fast, reads the OpenXML zip archive directly. Returns SHA-256 fingerprints,
counts of structural elements, and table dimensions so that callers can
establish pre- and post-operation invariants.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
import hashlib
import io
import json
from pathlib import Path
import re
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

from .ooxml import body_text, parse_table, top_level_prose_text

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = f"{{{NS['w']}}}"


@dataclass(frozen=True)
class TableAudit:
    index: int
    rows: int
    columns: int
    cell_hashes: tuple[str, ...]
    vertical_merges: tuple[tuple[int, int, int, str], ...] = ()
    horizontal_spans: tuple[tuple[int, int, int], ...] = ()


def is_file_locked(path: str | Path) -> tuple[bool, str | None]:
    """Check if a file is currently locked by another process (e.g. Word)."""
    p = Path(path).resolve()
    if not p.exists():
        return False, None
    lock_file = p.parent / f"~${p.name}"
    if lock_file.exists():
        return True, f"Word owner lock file exists: {lock_file.name}"
    try:
        with open(p, "a+b") as f:
            pass
    except (IOError, PermissionError) as e:
        return True, f"File is locked by another process: {e}"
    return False, None


@dataclass(frozen=True)
class Audit:
    path: str
    sha256: str
    body_text_sha256: str
    prose_text_sha256: str
    paragraph_count: int
    table_count: int
    tables: tuple[TableAudit, ...]
    bookmark_names: tuple[str, ...]
    bookmark_references: dict[str, tuple[str, ...]]
    field_count: int
    zotero_item_fields: int
    zotero_bibliography_fields: int
    zotero_pref_properties: int
    revision_count: int
    comment_count: int
    live_comment_count: int
    orphan_comment_ids: tuple[str, ...]
    drawing_count: int
    media_count: int
    live_media_count: int
    dangling_media_targets: tuple[str, ...]
    structural_errors: tuple[str, ...] = ()
    is_locked: bool = False
    lock_reason: str | None = None
    citation_readiness: dict[str, Any] = field(default_factory=dict)
    paragraph_count_basis: str = "raw OOXML paragraph elements (w:p); not Word COM paragraph numbering"

    def to_dict(self) -> dict:
        return asdict(self)

    def to_agent_dict(self) -> dict:
        """The audit as sent across the MCP boundary, without verifier-only data.

        `cell_hashes` exists so `tables.verify_*` and the benchmark verifier can
        prove cell text did not change. It is by far the largest thing in the
        audit -- 19,788 of the 20,202 characters under `tables` for the v42
        manuscript, 98% -- and it is SHA-256 hex, which no model can act on and
        which tokenizes close to worst case.

        A tool result is not a transient cost: it is appended to the
        conversation and every later turn carries it. Replacing the per-cell
        digests with one digest per table keeps the integrity claim expressible
        (two audits agree iff every cell agrees) and cuts the whole result by
        about 91% on the current manuscript.

        Everything else is preserved deliberately. In particular
        `bookmark_names` stays: `validate_release` refuses names absent from the
        audit, so an agent needs them to build a safe `clean_bookmarks` call,
        and forcing a second `list_bookmarks` round trip would cost more
        context than dropping them here saves.

        Core, `tables.py`, and every verifier keep using `to_dict`.
        """

        data = asdict(self)
        data["tables"] = [
            {key: value for key, value in table.items() if key != "cell_hashes"}
            | {"cell_text_sha256": cell_text_digest(table["cell_hashes"])}
            for table in data["tables"]
        ]
        return data

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), indent=indent, ensure_ascii=False)


def cell_text_digest(cell_hashes: "tuple[str, ...] | list[str]") -> str:
    """One digest standing in for a table's per-cell hashes.

    Order-sensitive and separator-delimited, so two tables agree on this digest
    if and only if they agree cell for cell in the same layout.
    """

    return hashlib.sha256(b"\x1f".join(h.encode("ascii") for h in cell_hashes)).hexdigest()


def file_sha256(path: str | Path) -> str:
    hasher = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest().upper()


def extract_body_text(root: ET.Element) -> str:
    return body_text(root)


def extract_prose_text(root: ET.Element) -> str:
    return top_level_prose_text(root)


def _part_structure_errors(xml_text: str, part_name: str) -> list[str]:
    """Validate compatibility prefixes and complex fields in one XML part."""
    errors: list[str] = []
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        return [f"{part_name}: invalid XML: {exc}"]

    # ET expands namespaced attributes to their URI, but QName-valued
    # attributes (mc:Ignorable and friends) retain prefix tokens in their
    # values.  iterparse gives us start-ns events so aliases and nested scope
    # are validated without assuming the conventional `mc` prefix.
    mc_uri = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    qname_attrs = {"Ignorable", "MustUnderstand", "PreserveElements", "PreserveAttributes", "ProcessContent", "Requires"}
    pending_ns: list[tuple[str, str]] = []
    scopes: list[dict[str, str]] = []
    try:
        for event, node in ET.iterparse(io.StringIO(xml_text), events=("start-ns", "start", "end")):
            if event == "start-ns":
                prefix, uri = node
                pending_ns.append((prefix or "", uri))
                continue
            if event == "start":
                scope = dict(scopes[-1]) if scopes else {}
                scope.update(dict(pending_ns))
                pending_ns.clear()
                local_tag = node.tag.rsplit("}", 1)[-1]
                tag_uri = node.tag[1:].split("}", 1)[0] if node.tag.startswith("{") else ""
                for attr, value in node.attrib.items():
                    local_attr = attr.rsplit("}", 1)[-1]
                    attr_uri = attr[1:].split("}", 1)[0] if attr.startswith("{") else ""
                    if local_attr not in qname_attrs:
                        continue
                    if local_attr == "Requires" and not (tag_uri == mc_uri and local_tag == "Choice"):
                        continue
                    if local_attr != "Requires" and attr_uri != mc_uri:
                        continue
                    for token in value.split():
                        # Ignorable and MustUnderstand contain prefix lists;
                        # the remaining values are whitespace-separated QNames.
                        prefix = token if local_attr in {"Ignorable", "MustUnderstand"} else token.split(":", 1)[0]
                        if ":" not in token and local_attr not in {"Ignorable", "MustUnderstand"}:
                            continue  # an unprefixed QName is valid in no namespace
                        if prefix and prefix not in scope:
                            label = "mc:Ignorable" if local_attr == "Ignorable" else local_attr
                            errors.append(f"{part_name}: undeclared {label} prefix: {prefix}")
                scopes.append(scope)
            elif event == "end" and scopes:
                scopes.pop()
    except ET.ParseError as exc:
        return [f"{part_name}: invalid XML: {exc}"]
    stack: list[bool] = []
    for node in root.iter():
        if node.tag != W + "fldChar":
            continue
        kind = node.get(W + "fldCharType", "")
        if kind == "begin":
            stack.append(False)
        elif kind == "separate":
            if not stack:
                errors.append(f"{part_name}: complex field separate without begin")
            elif stack[-1]:
                errors.append(f"{part_name}: complex field has duplicate separate")
            else:
                stack[-1] = True
        elif kind == "end":
            if not stack:
                errors.append(f"{part_name}: complex field end without begin")
            else:
                stack.pop()
    if stack:
        errors.append(f"{part_name}: {len(stack)} complex field(s) are unclosed")
    return errors


def document_structure_errors(path: str | Path) -> tuple[str, ...]:
    """Return deterministic OOXML preflight errors without starting Word."""
    resolved = Path(path).resolve(strict=True)
    errors: list[str] = []
    try:
        archive = zipfile.ZipFile(resolved)
    except (zipfile.BadZipFile, OSError) as exc:
        return (f"package: invalid ZIP: {exc}",)
    with archive:
        names = set(archive.namelist())
        if "word/document.xml" not in names:
            errors.append("package: missing word/document.xml")
        for name in sorted(names):
            if name.endswith("/") or not name.endswith(".xml"):
                continue
            try:
                xml_text = archive.read(name).decode("utf-8")
            except UnicodeDecodeError as exc:
                errors.append(f"{name}: invalid UTF-8: {exc}")
                continue
            errors.extend(_part_structure_errors(xml_text, name))
    return tuple(errors)


def bookmark_references(root: ET.Element) -> dict[str, tuple[str, ...]]:
    refs: dict[str, list[str]] = {}

    for hyperlink in root.findall(".//w:hyperlink", NS):
        anchor = hyperlink.get(W + "anchor")
        if anchor:
            refs.setdefault(anchor, []).append("hyperlink")

    for instr in root.findall(".//w:instrText", NS):
        text = instr.text or ""
        match = re.search(r'REF\s+([A-Za-z0-9_]+)', text)
        if match:
            refs.setdefault(match.group(1), []).append("ref_field")

    return {k: tuple(sorted(v)) for k, v in sorted(refs.items())}


def audit_docx(path: str | Path) -> Audit:
    resolved = Path(path).resolve(strict=True)
    with zipfile.ZipFile(resolved) as z:
        names = set(z.namelist())
        xml_text = z.read("word/document.xml").decode("utf-8")
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8") if "word/_rels/document.xml.rels" in names else ""
        comments_xml = z.read("word/comments.xml").decode("utf-8") if "word/comments.xml" in names else None
        custom = z.read("docProps/custom.xml").decode("utf-8") if "docProps/custom.xml" in names else ""

    root = ET.fromstring(xml_text)
    structural_errors = list(document_structure_errors(resolved))
    comments = ET.fromstring(comments_xml) if comments_xml is not None else None

    text = body_text(root)
    prose = top_level_prose_text(root)
    tables = tuple(
        TableAudit(
            index=idx,
            rows=t.rows,
            columns=t.columns,
            cell_hashes=t.cell_hashes,
            vertical_merges=t.vertical_merges,
            horizontal_spans=t.horizontal_spans,
        )
        for idx, t in enumerate((parse_table(tbl) for tbl in root.findall(".//w:tbl", NS)), start=1)
    )

    bookmarks = tuple(
        el.get(W + "name", "")
        for el in root.findall(".//w:bookmarkStart", NS)
        if el.get(W + "name") and el.get(W + "name") != "_GoBack"
    )

    # Word can retain comments in word/comments.xml whose anchors in the body
    # were deleted. They are invisible in Word's UI and must not block an
    # edit; counting them as present made an unachievable postcondition.
    comment_ids = (
        tuple(c.get(W + "id", "") for c in comments.findall(".//w:comment", NS))
        if comments is not None
        else ()
    )
    anchored = set(re.findall(r'<w:commentReference w:id="(\d+)"', xml_text))
    orphan_comments = tuple(cid for cid in comment_ids if cid not in anchored)

    # Likewise a media part is LIVE only if some relationship pointing at it is
    # actually used by the body.
    media_rels: dict[str, str] = {}
    if rels:
        rel_root = ET.fromstring(rels)
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        for relationship in rel_root.findall(f"{{{rel_ns}}}Relationship"):
            target = relationship.get("Target", "")
            if target.startswith("media/"):
                media_rels[relationship.get("Id", "")] = target
    used_rids = {
        value
        for element in root.iter()
        for attr, value in element.attrib.items()
        if attr.startswith("{" + NS["r"] + "}") and attr.rsplit("}", 1)[-1] in {"embed", "id", "link"}
    }
    dangling_media = tuple(
        sorted(target for rid, target in media_rels.items() if rid not in used_rids)
    )
    media_total = sum(
        1 for name in names if name.startswith("word/media/") and not name.endswith("/")
    )
    locked, lock_reason = is_file_locked(resolved)
    pref_count = len(re.findall(r"ZOTERO_PREF_\d+", custom))
    readiness_ready = bool(pref_count > 0 and not structural_errors)
    readiness_status = (
        "structurally_invalid" if structural_errors else
        "initialized" if pref_count > 0 else "virgin_document"
    )
    readiness = {
        "ready": readiness_ready,
        "status": readiness_status,
        "structural_errors": list(structural_errors),
        "message": (
            "Zotero preferences are present and offline structural checks passed; "
            "Word readability and successful Zotero refresh are not verified by this audit."
            if readiness_ready
            else (
                "Document has structural OOXML errors; repair the package before citation operations."
                if structural_errors
            else "Document lacks Zotero preferences. Open in Word, insert one citation manually, save, and close Word. Then rerun audit_document, confirm citation_readiness.ready is true, and retry with the fresh sha256 before inserting citations programmatically."
            )
        ),
    }

    return Audit(
        path=str(resolved),
        sha256=file_sha256(resolved),
        body_text_sha256=hashlib.sha256(text.encode("utf-8")).hexdigest().upper(),
        prose_text_sha256=hashlib.sha256(prose.encode("utf-8")).hexdigest().upper(),
        paragraph_count=len(root.findall(".//w:p", NS)),
        paragraph_count_basis="raw OOXML paragraph elements (w:p); not Word COM paragraph numbering",
        table_count=len(tables),
        tables=tables,
        bookmark_names=bookmarks,
        bookmark_references=bookmark_references(root),
        field_count=xml_text.count('w:fldCharType="begin"'),
        zotero_item_fields=xml_text.count("ZOTERO_ITEM"),
        zotero_bibliography_fields=xml_text.count("ZOTERO_BIBL"),
        zotero_pref_properties=pref_count,
        revision_count=sum(
            len(root.findall(f".//w:{name}", NS))
            for name in ("ins", "del", "moveFrom", "moveTo")
        ),
        comment_count=len(comment_ids),
        live_comment_count=len(comment_ids) - len(orphan_comments),
        orphan_comment_ids=orphan_comments,
        drawing_count=len(root.findall(".//w:drawing", NS)),
        media_count=media_total,
        live_media_count=media_total - len(dangling_media),
        dangling_media_targets=dangling_media,
        structural_errors=tuple(structural_errors),
        is_locked=locked,
        lock_reason=lock_reason,
        citation_readiness=readiness,
    )
