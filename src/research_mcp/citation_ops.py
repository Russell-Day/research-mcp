"""Citation editing operations wrapping isolated Word sessions and Zotero."""

from __future__ import annotations

from dataclasses import dataclass
import io
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Sequence

from .audit import audit_docx
from .constants import WD_FIELD_ADDIN
from .citations import (
    insert_bibliography_unrendered,
    insert_unrendered,
    refresh_file_and_wait,
    renew_zotero_session,
)
from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import W
from .word import IsolatedWordDocument, staged_document
from .zotero import ZoteroClient

DEFAULT_ITEM_DONOR = Path(__file__).resolve().parent / "data" / "zotero_item_donor.docx"
if not DEFAULT_ITEM_DONOR.exists():
    _test_donor = Path(__file__).resolve().parent.parent.parent / "tests" / "data" / "zotero_item_donor.docx"
    if _test_donor.exists():
        DEFAULT_ITEM_DONOR = _test_donor


def insert_citation_to_version(
    source: str | Path,
    output: str | Path | None = None,
    *,
    paragraph_index: int,
    anchor: str,
    item_keys: Sequence[str],
    item_donor: str | Path | None = None,
    bibliography_donor: str | Path | None = None,
    bibliography_paragraph_index: int | None = None,
    source_sha256: str | None = None,
) -> dict[str, Any]:
    """Add one citation cluster and optionally one bibliography to a new copy or in-place.

    `anchor` guarantees where the citation lands. When `source_sha256` is
    supplied, verifies that the source document has not changed since inspection.
    """

    if not item_keys:
        raise RefusedUnsafe("item_keys requires at least one item key")

    target_output = output if output is not None else source
    out_p = Path(target_output).resolve()
    src_p = Path(source).resolve()
    if out_p != src_p and out_p.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {target_output}")

    bibliography_requested = (
        bibliography_donor is not None or bibliography_paragraph_index is not None
    )
    if bibliography_requested and (
        bibliography_donor is None or bibliography_paragraph_index is None
    ):
        raise RefusedUnsafe(
            "bibliography_donor and bibliography_paragraph_index must be supplied together"
        )

    source_audit = audit_docx(source)
    if source_sha256 is not None and source_sha256.upper() != source_audit.sha256.upper():
        raise RefusedUnsafe(
            f"source document SHA-256 mismatch: expected {source_sha256.upper()}, "
            f"actual {source_audit.sha256.upper()}. Reinspect the document with inspect_document."
        )

    donor_path = Path(item_donor) if item_donor is not None else DEFAULT_ITEM_DONOR
    if not donor_path.exists():
        raise RefusedUnsafe(f"item donor document not found: {donor_path}")

    with staged_document(source, target_output) as output_path:
        if source_sha256 is not None and audit_docx(output_path).sha256.upper() != source_sha256.upper():
            raise RefusedUnsafe(
                f"source changed after its SHA-256 was checked (expected {source_sha256.upper()}); "
                "reinspect the document with inspect_document")
        with IsolatedWordDocument(output_path) as opened:
            renew_zotero_session(opened.document)
            reference = insert_unrendered(
                opened.document,
                donor_path=donor_path,
                paragraph_index=int(paragraph_index),
                anchor=anchor,
                item_keys=tuple(item_keys),
                zotero=ZoteroClient(),
            )
            if bibliography_requested:
                insert_bibliography_unrendered(
                    opened.document,
                    donor_path=bibliography_donor,
                    paragraph_index=int(bibliography_paragraph_index),
                )
            opened.save()

        rendered, bibliography_results, refresh_warnings = refresh_file_and_wait(
            output_path, (reference.citation_id,)
        )
        final = audit_docx(output_path)
        expected_items = source_audit.zotero_item_fields + 1
        expected_bibliographies = source_audit.zotero_bibliography_fields + int(
            bibliography_requested
        )
        if final.zotero_item_fields != expected_items:
            raise VerificationFailed(
                "versioned citation count mismatch",
                expected=expected_items,
                actual=final.zotero_item_fields,
            )
        if final.zotero_bibliography_fields != expected_bibliographies:
            raise VerificationFailed(
                "versioned bibliography count mismatch",
                expected=expected_bibliographies,
                actual=final.zotero_bibliography_fields,
            )

        verify_protected_invariants(source_audit, final, "citation insertion")

        return {
            "success": True,
            "source": str(Path(source).resolve()),
            "output": str(Path(output_path).resolve()),
            "citation_id": reference.citation_id,
            "paragraph_index": int(paragraph_index),
            "rendered_citation": rendered.get(reference.citation_id),
            "bibliography_rendered": (
                bibliography_results.get(reference.citation_id)
                if bibliography_requested
                else None
            ),
            "source_sha256": source_audit.sha256,
            "output_sha256": final.sha256,
            "warnings": list(refresh_warnings),
        }


# --------------------------------------------------------------------------
# Shared verification helpers (plan 602: edit_citation_cluster, refresh_citations).
# --------------------------------------------------------------------------

PROTECTED_INVARIANTS = (
    "table_count",
    "live_comment_count",
    "drawing_count",
    "media_count",
    "live_media_count",
)


def verify_protected_invariants(before, after, what: str) -> None:
    """Tables, comments, drawings and media must not change during a citation write."""

    for name in PROTECTED_INVARIANTS:
        if getattr(before, name) != getattr(after, name):
            raise VerificationFailed(
                f"protected invariant {name} changed during {what}",
                expected=getattr(before, name),
                actual=getattr(after, name),
            )


_ZOTERO_PREFIXES = ("ADDIN ZOTERO_ITEM", "ADDIN ZOTERO_BIBL")


def _walk_body_fields(path: str | Path):
    """Yield (event, value) over the body in document order: ("p", None) at each paragraph start,
    ("text", (chars, deleted)) for w:t/w:delText/w:tab, with `zotero` telling whether a Zotero
    field result (citation or bibliography) is open. Field begin/separate/end are tracked with a
    stack, so a result spanning runs and paragraphs (the bibliography) is handled."""

    with zipfile.ZipFile(Path(path)) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find(W + "body")
    stack: list[dict] = []
    deleted_depth: dict[int, bool] = {}

    def visit(node: ET.Element, deleted: bool):
        tag = node.tag
        if tag == W + "p":
            yield "p", None, any(f["zotero"] and f["separated"] for f in stack)
        elif tag == W + "fldChar":
            kind = node.get(W + "fldCharType", "")
            if kind == "begin":
                stack.append({"instr": [], "separated": False, "zotero": False})
            elif kind == "separate" and stack:
                field = stack[-1]
                field["separated"] = True
                field["zotero"] = "".join(field["instr"]).strip().startswith(_ZOTERO_PREFIXES)
            elif kind == "end" and stack:
                stack.pop()
        elif tag in (W + "instrText", W + "delInstrText"):
            if stack and not stack[-1]["separated"]:
                stack[-1]["instr"].append(node.text or "")
        elif tag in (W + "t", W + "delText", W + "tab"):
            text = "\t" if tag == W + "tab" else (node.text or "")
            yield "text", (text, deleted or tag == W + "delText"), any(
                f["zotero"] and f["separated"] for f in stack)
        child_deleted = deleted or tag in (W + "del", W + "moveFrom")
        for child in node:
            yield from visit(child, child_deleted)

    if body is not None:
        yield from visit(body, False)


def text_outside_zotero_fields(path: str | Path) -> str:
    """The body's text with every Zotero citation and bibliography result left out.

    Paragraph breaks inside a Zotero result are left out too, so a bibliography that gains or loses
    entries does not change it. Deleted (tracked) text is kept, marked `[-...-]`, so a refresh that
    rewrote tracked text shows up here."""

    chars: list[tuple[str, bool]] = []
    for event, value, in_zotero in _walk_body_fields(path):
        if in_zotero:
            continue
        if event == "p":
            chars.append(("\n", False))
        else:
            text, deleted = value
            chars.extend((ch, deleted) for ch in text)
    out: list[str] = []
    open_del = False
    for ch, deleted in chars:
        if deleted != open_del:
            out.append("[-" if deleted else "-]")
            open_del = deleted
        out.append(ch)
    if open_del:
        out.append("-]")
    return "".join(out)


def first_text_difference(before: str, after: str, context: int = 40) -> str | None:
    if before == after:
        return None
    index = next((i for i, (a, b) in enumerate(zip(before, after)) if a != b), min(len(before), len(after)))
    lo = max(0, index - context)
    return (f"at character {index}: before {before[lo:index + context]!r}, "
            f"after {after[lo:index + context]!r}")


def bibliography_entries(path: str | Path) -> list[str]:
    """The Zotero bibliography's entries as text, one per result paragraph, in order (every
    ZOTERO_BIBL field; empty when the document has none). Shared with refresh_citations' diff."""

    from .ooxml import zotero_package_fields

    entries: list[str] = []
    for result in zotero_package_fields(path).bibliography_results:
        entries.extend(line.strip() for line in result.split("\n") if line.strip())
    return entries


# Revision snapshot: every revision element in every word/*.xml part, as raw bytes and as a
# semantic signature (type, id, author, date, text, formatting) that ignores serialization.

_REV_LOCALS = ("ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd",
               "moveToRangeStart", "moveToRangeEnd", "rPrChange", "pPrChange", "tblPrChange", "trPrChange",
               "tcPrChange", "sectPrChange", "tblGridChange", "tblPrExChange", "numberingChange",
               "cellIns", "cellDel", "cellMerge")
# Bookmark markers inside a revision are compared (id, name and their other attributes): losing one
# loses a cross-reference target. Only proofing and rendering markers are ignored.
_IGNORED_CHILDREN = {W + "proofErr", W + "lastRenderedPageBreak"}


def _ignored_attr(name: str) -> bool:
    local = name.rsplit("}", 1)[-1]
    return local.lower().startswith("rsid") or local in ("paraId", "textId")


def _canonical(node: ET.Element, *, strict: bool) -> tuple:
    attrs = tuple(sorted((k, v) for k, v in node.attrib.items() if strict or not _ignored_attr(k)))
    children = tuple(_canonical(child, strict=strict) for child in node
                     if strict or child.tag not in _IGNORED_CHILDREN)
    text = node.text or ""
    if node.tag not in (W + "t", W + "delText", W + "instrText", W + "delInstrText"):
        text = text.strip()
    return (node.tag, attrs, text, children, (node.tail or "").strip() if strict else "")


@dataclass(frozen=True)
class RevisionRecord:
    part: str
    type: str
    id: str | None
    author: str | None
    date: str | None
    text: str
    raw: str
    signature: tuple
    strict_signature: tuple


def _raw_revisions(xml: str) -> list[str]:
    """Raw bytes of each revision element, in document order (prefixes bound to WordprocessingML on the
    part's root element; Word always declares them there)."""

    import re

    from .revisions import attr_value, parse

    tree = parse(xml)
    if not tree.children:
        return []
    head = tree.children[0].head(xml)
    prefixes = {p for p in re.findall(r"xmlns:([\w.\-]+)\s*=", head) if attr_value(head, f"xmlns:{p}") == W[1:-1]}
    return [xml[node.start:node.end] for node in tree.iter()
            if ":" in node.name and node.local in _REV_LOCALS and node.name.split(":", 1)[0] in prefixes]


def revision_snapshot(path: str | Path) -> list[RevisionRecord]:
    """Every revision element of every word/*.xml part, in document order."""

    wanted = {W + name for name in _REV_LOCALS}
    records: list[RevisionRecord] = []
    with zipfile.ZipFile(Path(path)) as archive:
        for part in sorted(archive.namelist()):
            if not (part.startswith("word/") and part.endswith(".xml")):
                continue
            data = archive.read(part)
            try:
                root = ET.fromstring(data)
            except ET.ParseError as exc:
                raise RefusedUnsafe(f"cannot scan {part} for tracked revisions: {exc}") from exc
            elements = [el for el in root.iter() if el.tag in wanted]
            if not elements:
                continue
            raws = _raw_revisions(data.decode("utf-8"))
            if len(raws) != len(elements):
                raws = [""] * len(elements)  # unusual prefixes: compare at the semantic level only
            for el, raw in zip(elements, raws):
                text = "".join(
                    "\t" if child.tag == W + "tab" else (child.text or "")
                    for child in el.iter()
                    if child.tag in (W + "t", W + "delText", W + "instrText", W + "delInstrText", W + "tab"))
                records.append(RevisionRecord(
                    part=part, type=el.tag.rsplit("}", 1)[-1], id=el.get(W + "id"), author=el.get(W + "author"),
                    date=el.get(W + "date"), text=text, raw=raw,
                    signature=_canonical(el, strict=False), strict_signature=_canonical(el, strict=True)))
    return records


def compare_revision_snapshots(before: list[RevisionRecord], after: list[RevisionRecord],
                               what: str) -> dict[str, Any]:
    """Every revision must survive: byte-identical, or differ only in serialization.

    Serialization means attribute order, namespace prefixes, whitespace between elements, rsid and
    paraId/textId attributes, and proofing / last-rendered-page-break markers inside the revision
    (bookmark markers are compared). Type, id, author, date, text and formatting (every other element and attribute) must be
    equal; otherwise VerificationFailed names the first changed revisions."""

    if len(before) != len(after):
        def counts(records):
            out: dict[str, int] = {}
            for record in records:
                out[f"{record.part}:{record.type}"] = out.get(f"{record.part}:{record.type}", 0) + 1
            return out
        raise VerificationFailed(
            f"{what}: tracked revision count changed from {len(before)} to {len(after)}",
            expected=counts(before), actual=counts(after))
    changed: list[str] = []
    byte_identical = 0
    attribute_order_only = 0
    ignored_only = 0
    for old, new in zip(before, after):
        if old.part != new.part or old.signature != new.signature:
            changed.append(
                f"{old.part} w:{old.type} id={old.id} by {old.author!r} ({old.date}) {old.text[:60]!r} -> "
                f"{new.part} w:{new.type} id={new.id} by {new.author!r} ({new.date}) {new.text[:60]!r}")
            continue
        if old.raw and old.raw == new.raw:
            byte_identical += 1
        elif old.strict_signature == new.strict_signature:
            attribute_order_only += 1
        else:
            ignored_only += 1
    if changed:
        shown = "; ".join(changed[:5]) + (f"; and {len(changed) - 5} more" if len(changed) > 5 else "")
        raise VerificationFailed(f"{what}: {len(changed)} tracked revision(s) changed: {shown}",
                                 expected=len(before), actual=len(changed))
    level = "byte_identical" if byte_identical == len(before) else "serialization_only"
    return {
        "revisions_checked": len(before),
        "level": level,
        "byte_identical": byte_identical,
        "serialization_changed": attribute_order_only,
        "markers_or_rsids_changed": ignored_only,
    }
