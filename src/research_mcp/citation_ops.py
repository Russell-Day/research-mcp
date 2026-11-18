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
    tracked: bool = False,
    author: str = "Claude",
    now: Any = None,
) -> dict[str, Any]:
    """Add one citation cluster and optionally one bibliography to a new copy or in-place.

    `anchor` guarantees where the citation lands. When `source_sha256` is
    supplied, verifies that the source document has not changed since inspection.

    Plan 864 decision 16: allowed on documents with tracked revisions. The anchor must lie outside every
    revision (refused, named, before Word starts), and every existing revision must survive Word's save and
    the Zotero refresh (`compare_revision_snapshots`), else nothing is written. `tracked=True` then wraps the
    new field's runs in one `w:ins` by `author` (direct XML, `wrap_citation_insertion`), so Reject removes the
    citation and Accept keeps it; the bibliography Zotero updated stays untracked (a warning says so).
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
        revisions_before = revision_snapshot(output_path)
        outside_before = text_outside_zotero_fields(output_path)
        refuse_anchor_in_revision(output_path, int(paragraph_index), anchor)
        with IsolatedWordDocument(output_path) as opened:
            renew_zotero_session(opened.document)
            # Word's own Track Changes (w:trackRevisions) would record the paste under Word's user name; the
            # tracked form is research-mcp's w:ins below. The setting itself is restored before the save.
            word_tracking = _suspend_word_tracking(opened.document)
            try:
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
            finally:
                if word_tracking:
                    opened.document.TrackRevisions = True
            opened.save()
        compare_revision_snapshots(revisions_before, revision_snapshot(output_path), f"{INSERT_OP} (Word's save)")

        rendered, bibliography_results, refresh_warnings = refresh_file_and_wait(
            output_path, (reference.citation_id,)
        )
        revision_check = compare_revision_snapshots(
            revisions_before, revision_snapshot(output_path), f"{INSERT_OP} (after the Zotero refresh)")
        difference = first_text_difference(outside_before, text_outside_zotero_fields(output_path))
        if difference:
            raise VerificationFailed(f"{INSERT_OP}: text outside the Zotero fields changed {difference}")
        # refresh_file_and_wait's third value is the worker's other result fields (timings and the like);
        # only a "warnings" entry there is a warning.
        extra = refresh_warnings.get("warnings") if isinstance(refresh_warnings, dict) else None
        warnings = [str(w) for w in (extra if isinstance(extra, (list, tuple)) else [extra] if extra else [])]
        wrapped = None
        if tracked:
            wrapped = wrap_citation_insertion(output_path, citation_id=reference.citation_id, author=author,
                                              now=now)
            warnings.append(TRACKED_BIBLIOGRAPHY_WARNING)
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
            # refresh_file_and_wait returns every bibliography field's rendered text, in document order.
            "bibliography_rendered": list(bibliography_results) if bibliography_requested else None,
            "source_sha256": source_audit.sha256,
            "output_sha256": final.sha256,
            "tracked": bool(tracked),
            "revision_id": wrapped["revision_id"] if wrapped else None,
            "author": author if tracked else None,
            "revisions_checked": revision_check["revisions_checked"],
            "revision_check": revision_check,
            "warnings": warnings,
        }


INSERT_OP = "insert_zotero_citation"
TRACKED_BIBLIOGRAPHY_WARNING = (
    "the citation is a tracked insertion but the bibliography Zotero updated for it is not tracked: if the "
    "citation is rejected, run refresh_citations after resolving it")
_REVISION_WRAPPERS = ("w:ins", "w:del", "w:moveFrom", "w:moveTo")


def _suspend_word_tracking(document) -> bool:
    """Turn Word's own Track Changes off for the insert; True when it was on (the caller turns it back on)."""

    try:
        was_on = bool(document.TrackRevisions)
    except Exception:  # noqa: BLE001 - a document without the property (fakes) has no Word tracking
        return False
    if was_on:
        document.TrackRevisions = False
    return was_on


def _describe_revision(xml: str, node) -> str:
    return (f"w:{node.local} id={node.attr(xml, 'w:id')} by {node.attr(xml, 'w:author')!r} "
            f"({node.attr(xml, 'w:date')})")


def refuse_anchor_in_revision(path: str | Path, paragraph_index: int, anchor: str, op: str = INSERT_OP) -> None:
    """Refuse (naming the revision) an anchor whose text lies inside or crosses a tracked revision.

    The paragraph's text is read as Word's Range.Text reads it with markup shown (inserted and deleted text,
    no field codes), the text insert_unrendered matches the anchor in. A run carrying a tracked formatting
    change (w:rPrChange) counts as a revision too: inserting before its text could split it. With no single
    match the check is left to Word, which refuses an anchor that is missing or repeated."""

    from .revisions import load, unescape

    if not anchor:
        return
    _, doc = load(path)
    xml = doc.xml
    paragraph = doc.paragraph(int(paragraph_index), op)
    chars: list[tuple[str, Any]] = []

    def owner(node) -> Any:
        wrapper = next((a for a in node.ancestors() if a.name in _REVISION_WRAPPERS), None)
        if wrapper is not None:
            return wrapper
        run = next((a for a in node.ancestors() if a.name == "w:r"), None)
        rpr = run.child("w:rPr") if run is not None else None
        change = next((c for c in rpr.children if c.name == "w:rPrChange"), None) if rpr is not None else None
        return change

    for node in paragraph.iter():
        if node.name in ("w:t", "w:delText"):
            text = "" if node.closed else unescape(node.inner(xml))
        elif node.name == "w:tab":
            text = "\t"
        else:
            continue
        if any(a.name in ("w:pPr", "w:txbxContent") for a in node.ancestors()):
            continue  # tab stops, and text boxes (their own paragraphs)
        rev = owner(node)
        chars.extend((ch, rev) for ch in text)
    text = "".join(ch for ch, _ in chars)
    start = text.find(anchor)
    if start < 0 or text.find(anchor, start + 1) >= 0:
        return
    span = chars[start:start + len(anchor)]
    hit = next((rev for _, rev in span if rev is not None), None)
    if hit is None:
        return
    inside = all(rev is hit for _, rev in span)
    where = "lies inside" if inside else "crosses"
    if hit.name == "w:rPrChange":
        label = (f"a tracked formatting change (w:rPrChange id={hit.attr(xml, 'w:id')} by "
                 f"{hit.attr(xml, 'w:author')!r})")
    else:
        label = f"tracked revision {_describe_revision(xml, hit)}"
    raise RefusedUnsafe(
        f"{op}: the anchor {anchor!r} in paragraph {paragraph_index} {where} {label}; the citation goes "
        "right before the anchor, so pick anchor text outside tracked changes (list_revisions shows them), or "
        "accept/reject that revision first")


def wrap_citation_insertion(path: str | Path, *, citation_id: str, author: str, now: Any = None) -> dict[str, Any]:
    """Wrap the runs of citation `citation_id` (fldChar begin .. end) in one `w:ins` by `author`, in place.

    The mirror of delete_citation's `w:del`. Checked before `path` is replaced: bytes outside the paragraph
    and every other revision unchanged; one new id, unique in the package; in memory, Reject of the new
    insertion gives the paragraph without the field and Accept gives it back unchanged (canonical XML); on the
    written file, the MCP's own accept/reject leave every citation (Accept) or every other citation (Reject)
    with the same field code JSON."""

    import copy
    import shutil
    import tempfile

    from .citation_delete import _field_runs, _identity
    from .cluster_ops import clusters_in, refuse_unsafe_cluster
    from .revisions import (
        _attr_escape, _check_stamps, _fragment, _revision_raw, _verify_written, et_resolve, load, next_revision_id,
        parse, read_parts, resolve_revisions, revision_stamp, rezip, REV_TAGS,
    )
    from .sections import canonical

    path = Path(path)
    author = (author or "").strip() or "Claude"
    parts = read_parts(path)
    _, doc = load(path)
    xml = doc.xml
    clusters = clusters_in(xml, doc.body)
    mine = [c for c in clusters if c.citation_id == citation_id]
    if len(mine) != 1:
        raise VerificationFailed(f"{INSERT_OP}: the new citation {citation_id} is not in the document exactly once",
                                 expected=1, actual=len(mine))
    cluster = mine[0]
    fix = ("insert it with tracked: false (tracking off for this call), or insert it in Word with Track Changes "
           "on; nothing was written")
    refuse_unsafe_cluster(xml, doc.body, cluster, op=INSERT_OP)
    paragraph = next((a for a in cluster.field.begin.ancestors() if a.name == "w:p"), None)
    if paragraph is None:
        raise VerificationFailed(f"{INSERT_OP}: the new citation {citation_id} is not in a paragraph")
    first, last = _field_runs(xml, paragraph, cluster, op=INSERT_OP, fix=fix)
    begin_run, end_run = paragraph.children[first], paragraph.children[last]
    new_id = next_revision_id(parts)
    attrs = f'w:id="{new_id}" w:author="{_attr_escape(author)}" {revision_stamp(xml, now)}'
    old_paragraph = xml[paragraph.start:paragraph.end]
    new_paragraph = (xml[paragraph.start:begin_run.start] + f"<w:ins {attrs}>" + xml[begin_run.start:end_run.end]
                     + "</w:ins>" + xml[end_run.end:paragraph.end])
    new_xml = xml[:paragraph.start] + new_paragraph + xml[paragraph.end:]

    # -- checked before anything is written
    if new_xml[:paragraph.start] != xml[:paragraph.start] or \
            new_xml[paragraph.start + len(new_paragraph):] != xml[paragraph.end:]:
        raise VerificationFailed(f"{INSERT_OP}: bytes outside the paragraph changed")
    old_tree, new_tree = parse(old_paragraph), parse(new_paragraph)
    old_revs = [_revision_raw(old_paragraph, n) for n in old_tree.iter() if n.local in REV_TAGS and n.name.startswith("w:")]
    kept, fresh = [], []
    for n in new_tree.iter():
        if n.local in REV_TAGS and n.name.startswith("w:"):
            if n.local == "ins" and n.attr(new_paragraph, "w:id") == str(new_id):
                fresh.append(n)
            else:
                kept.append(_revision_raw(new_paragraph, n))
    if kept != old_revs:
        raise VerificationFailed(f"{INSERT_OP}: a pre-existing revision in the paragraph changed")
    if len(fresh) != 1:
        raise VerificationFailed(f"{INSERT_OP}: expected exactly one new insertion, found {len(fresh)}")
    _check_stamps(new_paragraph, fresh, author, INSERT_OP)
    is_new = lambda el: el.get(W + "id") == str(new_id)  # noqa: E731
    new_p = _fragment(xml, new_paragraph)
    accepted, rejected = copy.deepcopy(new_p), copy.deepcopy(new_p)
    et_resolve(accepted, is_new, "accept")
    et_resolve(rejected, is_new, "reject")
    if canonical(accepted) != canonical(_fragment(xml, old_paragraph)):
        raise VerificationFailed(f"{INSERT_OP}: accepting the insertion does not give the paragraph with the citation")
    without = xml[paragraph.start:begin_run.start] + xml[end_run.end:paragraph.end]
    if canonical(rejected) != canonical(_fragment(xml, without)):
        raise VerificationFailed(f"{INSERT_OP}: rejecting the insertion does not give the paragraph without it")

    revisions_before = revision_snapshot(path)
    want_all = [_identity(item) for item in clusters]
    want_others = [item for item in want_all if item != _identity(cluster)]
    with tempfile.TemporaryDirectory(prefix="research-mcp-inscite-") as temp:
        staged = Path(temp) / "wrapped.docx"
        rezip(path, staged, new_xml)
        _verify_written(path, staged, new_xml, new_id, INSERT_OP)
        after = [r for r in revision_snapshot(staged)
                 if not (r.part == "word/document.xml" and r.type == "ins" and r.id == str(new_id))]
        compare_revision_snapshots(revisions_before, after, f"{INSERT_OP} (tracked insertion)")
        accept_p, reject_p = Path(temp) / "accept.docx", Path(temp) / "reject.docx"
        try:
            resolve_revisions(staged, accept_p, mode="accept", ids=[str(new_id)])
            resolve_revisions(staged, reject_p, mode="reject", ids=[str(new_id)])
        except (RefusedUnsafe, VerificationFailed) as exc:
            raise VerificationFailed(f"{INSERT_OP}: the new insertion cannot be resolved cleanly: {exc}") from exc
        for resolved, expected, what in ((accept_p, want_all, "accepting"), (reject_p, want_others, "rejecting")):
            _, after_doc = load(resolved)
            got = [_identity(item) for item in clusters_in(after_doc.xml, after_doc.body)]
            if got != expected:
                raise VerificationFailed(f"{INSERT_OP}: {what} the insertion does not leave the expected citations")
        shutil.copyfile(staged, path)
    index = next((number for number, node in doc.para_nodes.items() if node is paragraph), None)
    return {"revision_id": str(new_id), "author": author, "paragraph_index": index}


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


_DATE_ATTRS = (W + "date", "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc")


def _minute(value: str) -> str:
    """A revision date as Word writes it on save: seconds dropped ("2026-11-15T21:49:42Z" -> "...21:49:00Z")."""

    import re

    match = re.fullmatch(r"(\d{4}-\d\d-\d\dT\d\d:\d\d)(:\d\d(?:\.\d+)?)?(Z?)", value or "")
    return f"{match.group(1)}:00{match.group(3)}" if match else value


_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _same_after_word_save(old: tuple, new: tuple) -> bool:
    """`new` is `old` as Word's save writes it (seen live on MCP-made revisions, plan 864 S5): a w:id may change
    value (Word renumbers revision, bookmark and comment ids; `compare_revision_snapshots` then requires every id
    to map one to one) but never appear or disappear, a date attribute may lose its seconds (Word cuts w:date to
    the minute, and keeps w16du:dateUtc), and xml:space="preserve" may be dropped from text with no leading or
    trailing whitespace. Everything else (tags, text, author, formatting, order) must be equal."""

    old_tag, old_attrs, old_text, old_children, old_tail = old
    new_tag, new_attrs, new_text, new_children, new_tail = new
    if old_tag != new_tag or old_text != new_text or old_tail != new_tail or len(old_children) != len(new_children):
        return False
    spaced = old_text != old_text.strip()
    renumbered = _id_space(old_tag) in _RENUMBERED_SPACES

    def kept(attrs):
        return {k: ("" if k == W + "id" and renumbered else v) for k, v in attrs
                if not (k == _XML_SPACE and not spaced)}

    a, b = kept(old_attrs), kept(new_attrs)
    if a.keys() != b.keys():
        return False
    if any(b[k] != v and not (k in _DATE_ATTRS and b[k] == _minute(v)) for k, v in a.items()):
        return False
    return all(_same_after_word_save(x, y) for x, y in zip(old_children, new_children))


# The id spaces Word renumbers on save. Any other w:id (footnoteReference, endnoteReference, ...) names another
# object and must keep its value.
_RENUMBERED_SPACES = ("revision", "bookmark", "comment")
_MARKER_SPACES = ("bookmark", "comment")
_PARTNER_CHANGE = ("a bookmark or comment marker inside it lost or gained a partner marker (start, end or reference) "
                   "in Word's save")


def _id_space(tag: str) -> str:
    """Which id space a w:id belongs to: revisions, bookmarks (start and end share one) or comments (range
    start, end and reference share one); anything else by its own element name."""

    local = tag.rsplit("}", 1)[-1]
    if local in _REV_LOCALS:
        return "revision"
    if local in ("bookmarkStart", "bookmarkEnd"):
        return "bookmark"
    if local in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        return "comment"
    return local


def _ids_in(signature: tuple) -> list[tuple[str, str]]:
    """Every renumberable (id space, w:id) in a revision's signature, nested markers included, in document order."""

    tag, attrs, _, children, _ = signature
    own = [(_id_space(tag), v) for k, v in attrs if k == W + "id" and _id_space(tag) in _RENUMBERED_SPACES]
    return own + [i for child in children for i in _ids_in(child)]


def _word_save_only(old: "RevisionRecord", new: "RevisionRecord") -> bool:
    """True when `new` is `old` as Word's save writes it (`_same_after_word_save`)."""

    return old.part == new.part and old.type == new.type and _same_after_word_save(old.signature, new.signature)


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
    # For each bookmark or comment id nested in this revision: every marker of the part with that id (inside or
    # outside revisions), as sorted element names. A renumbered nested marker must take its partners with it.
    partners: tuple = ()
    # True when this w:ins/w:del directly follows its previous sibling, a revision Word's save merges it into
    # (`_merges_with_previous`): the two form one touching group.
    merges_with_prev: bool = False


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


# Word's save merges touching insertions or deletions (plan 864 S4 fix, live): adjacent sibling w:ins (or w:del)
# elements of one author whose dates are equal to the minute become one revision, its runs merged where their
# formatting is equal. Only such groups with nothing between them (a proofing marker aside) are tolerated: Word
# also merges across a bookmark or comment marker or an empty run, which still refuses.
_MERGED_LOCALS = ("ins", "del")
_TEXT_TAGS = (W + "t", W + "delText", W + "instrText", W + "delInstrText")


def _merge_key(el: ET.Element) -> dict:
    """A revision's attributes as Word's merge compares them: no id or rsids, dates to the minute."""

    return {k: (_minute(v) if k in _DATE_ATTRS else v) for k, v in el.attrib.items()
            if k != W + "id" and not _ignored_attr(k)}


def _merges_with_previous(el: ET.Element, parent: ET.Element | None, wanted: set[str]) -> bool:
    """True when `el` (a w:ins or w:del holding no other revision) follows a sibling of the same type, author and
    date to the minute (every other attribute equal too) that holds no other revision, with nothing between them
    but proofing markers: no text either, not even whitespace."""

    if parent is None or el.tag.rsplit("}", 1)[-1] not in _MERGED_LOCALS or el.get(W + "author") is None:
        return False
    if any(node is not el and node.tag in wanted for node in el.iter()):
        return False
    siblings = list(parent)
    index = siblings.index(el) - 1
    while index >= 0 and siblings[index].tag == W + "proofErr":
        if siblings[index].tail:
            return False
        index -= 1
    if index < 0:
        return False
    prev = siblings[index]
    if prev.tail or prev.tag != el.tag or _merge_key(prev) != _merge_key(el):
        return False
    return not any(node is not prev and node.tag in wanted for node in prev.iter())


def revision_snapshot(path: str | Path) -> list[RevisionRecord]:
    """Every revision element of every word/*.xml part, in document order."""

    wanted = {W + name for name in _REV_LOCALS}
    records: list[RevisionRecord] = []
    with zipfile.ZipFile(Path(path)) as archive:
        # A comment marker's partners include its definition in comments.xml (author, initials, text): renumbering
        # the markers without the definition, or the other way round, detaches the comment.
        definitions: dict[str, str] = {}
        if "word/comments.xml" in archive.namelist():
            try:
                for comment in ET.fromstring(archive.read("word/comments.xml")).iter(W + "comment"):
                    body = "".join(node.text or "" for node in comment.iter(W + "t"))
                    definitions[comment.get(W + "id")] = (
                        f"definition {comment.get(W + 'author')!r} {comment.get(W + 'initials')!r} {body!r}")
            except ET.ParseError as exc:
                raise RefusedUnsafe(f"cannot scan word/comments.xml for comment definitions: {exc}") from exc
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
            markers: dict[tuple[str, str], list[str]] = {}
            for node in root.iter():
                if _id_space(node.tag) in _MARKER_SPACES and node.get(W + "id") is not None:
                    markers.setdefault((_id_space(node.tag), node.get(W + "id")), []).append(
                        node.tag.rsplit("}", 1)[-1])
            parents = {child: node for node in root.iter() for child in node}
            for el, raw in zip(elements, raws):
                nested = [(_id_space(node.tag), node.get(W + "id")) for node in el.iter()
                          if node is not el and _id_space(node.tag) in _MARKER_SPACES and node.get(W + "id") is not None]
                partners = tuple(
                    (key, tuple(sorted(markers.get(key, ())))
                     + ((definitions.get(key[1], "no definition"),) if key[0] == "comment" else ()))
                    for key in nested)
                text = "".join(
                    "\t" if child.tag == W + "tab" else (child.text or "")
                    for child in el.iter()
                    if child.tag in (W + "t", W + "delText", W + "instrText", W + "delInstrText", W + "tab"))
                records.append(RevisionRecord(
                    part=part, type=el.tag.rsplit("}", 1)[-1], id=el.get(W + "id"), author=el.get(W + "author"),
                    date=el.get(W + "date"), text=text, raw=raw,
                    signature=_canonical(el, strict=False), strict_signature=_canonical(el, strict=True),
                    partners=partners, merges_with_prev=_merges_with_previous(el, parents.get(el), wanted)))
    return records


def _atoms(signature: tuple) -> list[tuple]:
    """A w:ins/w:del's content as Word's merge keeps it: one item per character (run attributes, run properties,
    text element and its attributes but xml:space, character), one per run with no content, and one per other run
    child or other element (an empty text element counts as one too). Runs
    may be split or joined; the formatting of every character may not change. xml:space is checked on the merged
    revision's own text elements (`_space_kept`)."""

    atoms: list[tuple] = []
    for child in signature[3]:
        tag, attrs, _, children, _ = child
        if tag != W + "r":
            atoms.append(("node", child))
            continue
        rpr = next((c for c in children if c[0] == W + "rPr"), None)
        if all(c[0] == W + "rPr" for c in children):
            # A run with no content (formatting only) is kept as one item: dropping it in a merge refuses, as the
            # one-to-one comparison refuses it.
            atoms.append(("empty run", attrs, rpr))
            continue
        for c in children:
            if c[0] == W + "rPr":
                continue
            if c[0] in _TEXT_TAGS:
                text_attrs = tuple((k, v) for k, v in c[1] if k != _XML_SPACE)
                atoms.extend(("char", attrs, rpr, c[0], text_attrs, ch) for ch in c[2])
                if not c[2]:
                    # an empty text element is kept as one item too (its run's formatting and its attributes)
                    atoms.append(("empty text", attrs, rpr, c[0], text_attrs))
            else:
                atoms.append(("run child", attrs, rpr, c))
    return atoms


def _space_kept(signature: tuple) -> bool:
    """Every text element of a merged revision whose text starts or ends with whitespace keeps it
    (xml:space="preserve"), the rule `_same_after_word_save` applies: Word drops xml:space only from text with no
    outer whitespace."""

    return all(dict(c[1]).get(_XML_SPACE) == "preserve"
               for child in signature[3] if child[0] == W + "r"
               for c in child[3] if c[0] in _TEXT_TAGS and c[2] != c[2].strip())


def _same_atom(old: tuple, new: tuple) -> bool:
    if old[0] != new[0]:
        return False
    if old[0] in ("char", "empty run", "empty text"):
        return old == new
    if old[0] == "node":
        return _same_after_word_save(old[1], new[1])
    return old[1:3] == new[1:3] and _same_after_word_save(old[3], new[3])


def _merge_group(before: list[RevisionRecord], i: int) -> int:
    """How many records form the touching group that starts at `before[i]` (1 when none does)."""

    if before[i].merges_with_prev:
        return 1
    k = 1
    while i + k < len(before) and before[i + k].merges_with_prev:
        k += 1
    return k


def compare_revision_snapshots(before: list[RevisionRecord], after: list[RevisionRecord],
                               what: str) -> dict[str, Any]:
    """Every revision must survive: byte-identical, or differ only in serialization.

    Serialization means attribute order, namespace prefixes, whitespace between elements, rsid and
    paraId/textId attributes, and proofing / last-rendered-page-break markers inside the revision
    (bookmark markers are compared). It also covers what Word's save writes (`_same_after_word_save`): ids
    renumbered (revision, bookmark and comment ids only; any other w:id, e.g. a footnoteReference, must keep its
    value), provided every w:id is still there, each (part, id space, old id) keeps one new id across all
    revisions and the markers nested in them, with no two old ids sharing one, and every nested bookmark or comment
    marker keeps the same partner markers in the part under its new id; `w:date` cut to the minute;
    `xml:space` dropped from text with no outer whitespace; and touching revisions merged: a group of adjacent
    sibling w:ins (or w:del) of one author and minute with nothing between them (`merges_with_prev`) may become
    one revision of that type, author and minute holding the group's content in order with the same formatting on
    every character (runs split or joined), its one id standing for every id of the group. Type, author, text and
    formatting (every other element and attribute) must be equal; otherwise VerificationFailed names the first
    changed revisions."""

    def count_error() -> VerificationFailed:
        def counts(records):
            out: dict[str, int] = {}
            for record in records:
                out[f"{record.part}:{record.type}"] = out.get(f"{record.part}:{record.type}", 0) + 1
            return out
        merged_note = (f" ({word_merged} touching group(s) of {merged_revisions} revisions merged by Word's save "
                       "accounted for)" if word_merged else "")
        return VerificationFailed(
            f"{what}: tracked revision count changed from {len(before)} to {len(after)}{merged_note}",
            expected=counts(before), actual=counts(after))

    changed: list[str] = []
    byte_identical = 0
    attribute_order_only = 0
    ignored_only = 0
    word_saved = 0
    word_merged = 0
    merged_revisions = 0
    # (part, id space, old id) -> new id, over every revision and every marker nested in one. Word's save may
    # renumber ids, but each old id must keep one new id everywhere it appears (a moveFrom/moveTo range start
    # and end, a bookmark's start and end, a kept revision next to a renumbered one) and no two old ids may
    # share a new one. The ids of a merged group's later members map to the merged revision's id in `absorbed`.
    id_map: dict[tuple[str, str, str], str] = {}
    absorbed: dict[tuple[str, str, str], str] = {}

    def consistent(part: str, olds: list, news: list) -> bool:
        if [space for space, _ in olds] != [space for space, _ in news]:
            return False
        return all(id_map.setdefault((part, space, a), b) == b for (space, a), (_, b) in zip(olds, news))

    def partners_kept(old: RevisionRecord, new: RevisionRecord) -> bool:
        # Each bookmark or comment marker nested in the revision still has the same partner markers in the part
        # under its (possibly new) id: a renumbered start whose end kept the old id, or an outside marker that
        # took a nested marker's id, changes them.
        return [kinds for _, kinds in old.partners] == [kinds for _, kinds in new.partners]

    def merged_into(group: list[RevisionRecord], new: RevisionRecord) -> bool:
        """`new` is the touching `group` merged by Word's save (checked in full; ids recorded when it is)."""

        first = group[0]
        if new.part != first.part or new.type != first.type or new.author != first.author:
            return False
        if not _same_after_word_save(first.signature[:3] + ((),) + first.signature[4:],
                                     new.signature[:3] + ((),) + new.signature[4:]):
            return False
        old_atoms = [_atoms(record.signature) for record in group]
        if not all(old_atoms):
            return False  # a member with no content: not a merge Word's save is known to make
        flat = [atom for atoms in old_atoms for atom in atoms]
        new_atoms = _atoms(new.signature)
        if len(flat) != len(new_atoms) or not all(_same_atom(a, b) for a, b in zip(flat, new_atoms)):
            return False
        if not _space_kept(new.signature):
            return False
        if [kinds for record in group for _, kinds in record.partners] != [kinds for _, kinds in new.partners]:
            return False
        olds = [i for record in group for i in _ids_in(record.signature)]
        own = [("revision", record.id) for record in group]
        nested = [i for i in olds if i not in own[1:]]
        news = _ids_in(new.signature)
        if not news or news[0][0] != "revision" or nested[:1] != own[:1]:
            return False
        # An old id absorbed by two groups (two revisions sharing one id) must go to one merged revision.
        clash = [old_id for _, old_id in own[1:]
                 if absorbed.get((first.part, "revision", old_id), news[0][1]) != news[0][1]]
        if clash:
            changed.append(f"{first.part} w:{first.type} id={clash[0]} by {first.author!r} was merged by Word's save "
                           "into two different revisions")
            return True  # consumed; the error above refuses the comparison
        if not consistent(first.part, nested, news):
            return False
        for _, old_id in own[1:]:
            absorbed[(first.part, "revision", old_id)] = news[0][1]
        return True

    i = j = 0
    while i < len(before) and j < len(after):
        old, new = before[i], after[j]
        k = _merge_group(before, i)
        if k > 1 and merged_into(before[i:i + k], new):
            word_merged += 1
            merged_revisions += k
            i, j = i + k, j + 1
            continue
        i, j = i + 1, j + 1
        if old.part == new.part and old.signature != new.signature and _word_save_only(old, new):
            # Word's save (plan 864 S5, live): ids renumbered, seconds dropped from w:date.
            if not partners_kept(old, new):
                changed.append(f"{old.part} w:{old.type} id={old.id} by {old.author!r}: {_PARTNER_CHANGE}")
                continue
            if consistent(old.part, _ids_in(old.signature), _ids_in(new.signature)):
                word_saved += 1
                continue
        if old.part != new.part or old.signature != new.signature:
            changed.append(
                f"{old.part} w:{old.type} id={old.id} by {old.author!r} ({old.date}) {old.text[:60]!r} -> "
                f"{new.part} w:{new.type} id={new.id} by {new.author!r} ({new.date}) {new.text[:60]!r}")
            continue
        ids = _ids_in(old.signature)
        if not partners_kept(old, new):
            changed.append(f"{old.part} w:{old.type} id={old.id} by {old.author!r}: {_PARTNER_CHANGE}")
            continue
        if not consistent(old.part, ids, ids):
            changed.append(f"{old.part} w:{old.type} id={old.id} by {old.author!r} keeps an id that Word's save "
                           "renumbered in another revision")
            continue
        if old.raw and old.raw == new.raw:
            byte_identical += 1
        elif old.strict_signature == new.strict_signature:
            attribute_order_only += 1
        else:
            ignored_only += 1
    if i < len(before) or j < len(after):
        raise count_error()
    if not changed and any(key in id_map for key in absorbed):
        changed.append("a revision merged by Word's save shares its id with another revision")
    if not changed and len({(part, space, b) for (part, space, _), b in id_map.items()}) != len(id_map):
        changed.append("ids no longer map one to one (two revisions or markers share an id after the save)")
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
        "word_renumbered_or_dates_to_minute": word_saved,
        "word_merged": word_merged,
        "word_merged_revisions": merged_revisions,
    }
