"""Change the items of one existing Zotero citation cluster (plan 602 S2).

Direct OOXML: the cluster's field code JSON (`w:instrText`) and its result text are rewritten in
`word/document.xml` by splicing the raw text (every other byte is copied), then Zotero renders the
cluster through the supervised refresh (`citations.refresh_file_and_wait`). Allowed on documents
with tracked changes: every revision element is compared before the XML edit and after Word's save.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
import shutil
import tempfile
from typing import Any, Sequence
from xml.sax.saxutils import escape

from .audit import audit_docx
from .citation_ops import (
    bibliography_entries,
    compare_revision_snapshots,
    first_text_difference,
    revision_snapshot,
    text_outside_zotero_fields,
    verify_protected_invariants,
)
from .citations import PLACEHOLDER, refresh_file_and_wait
from .errors import RefusedUnsafe, VerificationFailed
from .revisions import MOVE_RANGES, REV_TAGS, Node, attr_value, load, parse, unescape
from .word import staged_document
from .zotero import ZoteroClient

OP = "edit_citation_cluster"
_WRAPPERS = ("w:ins", "w:del", "w:moveFrom", "w:moveTo")
_REVISION_LOCALS = set(REV_TAGS) | set(MOVE_RANGES)
_ITEM_PREFIX = "ADDIN ZOTERO_ITEM"


@dataclass
class ZoteroField:
    """One complex field in the body, read from the raw tree."""

    begin: Node
    separate: Node | None = None
    end: Node | None = None
    instr: list[Node] = field(default_factory=list)
    result_t: list[Node] = field(default_factory=list)
    nested: list[str] = field(default_factory=list)
    del_instr: bool = False

    def instruction(self, xml: str) -> str:
        return "".join("" if node.closed else unescape(node.inner(xml)) for node in self.instr)

    def rendered(self, xml: str) -> str:
        return "".join("" if node.closed else unescape(node.inner(xml)) for node in self.result_t)


@dataclass
class Cluster:
    field: ZoteroField
    prefix: str
    payload: dict[str, Any]
    suffix: str
    citation_id: str
    item_keys: list[str]
    rendered: str

    def describe(self) -> str:
        return f"{self.citation_id} {self.rendered!r} items {self.item_keys}"


def item_key(entry: dict[str, Any]) -> str | None:
    """Zotero item key of a citationItems entry: the last segment of its first URI."""

    uris = entry.get("uris") if isinstance(entry, dict) else None
    if not isinstance(uris, list) or not uris:
        uri = entry.get("uri") if isinstance(entry, dict) else None
        uris = uri if isinstance(uri, list) else ([uri] if uri else [])
    if not uris:
        return None
    return str(uris[0]).rstrip("/").rsplit("/", 1)[-1] or None


def body_fields(xml: str, body: Node) -> list[ZoteroField]:
    """Every complex field in `body`, in document order of their ends (begin/separate/end stack)."""

    stack: list[ZoteroField] = []
    done: list[ZoteroField] = []
    for node in body.iter():
        name = node.name
        if name == "w:fldChar":
            kind = node.attr(xml, "w:fldCharType")
            if kind == "begin":
                for parent in stack:
                    parent.nested.append("complex field")
                stack.append(ZoteroField(begin=node))
            elif kind == "separate" and stack:
                stack[-1].separate = node
            elif kind == "end" and stack:
                current = stack.pop()
                current.end = node
                done.append(current)
        elif name == "w:fldSimple":
            for parent in stack:
                parent.nested.append("w:fldSimple " + (node.attr(xml, "w:instr") or "").strip()[:40])
        elif not stack:
            continue
        elif name == "w:instrText" and stack[-1].separate is None:
            stack[-1].instr.append(node)
        elif name == "w:delInstrText":
            stack[-1].del_instr = True
        elif name == "w:t" and stack[-1].separate is not None:
            stack[-1].result_t.append(node)
    return done


def _cluster(xml: str, item: ZoteroField) -> Cluster | None:
    instruction = item.instruction(xml)
    if not instruction.strip().startswith(_ITEM_PREFIX) or item.separate is None or item.end is None:
        return None
    start = instruction.find("{")
    if start < 0:
        return None
    try:
        payload, end = json.JSONDecoder().raw_decode(instruction[start:])
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict):
        return None
    entries = payload.get("citationItems") or []
    return Cluster(
        field=item,
        prefix=instruction[:start],
        payload=payload,
        suffix=instruction[start + end:],
        citation_id=str(payload.get("citationID") or ""),
        item_keys=[item_key(entry) or "?" for entry in entries],
        rendered=item.rendered(xml),
    )


def clusters_in(xml: str, body: Node, span: tuple[int, int] | None = None) -> list[Cluster]:
    found = []
    for item in body_fields(xml, body):
        if span is not None and not span[0] <= item.begin.start < span[1]:
            continue
        cluster = _cluster(xml, item)
        if cluster is not None:
            found.append(cluster)
    return found


def _norm(text: str) -> str:
    return " ".join(text.split())


def select_cluster(clusters: list[Cluster], paragraph_index: int, *, citation_id: str | None,
                   current_text: str | None) -> Cluster:
    if (citation_id is None) == (current_text is None):
        raise RefusedUnsafe(f"{OP}: give exactly one of citation_id or current_text")
    if citation_id is not None:
        matches = [c for c in clusters if c.citation_id == citation_id]
        wanted = f"citation_id {citation_id!r}"
    else:
        matches = [c for c in clusters if _norm(c.rendered) == _norm(current_text or "")]
        wanted = f"current_text {current_text!r}"
    if len(matches) == 1:
        return matches[0]
    listed = "; ".join(c.describe() for c in clusters) or "none"
    if not matches:
        raise RefusedUnsafe(
            f"{OP}: no Zotero citation in paragraph {paragraph_index} matches {wanted}; "
            f"citations in that paragraph: {listed}")
    raise RefusedUnsafe(
        f"{OP}: {len(matches)} Zotero citations in paragraph {paragraph_index} match {wanted}; "
        f"pass citation_id instead. Candidates: {'; '.join(c.describe() for c in matches)}")


def _describe(xml: str, node: Node) -> str:
    return (f"w:{node.local} id={node.attr(xml, 'w:id')} by {node.attr(xml, 'w:author')!r} "
            f"({node.attr(xml, 'w:date')})")


def refuse_unsafe_cluster(xml: str, body: Node, cluster: Cluster, op: str = OP) -> None:
    """The field must sit outside every tracked revision and contain no nested field."""

    item = cluster.field
    if item.nested:
        raise RefusedUnsafe(
            f"{op}: citation {cluster.citation_id} contains a nested field ({item.nested[0]}); edit it in Word")
    if item.del_instr:
        raise RefusedUnsafe(f"{op}: citation {cluster.citation_id} has deleted (tracked) field code text")
    for node in [item.begin, item.separate, item.end, *item.instr, *item.result_t]:
        wrapper = next((a for a in node.ancestors() if a.name in _WRAPPERS), None)
        if wrapper is not None:
            raise RefusedUnsafe(
                f"{op}: citation {cluster.citation_id} sits inside tracked revision {_describe(xml, wrapper)}; "
                "accept or reject it first")
    lo, hi = item.begin.start, item.end.end
    for node in body.iter():
        if node.start >= hi:
            break
        if node.start >= lo and node.name.startswith("w:") and node.local in _REVISION_LOCALS:
            raise RefusedUnsafe(
                f"{op}: tracked revision {_describe(xml, node)} lies inside citation {cluster.citation_id}; "
                "accept or reject it first")


def _set_text(xml: str, node: Node, text: str) -> tuple[int, int, str]:
    """Splice that sets one w:instrText / w:t element's text (XML-escaped, space preserved)."""

    head = node.head(xml)
    if node.closed:
        head = head[:-2].rstrip() + ">"
    if text and attr_value(head, "xml:space") is None:
        head = head[:-1] + ' xml:space="preserve">'
    return node.start, node.end, f"{head}{escape(text)}</{node.name}>"


def _new_payload(cluster: Cluster, item_keys: Sequence[str], zotero: ZoteroClient) -> tuple[dict, list[str]]:
    warnings: list[str] = []
    existing: dict[str, dict] = {}
    for entry, key in zip(cluster.payload.get("citationItems") or [], cluster.item_keys):
        existing.setdefault(key, entry)
    entries = []
    for key in item_keys:
        if key in existing:
            try:
                data = zotero.item(key)
            except Exception as exc:
                raise RefusedUnsafe(f"{OP}: Zotero item {key!r} is not in the local library: {exc}") from exc
            if data.get("deleted"):
                raise RefusedUnsafe(f"{OP}: Zotero item {key!r} is in the trash; restore it or drop it from item_keys")
            entries.append(existing[key])
            continue
        try:
            csl = zotero.csl(key)
        except RefusedUnsafe:
            raise
        except Exception as exc:
            raise RefusedUnsafe(f"{OP}: Zotero item {key!r} is not in the local library: {exc}") from exc
        uri = csl.get("id")
        if not isinstance(uri, str) or item_key({"uris": [uri]}) != key:
            raise RefusedUnsafe(f"{OP}: Zotero returned CSL for {key!r} without a matching item URI ({uri!r})")
        entries.append({"id": uri, "uris": [uri], "itemData": csl})
    payload = dict(cluster.payload)
    properties = dict(payload.get("properties") or {})
    properties["formattedCitation"] = PLACEHOLDER
    properties["plainCitation"] = PLACEHOLDER
    if properties.pop("dontUpdate", None):
        warnings.append("dropped the citation's dontUpdate flag (a hand-edited citation) so Zotero renders it")
    payload["properties"] = properties
    payload["citationItems"] = entries
    return payload, warnings


def plan_xml_edit(xml: str, cluster: Cluster, payload: dict[str, Any]) -> str:
    """The document.xml text with the cluster's instr JSON and result set; nothing else changes."""

    return apply_splices(xml, cluster_splices(xml, cluster, payload))


def cluster_splices(xml: str, cluster: Cluster, payload: dict[str, Any], op: str = OP) -> list[tuple[int, int, str]]:
    """Splices that write `payload` into the cluster's field code and `{Citation}` into its result."""

    item = cluster.field
    if not item.instr:
        raise RefusedUnsafe(f"{op}: citation {cluster.citation_id} has no w:instrText")
    if not item.result_t:
        raise RefusedUnsafe(f"{op}: citation {cluster.citation_id} has no result text run to render into")
    instruction = cluster.prefix + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + cluster.suffix
    splices = [_set_text(xml, item.instr[0], instruction)]
    splices += [_set_text(xml, node, "") for node in item.instr[1:] if not node.closed]
    splices.append(_set_text(xml, item.result_t[0], PLACEHOLDER))
    splices += [_set_text(xml, node, "") for node in item.result_t[1:] if not node.closed]
    return splices


def apply_splices(xml: str, splices: list[tuple[int, int, str]]) -> str:
    """Apply (start, end, text) replacements; they must not overlap."""

    ordered = sorted(splices)
    for (_, end, _), (start, _, _) in zip(ordered, ordered[1:]):
        if start < end:
            raise VerificationFailed("internal error: overlapping XML splices")
    for start, end, text in reversed(ordered):
        xml = xml[:start] + text + xml[end:]
    return xml


def _write_document_xml(path: Path, document_xml: str) -> None:
    """Replace word/document.xml in `path` (entry order and compression kept, every other part copied)."""

    from .revisions import rezip

    with tempfile.TemporaryDirectory(prefix="research-mcp-cluster-") as temp:
        staged = Path(temp) / path.name
        rezip(path, staged, document_xml)
        shutil.copyfile(staged, path)


def _clusters_of(path: Path) -> tuple[str, Any, list[Cluster]]:
    _, doc = load(path)
    return doc.xml, doc, clusters_in(doc.xml, doc.body)


def edit_cluster_to_version(
    source: str | Path,
    output: str | Path,
    *,
    source_sha256: str,
    paragraph_index: int,
    item_keys: Sequence[str],
    citation_id: str | None = None,
    current_text: str | None = None,
    max_wait_seconds: float = 180.0,
    zotero: ZoteroClient | None = None,
) -> dict[str, Any]:
    """Set one citation cluster's items to exactly `item_keys` and render it (versioned or in place)."""

    keys = [str(key).strip() for key in item_keys]
    if not keys:
        raise RefusedUnsafe(f"{OP}: item_keys is empty; to delete a citation, edit the text instead")
    if any(not key for key in keys):
        raise RefusedUnsafe(f"{OP}: item_keys holds an empty key")
    duplicates = sorted({key for key in keys if keys.count(key) > 1})
    if duplicates:
        raise RefusedUnsafe(f"{OP}: item_keys lists {duplicates} more than once")

    src = Path(source).resolve(strict=True)
    out = Path(output).resolve()
    if out != src and out.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    expected = source_sha256.upper()
    actual = audit_docx(src).sha256.upper()
    if expected != actual:
        raise RefusedUnsafe(
            f"source document SHA-256 mismatch: expected {expected}, "
            f"actual {actual}. Reinspect the document with inspect_document.")
    zotero = zotero or ZoteroClient()

    with staged_document(src, out) as staged:
        # Everything below reads the staged copy, pinned to the caller's sha: a save to the source after
        # the check above refuses here (and publish_inplace refuses one made after staging).
        source_audit = audit_docx(staged)
        if source_audit.sha256.upper() != expected:
            raise RefusedUnsafe(
                f"{OP}: the source changed after its SHA-256 was checked (expected {expected}, "
                f"staged {source_audit.sha256.upper()}). Reinspect the document with inspect_document.")
        xml, doc, _ = _clusters_of(Path(staged))
        paragraph = doc.paragraph(int(paragraph_index), OP)
        cluster = select_cluster(clusters_in(xml, doc.body, (paragraph.start, paragraph.end)),
                                 int(paragraph_index), citation_id=citation_id, current_text=current_text)
        refuse_unsafe_cluster(xml, doc.body, cluster)
        payload, warnings = _new_payload(cluster, keys, zotero)
        new_xml = plan_xml_edit(xml, cluster, payload)
        revisions_before = revision_snapshot(staged)
        outside_before = text_outside_zotero_fields(staged)
        bibliography_before = bibliography_entries(staged)
        target_id = cluster.citation_id

        _write_document_xml(Path(staged), new_xml)
        staged_xml, _, staged_clusters = _clusters_of(Path(staged))
        if staged_xml != new_xml:
            raise VerificationFailed(f"{OP}: the written word/document.xml is not the planned one")
        compare_revision_snapshots(revisions_before, revision_snapshot(staged), f"{OP} (XML edit)")
        mine = [c for c in staged_clusters if c.citation_id == target_id]
        if len(mine) != 1 or mine[0].item_keys != keys or mine[0].rendered != PLACEHOLDER:
            raise VerificationFailed(f"{OP}: the edited field code did not read back",
                                     expected=keys, actual=[c.item_keys for c in mine])

        rendered, _, refresh_info = refresh_file_and_wait(staged, (target_id,), max_wait_seconds=max_wait_seconds)

        final = audit_docx(staged)
        revision_check = compare_revision_snapshots(revisions_before, revision_snapshot(staged),
                                                    f"{OP} (after the Zotero refresh)")
        if final.zotero_item_fields != source_audit.zotero_item_fields:
            raise VerificationFailed(f"{OP}: Zotero citation count changed",
                                     expected=source_audit.zotero_item_fields, actual=final.zotero_item_fields)
        if final.zotero_bibliography_fields != source_audit.zotero_bibliography_fields:
            raise VerificationFailed(f"{OP}: Zotero bibliography count changed",
                                     expected=source_audit.zotero_bibliography_fields,
                                     actual=final.zotero_bibliography_fields)
        verify_protected_invariants(source_audit, final, OP)
        _, _, after_clusters = _clusters_of(Path(staged))
        mine = [c for c in after_clusters if c.citation_id == target_id]
        if len(mine) != 1:
            raise VerificationFailed(f"{OP}: citation {target_id} is not in the refreshed document exactly once",
                                     expected=1, actual=len(mine))
        after = mine[0]
        if after.item_keys != keys:
            raise VerificationFailed(f"{OP}: the refreshed citation has other items", expected=keys,
                                     actual=after.item_keys)
        unrendered = [c.citation_id for c in after_clusters if PLACEHOLDER in c.rendered or not c.rendered.strip()]
        if unrendered:
            raise VerificationFailed(f"{OP}: citations left unrendered after the refresh: {unrendered}")
        difference = first_text_difference(outside_before, text_outside_zotero_fields(staged))
        if difference:
            raise VerificationFailed(f"{OP}: text outside the Zotero fields changed {difference}")
        bibliography_after = bibliography_entries(staged)
        if any("{Bibliography}" in entry for entry in bibliography_after):
            raise VerificationFailed(f"{OP}: the bibliography was left unrendered")
        for name, value in refresh_info.items():
            if name != "output_sha256" and not name.endswith("seconds"):
                warnings.append(f"refresh {name}: {value}")

    return {
        "success": True,
        "source": str(src),
        "output": str(out),
        "source_sha256": source_audit.sha256,
        "output_sha256": audit_docx(out).sha256,
        "paragraph_index": int(paragraph_index),
        "citation_id": target_id,
        "item_keys_before": list(cluster.item_keys),
        "item_keys_after": list(after.item_keys),
        "rendered_citation_before": cluster.rendered,
        "rendered_citation_after": after.rendered,
        "refresh_rendered": rendered.get(target_id),
        "bibliography_before": bibliography_before,
        "bibliography_after": bibliography_after,
        "revisions_checked": revision_check["revisions_checked"],
        "revision_check": revision_check,
        "warnings": warnings,
    }
