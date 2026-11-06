"""Re-render every Zotero citation and the bibliography, and report what changed (plan 602 S3).

Direct OOXML, like `cluster_ops`: every citation's result and its formatted/plainCitation become
`{Citation}` and every ZOTERO_BIBL result becomes `{Bibliography}` in `word/document.xml` (splices,
every other byte copied), then Zotero renders them all through the supervised refresh. Allowed on
documents with tracked changes: every revision element is compared after the XML step and after Word's
save. `preview=True` runs the same refresh on a temporary copy and discards it.

Preview method: a full refresh of a staged temporary copy. The local API (`format=bib`) matched v28's 26
entries exactly, but it renders without the document's `custom` entries, embedded-only items, and
disambiguation across request chunks; the staged copy is the same code path as the real refresh, so its
diff is exactly what the refresh would write.
"""

from __future__ import annotations

from difflib import SequenceMatcher
import json
from pathlib import Path
import re
import shutil
import tempfile
from typing import Any
import unicodedata

from .audit import audit_docx
from .citation_ops import (
    bibliography_entries,
    compare_revision_snapshots,
    first_text_difference,
    revision_snapshot,
    text_outside_zotero_fields,
    verify_protected_invariants,
)
from .citations import BIBLIOGRAPHY_PLACEHOLDER, PLACEHOLDER, refresh_file_and_wait
from .cluster_ops import (
    _WRAPPERS,
    Cluster,
    ZoteroField,
    _clusters_of,
    _cluster,
    _describe,
    _write_document_xml,
    apply_splices,
    body_fields,
    cluster_splices,
    item_key,
    refuse_unsafe_cluster,
)
from .errors import RefusedUnsafe, VerificationFailed
from .revisions import MOVE_RANGES, REV_TAGS, Node, load
from .word import staged_document

OP = "refresh_citations"
_BIBL_PREFIX = "ADDIN ZOTERO_BIBL"
# Nothing in these may sit inside a bibliography result: the XML step deletes the result's runs and
# paragraphs, so a revision, bookmark, comment anchor, section break or field there would be lost.
_BIBL_FORBIDDEN = (set(REV_TAGS) | set(MOVE_RANGES) | {
    "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "commentReference",
    "sectPr", "fldChar", "instrText", "delInstrText", "fldSimple", "sdt", "footnoteReference",
    "endnoteReference", "drawing", "object", "pict", "tbl"})


# -- bibliography diff ---------------------------------------------------------------------------------

_LABEL = re.compile(r"^\s*(\[\d+\]|\(\d+\)|\d+\.?)(?=\s)")
_NUMBER_RANGE = re.compile(r"(\d+)\s*[-\u2013\u2014]\s*(\d+)|(\d+)")


def split_label(entry: str) -> tuple[str, str]:
    """(label, body) of one entry: "18.\tDay R..." -> ("18.", "Day R..."); whitespace collapsed.
    Only a candidate: `split_labels` keeps it as a label when the whole list is numbered 1..n."""

    match = _LABEL.match(entry)
    label = match.group(1) if match else ""
    body = entry[match.end():] if match else entry
    return label, " ".join(body.split())


def split_labels(entries: list[str]) -> list[tuple[str, str]]:
    """(label, body) of each entry. Leading numbers count as labels only when every entry has one and
    they run 1..n; otherwise (an author-date list starting "2020 Smith ...") they stay entry text."""

    parts = [split_label(entry) for entry in entries]
    numbers = [re.sub(r"\D", "", label) for label, _ in parts]
    if numbers != [str(n) for n in range(1, len(entries) + 1)]:
        return [("", " ".join(entry.split())) for entry in entries]
    return parts


def first_cited_numbers(rendered: list[str]) -> list[int]:
    """Citation numbers in order of first appearance, from the rendered citations in document order
    ("[1,3-5]" -> 1, 3, 4, 5)."""

    seen: list[int] = []
    for text in rendered:
        for match in _NUMBER_RANGE.finditer(text):
            low, high, single = match.groups()
            numbers = range(int(low), int(high) + 1) if low and int(high) - int(low) < 1000 else [int(single or low)]
            seen += [n for n in numbers if n not in seen]
    return seen


def _norm_text(text: str) -> str:
    """Casefolded words only: markup, punctuation and spacing dropped ("<i>In vivo</i>: A-B" -> "in vivo a b")."""

    text = unicodedata.normalize("NFKC", re.sub(r"<[^>]+>", "", text)).casefold()
    return " ".join(re.sub(r"[\W_]+", " ", text).split())


def _item_probe(data: dict[str, Any]) -> str:
    """The text an entry must show for this item: its title, or with no title its first author's family name."""

    title = _norm_text(str(data.get("title") or ""))
    if not title:
        authors = data.get("author") or []
        title = _norm_text(str(authors[0].get("family") or "")) if authors and isinstance(authors[0], dict) else ""
    return title


def entry_items(after: list[str], cited_order: list[str], bibl_data: dict[str, Any],
                rendered: list[str], item_data: dict[str, dict[str, Any]] | None = None) -> list[str | None]:
    """Item key of each rendered entry, when the order fixes it and the text confirms it; else all None.

    Entry i is the i-th distinct cited item (omitted items left out) only in citation-sequence order.
    Labels 1..n alone do not prove that (a numbered style can sort alphabetically), so the rendered
    citations must also number the items 1, 2, 3, ... in order of first appearance, n entries for n
    cited items, and no uncited items added. Even then a citation's numbers are sorted ("[1, 2]" for
    Zhao then Adams), so each key is confirmed against its entry: the entry must contain the item's
    title from the citation's itemData (or its first author's family name), and that must be the only
    way to match items to entries: if some items could swap entries (two items both titled
    "Editorial", each title in both entries) the text proves nothing. A title that also appears in
    another entry is fine as long as no swap is possible. One miss and every entry gets no key;
    entries are then matched by text alone.
    """

    omitted = {item_key(entry) for entry in bibl_data.get("omitted") or [] if isinstance(entry, dict)}
    order = [key for key in cited_order if key not in omitted]
    expected = list(range(1, len(order) + 1))
    labelled = [label for label, _ in split_labels(after)]
    if (bibl_data.get("uncited") or len(after) != len(order) or not all(labelled)
            or first_cited_numbers(rendered) != expected):
        return [None] * len(after)
    item_data = item_data or {}
    entries = [f" {_norm_text(entry)} " for entry in after]
    others: list[list[int]] = []  # entries other than its own that also contain item i's text
    for i, key in enumerate(order):
        probe = _item_probe(item_data.get(key) or {})
        hits = [j for j, entry in enumerate(entries) if probe and f" {probe} " in entry]
        if i not in hits:
            return [None] * len(after)
        others.append([j for j in hits if j != i])
    if _has_cycle(others):  # some items could swap entries: the match is not unique
        return [None] * len(after)
    return list(order)


def _has_cycle(edges: list[list[int]]) -> bool:
    """Whether the directed graph i -> edges[i] has a cycle (iterative DFS, white/grey/black)."""

    state = [0] * len(edges)
    for start in range(len(edges)):
        if state[start]:
            continue
        stack = [(start, iter(edges[start]))]
        state[start] = 1
        while stack:
            node, children = stack[-1]
            child = next(children, None)
            if child is None:
                state[node] = 2
                stack.pop()
            elif state[child] == 1:
                return True
            elif state[child] == 0:
                state[child] = 1
                stack.append((child, iter(edges[child])))
    return False


def diff_bibliography(before: list[str], after: list[str],
                      after_keys: list[str | None] | None = None) -> dict[str, Any]:
    """Pair each entry before the refresh with one after it, then sort the pairs into changed and unchanged.

    Bodies (entry text without its number label, whitespace collapsed) are paired in three passes:
    1. identical bodies, in order (so duplicate-looking entries pair first-to-first); a body that
       occurs more than once before the refresh (say a hand edit copied another entry) pairs at its
       own position first, so the copy that the refresh overwrote is the one reported changed;
    2. the most similar remaining pair, similarity >= 0.6 (difflib ratio), nearer position on ties;
    3. a remaining pair at the same position with similarity >= 0.3 (a heavy hand edit of one entry).
    Unpaired entries before are `dropped`, unpaired after are `added`. A pair whose body differs is
    `changed`; a pair that only got a new number is unchanged (counted in `renumbered_count`).
    """

    keys = after_keys or [None] * len(after)
    old_parts, new_parts = split_labels(before), split_labels(after)
    old = [body for _, body in old_parts]
    new = [body for _, body in new_parts]
    pairs: dict[int, int] = {}
    taken: set[int] = set()

    repeated = {body for body in old if old.count(body) > 1}
    for i, body in enumerate(old):
        if body in repeated and i < len(new) and new[i] == body:
            pairs[i] = i
            taken.add(i)
    for i, body in enumerate(old):
        if i in pairs:
            continue
        j = next((j for j, other in enumerate(new) if j not in taken and other == body), None)
        if j is not None:
            pairs[i] = j
            taken.add(j)

    candidates = []
    for i, body in enumerate(old):
        if i in pairs:
            continue
        for j, other in enumerate(new):
            if j in taken:
                continue
            matcher = SequenceMatcher(None, body, other, autojunk=False)
            if matcher.real_quick_ratio() < 0.3 or matcher.quick_ratio() < 0.3:
                continue
            candidates.append((matcher.ratio(), -abs(i - j), i, j))
    candidates.sort(reverse=True)
    for ratio, _, i, j in candidates:
        if ratio >= 0.6 and i not in pairs and j not in taken:
            pairs[i] = j
            taken.add(j)
    for ratio, _, i, j in candidates:
        if ratio >= 0.3 and i == j and i not in pairs and j not in taken:
            pairs[i] = j
            taken.add(j)

    changed, unchanged, renumbered = [], 0, 0
    for i, j in sorted(pairs.items()):
        if old[i] == new[j]:
            unchanged += 1
            renumbered += old_parts[i][0] != new_parts[j][0]
        else:
            changed.append({"item_key": keys[j], "before_index": i, "after_index": j,
                            "before": before[i], "after": after[j]})
    return {
        "changed": changed,
        "dropped": [{"before_index": i, "text": before[i]} for i in range(len(before)) if i not in pairs],
        "added": [{"after_index": j, "item_key": keys[j], "text": after[j]}
                  for j in range(len(after)) if j not in taken],
        "unchanged_count": unchanged,
        "renumbered_count": renumbered,
    }


# -- XML step ------------------------------------------------------------------------------------------


def _run_of(node: Node, what: str) -> Node:
    run = next((a for a in node.ancestors() if a.name == "w:r"), None)
    if run is None or run.parent is None or run.parent.name != "w:p":
        raise RefusedUnsafe(f"{OP}: the bibliography field's {what} is not in a plain paragraph run; refresh it in Word")
    return run


def bibliography_splice(xml: str, body: Node, item: ZoteroField) -> tuple[int, int, str]:
    """Splice that replaces a ZOTERO_BIBL result (runs and paragraphs) with one `{Bibliography}` run.

    The result runs from the run after `separate` to the run before `end`; both runs must sit in
    paragraphs of the same container, with only paragraphs in between, and nothing in that span that
    the cut would lose (revisions, bookmarks, comment anchors, section breaks, fields, tables).
    """

    if item.separate is None or item.end is None:
        raise RefusedUnsafe(f"{OP}: the bibliography field has no result")
    if item.nested or item.del_instr:
        raise RefusedUnsafe(f"{OP}: the bibliography field holds a nested field or deleted field code; refresh it in Word")
    for node in [item.begin, item.separate, item.end, *item.instr]:
        wrapper = next((a for a in node.ancestors() if a.name in _WRAPPERS), None)
        if wrapper is not None:
            raise RefusedUnsafe(f"{OP}: the bibliography field sits inside tracked revision "
                                f"{_describe(xml, wrapper)}; accept or reject it first")
    first, last = _run_of(item.separate, "separator"), _run_of(item.end, "end")
    container = first.parent.parent
    if last.parent.parent is not container:
        raise RefusedUnsafe(f"{OP}: the bibliography result spans different containers; refresh it in Word")
    lo, hi = first.end, last.start
    for node in body.iter():
        if node.start >= hi:
            break
        if node.start < lo:
            continue
        if node.name.startswith("w:") and node.local in _BIBL_FORBIDDEN:
            what = (f"tracked revision {_describe(xml, node)}" if node.local in REV_TAGS or node.local in MOVE_RANGES
                    else f"a w:{node.local}")
            raise RefusedUnsafe(f"{OP}: the bibliography result holds {what}; the refresh would remove it. "
                                "Fix it in Word first")
        if node.parent is container and node.name != "w:p":
            raise RefusedUnsafe(f"{OP}: the bibliography result holds a {node.name} between its paragraphs")
    style_run = next((a for a in item.result_t[0].ancestors() if a.name == "w:r"), None) if item.result_t else None
    rpr = style_run.child("w:rPr") if style_run is not None else None
    rpr_xml = xml[rpr.start:rpr.end] if rpr is not None else ""
    return lo, hi, f"<w:r>{rpr_xml}<w:t>{BIBLIOGRAPHY_PLACEHOLDER}</w:t></w:r>"


def plan_refresh_xml(xml: str, body: Node) -> tuple[str, list[Cluster], list[Cluster], int]:
    """(new xml, clusters reset, clusters kept because Zotero marks them hand-edited, bibliography count)."""

    splices: list[tuple[int, int, str]] = []
    reset: list[Cluster] = []
    kept: list[Cluster] = []
    bibliographies = 0
    for item in body_fields(xml, body):
        instruction = item.instruction(xml).strip()
        if instruction.startswith(_BIBL_PREFIX):
            splices.append(bibliography_splice(xml, body, item))
            bibliographies += 1
            continue
        cluster = _cluster(xml, item)
        if cluster is None:
            continue
        refuse_unsafe_cluster(xml, body, cluster, OP)
        properties = dict(cluster.payload.get("properties") or {})
        if properties.get("dontUpdate"):
            # Zotero keeps a citation the user edited by hand ("dontUpdate") as is on refresh; so do we.
            kept.append(cluster)
            continue
        properties["formattedCitation"] = PLACEHOLDER
        properties["plainCitation"] = PLACEHOLDER
        splices += cluster_splices(xml, cluster, {**cluster.payload, "properties": properties}, OP)
        reset.append(cluster)
    return apply_splices(xml, splices), reset, kept, bibliographies


def _bibl_data(xml: str, body: Node) -> dict[str, Any]:
    for item in body_fields(xml, body):
        instruction = item.instruction(xml).strip()
        if instruction.startswith(_BIBL_PREFIX):
            start = instruction.find("{")
            try:
                data, _ = json.JSONDecoder().raw_decode(instruction[start:]) if start >= 0 else ({}, 0)
            except json.JSONDecodeError:
                data = {}
            return data if isinstance(data, dict) else {}
    return {}


def _bibliography_results(xml: str, body: Node) -> list[str]:
    return [item.rendered(xml) for item in body_fields(xml, body)
            if item.instruction(xml).strip().startswith(_BIBL_PREFIX)]


# -- refresh -------------------------------------------------------------------------------------------


def _refresh_staged(staged: Path, expected: str, max_wait_seconds: float) -> dict[str, Any]:
    source_audit = audit_docx(staged)
    if source_audit.sha256.upper() != expected:
        raise RefusedUnsafe(
            f"{OP}: the source changed after its SHA-256 was checked (expected {expected}, "
            f"staged {source_audit.sha256.upper()}). Reinspect the document with inspect_document.")
    _, doc = load(staged)
    xml = doc.xml
    new_xml, reset, kept, bibliographies = plan_refresh_xml(xml, doc.body)
    if not reset and not bibliographies:
        raise RefusedUnsafe(f"{OP}: the document has no Zotero citation or bibliography to refresh")
    cited_order: list[str] = []
    _, _, all_clusters = _clusters_of(staged)
    item_data: dict[str, dict[str, Any]] = {}
    for cluster in all_clusters:
        cited_order += [key for key in cluster.item_keys if key not in cited_order]
        for entry in cluster.payload.get("citationItems") or []:
            if isinstance(entry, dict) and isinstance(entry.get("itemData"), dict):
                item_data.setdefault(item_key(entry), entry["itemData"])
    bibl_data = _bibl_data(xml, doc.body)
    revisions_before = revision_snapshot(staged)
    outside_before = text_outside_zotero_fields(staged)
    bibliography_before = bibliography_entries(staged)
    before = {c.citation_id: c for c in all_clusters}
    reset_ids = [c.citation_id for c in reset]

    _write_document_xml(staged, new_xml)
    _, staged_doc = load(staged)
    if staged_doc.xml != new_xml:
        raise VerificationFailed(f"{OP}: the written word/document.xml is not the planned one")
    compare_revision_snapshots(revisions_before, revision_snapshot(staged), f"{OP} (XML edit)")
    staged_clusters = {c.citation_id: c for c in _clusters_of(staged)[2]}
    not_reset = [cid for cid in reset_ids
                 if cid not in staged_clusters or staged_clusters[cid].rendered != PLACEHOLDER
                 or staged_clusters[cid].item_keys != before[cid].item_keys]
    if not_reset or any(r != BIBLIOGRAPHY_PLACEHOLDER for r in _bibliography_results(staged_doc.xml, staged_doc.body)):
        raise VerificationFailed(f"{OP}: the XML reset did not read back", actual=not_reset)

    _, _, refresh_info = refresh_file_and_wait(staged, reset_ids, max_wait_seconds=max_wait_seconds)

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
    after = {c.citation_id: c for c in _clusters_of(staged)[2]}
    if set(after) != set(before):
        raise VerificationFailed(f"{OP}: the citation ids changed during the refresh",
                                 expected=sorted(before), actual=sorted(after))
    moved = [cid for cid in before if after[cid].item_keys != before[cid].item_keys]
    if moved:
        raise VerificationFailed(f"{OP}: the refresh changed the items of citations {moved}")
    unrendered = [cid for cid, c in after.items() if PLACEHOLDER in c.rendered or not c.rendered.strip()]
    if unrendered:
        raise VerificationFailed(f"{OP}: citations left unrendered after the refresh: {unrendered}")
    difference = first_text_difference(outside_before, text_outside_zotero_fields(staged))
    if difference:
        raise VerificationFailed(f"{OP}: text outside the Zotero fields changed {difference}")
    bibliography_after = bibliography_entries(staged)
    if any(BIBLIOGRAPHY_PLACEHOLDER in entry for entry in bibliography_after):
        raise VerificationFailed(f"{OP}: the bibliography was left unrendered")

    warnings = [f"citation {c.citation_id} {c.rendered!r} is marked edited by hand in Word (dontUpdate); "
                "Zotero keeps its text, so it was not re-rendered" for c in kept]
    for name, value in refresh_info.items():
        if name != "output_sha256" and not name.endswith("seconds"):
            warnings.append(f"refresh {name}: {value}")
    diff = diff_bibliography(bibliography_before, bibliography_after,
                             entry_items(bibliography_after, cited_order, bibl_data,
                                         [c.rendered for c in after.values()], item_data))
    return {
        "source_sha256": source_audit.sha256,
        "citations_refreshed": len(reset_ids),
        "citations_changed": [
            {"citation_id": cid, "item_keys": after[cid].item_keys,
             "before": before[cid].rendered, "after": after[cid].rendered}
            for cid in before if " ".join(before[cid].rendered.split()) != " ".join(after[cid].rendered.split())],
        "bibliography_changed": diff["changed"],
        "bibliography_dropped": diff["dropped"],
        "bibliography_added": diff["added"],
        "bibliography_unchanged_count": diff["unchanged_count"],
        "bibliography_renumbered_count": diff["renumbered_count"],
        "bibliography_before": bibliography_before,
        "bibliography_after": bibliography_after,
        "custom_entries": len(bibl_data.get("custom") or []),
        "revisions_checked": revision_check["revisions_checked"],
        "revision_check": revision_check,
        "warnings": warnings,
    }


def refresh_to_version(
    source: str | Path,
    output: str | Path,
    *,
    source_sha256: str,
    preview: bool = False,
    max_wait_seconds: float = 300.0,
) -> dict[str, Any]:
    """Re-render every citation and the bibliography into `output` (or in place); `preview` writes nothing."""

    src = Path(source).resolve(strict=True)
    out = Path(output).resolve()
    if not preview and out != src and out.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output}")
    expected = source_sha256.upper()
    actual = audit_docx(src).sha256.upper()
    if expected != actual:
        raise RefusedUnsafe(
            f"source document SHA-256 mismatch: expected {expected}, "
            f"actual {actual}. Reinspect the document with inspect_document.")

    if preview:
        with tempfile.TemporaryDirectory(prefix="research-mcp-refresh-preview-") as temp:
            staged = Path(temp) / src.name
            shutil.copyfile(src, staged)
            result = _refresh_staged(staged, expected, max_wait_seconds)
        return {"success": True, "preview": True, "source": str(src), "output": None, "output_sha256": None,
                **result}

    with staged_document(src, out) as staged:
        result = _refresh_staged(Path(staged), expected, max_wait_seconds)
    return {"success": True, "preview": False, "source": str(src), "output": str(out),
            "output_sha256": audit_docx(out).sha256, **result}
