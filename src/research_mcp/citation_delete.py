"""Delete one whole Zotero citation field, as a tracked deletion by default (plan 831 S4).

Direct OOXML, never through Word: every run of the field (fldChar begin .. end) goes into one `w:del` by
the author, with `w:instrText` retagged `w:delInstrText` and `w:t` retagged `w:delText`. That is the exact
inverse of what `reject_revisions` restores for a whole field inside a `w:del`, so rejecting the deletion in
Word or through the MCP gives the live field back (same field code JSON, same result runs). When the field is
preceded by one space and followed by a space, punctuation or the paragraph end, that space goes into the same
`w:del` ("legend 22." reads "legend." after Accept).

No Zotero refresh: `refresh_citations` refuses citations inside revisions, and the bibliography must not change
while the deletion is only a proposal. The result names the items cited only by this citation
(`orphaned_items`) and their current bibliography entries: they leave the bibliography at the next
`refresh_citations` after the deletion is accepted. `tracked=False` removes the field outright (same warning).
"""

from __future__ import annotations

import copy
import datetime as dt
import json
import re
import tempfile
from pathlib import Path
from typing import Any

from .cluster_ops import _ITEM_PREFIX, Cluster, clusters_in, item_key, refuse_unsafe_cluster, select_cluster
from .errors import RefusedUnsafe, VerificationFailed
from .ooxml import W
from .revisions import (
    DEFAULT_AUTHOR,
    REV_TAGS,
    Node,
    _attr_escape,
    _check_stamps,
    _detached_comments,
    _et_final_chars,
    _fragment,
    _result,
    _revision_raw,
    _verify_written,
    et_resolve,
    load,
    next_revision_id,
    parse,
    read_parts,
    revision_stamp,
    resolve_revisions,
    rezip,
    unescape,
)

OP = "delete_citation"
# Elements a field's runs may sit between inside the paragraph (anything else is refused, named).
_BETWEEN_OK = ("w:r", "w:proofErr")
_MARKERS = {"w:bookmarkStart": "bookmark", "w:bookmarkEnd": "bookmark", "w:commentRangeStart": "comment",
            "w:commentRangeEnd": "comment", "w:permStart": "permission range", "w:permEnd": "permission range"}
_RUN_REFUSED = ("w:drawing", "w:pict", "w:object", "w:commentReference", "w:footnoteReference",
                "w:endnoteReference")
# A deleted space is taken along only when the citation is followed by one of these (or the paragraph end).
_FOLLOW_OK = set(" .,;:!?)]}’”")


def _retag(fragment: str) -> str:
    """w:t -> w:delText and w:instrText -> w:delInstrText (start, end and self-closing tags only)."""

    fragment = re.sub(r"<w:t(?=[\s/>])", "<w:delText", fragment)
    fragment = fragment.replace("</w:t>", "</w:delText>")
    fragment = re.sub(r"<w:instrText(?=[\s/>])", "<w:delInstrText", fragment)
    return fragment.replace("</w:instrText>", "</w:delInstrText>")


def _text_of(xml: str, run: Node) -> list[Node]:
    return [item for item in run.children if item.name != "w:rPr"]


def _field_runs(xml: str, paragraph: Node, cluster: Cluster, *, op: str = OP,
                fix: str = "delete it in Word") -> tuple[int, int]:
    """Indices (in paragraph.children) of the field's begin and end runs; refuses any shape Word could not
    take as one revision of whole runs (a deletion here; insert_zotero_citation's tracked insertion uses it
    too, with its own `op` and `fix`)."""

    field = cluster.field
    begin_run, end_run = field.begin.parent, field.end.parent  # type: ignore[union-attr]
    cid = cluster.citation_id
    if end_run is None or not any(node is paragraph for node in end_run.ancestors()):
        raise RefusedUnsafe(f"{op}: citation {cid} spans paragraphs; {fix}")
    for run in (begin_run, end_run):
        if run is None or run.name != "w:r" or run.parent is not paragraph:
            holder = run.parent.name if run is not None and run.parent is not None else "?"
            raise RefusedUnsafe(f"{op}: citation {cid} sits inside <{holder}> rather than directly in the "
                                f"paragraph; {fix}")
    first, last = paragraph.children.index(begin_run), paragraph.children.index(end_run)
    # The whole begin and end runs go into the deletion, so they must hold nothing outside the field (review F1).
    outside = _text_of(xml, begin_run)[:_text_of(xml, begin_run).index(field.begin)] + \
        _text_of(xml, end_run)[_text_of(xml, end_run).index(field.end) + 1:]
    if outside:
        raise RefusedUnsafe(
            f"{op}: citation {cid} shares a run with <{outside[0].name}> outside the field (its begin or end "
            f"marker is not in a run of its own); {fix}")
    for node in paragraph.children[first:last + 1]:
        if node.name in _MARKERS:
            raise RefusedUnsafe(
                f"{op}: a {_MARKERS[node.name]} marker lies inside citation {cid} (a {_MARKERS[node.name]} "
                f"that covers only part of it); {fix}")
        if node.name not in _BETWEEN_OK:
            raise RefusedUnsafe(f"{op}: citation {cid} holds <{node.name}>; {fix}")
        bad = next((item.name for item in node.iter() if item.name in _RUN_REFUSED), None)
        if bad:
            raise RefusedUnsafe(f"{op}: citation {cid} holds <{bad}>; {fix}")
    return first, last


def _next_char(xml: str, paragraph: Node, after: int) -> str | None:
    """First visible character after paragraph.children[after] (None at the paragraph end)."""

    def visible(node: Node) -> str | None:
        if node.name in ("w:del", "w:moveFrom"):
            return None  # deleted text, tabs and breaks are not in the accepted view
        if node.name == "w:t" and not node.closed and node.inner(xml):
            return unescape(node.inner(xml))[:1]
        if node.name in ("w:tab", "w:br", "w:cr"):
            return "\t" if node.name == "w:tab" else "\v"
        for child in node.children:
            found = visible(child)
            if found is not None:
                return found
        return None

    for node in paragraph.children[after + 1:]:
        found = visible(node)
        if found is not None:
            return found
    return None


def _space_split(xml: str, paragraph: Node, first: int, last: int) -> tuple[Node, Node, str] | None:
    """(previous run, its last w:t, run head) when one trailing space before the field can go with it."""

    if first == 0:
        return None
    prev = paragraph.children[first - 1]
    if prev.name != "w:r":
        return None
    rpr = prev.child("w:rPr")
    if rpr is not None and any(item.local == "rPrChange" for item in rpr.iter()):
        return None  # splitting a run with a tracked formatting change is refused elsewhere; keep the space
    content = _text_of(xml, prev)
    if not content or content[-1].name != "w:t" or content[-1].closed:
        return None
    raw = content[-1].inner(xml)
    if not raw.endswith(" ") or raw.endswith("  "):
        return None
    following = _next_char(xml, paragraph, last)
    if following is not None and following not in _FOLLOW_OK:
        return None
    return prev, content[-1], prev.head(xml)


def _split_prev_run(xml: str, prev: Node, text_node: Node, tracked: bool, del_attrs: str) -> tuple[str, str]:
    """(the previous run without its trailing space, the deleted space run or "")."""

    raw = text_node.inner(xml)[:-1]
    if raw:
        kept_t = xml[text_node.start:text_node.head_end] + raw + f"</{text_node.name}>"
    else:
        kept_t = ""
    kept = xml[prev.start:text_node.start] + kept_t + xml[text_node.end:prev.end]
    if not raw and not [item for item in _text_of(xml, prev) if item is not text_node]:
        kept = ""  # the run held only that space
    rpr = prev.child("w:rPr")
    space = (f'{prev.head(xml)}{rpr and xml[rpr.start:rpr.end] or ""}'
             f'<w:delText xml:space="preserve"> </w:delText></w:r>') if tracked else ""
    return kept, space


def _paragraph_chars(xml: str, fragment: str) -> str:
    return "".join(ch for ch, _ in _et_final_chars(_fragment(xml, fragment)))


def _simple_citations(xml: str, body: Node, span: tuple[int, int] | None = None) -> list[list[str]]:
    """Item keys of each Zotero citation stored as a simple field (`w:fldSimple`) in `body` (or in `span`).

    Zotero's Word plugin writes complex fields; a simple one comes from another tool. delete_citation does not
    delete it, but it still cites its items, so it counts when deciding which items a deletion orphans.
    """

    found = []
    for node in body.iter():
        if node.name != "w:fldSimple" or (span is not None and not span[0] <= node.start < span[1]):
            continue
        instruction = unescape(node.attr(xml, "w:instr") or "").strip()
        if not instruction.startswith(_ITEM_PREFIX):
            continue
        keys: list[str] = []
        start = instruction.find("{")
        try:
            payload = json.JSONDecoder().raw_decode(instruction[start:])[0] if start >= 0 else {}
        except json.JSONDecodeError:
            payload = {}
        for entry in (payload.get("citationItems") or []) if isinstance(payload, dict) else []:
            key = item_key(entry)
            if key:
                keys.append(key)
        found.append(keys)
    return found


def _orphans(xml: str, body: Node, cluster: Cluster, source: Path) -> list[dict[str, Any]]:
    from .citation_ops import bibliography_entries

    others = {key for item in clusters_in(xml, body) if item.field.begin is not cluster.field.begin
              for key in item.item_keys}
    others.update(key for keys in _simple_citations(xml, body) for key in keys)
    entries = bibliography_entries(source)
    found = []
    items = cluster.payload.get("citationItems") or []
    for position, key in enumerate(cluster.item_keys):
        if key in others or key == "?":
            continue
        data = items[position].get("itemData") if position < len(items) and isinstance(items[position], dict) \
            else None
        title = str((data or {}).get("title") or "").strip()
        norm = " ".join(title.casefold().split())
        matches = [entry for entry in entries if norm and norm in " ".join(entry.casefold().split())]
        found.append({"item_key": key, "title": title or None,
                      "bibliography_entry": matches[0] if len(matches) == 1 else None})
    return found


def delete_citation(source, output, *, paragraph_index: int, citation_id: str | None = None,
                    current_text: str | None = None, author: str = DEFAULT_AUTHOR, tracked: bool = True,
                    now: dt.datetime | None = None) -> dict:
    source, output = Path(source), Path(output)
    author = (author or "").strip() or DEFAULT_AUTHOR
    parts = read_parts(source)
    _, doc = load(source)
    xml = doc.xml
    paragraph = doc.paragraph(int(paragraph_index), OP)
    clusters = clusters_in(xml, doc.body, (paragraph.start, paragraph.end))
    try:
        if citation_id is None and current_text is None:
            # Plan 864 decision 18: no selector refuses here, listing what the paragraph holds.
            listed = "; ".join(c.describe() for c in clusters) or "none"
            raise RefusedUnsafe(
                f"{OP}: say which citation to delete with current_text (its rendered text, e.g. \"8\") or "
                f"citation_id; Zotero citations in paragraph {paragraph_index}: {listed}")
        cluster = select_cluster(clusters, int(paragraph_index), citation_id=citation_id, current_text=current_text,
                                 op=OP)
    except RefusedUnsafe as exc:
        simple = _simple_citations(xml, doc.body, (paragraph.start, paragraph.end))
        if not simple:
            raise
        raise RefusedUnsafe(
            f"{exc} (paragraph {paragraph_index} also holds {len(simple)} Zotero citation(s) stored as a simple "
            f"field, w:fldSimple, which {OP} does not delete: Zotero's Word plugin writes complex fields, so this "
            f"one came from another tool; delete it in Word)") from None
    refuse_unsafe_cluster(xml, doc.body, cluster, op=OP)
    first, last = _field_runs(xml, paragraph, cluster)
    begin_run, end_run = paragraph.children[first], paragraph.children[last]

    first_id = next_revision_id(parts)
    del_attrs = f'w:id="{first_id}" w:author="{_attr_escape(author)}" {revision_stamp(xml, now)}'
    split = _space_split(xml, paragraph, first, last)
    field_xml = xml[begin_run.start:end_run.end]
    if split is not None:
        prev, text_node, _ = split
        kept, space_run = _split_prev_run(xml, prev, text_node, tracked, del_attrs)
        lead, start = kept, prev.start
    else:
        lead, space_run, start = "", "", begin_run.start
    removed = (_retag(field_xml) if tracked else "")
    middle = f"<w:del {del_attrs}>{space_run}{removed}</w:del>" if tracked else ""
    new_paragraph = (xml[paragraph.start:start] + lead + middle + xml[end_run.end:paragraph.end])
    new_xml = xml[:paragraph.start] + new_paragraph + xml[paragraph.end:]

    _verify(doc, paragraph, new_paragraph, new_xml, begin_run, end_run, first_id, author, tracked,
            space=split is not None)
    rezip(source, output, new_xml)
    try:
        _verify_written(source, output, new_xml, first_id, OP)
        detached = _verify_resolution(source, output, cluster, first_id, tracked)
    except Exception:
        output.unlink(missing_ok=True)
        raise

    orphaned = _orphans(xml, doc.body, cluster, source)
    warnings = []
    if orphaned:
        when = ("at the next refresh_citations after this deletion is accepted" if tracked
                else "at the next refresh_citations")
        warnings.append(f"{len(orphaned)} item(s) cited only by this citation leave the bibliography {when}: "
                        + "; ".join(item["title"] or item["item_key"] for item in orphaned))
    if detached:
        warnings.append(("accepting the deletion leaves" if tracked else "the deletion leaves")
                        + f" comment(s) {detached} with no text (Word would delete them)")
    old_text = _paragraph_chars(xml, xml[paragraph.start:paragraph.end])
    _, after_doc = load(output)
    after_node = after_doc.paragraph(int(paragraph_index), OP)
    accepted = copy.deepcopy(_fragment(after_doc.xml, after_doc.xml[after_node.start:after_node.end]))
    et_resolve(accepted, lambda el: el.get(W + "id") == str(first_id), "accept")
    return _result(source, output, OP, doc, paragraph_index=int(paragraph_index), tracked=tracked,
                   citation_id=cluster.citation_id, item_keys=cluster.item_keys, rendered=cluster.rendered,
                   revision_id=str(first_id) if tracked else None, author=author if tracked else None,
                   also_deleted=" " if split is not None else "",
                   text_before=old_text, text_after="".join(ch for ch, _ in _et_final_chars(accepted)),
                   orphaned_items=orphaned, warnings=warnings,
                   message=(f"Deleted citation {cluster.rendered!r} ({cluster.citation_id}) in paragraph "
                            f"{paragraph_index}" + (f" as tracked deletion {first_id} by {author}" if tracked
                                                    else " (untracked)")
                            + ". The bibliography is not refreshed: " + (
                                "it changes at refresh_citations after the deletion is accepted."
                                if tracked else "run refresh_citations when ready.")))


def _verify(doc, paragraph: Node, new_paragraph: str, new_xml: str, begin_run: Node, end_run: Node,
            new_id: int, author: str, tracked: bool, *, space: bool) -> None:
    """Checked before anything is written: bytes outside the paragraph, every other revision, Reject gives
    the input paragraph back, Accept gives the paragraph text without the citation (and the one space)."""

    from .sections import canonical

    xml = doc.xml
    old_paragraph = xml[paragraph.start:paragraph.end]
    if new_xml[:paragraph.start] != xml[:paragraph.start] or \
            new_xml[paragraph.start + len(new_paragraph):] != xml[paragraph.end:]:
        raise VerificationFailed(f"{OP}: bytes outside the paragraph changed")
    old_tree, new_tree = parse(old_paragraph), parse(new_paragraph)
    old_revs = [_revision_raw(old_paragraph, item) for item in old_tree.iter()
                if item.local in REV_TAGS and item.name.startswith("w:")]
    kept, fresh = [], []
    for item in new_tree.iter():
        if item.local in REV_TAGS and item.name.startswith("w:"):
            if tracked and item.local == "del" and item.attr(new_paragraph, "w:id") == str(new_id):
                fresh.append(item)
            else:
                kept.append(_revision_raw(new_paragraph, item))
    if kept != old_revs:
        raise VerificationFailed(f"{OP}: a pre-existing revision in the paragraph changed")
    if tracked:
        if len(fresh) != 1:
            raise VerificationFailed(f"{OP}: expected exactly one new deletion, found {len(fresh)}")
        _check_stamps(new_paragraph, fresh, author, OP)
    old_p, new_p = _fragment(xml, old_paragraph), _fragment(xml, new_paragraph)
    is_new = lambda el: tracked and el.get(W + "id") == str(new_id)  # noqa: E731
    if tracked:
        rejected = copy.deepcopy(new_p)
        et_resolve(rejected, is_new, "reject")
        if canonical(rejected) != canonical(old_p):
            raise VerificationFailed(f"{OP}: rejecting the deletion does not give back the paragraph")
    accepted = copy.deepcopy(new_p)
    et_resolve(accepted, is_new, "accept")
    old_text = _paragraph_chars(xml, old_paragraph)
    k = len(_paragraph_chars(xml, xml[paragraph.start:begin_run.start] + "</w:p>"))
    m = len(_paragraph_chars(xml, xml[paragraph.start:end_run.end] + "</w:p>"))
    s = 1 if space else 0
    if space and old_text[k - 1:k] != " ":
        raise VerificationFailed(f"{OP}: the space before the citation is not where it was planned")
    expected = old_text[:k - s] + old_text[m:]
    got = "".join(ch for ch, _ in _et_final_chars(accepted))
    if got != expected:
        raise VerificationFailed(f"{OP}: accepting the deletion does not give the paragraph without the citation",
                                 expected=expected, actual=got)


def _identity(cluster: Cluster) -> tuple:
    return cluster.citation_id, tuple(cluster.item_keys), cluster.prefix, repr(cluster.payload), cluster.suffix


def _verify_resolution(source: Path, output: Path, cluster: Cluster, new_id: int, tracked: bool) -> list:
    """Through the MCP's own accept/reject: Accept removes exactly this citation; Reject gives back every
    citation as it was (field code JSON included). Returns the comments Accept would leave without text."""

    _, before = load(source)
    want = [_identity(item) for item in clusters_in(before.xml, before.body)]
    gone = _identity(cluster)
    remaining = list(want)
    remaining.remove(gone)
    if not tracked:
        _, after = load(output)
        got = [_identity(item) for item in clusters_in(after.xml, after.body)]
        if got != remaining:
            raise VerificationFailed(f"{OP}: the document's citations after the deletion are not the others")
        return _detached_comments(before.xml, after.xml)  # review F2: same warning as an accepted deletion
    with tempfile.TemporaryDirectory(prefix="research-mcp-delcite-") as temp:
        accepted, rejected = Path(temp) / "accept.docx", Path(temp) / "reject.docx"
        try:
            done = resolve_revisions(output, accepted, mode="accept", ids=[str(new_id)])
            resolve_revisions(output, rejected, mode="reject", ids=[str(new_id)])
        except (RefusedUnsafe, VerificationFailed) as exc:
            raise VerificationFailed(f"{OP}: the new deletion cannot be resolved cleanly: {exc}") from exc
        for path, expected, what in ((accepted, remaining, "accepting"), (rejected, want, "rejecting")):
            _, after = load(path)
            got = [_identity(item) for item in clusters_in(after.xml, after.body)]
            if got != expected:
                raise VerificationFailed(f"{OP}: {what} the deletion does not leave the expected citations")
        return list(done.get("detached_comments") or [])


def apply_operation(source, output, operation) -> dict:
    return delete_citation(source, output, paragraph_index=operation.paragraph_index,
                           citation_id=operation.citation_id, current_text=operation.current_text,
                           author=operation.author or DEFAULT_AUTHOR,
                           tracked=True if operation.tracked is None else operation.tracked)
