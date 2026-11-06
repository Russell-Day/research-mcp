"""Restore live Zotero citations from v42's own plain-text markers.

v42 carries its citation map in the prose as `(Author Year [N])` markers whose
`[N]` indexes the numbered plain bibliography. That makes restoration
deterministic: the document states both the exact insertion point and which
reference belongs there. No alignment against an older version is required.

Ordering rule: markers are applied in descending (paragraph, offset) order so
that each replacement cannot invalidate the offsets of the markers still to be
processed.
"""

from __future__ import annotations

from dataclasses import dataclass, field as dataclass_field
import difflib
import html
import json
from pathlib import Path
import re
from typing import Any, Sequence

from .citations import (
    BIBLIOGRAPHY_CODE,
    BIBLIOGRAPHY_PLACEHOLDER,
    PLACEHOLDER,
    SCHEMA,
    _bibliography_fields,
    _has_zotero_preferences,
    _item_fields,
    new_citation_id,
)
from .constants import WD_FIELD_ADDIN
from .errors import RefusedUnsafe, VerificationFailed
from .zotero import ZoteroClient

# Innermost parenthesised group carrying at least one [N]. Using [^()] keeps a
# nested opening paren (v42 paragraph 13 has one) from swallowing the marker.
MARKER_PATTERN = re.compile(r"\(([^()]*?\[\d+\][^()]*?)\)")
REFERENCE_PATTERN = re.compile(r"\[(\d+)\]")
BIB_NUMBER_PATTERN = re.compile(r"^\s*(\d+)\s*\.")
DOI_PATTERN = re.compile(r"doi:\s*(10\.\S+)", re.IGNORECASE)
TITLE_MATCH_THRESHOLD = 0.72


@dataclass(frozen=True)
class Marker:
    paragraph_index: int
    start: int
    end: int
    text: str
    references: tuple[int, ...]


@dataclass(frozen=True)
class BibliographyEntry:
    paragraph_index: int
    number: int
    text: str


@dataclass
class RestorePlan:
    markers: list[Marker] = dataclass_field(default_factory=list)
    entries: list[BibliographyEntry] = dataclass_field(default_factory=list)
    reference_keys: dict[int, str] = dataclass_field(default_factory=dict)
    unresolved: list[int] = dataclass_field(default_factory=list)

    @property
    def total_references(self) -> int:
        return sum(len(marker.references) for marker in self.markers)

    def to_json(self) -> str:
        return json.dumps(
            {
                "markers": [
                    {
                        "paragraph_index": m.paragraph_index,
                        "start": m.start,
                        "end": m.end,
                        "text": m.text,
                        "references": list(m.references),
                        "item_keys": [self.reference_keys.get(n) for n in m.references],
                    }
                    for m in self.markers
                ],
                "bibliography_paragraphs": [e.paragraph_index for e in self.entries],
                "reference_keys": {str(k): v for k, v in sorted(self.reference_keys.items())},
                "marker_count": len(self.markers),
                "total_references": self.total_references,
                "distinct_references": len(self.reference_keys),
                "unresolved": self.unresolved,
            },
            indent=2,
        )


def _paragraph_text(paragraph) -> str:
    return str(paragraph.Range.Text).rstrip("\r\x07")


def find_markers(document) -> list[Marker]:
    """Locate every citation marker, addressed the way insert_unrendered expects."""
    markers: list[Marker] = []
    for index in range(1, int(document.Paragraphs.Count) + 1):
        paragraph = document.Paragraphs(index)
        if paragraph.Range.Information(12):  # wdWithInTable
            continue
        if str(paragraph.Style.NameLocal).lower().startswith("bibliograph"):
            continue
        text = _paragraph_text(paragraph)
        for match in MARKER_PATTERN.finditer(text):
            numbers = tuple(int(n) for n in REFERENCE_PATTERN.findall(match.group(1)))
            if not numbers:
                continue
            markers.append(
                Marker(
                    paragraph_index=index,
                    start=match.start(),
                    end=match.end(),
                    text=match.group(0),
                    references=numbers,
                )
            )
    return markers


def find_bibliography_entries(document) -> list[BibliographyEntry]:
    entries: list[BibliographyEntry] = []
    for index in range(1, int(document.Paragraphs.Count) + 1):
        paragraph = document.Paragraphs(index)
        if not str(paragraph.Style.NameLocal).lower().startswith("bibliograph"):
            continue
        text = _paragraph_text(paragraph)
        number = BIB_NUMBER_PATTERN.match(text)
        if not number:
            continue
        entries.append(
            BibliographyEntry(
                paragraph_index=index, number=int(number.group(1)), text=text
            )
        )
    return entries


def resolve_reference_keys(
    entries: Sequence[BibliographyEntry], zotero: ZoteroClient
) -> tuple[dict[int, str], list[int]]:
    """Map each [N] to a Zotero item key by DOI, falling back to title."""
    catalogue = zotero.top_level()
    by_doi: dict[str, str] = {}
    titles: list[tuple[str, str]] = []
    for item in catalogue:
        data = item.get("data", item)
        key = data.get("key")
        if not key:
            continue
        doi = str(data.get("DOI") or "").strip().lower()
        if doi:
            by_doi[doi] = key
        titles.append((str(data.get("title") or ""), key))

    resolved: dict[int, str] = {}
    unresolved: list[int] = []
    for entry in entries:
        doi_match = DOI_PATTERN.search(entry.text)
        if doi_match:
            doi = html.unescape(doi_match.group(1)).strip().lower().rstrip(".")
            if doi in by_doi:
                resolved[entry.number] = by_doi[doi]
                continue
        probe = re.sub(r"^\s*\d+\s*\.", "", entry.text)
        probe = re.split(r"\.\s", probe, maxsplit=2)[-1] if "." in probe else probe
        best_key, best_score = None, 0.0
        for title, key in titles:
            if not title:
                continue
            score = difflib.SequenceMatcher(None, probe.lower(), title.lower()).ratio()
            containment = title.lower() in entry.text.lower()
            score = 1.0 if containment and len(title) > 20 else score
            if score > best_score:
                best_key, best_score = key, score
        if best_key and best_score >= TITLE_MATCH_THRESHOLD:
            resolved[entry.number] = best_key
        else:
            unresolved.append(entry.number)
    return resolved, unresolved


def build_plan(document, zotero: ZoteroClient) -> RestorePlan:
    markers = find_markers(document)
    entries = find_bibliography_entries(document)
    keys, unresolved = resolve_reference_keys(entries, zotero)
    plan = RestorePlan(
        markers=markers, entries=entries, reference_keys=keys, unresolved=unresolved
    )
    cited = {n for marker in markers for n in marker.references}
    missing = sorted(n for n in cited if n not in keys)
    if missing:
        plan.unresolved = sorted(set(plan.unresolved) | set(missing))
    return plan


def _build_payload(csl_items: Sequence[dict[str, Any]]) -> dict[str, Any]:
    citation_items = []
    for item in csl_items:
        uri = item.get("id")
        if not isinstance(uri, str) or not uri:
            raise RefusedUnsafe("CSL item lacks a stable id URI")
        citation_items.append({"id": uri, "uris": [uri], "itemData": item})
    return {
        "citationID": new_citation_id(),
        "properties": {
            "unsorted": False,
            "formattedCitation": PLACEHOLDER,
            "plainCitation": PLACEHOLDER,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": SCHEMA,
    }


def replace_markers(
    document,
    plan: RestorePlan,
    *,
    donor_source,
    zotero: ZoteroClient,
) -> list[str]:
    """Replace each marker with one live citation field. Returns citation IDs.

    `donor_source` is a Range holding one Zotero item field, captured once from
    the donor document so it is not reopened 36 times.
    """
    if not _has_zotero_preferences(document):
        raise RefusedUnsafe("target document has no ZOTERO_PREF_* properties")

    csl_cache: dict[str, dict[str, Any]] = {}
    citation_ids: list[str] = []
    ordered = sorted(
        plan.markers, key=lambda m: (m.paragraph_index, m.start), reverse=True
    )
    for marker in ordered:
        keys = [plan.reference_keys[n] for n in marker.references]
        for key in keys:
            if key not in csl_cache:
                csl_cache[key] = zotero.csl(key)
        payload = _build_payload([csl_cache[key] for key in keys])

        paragraph = document.Paragraphs(marker.paragraph_index).Range
        absolute_start = paragraph.Start + marker.start
        absolute_end = paragraph.Start + marker.end
        span = document.Range(absolute_start, absolute_end)
        if str(span.Text) != marker.text:
            raise VerificationFailed(
                "marker text moved before replacement",
                expected=marker.text,
                actual=str(span.Text),
            )
        span.Delete()

        point = document.Range(absolute_start, absolute_start)
        point.FormattedText = donor_source.FormattedText
        pasted = [f for f in point.Fields if int(f.Type) == WD_FIELD_ADDIN]
        if len(pasted) != 1:
            raise VerificationFailed(
                "marker replacement did not paste exactly one field",
                expected=1,
                actual=len(pasted),
            )
        field = pasted[0]
        field.Code.Text = "ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(
            payload, separators=(",", ":")
        )
        field.Result.Text = PLACEHOLDER
        citation_ids.append(payload["citationID"])
    return citation_ids


def replace_plain_bibliography(document, plan: RestorePlan, *, donor_source) -> None:
    """Delete the plain bibliography paragraphs and clone one live field there."""
    if _bibliography_fields(document):
        raise RefusedUnsafe("target document already contains a Zotero bibliography")
    if not plan.entries:
        raise RefusedUnsafe("no plain bibliography paragraphs were found")

    indices = sorted(entry.paragraph_index for entry in plan.entries)
    first = indices[0]
    start = document.Paragraphs(first).Range.Start
    end = document.Paragraphs(indices[-1]).Range.End
    document.Range(start, end).Delete()

    point = document.Range(start, start)
    point.FormattedText = donor_source.FormattedText
    pasted = [
        f
        for f in point.Fields
        if int(f.Type) == WD_FIELD_ADDIN
        and str(f.Code.Text).strip().startswith("ADDIN ZOTERO_BIBL")
    ]
    if len(pasted) != 1:
        raise VerificationFailed(
            "bibliography replacement did not paste exactly one field",
            expected=1,
            actual=len(pasted),
        )
    field = pasted[0]
    # Never keep the donor's rendered entries; Zotero throws reconciling a
    # bibliography that lists items this document does not cite.
    field.Code.Text = BIBLIOGRAPHY_CODE
    field.Result.Text = BIBLIOGRAPHY_PLACEHOLDER


def assert_no_markers_remain(document) -> None:
    for index in range(1, int(document.Paragraphs.Count) + 1):
        paragraph = document.Paragraphs(index)
        if paragraph.Range.Information(12):
            continue
        text = _paragraph_text(paragraph)
        for match in MARKER_PATTERN.finditer(text):
            if REFERENCE_PATTERN.search(match.group(1)):
                raise VerificationFailed(
                    "a plain-text citation marker survived restoration",
                    expected="no (Author Year [N]) markers",
                    actual=f"paragraph {index}: {match.group(0)}",
                )
