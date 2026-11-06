"""Live Zotero field construction with narrow placement and verification."""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import random
import re
import string
import time
from typing import Any, Sequence

from .constants import WD_FIELD_ADDIN
from .editing import (
    _WD_COLLAPSE_START,
    _field_spans,
    _refuse_field_overlap,
    _refuse_unsafe_candidate,
    locate_anchor,
)
from .errors import DocumentDamaged, RefusedUnsafe, VerificationFailed, ZoteroModal
from .integration import dispatch_refresh
from .zotero import ZoteroClient


PLACEHOLDER = "{Citation}"
BIBLIOGRAPHY_PLACEHOLDER = "{Bibliography}"
BIBLIOGRAPHY_CODE = 'ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY'
SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
VIRGIN_DOCUMENT_HELP = (
    "Target document has no ZOTERO_PREF_* properties (it is an uninitialized / virgin document). "
    "To fix: Please open this document in Microsoft Word, insert at least one Zotero citation "
    "manually (or choose 'Document Preferences' from the Zotero tab), and save it. "
    "This initializes the document's citation style and session XML, avoiding interactive Word COM modal freezes."
)


@dataclass(frozen=True)
class CitationRef:
    citation_id: str
    result: str
    item_keys: tuple[str, ...]


def new_citation_id() -> str:
    return "".join(random.SystemRandom().choices(string.ascii_letters + string.digits, k=8))


def build_payload(csl_items: Sequence[dict[str, Any]], *, citation_id: str | None = None) -> dict[str, Any]:
    if not csl_items:
        raise RefusedUnsafe("a citation requires at least one item")
    citation_items = []
    for item in csl_items:
        uri = item.get("id")
        if not isinstance(uri, str) or not uri:
            raise RefusedUnsafe("CSL item lacks a stable id URI")
        citation_items.append({"id": uri, "uris": [uri], "itemData": item})
    return {
        "citationID": citation_id or new_citation_id(),
        "properties": {
            "unsorted": False,
            "formattedCitation": PLACEHOLDER,
            "plainCitation": PLACEHOLDER,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": SCHEMA,
    }


def parse_zotero_field(field) -> dict[str, Any]:
    code = str(field.Code.Text).strip()
    start = code.find("{")
    if start < 0:
        raise ValueError("Zotero field has no JSON payload")
    return json.loads(code[start:])


def _item_fields(document) -> list:
    return [
        field
        for field in document.Fields
        if int(field.Type) == WD_FIELD_ADDIN
        and str(field.Code.Text).strip().startswith("ADDIN ZOTERO_ITEM")
    ]


def _bibliography_fields(document) -> list:
    return [
        field
        for field in document.Fields
        if int(field.Type) == WD_FIELD_ADDIN
        and str(field.Code.Text).strip().startswith("ADDIN ZOTERO_BIBL")
    ]


def _reset_bibliography_field(field) -> None:
    field.Code.Text = BIBLIOGRAPHY_CODE
    field.Result.Text = BIBLIOGRAPHY_PLACEHOLDER
    if str(field.Code.Text).strip() != BIBLIOGRAPHY_CODE or str(field.Result.Text).strip() != BIBLIOGRAPHY_PLACEHOLDER:
        raise VerificationFailed(
            "bibliography field did not reset to its canonical placeholder",
            expected={"code": BIBLIOGRAPHY_CODE, "result": BIBLIOGRAPHY_PLACEHOLDER},
            actual={"code": str(field.Code.Text).strip(), "result": str(field.Result.Text).strip()},
        )


def _has_zotero_preferences(document) -> bool:
    for index in range(1, document.CustomDocumentProperties.Count + 1):
        if str(document.CustomDocumentProperties(index).Name).startswith("ZOTERO_PREF_"):
            return True
    return False


def renew_zotero_session(document) -> tuple[str, str]:
    """Assign a unique Zotero document session to a newly versioned copy."""

    properties = []
    for index in range(1, document.CustomDocumentProperties.Count + 1):
        prop = document.CustomDocumentProperties(index)
        match = re.fullmatch(r"ZOTERO_PREF_(\d+)", str(prop.Name))
        if match:
            properties.append((int(match.group(1)), prop))
    properties.sort(key=lambda item: item[0])
    if not properties:
        raise RefusedUnsafe(VIRGIN_DOCUMENT_HELP)
    parts = [str(prop.Value) for _, prop in properties]
    combined = "".join(parts)
    match = re.search(r'<session id="([^"]+)"\s*/>', combined)
    if match is None or len(re.findall(r'<session id="[^"]+"\s*/>', combined)) != 1:
        raise RefusedUnsafe("Zotero preferences must contain exactly one session id")
    old_id = match.group(1)
    new_id = new_citation_id()
    updated = combined[: match.start(1)] + new_id + combined[match.end(1) :]
    offset = 0
    for position, (_, prop) in enumerate(properties):
        length = len(parts[position]) if position < len(properties) - 1 else len(updated) - offset
        prop.Value = updated[offset : offset + length]
        offset += length
    reread = "".join(str(prop.Value) for _, prop in properties)
    if f'<session id="{new_id}"' not in reread or old_id == new_id:
        raise VerificationFailed("Zotero session renewal did not read back")
    return old_id, new_id


def insert_unrendered(
    document,
    *,
    donor_path: str | Path,
    paragraph_index: int,
    anchor: str,
    item_keys: Sequence[str],
    zotero: ZoteroClient,
) -> CitationRef:
    """Insert one unrendered citation field immediately before `anchor`.

    `anchor` must occur exactly once in `paragraph_index`'s text -- checked
    against Word's own paragraph text, not a computed offset, because a
    paragraph holding a field has `Range` coordinates that do not match its
    visible text (`locate_anchor`'s docstring in `editing.py` has the
    measured example). Match scope is this paragraph only; a duplicate of
    `anchor` elsewhere in the document is not a conflict.
    """
    if not _has_zotero_preferences(document):
        raise RefusedUnsafe(VIRGIN_DOCUMENT_HELP)
    if not 1 <= paragraph_index <= document.Paragraphs.Count:
        raise RefusedUnsafe("paragraph index is out of range")
    if not anchor:
        raise RefusedUnsafe("anchor must not be empty")
    paragraph_text = str(document.Paragraphs(paragraph_index).Range.Text).rstrip("\r\x07")
    occurrences = paragraph_text.count(anchor)
    if occurrences != 1:
        raise RefusedUnsafe(
            f"anchor must occur exactly once in paragraph {paragraph_index}; "
            f"it occurs {occurrences} times"
        )
    candidate = locate_anchor(document, paragraph_index, anchor)
    _refuse_field_overlap(
        _field_spans(document, paragraph_index), int(candidate.Start), int(candidate.End)
    )
    _refuse_unsafe_candidate(document, candidate)
    point = candidate.Duplicate
    point.Collapse(_WD_COLLAPSE_START)
    if point.Start != point.End:
        raise RefusedUnsafe("citation insertion point must be collapsed")

    csl_items = [zotero.csl(key) for key in item_keys]
    payload = build_payload(csl_items)
    donor = document.Application.Documents.Open(
        str(Path(donor_path).resolve(strict=True)), ReadOnly=True, AddToRecentFiles=False, Visible=False
    )
    try:
        donors = _item_fields(donor)
        if len(donors) != 1:
            raise RefusedUnsafe(f"donor must contain exactly one Zotero item field, found {len(donors)}")
        donor_field = donors[0]
        source = donor.Range(donor_field.Code.Start - 1, donor_field.Result.End + 1)
        point.FormattedText = source.FormattedText
    finally:
        donor.Close(0)

    pasted = [field for field in point.Fields if int(field.Type) == WD_FIELD_ADDIN]
    if len(pasted) != 1:
        raise VerificationFailed("paste range did not contain exactly one ADDIN field", expected=1, actual=len(pasted))
    field = pasted[0]
    field.Code.Text = "ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, separators=(",", ":"))
    field.Result.Text = PLACEHOLDER
    reread = parse_zotero_field(field)
    if reread.get("citationID") != payload["citationID"] or str(field.Result.Text).strip() != PLACEHOLDER:
        raise VerificationFailed("citation field write did not read back")
    return CitationRef(payload["citationID"], PLACEHOLDER, tuple(item_keys))


def insert_bibliography_unrendered(
    document,
    *,
    donor_path: str | Path,
    paragraph_index: int,
) -> None:
    """Clone one live bibliography field at an empty target paragraph."""

    if not _has_zotero_preferences(document):
        raise RefusedUnsafe(VIRGIN_DOCUMENT_HELP)
    if _bibliography_fields(document):
        raise RefusedUnsafe("target document already contains a Zotero bibliography")
    if not 1 <= paragraph_index <= document.Paragraphs.Count:
        raise RefusedUnsafe("paragraph index is out of range")
    paragraph = document.Paragraphs(paragraph_index).Range
    if str(paragraph.Text).rstrip("\r\x07").strip():
        raise RefusedUnsafe("bibliography target paragraph must be empty")
    point = document.Range(paragraph.Start, paragraph.Start)

    donor = document.Application.Documents.Open(
        str(Path(donor_path).resolve(strict=True)), ReadOnly=True, AddToRecentFiles=False, Visible=False
    )
    try:
        donors = _bibliography_fields(donor)
        if len(donors) != 1:
            raise RefusedUnsafe(
                f"donor must contain exactly one Zotero bibliography field, found {len(donors)}"
            )
        donor_field = donors[0]
        source = donor.Range(donor_field.Code.Start - 1, donor_field.Result.End + 1)
        point.FormattedText = source.FormattedText
    finally:
        donor.Close(0)

    pasted = [
        field
        for field in point.Fields
        if int(field.Type) == WD_FIELD_ADDIN
        and str(field.Code.Text).strip().startswith("ADDIN ZOTERO_BIBL")
    ]
    if len(pasted) != 1:
        raise VerificationFailed(
            "paste range did not contain exactly one ADDIN field",
            expected=1,
            actual=len(pasted),
        )
    field = pasted[0]
    _reset_bibliography_field(field)
    reread = _bibliography_fields(document)
    if len(reread) != 1:
        raise VerificationFailed("bibliography field write did not read back")


def refresh_file_and_wait(
    path: str | Path,
    expected_citation_ids: Sequence[str],
    *,
    poll_seconds: float = 0.2,
    max_wait_seconds: float = 75.0,
    timeout: float | None = None,
) -> tuple[dict[str, str], list[str], dict[str, Any]]:
    """Drive ZoteroRefresh through a supervised separate process."""

    target = Path(path).resolve(strict=True)
    wait_timeout = timeout if timeout is not None else max_wait_seconds
    res = dispatch_refresh(
        target,
        citation_ids=tuple(expected_citation_ids),
        dispatch_timeout=wait_timeout,
        poll_interval=poll_seconds,
    )
    rendered = res.get("rendered", {})
    bibliography_results = res.get("bibliography_results", [])
    warnings = {
        k: v for k, v in res.items() if k not in ("rendered", "bibliography_results")
    }

    missing = [
        cid
        for cid in expected_citation_ids
        if cid not in rendered or rendered[cid] == PLACEHOLDER
    ]
    if missing:
        raise VerificationFailed(
            f"citations did not render: {missing}",
            expected=list(expected_citation_ids),
            actual=rendered,
        )

    return rendered, bibliography_results, warnings
