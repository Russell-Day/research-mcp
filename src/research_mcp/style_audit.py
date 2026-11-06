"""Read-only OOXML style diagnostics for manuscript documents.

This module intentionally does not repair formatting or start Word.  It reports
effective character emphasis (including ``docDefaults`` and style inheritance)
and keeps uncertain field/revision text visible to callers.
"""

from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

from .audit import is_file_locked
from .ooxml import NS, W
from .paragraphs import _paragraph_nodes, _paragraph_text, read_snapshot
from io import BytesIO
from .styling import MAJOR_HEADINGS, PROFILES, _verify_saved_typography, classify_paragraph


_FALSE = {"0", "false", "off", "no"}
_TRUE = {"1", "true", "on", "yes"}
_LABELS = ("purpose:", "background:", "methods:", "results:", "conclusions:", "objective:")
DEFAULT_WINDOW = 40
MAX_WINDOW = 200


def _toggle(node: ET.Element | None, name: str) -> bool | None:
    if node is None:
        return None
    item = node.find(f"w:{name}", NS)
    if item is None:
        return None
    value = item.get(W + "val", "true").lower()
    if value in _FALSE:
        return False
    if value in _TRUE:
        return True
    return None


def _present(node: ET.Element | None, name: str) -> tuple[bool, bool]:
    if node is None:
        return False, False
    item = node.find(f"w:{name}", NS)
    if item is None:
        return False, False
    value = item.get(W + "val", "true").lower()
    return True, value not in _FALSE


def _props(node: ET.Element | None) -> dict[str, bool | None]:
    return {key: _toggle(node, key) for key in ("b", "i")}


def _merge(*layers: dict[str, bool | None]) -> dict[str, bool | None]:
    result: dict[str, bool | None] = {"b": None, "i": None}
    for layer in layers:
        for key, value in layer.items():
            if value is not None:
                result[key] = value
    return result


def _direct(base: dict[str, bool], node: ET.Element | None) -> dict[str, bool]:
    result = dict(base)
    for key in ("b", "i"):
        present, value = _present(node, key)
        if present:
            result[key] = value
    return result


def _style_layer(base: dict[str, bool], node: ET.Element | None) -> dict[str, bool]:
    """Apply style toggle properties (a true style value toggles)."""
    result = dict(base)
    for key in ("b", "i"):
        present, value = _present(node, key)
        if present:
            result[key] = (not result[key]) if value else result[key]
    return result


def _text(node: ET.Element) -> str:
    out: list[str] = []
    for child in node.iter():
        if child.tag == W + "t":
            out.append(child.text or "")
        elif child.tag == W + "tab":
            out.append("\t")
        elif child.tag in (W + "br", W + "cr"):
            out.append("\f" if child.get(W + "type") == "page" else "\v")
    return "".join(out)


def _profile_menu() -> dict[str, dict[str, Any]]:
    purposes = {"editing_compact": "Compact working draft", "compact_single": "Alias of editing_compact", "radiology_ai": "Configured Radiology AI typography preset", "standard_double": "General 12-point double-spaced layout"}
    return {name: {"purpose": purposes[name], "effects": asdict(profile), "default_emphasis": "preserve existing bold and italic"} for name, profile in PROFILES.items()}


def _is_heading(text: str, style_name: str | None) -> bool:
    clean = text.strip().lower().rstrip(":")
    return bool(style_name and style_name.lower().startswith(("heading", "title"))) or clean in MAJOR_HEADINGS


def audit_document_style(
    path: str | Path,
    profile: str | None = None,
    *,
    first_paragraph: int | None = None,
    last_paragraph: int | None = None,
    search_text: str | None = None,
    _unbounded: bool = False,
) -> dict[str, Any]:
    """Audit effective bold/italic spans in a bounded Word-compatible range."""
    snapshot = read_snapshot(path)
    if first_paragraph is not None and first_paragraph < 1:
        raise ValueError("first_paragraph must be >= 1")
    if last_paragraph is not None and last_paragraph < 1:
        raise ValueError("last_paragraph must be >= 1")
    if first_paragraph is not None and last_paragraph is not None and last_paragraph < first_paragraph:
        raise ValueError("last_paragraph must be >= first_paragraph")
    locked, lock_reason = is_file_locked(snapshot.path)
    with zipfile.ZipFile(BytesIO(snapshot.package)) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        styles = ET.fromstring(archive.read("word/styles.xml")) if "word/styles.xml" in archive.namelist() else ET.Element(W + "styles")

    by_id = {item.get(W + "styleId"): item for item in styles.findall("w:style", NS) if item.get(W + "styleId")}
    default_rpr = styles.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
    defaults = {key: (_toggle(default_rpr, key) is True) for key in ("b", "i")}
    style_cache: dict[tuple[str, str], dict[str, bool | None]] = {}
    limitations: list[str] = ["Emphasis audit covers main-document Latin-script bold/italic; conditional table styles, numbering-label formatting, text boxes, headers and footers are outside this check."]

    def style_props(style_id: str | None, kind: str, trail: tuple[str, ...] = (), base=None) -> dict[str, bool]:
        if not style_id:
            return dict(defaults if base is None else base)
        key = (style_id, kind)
        if base is None and key in style_cache:
            return style_cache[key]
        if style_id in trail:
            limitations.append(f"cyclic style inheritance at {style_id}")
            return {"b": False, "i": False}
        style = by_id.get(style_id)
        if style is None:
            limitations.append(f"missing style {style_id}")
            return dict(defaults)
        parent = style.find("w:basedOn", NS)
        inherited = style_props(parent.get(W + "val"), kind, (*trail, style_id), base) if parent is not None else dict(defaults if base is None else base)
        result = _style_layer(inherited, style.find(f"w:{kind}", NS))
        if base is None:
            style_cache[key] = result
        return result

    nodes = _paragraph_nodes(snapshot)
    default_style = next((key for key, node in by_id.items() if node.get(W + "type") == "paragraph" and node.get(W + "default") in ("1", "true", "on")), "Normal")
    body_text_nodes = [item for item in nodes if item[1] is not None and not item[2] and _text(item[1]).strip()]
    first_body_index = body_text_nodes[0][0] if body_text_nodes else None
    requested_first, requested_last = first_paragraph, last_paragraph
    if search_text is None and not _unbounded:
        if first_paragraph is None:
            first_paragraph = 1
        if last_paragraph is None:
            last_paragraph = min(first_paragraph + DEFAULT_WINDOW - 1, first_paragraph + MAX_WINDOW - 1)
        if last_paragraph - first_paragraph + 1 > MAX_WINDOW:
            last_paragraph = first_paragraph + MAX_WINDOW - 1
    paragraphs: list[dict[str, Any]] = []
    findings: list[dict[str, Any]] = []
    truncated = False
    for index, paragraph, in_table in nodes:
        if first_paragraph is not None and index < first_paragraph:
            continue
        if last_paragraph is not None and index > last_paragraph:
            continue
        if paragraph is None:
            continue
        visible = _paragraph_text(paragraph)
        if search_text is not None and search_text.lower() not in visible.lower():
            continue
        if len(paragraphs) >= MAX_WINDOW and not _unbounded:
            truncated = True
            continue
        ppr = paragraph.find("w:pPr", NS)
        pstyle = ppr.find("w:pStyle", NS) if ppr is not None else None
        style_id = pstyle.get(W + "val") if pstyle is not None else default_style
        style_node = by_id.get(style_id)
        name = style_node.find("w:name", NS).get(W + "val") if style_node is not None and style_node.find("w:name", NS) is not None else style_id
        inherited = style_props(style_id, "rPr")
        spans: list[dict[str, Any]] = []
        has_revision = any(node.tag in (W + "ins", W + "del", W + "moveFrom", W + "moveTo") for node in paragraph.iter())
        has_opaque = any(node.tag in (W + "drawing", W + "pict", W + "object") for node in paragraph.iter())
        has_field = any(node.tag in (W + "fldChar", W + "instrText", W + "fldSimple") for node in paragraph.iter())
        for run in paragraph.iter(W + "r"):
            text = _text(run)
            if not text:
                continue
            rpr = run.find("w:rPr", NS)
            rstyle = rpr.find("w:rStyle", NS) if rpr is not None else None
            effective = _direct(style_props(rstyle.get(W + "val"), "rPr", base=inherited) if rstyle is not None else inherited, rpr)
            span = {"text": text, "bold": effective["b"], "italic": effective["i"], "resolved": effective["b"] is not None and effective["i"] is not None}
            if spans and all(spans[-1][key] == span[key] for key in ("bold", "italic", "resolved")):
                spans[-1]["text"] += text
            else:
                spans.append(span)
        if has_opaque:
            limitations.append(f"paragraph {index} includes a drawing/object; visible text is inexact")
        if has_revision:
            limitations.append(f"paragraph {index} contains tracked revisions; visible text may differ in Word")
        if has_field:
            limitations.append(f"paragraph {index} contains fields; field results are not resolved")
        paragraphs.append({"index": index, "scope": "table" if in_table else "body", "style": name, "text": visible, "text_exact": not (has_revision or has_field or has_opaque), "spans": spans})
        all_bold = bool(spans) and all(span["bold"] is True for span in spans)
        role = classify_paragraph(visible, name or "", index == first_body_index, in_table)
        paragraphs[-1]["role"] = role
        paragraphs[-1]["role_inferred"] = True
        heading = role in ("title", "heading", "summary")
        if all_bold and not in_table and visible.strip() and not heading:
            recommendation = "label-only" if visible.lstrip().lower().startswith(_LABELS) else "review whole paragraph"
            findings.append({"severity": "warning", "code": "WHOLE_BODY_PARAGRAPH_BOLD", "paragraph_index": index, "scope": "body", "message": "Entire body paragraph resolves bold; this can result from paragraph-level or inherited formatting.", "recommendation": recommendation, "text_exact": not (has_revision or has_field or has_opaque)})
        elif any(span["bold"] is True for span in spans) and not all_bold:
            findings.append({"severity": "info", "code": "PARTIAL_BOLD_SPAN", "paragraph_index": index, "scope": "table" if in_table else "body", "message": "Bold is limited to spans; inspect the label and intended emphasis before changing it.", "recommendation": "label-only", "text_exact": not (has_revision or has_field or has_opaque)})

    result: dict[str, Any] = {"path": str(snapshot.path), "source_sha256": snapshot.sha256, "is_locked": locked, "profile": profile, "profile_menu": _profile_menu(), "paragraphs": paragraphs, "findings": findings, "limitations": sorted(set(limitations)), "paragraph_index_matches_word": True, "requested_range": {"first": requested_first, "last": requested_last, "search_text": search_text}, "returned_range": {"first": paragraphs[0]["index"] if paragraphs else None, "last": paragraphs[-1]["index"] if paragraphs else None}, "range_limit": MAX_WINDOW, "truncated": truncated}
    result["total_paragraph_count"] = len(nodes)
    result["finding_scope"] = "returned_paragraphs_only"
    result["window_truncated"] = search_text is None and bool(nodes) and (last_paragraph or len(nodes)) < len(nodes)
    result["truncated"] = result["truncated"] or result["window_truncated"]
    if lock_reason:
        result["lock_reason"] = lock_reason
    if profile is not None:
        if profile not in PROFILES:
            result["typography"] = {"valid": False, "verified": False, "message": f"unknown style profile '{profile}'"}
        else:
            try:
                _verify_saved_typography(snapshot.package, PROFILES[profile])
                result["typography"] = {"valid": True, "verified": True, "message": "saved typography verifier passed; this is partial verification"}
            except Exception as exc:
                result["typography"] = {"valid": False, "verified": True, "message": str(exc), "expected": getattr(exc, "expected", None), "actual": getattr(exc, "actual", None), "note": "first failure from partial saved-typography verification; not a comprehensive journal compliance determination"}
    return result


def snapshot_emphasis(path: str | Path) -> dict[tuple[int, int], tuple[str, bool, bool]]:
    """Return compact per-run emphasis data for a verified package snapshot."""
    result = audit_document_style(path, _unbounded=True)
    output: dict[tuple[int, int], tuple[str, bool, bool]] = {}
    for item in result["paragraphs"]:
        offset = 0
        for span in item["spans"]:
            for character in span["text"]:
                offset += 1
                output[(item["index"], offset)] = (character, bool(span["bold"]), bool(span["italic"]))
    return output
