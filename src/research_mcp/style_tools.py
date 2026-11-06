"""Agent-facing scope previews and atomic publication for inline formatting."""
from pathlib import Path
import os
import tempfile

from .audit import audit_docx
from .errors import RefusedUnsafe, VerificationFailed
from .inline_style import _validate_request
from .paragraphs import read_paragraphs
from .operation_process import run_operation
from .word import staged_document


def style_text(source: str, source_sha256: str, paragraph_index: int, text: str, *, bold: bool | None = None, italic: bool | None = None, output: str | None = None, preview_only: bool = True, allow_whole_paragraph: bool = False) -> dict:
    path = Path(source).resolve(strict=True)
    audit = audit_docx(path)
    if source_sha256.upper() != audit.sha256:
        raise RefusedUnsafe("Source changed; inspect again and supply the current source_sha256")
    if audit.is_locked:
        raise RefusedUnsafe("Source is locked; save and close it, then inspect again")
    if audit.structural_errors or audit.revision_count:
        raise RefusedUnsafe("Resolve structural errors or tracked revisions before inline styling")
    if bold is None and italic is None:
        raise RefusedUnsafe("Specify bold and/or italic; omitted properties stay unchanged")
    paragraphs = read_paragraphs(path)
    if not 1 <= paragraph_index <= len(paragraphs):
        raise RefusedUnsafe("Paragraph index is outside the document; inspect again")
    target = paragraphs[paragraph_index - 1]
    if not target.text_exact:
        raise RefusedUnsafe("This paragraph has inexact text; use a verified plain-text paragraph")
    _validate_request(target.text, text, allow_whole_paragraph=allow_whole_paragraph)
    start = target.text.index(text)
    plan = {
        "status": "preview", "source": str(path), "source_sha256": audit.sha256,
        "paragraph_index": paragraph_index, "selected_text": text,
        "context_before": target.text[max(0, start - 80):start],
        "context_after": target.text[start + len(text):start + len(text) + 80],
        "scope": "whole_paragraph" if text.strip() == target.text.strip() else "exact_text",
        "requested": {key: value for key, value in {"bold": bold, "italic": italic}.items() if value is not None},
        "characters_selected": len(text), "has_fields": target.has_fields,
        "field_overlap_check": "performed in Word before mutation" if target.has_fields else "no fields in this paragraph",
        "next_step": "If this matches the authorized edit, call with preview_only=False and the same SHA. No user reconfirmation is required for an already-authorized edit.",
    }
    if preview_only:
        return plan
    destination = Path(output).resolve() if output else path
    with staged_document(path, destination) as stage:
        if audit_docx(stage).sha256 != audit.sha256:
            raise RefusedUnsafe("Source changed while staging; inspect again")
        with tempfile.TemporaryDirectory(prefix="research-mcp-inline-", dir=stage.parent) as temporary:
            candidate = Path(temporary) / "candidate.docx"
            result = run_operation("inline_style", {"source": str(stage), "output": str(candidate), "paragraph_index": paragraph_index, "text": text, "bold": bold, "italic": italic, "allow_whole_paragraph": allow_whole_paragraph})
            if not candidate.is_file() or not result.get("saved_state_verified"):
                raise VerificationFailed("Inline worker did not supply a verified saved document")
            after = audit_docx(candidate)
            os.replace(candidate, stage)
    return {**plan, "status": "complete", "output": str(destination), "output_sha256": after.sha256, "verification": {key: result[key] for key in ("saved_state_verified", "field_instructions_unchanged", "verified_properties", "verification_note") if key in result}, "next_step": "Inspect formatting with audit_document_style; render if visual layout matters."}


def plan_document_style(path: str, profile: str, emphasis_policy: str) -> dict:
    from .style_audit import audit_document_style
    from .styling import PROFILES, classify_paragraph
    from dataclasses import asdict
    if profile not in PROFILES:
        raise RefusedUnsafe("Unknown style profile")
    audited = audit_document_style(path, profile=profile)
    paragraphs = read_paragraphs(path)
    import zipfile
    import xml.etree.ElementTree as ET
    from .ooxml import W
    with zipfile.ZipFile(path) as archive:
        styles = ET.fromstring(archive.read("word/styles.xml")) if "word/styles.xml" in archive.namelist() else ET.Element(W + "styles")
    names = {}
    for style in styles.findall(W + "style"):
        name = style.find(W + "name")
        names[style.get(W + "styleId")] = name.get(W + "val", "") if name is not None else style.get(W + "styleId", "")
    first = True
    decisions = []
    counts = {}
    for paragraph in paragraphs:
        if not paragraph.text.strip():
            continue
        role = classify_paragraph(paragraph.text, names.get(paragraph.style, paragraph.style or ""), first, paragraph.in_table)
        if not paragraph.in_table:
            first = False
        counts[role] = counts.get(role, 0) + 1
        if len(decisions) < 40:
            decisions.append({"paragraph_index": paragraph.index, "text_preview": paragraph.text[:120], "role": role, "emphasis_action": "bold_entire_paragraph" if emphasis_policy == "headings_and_summary" and role in ("title", "heading", "summary") else "preserve_existing_emphasis"})
    return {"status": "preview", "source": str(Path(path).resolve()), "source_sha256": audited["source_sha256"], "is_locked": audited["is_locked"], "profile": profile, "profile_settings": asdict(PROFILES[profile]), "emphasis_policy": emphasis_policy, "role_counts": counts, "decisions": decisions, "decisions_truncated": sum(counts.values()) > len(decisions), "findings": audited["findings"], "scope": "whole_document_layout", "limitations": ["First nonempty body paragraph is treated as a title; inspect this decision before applying.", "Typography presets do not verify full journal or anonymization compliance."], "next_step": "If these decisions match the authorized task, call apply_document_style with preview_only=False. Use style_text for phrase-only emphasis."}
