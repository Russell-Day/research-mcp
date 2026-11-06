"""Safe, exact-range inline emphasis for Word documents.

The operation in this module deliberately assigns only ``Range.Font.Bold`` or
``Range.Font.Italic``.  Assigning a paragraph range's Font is an easy way to
accidentally bold an entire paragraph, so paragraph-sized anchors are refused
unless the caller explicitly opts in.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .audit import audit_docx
from . import style_audit
from .editing import (
    FIND_LIMIT,
    _field_spans,
    _refuse_field_overlap,
    _refuse_unsafe_candidate,
    locate_anchor,
)
from .errors import RefusedUnsafe, VerificationFailed
from .paragraphs import read_paragraphs
from .word import IsolatedWordDocument, create_versioned_copy


_PARAGRAPH_MARKS = {"\r", "\x07"}
def _offline_formatting_snapshot(path: str | Path) -> dict[tuple[int, int], tuple[str, bool | None, bool | None]]:
    """Return the unbounded shared OOXML effective emphasis snapshot."""
    records = style_audit.snapshot_emphasis(path)
    if isinstance(records, dict):
        return {tuple(key): tuple(value) for key, value in records.items()}
    result: dict[tuple[int, int], tuple[str, bool | None, bool | None]] = {}
    offsets: dict[int, int] = {}
    for record in records:
        paragraph_index = int(record["paragraph_index"])
        offset = offsets.get(paragraph_index, 0)
        for character in str(record["text"]):
            offset += 1
            result[(paragraph_index, offset)] = (
                character, record.get("bold"), record.get("italic")
            )
        offsets[paragraph_index] = offset
    return result


def _field_snapshot(document) -> tuple[tuple[int, str, str], ...]:
    """Return field instructions/results so a style edit cannot rewrite fields."""

    fields = []
    for index in range(1, int(document.Fields.Count) + 1):
        field = document.Fields(index)
        fields.append((int(field.Type), str(field.Code.Text), str(field.Result.Text)))
    return tuple(fields)


def _verify_formatting(before, after, allowed_keys: set[tuple[int, int]], *, bold, italic) -> None:
    if not allowed_keys.issubset(before) or not allowed_keys.issubset(after):
        missing = sorted(allowed_keys - set(before)) + sorted(allowed_keys - set(after))
        raise VerificationFailed("matched visible characters are missing from the style snapshot", actual=missing)
    if set(before) != set(after):
        raise VerificationFailed("character layout changed during inline styling")
    changed = []
    for key in before:
        before_text, *before_font = before[key]
        after_text, *after_font = after[key]
        if before_text != after_text:
            raise VerificationFailed("text changed during inline styling", expected=before_text, actual=after_text)
        if key not in allowed_keys and before_font != after_font:
            changed.append((key, before_font, after_font))
            continue
        if key in allowed_keys:
            for pos, (old, new) in enumerate(zip(before_font, after_font)):
                allowed = (pos == 0 and bold is not None) or (pos == 1 and italic is not None)
                if old != new and not allowed:
                    changed.append((key, before_font, after_font))
                    break
                if pos == 0 and bold is not None and new is not bool(bold):
                    changed.append((key, before_font, after_font))
                    break
                if pos == 1 and italic is not None and new is not bool(italic):
                    changed.append((key, before_font, after_font))
                    break
    if changed:
        raise VerificationFailed(
            "inline styling changed formatting outside the requested properties",
            expected="only Bold/Italic in the matched range",
            actual=changed[:3],
        )


def _verify_protected(before, after) -> None:
    if after.structural_errors:
        raise VerificationFailed("saved document has structural errors", actual=after.structural_errors)
    if before.paragraph_count != after.paragraph_count:
        raise VerificationFailed("paragraph count changed during inline styling", expected=before.paragraph_count, actual=after.paragraph_count)
    for name in (
        "table_count", "tables", "bookmark_names", "field_count",
        "zotero_item_fields", "zotero_bibliography_fields", "revision_count",
        "comment_count", "live_comment_count", "drawing_count", "media_count",
        "live_media_count", "dangling_media_targets",
    ):
        if getattr(before, name) != getattr(after, name):
            raise VerificationFailed(
                f"{name} changed during inline styling",
                expected=getattr(before, name), actual=getattr(after, name),
            )
    if before.body_text_sha256 != after.body_text_sha256 or before.prose_text_sha256 != after.prose_text_sha256:
        raise VerificationFailed("document text changed during inline styling")


def _validate_request(paragraph_text: str, text: str, *, allow_whole_paragraph: bool) -> int:
    if not text:
        raise RefusedUnsafe("text must not be empty")
    if len(text) > FIND_LIMIT:
        raise RefusedUnsafe(f"text is {len(text)} characters; Word's search cannot exceed {FIND_LIMIT}")
    count = paragraph_text.count(text)
    if count != 1:
        raise RefusedUnsafe(f"text must occur exactly once in the paragraph; it occurs {count} times")
    visible = paragraph_text.rstrip("\r\x07")
    if text.strip() == visible.strip() and not allow_whole_paragraph:
        raise RefusedUnsafe("whole-paragraph styling is refused; set allow_whole_paragraph=True explicitly")
    return count


def style_exact_text(
    source: str | Path,
    output: str | Path,
    paragraph_index: int,
    text: str,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    allow_whole_paragraph: bool = False,
    visible: bool = False,
) -> dict[str, Any]:
    """Apply exact inline bold/italic formatting to one unique text occurrence.

    ``output`` is always a newly staged copy; an existing output is refused by
    ``create_versioned_copy``.  The saved copy is reopened read-only before the
    operation returns and its text, formatting, and field instructions are
    verified against the live pre-edit snapshot.
    """

    if bold is None and italic is None:
        raise RefusedUnsafe("specify bold and/or italic")
    paragraphs = read_paragraphs(source)
    if not 1 <= paragraph_index <= len(paragraphs):
        raise RefusedUnsafe(f"paragraph_index {paragraph_index} is outside 1..{len(paragraphs)}")
    target = paragraphs[paragraph_index - 1]
    if not target.text_exact:
        raise RefusedUnsafe(f"paragraph {paragraph_index} contains a drawing and cannot be styled safely")
    _validate_request(target.text, text, allow_whole_paragraph=allow_whole_paragraph)
    before_audit = audit_docx(source)
    if before_audit.revision_count:
        raise RefusedUnsafe("tracked revisions are not supported for inline styling")
    before_formatting = _offline_formatting_snapshot(source)
    _, output_path = create_versioned_copy(source, output)

    with IsolatedWordDocument(output_path, visible=visible) as session:
        document = session.document
        candidate = locate_anchor(document, paragraph_index, text)
        if str(candidate.Text) != text:
            raise RefusedUnsafe("Word resolved a range whose text differs from the requested anchor")
        paragraph_range = document.Paragraphs(paragraph_index).Range
        if int(candidate.Start) < int(paragraph_range.Start) or int(candidate.End) > int(paragraph_range.End) - 1:
            raise RefusedUnsafe("matched range includes a paragraph mark")
        _refuse_field_overlap(_field_spans(document, paragraph_index), int(candidate.Start), int(candidate.End))
        _refuse_unsafe_candidate(document, candidate)
        before_fields = _field_snapshot(document)
        allowed_keys = set()
        target_start = target.text.rstrip("\r\x07").find(text)
        for offset in range(target_start, target_start + len(text)):
            allowed_keys.add((paragraph_index, offset + 1))
        if target_start < 0 or not allowed_keys:
            raise RefusedUnsafe("matched range contains no editable characters")
        if bold is not None:
            candidate.Font.Bold = bool(bold)
        if italic is not None:
            candidate.Font.Italic = bool(italic)
        session.save()

    # Reopen the saved copy read-only, as COM state immediately before Save is
    # insufficient evidence that Word persisted the requested formatting.
    with IsolatedWordDocument(output_path, visible=False, read_only=True) as session:
        saved_document = session.document
        if _field_snapshot(saved_document) != before_fields:
            raise VerificationFailed("field instructions changed during inline styling")

    after_formatting = _offline_formatting_snapshot(output_path)
    _verify_formatting(before_formatting, after_formatting, allowed_keys, bold=bold, italic=italic)

    after_audit = audit_docx(output_path)
    _verify_protected(before_audit, after_audit)
    return {
        "source": str(Path(source).resolve()),
        "output": str(Path(output_path).resolve()),
        "operation": "style_exact_text",
        "paragraph_index": paragraph_index,
        "text": text,
        "occurrence_count": 1,
        "scope": {"paragraph_index": paragraph_index, "characters": len(allowed_keys)},
        "requested": {"bold": bold, "italic": italic},
        "allow_whole_paragraph": allow_whole_paragraph,
        "source_sha256": before_audit.sha256,
        "output_sha256": after_audit.sha256,
        "saved_state_verified": True,
        "field_instructions_unchanged": True,
        "verified_properties": ["bold", "italic"],
        "verification_note": "OOXML effective style audit verifies visible bold/italic per character; protected structure and text hashes are unchanged.",
    }
