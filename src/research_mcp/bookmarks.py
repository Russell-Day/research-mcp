"""Audit and release internal Word bookmarks with strict safety invariants."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Sequence

from .audit import Audit, audit_docx
from .errors import RefusedUnsafe, VerificationFailed
from .word import IsolatedWordDocument, staged_document

_RE_SAFE = re.compile(r"^(_BM_|bm_|_RMCP_)")


@dataclass(frozen=True)
class BookmarkInfo:
    name: str
    referenced_by: tuple[str, ...] = ()


def inventory(path: str | Path) -> list[BookmarkInfo]:
    """List all bookmarks in a document with their referencing fields."""
    audit = audit_docx(path)
    return [
        BookmarkInfo(
            name=name,
            referenced_by=audit.bookmark_references.get(name, ()),
        )
        for name in audit.bookmark_names
    ]


def is_safe_to_release(name: str) -> bool:
    """Return True if bookmark is a known safe artifact prefix."""
    return bool(_RE_SAFE.match(name))


def validate_release(audit: Audit, names: Sequence[str]) -> tuple[str, ...]:
    """Verify all requested bookmark names exist and meet safety invariants."""
    requested = list(names)
    if not requested:
        raise RefusedUnsafe("no bookmarks specified for release")

    existing = set(audit.bookmark_names)
    missing = [name for name in requested if name not in existing]
    if missing:
        raise RefusedUnsafe(
            f"requested bookmarks not found in document: {missing}"
        )

    # Invariant: Bookmark must not be referenced by any field (REF, PAGEREF, hyperlink, etc.)
    referenced = audit.bookmark_references
    has_refs = [name for name in requested if name in referenced and referenced[name]]
    if has_refs:
        raise RefusedUnsafe(
            f"bookmarks are referenced by document elements: {has_refs}"
        )

    return tuple(requested)


def _verify_protected_structure(before: Audit, after: Audit) -> None:
    for field in (
        "paragraph_count",
        "table_count",
        "zotero_item_fields",
        "zotero_bibliography_fields",
        "live_comment_count",
        "drawing_count",
        "media_count",
        "live_media_count",
    ):
        if getattr(before, field) != getattr(after, field):
            raise VerificationFailed(
                f"protected invariant {field} changed during bookmark release",
                expected=getattr(before, field),
                actual=getattr(after, field),
            )


def release_to_version(
    source: str | Path,
    output: str | Path | None = None,
    *,
    names: Sequence[str] = (),
    visible: bool = False,
) -> Audit:
    """Release requested unreferenced bookmarks in a versioned copy or in-place."""
    target_output = output if output is not None else source
    before = audit_docx(source)
    valid_names = validate_release(before, names)

    with staged_document(source, target_output) as output_path:
        with IsolatedWordDocument(output_path, visible=visible) as opened:
            doc = opened.document
            for name in valid_names:
                if doc.Bookmarks.Exists(name):
                    doc.Bookmarks(name).Delete()
            opened.save()

        after = audit_docx(output_path)
        _verify_protected_structure(before, after)
        expected_remaining = len(before.bookmark_names) - len(valid_names)
        if len(after.bookmark_names) != expected_remaining:
            raise VerificationFailed(
                "bookmark count mismatch after release",
                expected=expected_remaining,
                actual=len(after.bookmark_names),
            )
        return after
