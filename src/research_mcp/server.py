"""Thin MCP 2.0 boundary over the tested research_mcp core."""

from __future__ import annotations

import argparse
from dataclasses import asdict
import os
from pathlib import Path
from typing import Any, Literal

from mcp.server import MCPServer
from mcp.types import ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field, ValidationError

from .audit import audit_docx
from .bookmarks import inventory, release_to_version
from .citation_ops import insert_citation_to_version
from .cluster_ops import edit_cluster_to_version
from .refresh_ops import refresh_to_version
from .edit_batch import edit_document as _edit_document
from .edit_models import (
    AddCommentOperation,
    EditBatchResult,
    EditDocumentRequest,
    InsertParagraphOperation,
    InsertTableOperation,
    InsertTextOperation,
    ReplaceParagraphOperation,
    ReplaceSectionOperation,
    AcceptRevisionsOperation,
    ListRevisionsOperation,
    SweepCharactersOperation,
    RejectRevisionsOperation,
    TrackedReplaceParagraphOperation,
    TrackedReplaceTextOperation,
    ReplaceTableRowOperation,
    ReplaceTextOperation,
    ReplyCommentOperation,
    ResolveCommentOperation,
    ReopenCommentOperation,
    DeleteCommentOperation,
    SHA256,
)
from .errors import RefusedUnsafe
from .paragraphs import inspect_document as _inspect_document
from .revisions import list_revisions as _list_revisions
from . import publication
import time
from . import config, textclean
from .logger import get_log_dir, log_tool_call, record_agent_issue
from .schema_models import (
    AddZoteroItemResult,
    ApplyStyleResult,
    AuditDocumentResult,
    CheckpointDocumentResult,
    CleanBookmarksResult,
    DocumentHistoryResult,
    EditCitationClusterResult,
    RefreshCitationsResult,
    FormatTablesResult,
    InsertCitationResult,
    InspectDocumentResult,
    ListBookmarksResult,
    ListCollectionsResult,
    MCPMergeSpec,
    MCPTableFormatSpec,
    RenderDocumentResult,
    ReportIssueResult,
    RevertDocumentResult,
    SearchZoteroResult,
    SetProjectCollectionResult,
    UndoDocumentResult,
    UpdateZoteroItemResult,
)
from .styling import apply_journal_style
from .rendering import render_docx_pages
from .tables import MergeGroup, TableFormatSpec, format_to_version
from .word import staged_document
from .zotero import (
    ZoteroClient,
    add_item_by_doi,
    find_project_collection_details,
    update_item,
    write_project_collection,
)

SERVER_INSTRUCTIONS = """Safe read-before-write research document workflow:
1. Always inspect or audit a document before writing: use `inspect_document` or `audit_document` to inspect structure and get its current `sha256`. Check `is_locked` to verify the file is not currently open in Word.
2. Provide `source_sha256` when calling mutation tools (`edit_document`, `format_tables`, `clean_bookmarks`, `insert_zotero_citation`, `apply_document_style`) to guarantee you are modifying the exact version inspected.
3. Editing Cadence & Versioning Policy:
   - Routine edits default to IN-PLACE editing: pass `output == source`. Changes are staged, applied operations are verified, and the verified pre-edit state is stored in the undo store (outside OneDrive, under %LOCALAPPDATA%\\research_mcp\\checkpoints) before atomic replacement. No `.docx.bak` file is written. A store failure, a full store or a changed source prevents replacement. If Word holds the file open, save and close it, then inspect again.
   - Every write result returns `undo_token`, `sha256_before` and `sha256_after` (`undo_token` is null for a versioned write, where the source is untouched).
   - Undo in one call: `undo_document(path)` restores the state before the last MCP edit or batch; `undo_document(path, undo_token=...)` restores the state before that edit (and so also undoes every later one); `checkpoint_document(path, label)` then `undo_document(path, to=label)` returns to a named point. An undo is itself recorded, so `undo_document(path)` again reverses it. `document_history(path)` lists edits, checkpoints, undos, outside changes and pruned entries. `revert_document(path)` is the same as `undo_document(path)`.
   - Undo refuses when the document changed since the MCP last recorded it (the user, Word or OneDrive edited it); the refusal names both shas. `force=True` restores anyway after checkpointing the current state; `output=<new .docx path>` writes the earlier state as a new file and leaves the document alone. Undo also refuses while Word holds the file.
   - Undo limits: whole-file only (undoing edit 3 of 5 also undoes 4 and 5); tracks MCP edits only (hand edits are detected, not stored); it is not Word's Ctrl+Z; the store is on this PC only and is not synced (OneDrive version history is the outer fallback); 30 entries per document, 14 days and 1 GB in total, oldest dropped first; linked (not embedded) files are not restored; a legacy `.docx.bak` is used only when a document has no recorded edits, and is never moved or deleted.
   - Edit batches are all-or-nothing by default; allow_partial=True explicitly enables independent successes, but never splits a group: ops sharing `group`, the same table row or the same replace_section succeed or roll back together (result `groups`). rolled_back operations were never published.
   - Whole-line rewrites are safe: replace_text / replace_paragraph change only the characters that differ, so bold/italic runs in a résumé title line survive. Use replace_table_row (row found by a cell's text) or replace_section (heading as displayed, e.g. "WORK EXPERIENCE") to rewrite an entry's title, description lines and dates together; never use allow_partial to apply part of one entry.
   - Only create a new version file (e.g. `output="paper_v43.docx"`) in two scenarios:
     a) The user explicitly requested a new version (e.g., "save as v43", "bump version", "create new version").
     b) If you believe a major structural milestone or large rewrite warrants a new version, you MUST ask the user for explicit permission first before creating a new file. If the user didn't ask and hasn't given explicit permission, edit in place (`output == source`).
   - Overwriting a pre-existing file that is distinct from `source` is refused.
4. Locating content: Use `inspect_document(path, search_text="query")` to search across all paragraphs and find exact paragraph indices, styles, and text snippets without paging.
5. Document Styling and Visual Proofing:
   - Start with audit_document_style to inspect effective emphasis, suspicious whole-paragraph bold, profile choices and saved typography. Findings are not automatic editorial decisions.
   - Use style_text for exact phrase bold/italic. Preview the selected text, then apply with preview_only=False under existing user authorization. Never use a whole-paragraph font assignment for a phrase request.
   - Use `apply_document_style` to enforce target journal typography (e.g. `profile="radiology_ai"` for 11pt Times New Roman, 1-inch margins, double spacing, configured paragraph roles, and stripped page numbers).
   - Use `render_document_pages` to export high-resolution PNG page images (clean print layout without comment markup balloon gutters) to visually verify tables, figures, and formatting.
6. When citing or searching references:
   - Use `search_zotero` or `add_zotero_item_by_doi`.
   - A wrong reference (title, venue, volume, pages...) is fixed in Zotero with `update_zotero_item`, then the document is refreshed; never edit the reference-list text (a refresh overwrites it).
   - To change which references an existing citation cites (add, remove, reorder), call `edit_citation_cluster` (pick the cluster by `citation_id` or its rendered `current_text` in `paragraph_index`); it works on documents with tracked changes and refreshes through Zotero.
   - To re-render every citation and the bibliography (after fixing items in Zotero, or to check for hand edits), call `refresh_citations` (`preview=True` first writes nothing); it works on documents with tracked changes. A bibliography entry it lists as changed that you did not change in Zotero was probably edited by hand: fix the item with `update_zotero_item` and refresh; never edit the reference-list text.
   - Call `insert_zotero_citation`: `item_donor` is optional in python core and defaults to the bundled donor document.
   - Verify Zotero Desktop is running (port 23119).
   - If `audit_document` reports `citation_readiness.status == "virgin_document"`, open the document in Word, insert one citation manually, save, and close Word. Then rerun `audit_document`, confirm readiness, and retry with the fresh `sha256`. If status is `structurally_invalid`, repair the package before opening it in Word or attempting citation operations.
7. If you encounter any unexpected issues, formatting quirks, Word lock contention, or agent friction during execution, call `report_issue` to log a detailed report to logs/ISSUES.md for developer improvement.
8. Text you write is cleaned of invisible and unusual characters (results list them under `cleaned`; pass `keep` to hold one). Equations, fields and Zotero citations are never touched. Sweep an existing document only when asked (`sweep_characters`, tracked). If a character was mapped wrongly, call `report_issue(title=..., details=<why>, category="char_mapping", char=..., should_be=...)` and follow what it returns (give the user its `issue_url`, or follow its `file_with` steps when it has them).
"""

server = MCPServer(
    "research-mcp",
    title="Research Word and Zotero MCP",
    description="Safety-enforcing, versioned Word formatting and Zotero operations.",
    version="0.1.0",
    instructions=SERVER_INSTRUCTIONS,
)

_orig_call_tool = server._tool_manager.call_tool


async def _logged_call_tool(
    name: str,
    arguments: dict[str, Any],
    context: Any,
    convert_result: bool = False,
) -> Any:
    start = time.perf_counter()
    error = None
    result = None
    try:
        result = await _orig_call_tool(
            name, arguments, context, convert_result=convert_result
        )
        return result
    except Exception as exc:
        error = exc
        raise
    finally:
        duration_ms = (time.perf_counter() - start) * 1000.0
        log_tool_call(
            tool_name=name,
            arguments=arguments,
            result=result,
            duration_ms=duration_ms,
            error=error,
        )


server._tool_manager.call_tool = _logged_call_tool

ROOT_DIR: Path | None = None


def set_root_dir(path: str | Path | None) -> None:
    global ROOT_DIR
    if path is None:
        ROOT_DIR = None
    else:
        ROOT_DIR = Path(path).resolve()


def validate_path_in_root(path_str: str | Path) -> Path:
    target = Path(path_str).resolve()
    if ROOT_DIR is not None:
        try:
            target.relative_to(ROOT_DIR)
        except ValueError:
            raise RefusedUnsafe(
                f"path {path_str!r} is outside the allowed root directory {ROOT_DIR}"
            )
    return target


if config.get("RESEARCH_MCP_ROOT_DIR"):
    set_root_dir(config.get("RESEARCH_MCP_ROOT_DIR"))


# Flat envelope. Not a style choice -- a llama.cpp grammar limit.
class MCPEditOperation(BaseModel):
    model_config = ConfigDict(extra="forbid")
    op_id: str
    type: Literal[
        "replace_text",
        "insert_text",
        "insert_paragraph",
        "insert_table",
        "reply_comment",
        "add_comment",
        "resolve_comment",
        "reopen_comment",
        "delete_comment",
        "replace_paragraph",
        "replace_table_row",
        "replace_section",
        "tracked_replace_text",
        "tracked_replace_paragraph",
        "accept_revisions",
        "reject_revisions",
        "list_revisions",
        "sweep_characters",
    ]
    paragraph_index: int = 1
    find: str = Field(default="", max_length=255, description="Exact visible anchor; Word Find supports at most 255 characters.")
    replace: str = ""
    match_case: bool = True
    occurrence: int = 1
    expected_count: int = 1
    position: Literal["before", "after"] = "before"
    anchor: str = Field(default="", max_length=255, description="Exact visible anchor; Word Find supports at most 255 characters.")
    after_paragraph_index: int = 0
    text: str = ""
    style: str | None = None
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    descriptive_columns: list[int] = Field(default_factory=list)
    merges: list[dict[str, Any]] = Field(default_factory=list)
    comment_id: str = Field(default="", description="reply/resolve/reopen/delete_comment: the thread's comment_id from inspect_document (durable id); legacy:N or N also works when unambiguous.")
    disposition: Literal["addressed", "rejected", "no_change", "needs_information", "accepted"] = Field(default="addressed", description="reply_comment: addressed (needs depends_on: the op_id of the edit), rejected, no_change, needs_information. 'accepted' is an alias of addressed.")
    reply_text: str = ""
    resolve: bool = Field(default=False, description="reply_comment: also mark the thread resolved in the same write.")
    include_replies: bool = Field(default=True, description="delete_comment: true deletes the thread with its replies; false deletes only the given reply.")
    author: str = Field(default="", description="Author. Comments default to RESEARCH_MCP_AUTHOR from the config, else Word's user name, else \"Author\"; tracked edits default to Claude unless given.")
    initials: str = Field(default="", description="Comment initials (default: the configured or Word initials for the default author, else the author's initials).")
    depends_on: list[str] = Field(default_factory=list)
    heading: str = Field(default="", description="replace_section: the section heading as displayed (caps headings match in any case).")
    until_heading: str | None = Field(default=None, description="replace_section: optional heading that ends the section.")
    paragraphs: list[str] = Field(default_factory=list, description="replace_section on body paragraphs: the new paragraphs, one string each.")
    match: str = Field(default="", description="replace_table_row: text of one cell (or one line of a cell) that picks out the row.")
    cells: list[str] = Field(default_factory=list, description='replace_table_row: new text per cell; "\n" separates paragraphs in a cell.')
    table_index: int | None = Field(default=None, ge=1, description="replace_table_row: optional 1-based table number to search.")
    group: str | None = Field(default=None, description="Operations sharing a group are all-or-nothing together under allow_partial.")
    ids: list[int] | None = Field(default=None, description="accept/reject/list_revisions: w:id values of the revisions.")
    through_paragraph_index: int | None = Field(default=None, ge=1, description="accept/reject/list_revisions: last paragraph of the range starting at paragraph_index.")
    limit: int = Field(default=200, ge=1, le=2000, description="list_revisions: most revisions to return.")
    keep: list[str] = Field(default_factory=list, description='Write ops and sweep_characters: characters to leave as written (each a character or "U+XXXX"); everything else written is cleaned with textclean and listed in the result\'s `cleaned`.')
    preview_only: bool = Field(default=False, description="sweep_characters: report what would change; write nothing.")


READ_ONLY = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)

VERSIONED_WRITE = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)

OPEN_WORLD_READ = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=True,
)

OPEN_WORLD_WRITE = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=True,
)

UNDO_WRITE = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=False,
)

CHECKPOINT_WRITE = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=False,
)

DESTROY_FILE = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=True,
    idempotentHint=False,
    openWorldHint=True,
)


_OPERATION_MODELS = {
    "add_comment": AddCommentOperation,
    "replace_text": ReplaceTextOperation,
    "insert_text": InsertTextOperation,
    "insert_paragraph": InsertParagraphOperation,
    "insert_table": InsertTableOperation,
    "reply_comment": ReplyCommentOperation,
    "resolve_comment": ResolveCommentOperation,
    "reopen_comment": ReopenCommentOperation,
    "delete_comment": DeleteCommentOperation,
    "replace_paragraph": ReplaceParagraphOperation,
    "replace_table_row": ReplaceTableRowOperation,
    "replace_section": ReplaceSectionOperation,
    "tracked_replace_text": TrackedReplaceTextOperation,
    "tracked_replace_paragraph": TrackedReplaceParagraphOperation,
    "accept_revisions": AcceptRevisionsOperation,
    "reject_revisions": RejectRevisionsOperation,
    "list_revisions": ListRevisionsOperation,
    "sweep_characters": SweepCharactersOperation,
}
_BASE_FIELDS = {"op_id", "type", "depends_on", "group"}


def _field_contract(op_type: str) -> tuple[tuple[str, ...], tuple[str, ...]]:
    model = _OPERATION_MODELS.get(op_type)
    if model is None:
        return ((), ())
    required: list[str] = []
    optional: list[str] = []
    for name, field in model.model_fields.items():
        if name in _BASE_FIELDS:
            continue
        if field.is_required():
            required.append(name)
        else:
            optional.append(name)
    return (tuple(required), tuple(optional))


def _strict_operation_payload(op: MCPEditOperation):
    base = {"op_id": op.op_id, "type": op.type, "depends_on": op.depends_on, "group": op.group or None}
    given = op.model_fields_set
    model = _OPERATION_MODELS.get(op.type)
    if model is not None and "keep" in model.model_fields and op.keep:
        base["keep"] = list(op.keep)
    if op.type == "sweep_characters":
        kwargs = {**base, "preview_only": op.preview_only, "through_paragraph_index": op.through_paragraph_index}
        if "paragraph_index" in given or op.through_paragraph_index is not None:
            kwargs["paragraph_index"] = op.paragraph_index
        if "author" in given:
            kwargs["author"] = op.author
        return SweepCharactersOperation(**kwargs)
    if op.type in ("tracked_replace_text", "tracked_replace_paragraph"):
        # The envelope's author default belongs to comments; a tracked edit defaults to "Claude".
        kwargs = {**base, "paragraph_index": op.paragraph_index}
        if "author" in given:
            kwargs["author"] = op.author
        if op.type == "tracked_replace_text":
            return TrackedReplaceTextOperation(**kwargs, find=op.find, replace=op.replace)
        return TrackedReplaceParagraphOperation(**kwargs, text=op.text)
    if op.type in ("accept_revisions", "reject_revisions", "list_revisions"):
        # Only selectors the caller gave count; envelope defaults select nothing.
        kwargs = {**base, "ids": op.ids or None, "through_paragraph_index": op.through_paragraph_index,
                  "heading": op.heading or None, "until_heading": op.until_heading or None}
        if "paragraph_index" in given or op.through_paragraph_index is not None:
            kwargs["paragraph_index"] = op.paragraph_index
        if "author" in given:
            kwargs["author"] = op.author
        if op.type == "list_revisions":
            return ListRevisionsOperation(**kwargs, limit=op.limit)
        model = AcceptRevisionsOperation if op.type == "accept_revisions" else RejectRevisionsOperation
        return model(**kwargs)
    if op.type == "replace_paragraph":
        return ReplaceParagraphOperation(**base, paragraph_index=op.paragraph_index, text=op.text)
    if op.type == "replace_table_row":
        kwargs: dict[str, Any] = {**base, "cells": op.cells, "table_index": op.table_index}
        if op.match:
            kwargs["match"] = op.match
        return ReplaceTableRowOperation(**kwargs)
    if op.type == "replace_section":
        kwargs = {**base, "until_heading": op.until_heading or None}
        if op.heading:
            kwargs["heading"] = op.heading
        if op.rows:
            kwargs["rows"] = op.rows
        if op.paragraphs:
            kwargs["paragraphs"] = op.paragraphs
        return ReplaceSectionOperation(**kwargs)
    if op.type == "add_comment":
        return AddCommentOperation(
            **base, paragraph_index=op.paragraph_index, anchor=op.anchor,
            text=op.text, author=op.author or None, initials=op.initials or None,
        )
    if op.type in ("resolve_comment", "reopen_comment"):
        model = ResolveCommentOperation if op.type == "resolve_comment" else ReopenCommentOperation
        return model(**base, comment_id=op.comment_id)
    if op.type == "delete_comment":
        return DeleteCommentOperation(**base, comment_id=op.comment_id, include_replies=op.include_replies)
    if op.type == "replace_text":
        return ReplaceTextOperation(
            **base,
            paragraph_index=op.paragraph_index,
            find=op.find,
            replace=op.replace,
            match_case=op.match_case,
            occurrence=op.occurrence,
            expected_count=op.expected_count,
        )
    if op.type == "insert_text":
        kwargs: dict[str, Any] = {
            **base,
            "paragraph_index": op.paragraph_index,
            "position": op.position,
        }
        if op.anchor:
            kwargs["anchor"] = op.anchor
        if op.text:
            kwargs["text"] = op.text
        return InsertTextOperation(**kwargs)
    if op.type == "insert_paragraph":
        kwargs: dict[str, Any] = {
            **base,
            "after_paragraph_index": op.after_paragraph_index,
            "style": op.style,
        }
        if op.text:
            kwargs["text"] = op.text
        return InsertParagraphOperation(**kwargs)
    if op.type == "insert_table":
        kwargs: dict[str, Any] = {
            **base,
            "after_paragraph_index": op.after_paragraph_index,
            "headers": op.headers,
            "rows": op.rows,
            "style": op.style,
        }
        if op.descriptive_columns:
            kwargs["descriptive_columns"] = op.descriptive_columns
        if op.merges:
            kwargs["merges"] = op.merges
        return InsertTableOperation(**kwargs)
    if op.type == "reply_comment":
        kwargs: dict[str, Any] = {
            **base,
            "comment_id": op.comment_id,
            "disposition": "addressed" if op.disposition == "accepted" else op.disposition,
            "resolve": op.resolve,
            "author": op.author or None,
            "initials": op.initials or None,
        }
        if op.reply_text:
            kwargs["reply_text"] = op.reply_text
        return ReplyCommentOperation(**kwargs)
    raise ValueError(f"unsupported operation type: {op.type}")


def _describe_operation_error(
    operations: list[MCPEditOperation], exc: ValidationError
) -> str:
    lines = []
    for error in exc.errors():
        loc = error.get("loc", ())
        message = error.get("msg", "validation error")
        if len(loc) >= 1 and isinstance(loc[0], int) and loc[0] < len(operations):
            op = operations[loc[0]]
            field = loc[-1] if len(loc) > 1 else "operation"
            required, optional = _field_contract(op.type)
            msg_str = (
                f"operation {op.op_id!r} ({op.type}) invalid: field {field!r} {message}."
                f" {op.type} requires: {', '.join(required)}"
            )
            if optional:
                msg_str += f". Optional: {', '.join(optional)}"
            lines.append(msg_str)
        else:
            lines.append(message)
    return " | ".join(lines)


@server.tool(annotations=READ_ONLY, structured_output=True)
def audit_document(path: str) -> AuditDocumentResult:
    """Audit a DOCX package without opening or changing it in Word."""
    validate_path_in_root(path)
    return audit_docx(path).to_agent_dict()


@server.tool(annotations=READ_ONLY, structured_output=True)
def inspect_document(
    path: str,
    first_paragraph: int | None = None,
    last_paragraph: int | None = None,
    include_comments: bool = False,
    search_text: str | None = None,
) -> InspectDocumentResult:
    """Read a document's headings, search paragraphs, or read one bounded window of text."""
    validate_path_in_root(path)
    return _inspect_document(
        path,
        first_paragraph=first_paragraph,
        last_paragraph=last_paragraph,
        include_comments=include_comments,
        search_text=search_text,
    )


@server.tool(annotations=READ_ONLY, structured_output=True)
def list_bookmarks(path: str) -> ListBookmarksResult:
    """List every bookmark in a document and every field referencing it."""
    validate_path_in_root(path)
    return {"bookmarks": [asdict(item) for item in inventory(path)]}


@server.tool(annotations=READ_ONLY, structured_output=True)
def list_revisions(
    path: str,
    ids: list[int] | None = None,
    author: str | None = None,
    paragraph_index: int | None = None,
    through_paragraph_index: int | None = None,
    heading: str | None = None,
    until_heading: str | None = None,
    limit: int = 200,
) -> dict[str, Any]:
    """List a document's tracked changes: id, type, author, date, paragraph_index and text,
    plus each paragraph's accepted and original text. Read only. Optional selectors (all
    given must match): ids, author, paragraph_index (+ through_paragraph_index), heading
    (+ until_heading). Accept or reject them with edit_document's accept_revisions /
    reject_revisions operations."""
    validate_path_in_root(path)
    return _list_revisions(
        path,
        ids=[str(i) for i in ids] if ids else None,
        author=author,
        paragraph_index=paragraph_index,
        through_paragraph_index=through_paragraph_index,
        heading=heading,
        until_heading=until_heading,
        limit=limit,
    )


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def clean_bookmarks(
    source: str,
    output: str,
    exact_names: list[str],
    source_sha256: SHA256 | None = None,
) -> CleanBookmarksResult:
    """Delete only the supplied exact bookmark names at the requested output path; output may equal source for an in-place edit."""
    validate_path_in_root(source)
    validate_path_in_root(output)
    _refuse_tracked_revisions(source, "clean_bookmarks")
    if source_sha256 is not None:
        actual_sha = audit_docx(source).sha256
        if source_sha256.upper() != actual_sha.upper():
            raise RefusedUnsafe(
                f"source document SHA-256 mismatch: expected {source_sha256.upper()}, "
                f"actual {actual_sha.upper()}. Reinspect the document with inspect_document."
            )
    before = audit_docx(source)
    from .operation_process import run_operation

    with publication.capture(tool="clean_bookmarks", op=f"remove {len(exact_names)} bookmark(s)") as published:
        with staged_document(source, output) as stage:
            run_operation(
                "bookmarks",
                {"source": str(stage), "output": str(stage), "names": list(exact_names)},
            )
            after = audit_docx(stage)
    return {
        "source": str(Path(source).resolve()),
        "source_sha256": before.sha256,
        "output": str(Path(output).resolve()),
        "output_sha256": after.sha256,
        "removed": list(exact_names),
        "retained": list(after.bookmark_names),
        "protected_structure_unchanged": True,
        **publication.write_fields(published, before.sha256, after.sha256),
    }


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def format_tables(
    source: str,
    output: str,
    table_specs: list[MCPTableFormatSpec],
    source_sha256: SHA256 | None = None,
) -> FormatTablesResult:
    """Apply the fixed journal-neutral table policy at the requested output path; output may equal source for an in-place edit."""
    validate_path_in_root(source)
    validate_path_in_root(output)
    _refuse_tracked_revisions(source, "format_tables")
    if source_sha256 is not None:
        actual_sha = audit_docx(source).sha256
        if source_sha256.upper() != actual_sha.upper():
            raise RefusedUnsafe(
                f"source document SHA-256 mismatch: expected {source_sha256.upper()}, "
                f"actual {actual_sha.upper()}. Reinspect the document with inspect_document."
            )
    specs = tuple(
        TableFormatSpec(
            index=int(raw.index if isinstance(raw, MCPTableFormatSpec) else raw["index"]),
            descriptive_columns=tuple(
                int(value)
                for value in (
                    raw.descriptive_columns
                    if isinstance(raw, MCPTableFormatSpec)
                    else raw["descriptive_columns"]
                )
            ),
            merges=tuple(
                MergeGroup(
                    column=m.column if isinstance(m, MCPMergeSpec) else m["column"],
                    start_row=m.start_row if isinstance(m, MCPMergeSpec) else m["start_row"],
                    end_row=m.end_row if isinstance(m, MCPMergeSpec) else m["end_row"],
                    label=m.label if isinstance(m, MCPMergeSpec) else m.get("label", ""),
                )
                for m in (raw.merges if isinstance(raw, MCPTableFormatSpec) else raw.get("merges", []))
            ),
        )
        for raw in table_specs
    )
    before = audit_docx(source)
    from .operation_process import run_operation

    with publication.capture(tool="format_tables", op=f"format {len(specs)} table(s)") as published:
        with staged_document(source, output) as stage:
            run_operation(
                "tables",
                {
                    "source": str(stage),
                    "output": str(stage),
                    "specs": [asdict(spec) for spec in specs],
                },
            )
            after = audit_docx(stage)
    return {
        "source": str(Path(source).resolve()),
        "source_sha256": before.sha256,
        "output": str(Path(output).resolve()),
        "output_sha256": after.sha256,
        "formatted_tables": [spec.index for spec in specs],
        "protected_structure_unchanged": True,
        **publication.write_fields(published, before.sha256, after.sha256),
    }


@server.tool(annotations=READ_ONLY, structured_output=True)
def search_zotero(
    query: str,
    limit: int = 10,
    path: str | None = None,
    collection: str | None = None,
) -> SearchZoteroResult:
    """Search the local Zotero library and return compact candidate metadata.

    Scoped to one collection (folder) if `collection` is given, or a
    `.zotero-collection` marker is found walking up from `path`'s folder;
    otherwise searches the whole library. A marker is a plain text file whose
    first line is the exact collection name, with no quotes or trailing
    punctuation. The response reports the exact marker path when it supplied
    the collection.
    """
    if path is not None:
        validate_path_in_root(path)
    client = ZoteroClient()
    marker_details = find_project_collection_details(path) if path else None
    target = collection or (marker_details[0] if marker_details else None)
    collection_key = client.resolve_collection(target) if target else None
    return {
        "candidates": [
            asdict(item) for item in client.search(query, limit=limit, collection_key=collection_key)
        ],
        "collection_marker_path": (
            str(marker_details[1]) if marker_details and collection is None else None
        ),
    }


@server.tool(annotations=READ_ONLY, structured_output=True)
def list_collections() -> ListCollectionsResult:
    """List Zotero collections (folders) with key/name/parent.

    For choosing a collection name/key to pass explicitly, or for writing a
    `.zotero-collection` marker file to auto-scope a project's folder.
    """
    client = ZoteroClient()
    return {"collections": [item if isinstance(item, dict) else asdict(item) for item in client.collections()]}


@server.tool(annotations=DESTROY_FILE, structured_output=True)
def set_project_collection(
    path: str,
    collection: str,
    replace_existing: bool = False,
) -> SetProjectCollectionResult:
    """Pin a folder to a Zotero collection by writing `.zotero-collection`.

    The collection argument can be either the exact collection name or its key.
    Refuses if a `.zotero-collection` marker already exists in `path` unless
    `replace_existing=True`.
    """
    validate_path_in_root(path)
    client = ZoteroClient()
    resolved_key = client.resolve_collection(collection)
    if not resolved_key:
        raise RefusedUnsafe(
            f"collection {collection!r} not found in the local Zotero library"
        )
    resolved_name = None
    for c in client.collections():
        c_dict = c if isinstance(c, dict) else asdict(c)
        if c_dict.get("key") == resolved_key:
            resolved_name = c_dict.get("name")
            break
    if not resolved_name:
        resolved_name = collection
    written = write_project_collection(
        path,
        resolved_name,
        replace_existing=replace_existing,
    )
    return {
        "success": True,
        "marker_path": written["marker_path"],
        "collection": resolved_name,
        "collection_key": resolved_key,
        "previous_value": written.get("previous_value"),
        "replaced": written.get("replaced", False),
    }


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def insert_zotero_citation(
    source: str,
    output: str,
    paragraph_index: int,
    anchor: str,
    item_keys: list[str],
    item_donor: str,
    bibliography_donor: str | None = None,
    bibliography_paragraph_index: int | None = None,
    source_sha256: SHA256 | None = None,
) -> InsertCitationResult:
    """Insert one live citation cluster at the requested output path and verify it.

    `anchor` is the exact visible text the citation is inserted immediately
    before; it must occur exactly once in `paragraph_index` (a duplicate
    elsewhere in the document is not a conflict). Word resolves the anchor's
    position itself rather than using a numeric offset, because a paragraph
    holding a field has Range coordinates that do not match its visible text.

    Bibliography donor and paragraph must either both be omitted or both be
    supplied. The source is never overwritten.
    """
    validate_path_in_root(source)
    validate_path_in_root(output)
    validate_path_in_root(item_donor)
    if bibliography_donor is not None:
        validate_path_in_root(bibliography_donor)
    _refuse_tracked_revisions(source, "insert_zotero_citation")
    with publication.capture(tool="insert_zotero_citation", op=f"cite {', '.join(item_keys)}") as published:
        raw = insert_citation_to_version(
            source,
            output,
            item_donor=item_donor,
            paragraph_index=paragraph_index,
            anchor=anchor,
            item_keys=tuple(item_keys),
            bibliography_donor=bibliography_donor,
            bibliography_paragraph_index=bibliography_paragraph_index,
            source_sha256=source_sha256,
        )
    after = audit_docx(raw["output"])
    return {
        "source": raw["source"],
        "source_sha256": raw["source_sha256"],
        "output": raw["output"],
        "output_sha256": raw["output_sha256"],
        "item_donor": str(Path(item_donor).resolve()),
        "paragraph_index": int(paragraph_index),
        "anchor": anchor,
        "item_keys": list(item_keys),
        "rendered_citation_text": raw.get("rendered_citation") or "",
        "zotero_field_count": after.zotero_item_fields,
        **publication.write_fields(published, raw["source_sha256"], raw["output_sha256"]),
    }


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def edit_citation_cluster(
    source: str,
    source_sha256: SHA256,
    paragraph_index: int,
    item_keys: list[str],
    output: str,
    citation_id: str | None = None,
    current_text: str | None = None,
    max_wait_seconds: float = 180.0,
) -> EditCitationClusterResult:
    """Set one existing Zotero citation's items to exactly `item_keys` (in that order) and render it.

    Pick the citation in `paragraph_index` by `citation_id` (its field JSON citationID) or by its
    rendered text `current_text` (e.g. "(12)"); exactly one of the two. One call can add, remove and
    reorder; kept items keep their locator/prefix/suffix. Works on documents with tracked changes: every
    revision is compared before the edit and after Word's save, and any changed revision refuses. A
    citation inside or crossing a tracked revision, or holding a nested field, is refused. Zotero then
    renders the citation and the bibliography. `output == source` edits in place (undo_document reverses
    it); otherwise `output` must not exist.
    """
    validate_path_in_root(source)
    validate_path_in_root(output)
    with publication.capture(tool="edit_citation_cluster", op=f"cite {', '.join(item_keys)}") as published:
        raw = edit_cluster_to_version(
            source,
            output,
            source_sha256=source_sha256,
            paragraph_index=paragraph_index,
            item_keys=item_keys,
            citation_id=citation_id,
            current_text=current_text,
            max_wait_seconds=max_wait_seconds,
        )
    raw.pop("refresh_rendered", None)
    return {**raw, **publication.write_fields(published, raw["source_sha256"], raw["output_sha256"])}


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def refresh_citations(
    source: str,
    source_sha256: SHA256,
    output: str,
    preview: bool = False,
    max_wait_seconds: float = 300.0,
) -> RefreshCitationsResult:
    """Re-render every Zotero citation and the bibliography through Zotero, inserting nothing, and diff them.

    Every citation's text and the reference list are reset and rendered fresh from the Zotero library.
    The result lists the bibliography entries that changed (before -> after), were dropped or added, the
    unchanged count, and the citations whose text changed. An entry listed as changed that you did not
    change in Zotero was probably edited by hand: fix the item in Zotero (update_zotero_item) and refresh
    again; never edit the reference-list text. `preview=True` runs the same refresh on a temporary copy
    and writes nothing (the source and output stay untouched). Otherwise `output == source` refreshes in
    place (undo_document reverts it) or `output` must not exist. Works on documents with tracked changes:
    every revision is compared after the reset and after Word's save; a citation inside or crossing a
    revision, or a bibliography holding one, is refused. Citations marked hand-edited in Word (Zotero's
    dontUpdate) keep their text, as Zotero does; the bibliography's custom entries stay Zotero's. A large
    document takes minutes (v28: ~3 min).
    """
    validate_path_in_root(source)
    if not preview:
        validate_path_in_root(output)
    with publication.capture(tool="refresh_citations", op="refresh citations") as published:
        raw = refresh_to_version(source, output, source_sha256=source_sha256, preview=preview,
                                 max_wait_seconds=max_wait_seconds)
    if preview:
        return {**raw, **publication.write_fields([], None, None)}
    return {**raw, **publication.write_fields(published, raw["source_sha256"], raw["output_sha256"])}


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def edit_document(
    source: str,
    source_sha256: SHA256,
    operations: list[MCPEditOperation],
    output: str,
    allow_partial: bool = False,
) -> EditBatchResult:
    """Apply up to 20 verified Word edits atomically by default.

    Tracked changes are operation types here, not separate tools:
    list_revisions, tracked_replace_text, tracked_replace_paragraph,
    accept_revisions, reject_revisions and sweep_characters go in
    operations[].type.

    Obtain source_sha256 from inspect_document. Each operation runs in an
    isolated version. Any failure discards the batch unless allow_partial=True.
    The result gives per-operation status and the exact source/hash for retries,
    plus undo_token / sha256_before / sha256_after: undo_document(path) restores
    the state before an in-place batch in one call.

    Each type reads only its own fields; the rest of the envelope is ignored.
    replace_text: paragraph_index, find, replace. find may be a whole line
    spanning bold/italic runs: only the characters that differ are rewritten,
    unchanged characters keep their formatting, and new text takes the
    formatting of the character it replaces or extends. Inside caps/small-caps
    runs, find matches in any case ("WORK EXPERIENCE" finds stored
    "WorK Experience") and the stored case is kept.
    replace_paragraph: paragraph_index, text (the whole new paragraph; same
    minimal-diff rule).
    replace_table_row: match (one cell's text, or one line of it, unique in
    the document), cells (new text per cell, "\n" between paragraphs of a
    cell; cell count must stay the same); optional table_index.
    replace_section: heading (as displayed) and rows (cell texts per row, for
    a section laid out as table rows) or paragraphs (for body text); optional
    until_heading. Rows/paragraphs may be added or removed inside the section
    only; everything outside is verified unchanged, cell counts are kept and
    date cells must still read as dates.
    group (any type): ops sharing a group, a table row or a section are
    all-or-nothing together under allow_partial (a row named by text an
    earlier op wrote joins that row's group; a row/section op whose target is
    found nowhere refuses the partial publish); the result's `groups` lists
    each group as applied/rolled_back/failed, and `undo_call` is the one call
    that restores the pre-batch state.
    insert_text: paragraph_index,
    anchor, position, text. insert_paragraph: after_paragraph_index, text
    (style optional). insert_table: after_paragraph_index, headers, rows,
    descriptive_columns (all required).
    Comments (direct XML, never through Word; work on documents with tracked
    changes; every other thread stays byte-identical, and so does every
    revision except an insertion holding an anchor edge, which gains only
    the markers and a run split there (same author, date, text, format);
    author defaults to RESEARCH_MCP_AUTHOR, else Word's user name, else
    "Author"; override with author/initials):
    add_comment: paragraph_index, anchor (unique in the paragraph's visible,
    accepted-view text; refused if it lies only in deleted text or cuts a
    field), text ("\\n" = new comment paragraph). New comments are open.
    reply_comment: comment_id (the thread id inspect_document shows),
    disposition (addressed needs depends_on = the op_id of the edit;
    rejected, no_change, needs_information), reply_text; resolve=true also
    resolves the thread. A resolved thread must be reopened before a plain
    reply. resolve_comment / reopen_comment: comment_id (thread and replies).
    delete_comment: comment_id; include_replies (default true) deletes the
    thread with its replies; include_replies=false with a reply's id deletes
    only that reply. Comment ops never change manuscript prose.

    Tracked changes (direct XML, never through Word, so other authors'
    revisions stay byte-identical):
    tracked_replace_text: paragraph_index, find (once in the accepted text),
    replace; optional author (default "Claude"). tracked_replace_paragraph:
    paragraph_index, text (the whole accepted paragraph); optional author.
    The minimal diff is written as w:del/w:ins (a change inside a word
    becomes the whole word); new text takes its neighbour's formatting.
    Verified before saving: rejecting the new revisions gives the input back,
    accepting them gives exactly the requested text. Refused, naming the
    revision (id, type, author), when the edit deletes or splits text inside
    another revision; also refused inside field results and in paragraphs
    with drawings.
    accept_revisions / reject_revisions: selectors ids, author,
    paragraph_index (+ through_paragraph_index), heading (+ until_heading);
    every selector given must match, at least one is required. A move is
    resolved as a pair: either side pulls in its partner by w:name (result
    pulled_in, moved). A removed paragraph mark (accepted deletion, rejected
    insertion) merges the paragraph into the next one, which keeps its own
    properties as in Word (result merged, index_shifts); merges into a table,
    across a cell or section end, or past an unselected property change are
    refused. Bookmark and comment markers inside removed text are kept where
    it was; revisions inside it are listed in removed_with; resolving a move
    whose removed side holds an unselected revision (e.g. another author's
    deletion inside a rejected moveTo) is refused, naming that id (select it
    too or resolve in Word); a move nested inside another move is refused;
    comments left with an empty range are kept and
    listed in detached_comments. All refused revisions are reported in one
    error with their count and ids, grouped by reason. A whole
    field (e.g. a Zotero citation) inside a deletion, insertion or move is
    removed or restored with it; the Zotero check then only requires every
    citation in the result to exist already in the document. Rejecting a
    table, row or cell property change restores the old properties as Word
    does. Refused: table-cell and row revisions, rejecting section or
    numbering formatting changes, a table grid change (tblGridChange) or a
    row/cell change that alters the grid (gridSpan, vMerge, gridBefore/After),
    and changes that
    would remove or restore part of a field, a drawing, or text sharing a
    run with a comment reference.
    list_revisions (read only, nothing written, no undo; the standalone
    list_revisions tool does the same without an output path): same selectors,
    optional limit; returns id, type, author, date, paragraph_index and text
    per revision plus each paragraph's accepted and original text, in the
    operation's details.
    On a document with tracked revisions, replace_text and replace_paragraph
    are made as untracked direct XML and refused where they touch a revised
    run; the other Word-based edits are refused.

    Clean on write: all text an operation writes (replacements, inserted and
    paragraph text, table cells/headers, comment and reply text) passes through
    textclean first: non-breaking and odd spaces, zero-width and control
    characters are normalised or removed; dashes, curly quotes, ellipsis,
    arrows and minus signs are kept, as is "\v" (Word's manual line
    break). Replacing known text cleans only what
    differs from it. Optional keep (any write op): characters to leave alone,
    each a character or "U+XXXX". find/anchor/match/heading are never cleaned.
    Each operation's details and the batch's `cleaned` list what changed.
    sweep_characters: clean text already in the document as tracked changes
    (author "Claude"); optional paragraph_index (+ through_paragraph_index;
    default the whole body), keep, author, preview_only. Equations (OMML),
    every field (Zotero citations included; field text counts as field text
    if it is inside the field with tracked changes accepted or rejected),
    text inside existing tracked changes and paragraphs with drawings are
    skipped; details give `cleaned`, `skipped` (counts by reason) and
    `skipped_total`. A preview writes nothing: its would-be changes are in
    details `would_clean`, and `cleaned` stays empty. inspect_document reports
    the same count as `unusual_characters`.
    """
    validate_path_in_root(source)
    validate_path_in_root(output)
    strict_ops = []
    for op in operations:
        try:
            strict_ops.append(_strict_operation_payload(op))
        except ValidationError as exc:
            field = exc.errors()[0].get("loc", ["operation"])[-1]
            msg = exc.errors()[0].get("msg", "Field required")
            required, optional = _field_contract(op.type)
            msg_str = (
                f"operation {op.op_id!r} ({op.type}) invalid: field {field!r} {msg}."
                f" {op.type} requires: {', '.join(required)}"
            )
            if optional:
                msg_str += f". Optional: {', '.join(optional)}"
            raise RefusedUnsafe(msg_str) from exc

    try:
        request = EditDocumentRequest(
            source=source,
            source_sha256=source_sha256,
            output=output,
            operations=strict_ops,
            allow_partial=allow_partial,
        )
    except ValidationError as exc:
        raise RefusedUnsafe(_describe_operation_error(operations, exc)) from exc
    with publication.capture(tool="edit_document"):
        return _edit_document(request).model_dump(mode="json")


def _refuse_tracked_revisions(source: str, tool: str) -> None:
    """Style, bookmark, table and citation tools run through Word, and a Word save rewrites other authors' revision records; the audit's
    revision_count misses property-only revisions (*PrChange), so gate on the complete count (plan 501 F2)."""

    from .editing import revision_total

    count = revision_total(source)
    if count:
        raise RefusedUnsafe(
            f"{tool}: source contains {count} tracked revision elements (property changes included); {tool} "
            "runs through Word and is refused on a document with revisions. Resolve them first with "
            "accept_revisions/reject_revisions, or edit text as direct XML through edit_document"
        )


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def apply_document_style(
    source: str,
    output: str,
    source_sha256: SHA256,
    profile: Literal["radiology_ai", "editing_compact", "compact_single", "standard_double"] = "radiology_ai",
    emphasis_policy: Literal["preserve", "headings_and_summary"] = "preserve",
    preview_only: bool = False,
) -> dict[str, Any]:
    """Apply journal typography and layout policies at the requested output path.

    Use preview_only=True to inspect role decisions and effects without writing.
    Layout preserves existing bold/italic by default. Only the explicit
    headings_and_summary policy adds whole-paragraph bold to those roles.
    For phrase emphasis use style_text, not this document-wide tool.
    Profile settings are typography presets, not full journal compliance.
    """
    validate_path_in_root(source)
    validate_path_in_root(output)
    if source_sha256 is not None:
        actual_sha = audit_docx(source).sha256
        if source_sha256.upper() != actual_sha.upper():
            raise RefusedUnsafe(
                f"source document SHA-256 mismatch: expected {source_sha256.upper()}, "
                f"actual {actual_sha.upper()}. Reinspect the document with inspect_document."
            )
    if preview_only:
        from .style_tools import plan_document_style
        return plan_document_style(source, profile, emphasis_policy)
    _refuse_tracked_revisions(source, "apply_document_style")
    with publication.capture(tool="apply_document_style", op=f"profile {profile}") as published:
        result = apply_journal_style(source, output, profile_name=profile, emphasis_policy=emphasis_policy, source_sha256=source_sha256)
    return {**result, **publication.write_fields(published, result["source_sha256"], result["output_sha256"])}


@server.tool(annotations=READ_ONLY, structured_output=True)
def audit_document_style(path: str, profile: str | None = None, first_paragraph: int = 1, last_paragraph: int | None = None, search_text: str | None = None) -> dict[str, Any]:
    """Inspect effective emphasis spans and flag suspicious styling without changing Word.

    Choose a profile to check saved typography; findings are guardrails, not a
    declaration of journal compliance. Use paragraph indices with style_text.
    """
    validate_path_in_root(path)
    from .style_audit import audit_document_style as audit_style
    return audit_style(path, profile=profile, first_paragraph=first_paragraph, last_paragraph=last_paragraph, search_text=search_text)


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def style_text(source: str, source_sha256: SHA256, paragraph_index: int, text: str, bold: bool | None = None, italic: bool | None = None, output: str | None = None, preview_only: bool = True, allow_whole_paragraph: bool = False) -> dict[str, Any]:
    """Preview or apply bold/italic to one exact phrase, preserving emphasis elsewhere.

    Defaults to a read-only preview. After checking the selected text, use
    preview_only=False to apply an already-authorized change. Text must match
    once. Whole-paragraph selection requires allow_whole_paragraph=True.
    Omitted properties stay unchanged. Field-overlapping selections are refused.
    Output defaults to source with staged verification; undo_document undoes it.
    """
    validate_path_in_root(source)
    if output is not None:
        validate_path_in_root(output)
    _refuse_tracked_revisions(source, "style_text")
    from .style_tools import style_text as apply_inline
    with publication.capture(tool="style_text", op=f"style paragraph {paragraph_index}") as published:
        result = apply_inline(source, source_sha256, paragraph_index, text, bold=bold, italic=italic, output=output, preview_only=preview_only, allow_whole_paragraph=allow_whole_paragraph)
    if preview_only:
        return result
    return {**result, **publication.write_fields(published, result.get("source_sha256"), result.get("output_sha256"))}


# Not read-only: it writes PNGs into ``output_dir`` (overwriting page-N.png there), so it is left out of
# read-only modes such as `research-mcp chatgpt --read-only`.
@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def render_document_pages(
    path: str,
    output_dir: str | None = None,
    pages: list[int] | None = None,
    dpi: int = 200,
) -> RenderDocumentResult:
    """Render pages of a .docx file into clean, print-quality PNG images.

    Exports clean vector PDF suppressing comment markup sidebars,
    then rasterizes pages at the specified DPI for visual inspection.
    """
    validate_path_in_root(path)
    if output_dir is not None:
        validate_path_in_root(output_dir)
    from .operation_process import run_operation
    return run_operation("render", {"docx_path": path, "output_dir": output_dir, "pages": pages, "dpi": dpi}, timeout_seconds=180)


@server.tool(annotations=UNDO_WRITE, structured_output=True)
def undo_document(
    path: str,
    undo_token: str | None = None,
    to: str | None = None,
    force: bool = False,
    output: str | None = None,
) -> UndoDocumentResult:
    """Undo MCP edits in one call by restoring an earlier whole-file state.

    No arguments: restore the state before the last MCP edit or batch.
    undo_token: restore the state before that edit (later edits are undone too).
    to: restore the checkpoint with that label (see checkpoint_document).
    The undo is recorded, so calling undo_document(path) again reverses it.
    Refuses if the document changed since the MCP last recorded it (the message
    names both shas) unless force=True (the current state is checkpointed first)
    or output=<new .docx path> (writes the earlier state as a new file, leaving
    the document as it is). Refuses while Word holds the file. With no recorded
    edits, a legacy .docx.bak beside the document is the fallback (never deleted).
    Limits: whole-file, MCP edits only, not Word's Ctrl+Z, this PC only.
    """
    validate_path_in_root(path)
    if output is not None:
        validate_path_in_root(output)
    return publication.undo_document(path, undo_token=undo_token, to=to, force=force, output=output)


@server.tool(annotations=UNDO_WRITE, structured_output=True)
def revert_document(path: str, force: bool = False, output: str | None = None) -> RevertDocumentResult:
    """Undo the last MCP edit of a document: an alias of undo_document(path).

    Restores the state before the last in-place edit or batch from the undo
    store (falls back to a legacy .docx.bak only when the document has no
    recorded edits). Same stale guard and Word-lock refusal as undo_document.
    """
    validate_path_in_root(path)
    if output is not None:
        validate_path_in_root(output)
    return publication.undo_document(path, force=force, output=output, tool="revert_document")


@server.tool(annotations=CHECKPOINT_WRITE, structured_output=True)
def checkpoint_document(path: str, label: str) -> CheckpointDocumentResult:
    """Store the document's current state under a label in the undo store.

    undo_document(path, to=label) returns to it later. The document is not
    changed. Stored on this PC only, outside OneDrive; retention limits apply.
    """
    validate_path_in_root(path)
    return publication.checkpoint_document(path, label)


@server.tool(annotations=READ_ONLY, structured_output=True)
def document_history(path: str) -> DocumentHistoryResult:
    """List a document's recorded edits, checkpoints, undos and outside changes, newest first.

    Each entry id is an undo_token. Also lists pruned entries (dropped, with the
    reason), the retention limits, and whether the document changed outside the
    server since the last record.
    """
    validate_path_in_root(path)
    return publication.document_history(path)


@server.tool(annotations=OPEN_WORLD_WRITE, structured_output=True)
def report_issue(
    title: str,
    details: str,
    category: Literal[
        "tool_error",
        "formatting_defect",
        "zotero",
        "word_com",
        "lock",
        "friction",
        "feature_request",
        "char_mapping",
        "other",
    ] = "friction",
    document_path: str | None = None,
    context: dict[str, Any] | None = None,
    char: str | None = None,
    should_be: str | None = None,
    mapped_to: str | None = None,
    snippet: str | None = None,
) -> ReportIssueResult:
    """Report an issue, unexpected formatting defect, or workflow friction encountered during execution.

    Appends a structured entry to the project's ISSUES.md log so developers can trace and resolve it.

    Wrong character mapping (plan 551): when a write's `cleaned` list changed a character it should not
    have (or left one it should have changed), use category="char_mapping" with `char` (the character or
    "U+XXXX"), `should_be` ("keep", what it should become, or "" to remove it), optional `mapped_to` (what you saw) and an
    optional `snippet` (at most 80 characters of surrounding text, never a whole paragraph); `details` is
    the reason. Reports never change the mapping table; to protect a character now, pass `keep`.

    Every report is logged locally. The result also carries `issue_url`, a prefilled GitHub issue link
    for the user to open (with `issue_title`/`issue_body`, the same text), or, when the server is set up
    to file reports elsewhere, `improvement` and `file_with` steps to follow instead.
    """
    if document_path is not None:
        try:
            validate_path_in_root(document_path)
        except Exception:
            pass
    if category == "char_mapping":
        return _report_char_mapping(title, details, document_path, context, char, should_be, mapped_to, snippet)
    logged = record_agent_issue(
        title=title,
        details=details,
        category=category,
        document_path=document_path,
        context=context,
    )
    if config.issues_mode() == "agent":
        return logged
    body = f"Category: {category}\n\n{details.strip()}"
    return {**logged, **_github_issue(title, body)}


GITHUB_ISSUES = "https://github.com/Russell-Day/research-mcp/issues/new"
# GitHub answers a prefilled new-issue link up to about 6000 characters (500 from ~6500, 414 from ~9000).
ISSUE_URL_MAX = 6000


def _github_issue(title: str, body: str) -> dict[str, str]:
    from urllib.parse import urlencode

    title = " ".join(title.split())[:200]
    body = body.strip()

    def url(text: str) -> str:
        return f"{GITHUB_ISSUES}?{urlencode({'title': title, 'body': text})}"

    if len(url(body)) > ISSUE_URL_MAX:
        # Longest body prefix whose encoded link fits; encoded length grows with the prefix.
        low, high = 0, len(body)
        while low < high:
            mid = (low + high + 1) // 2
            if len(url(body[:mid])) <= ISSUE_URL_MAX:
                low = mid
            else:
                high = mid - 1
        body = body[:low]
    return {
        "issue_url": url(body),
        "issue_title": title,
        "issue_body": body,
        "message": "Logged locally. To report it upstream, open issue_url (a prefilled GitHub issue) and submit it.",
    }


def _report_char_mapping(title, details, document_path, context, char, should_be, mapped_to, snippet):
    if not char or should_be is None:
        raise ValueError('category char_mapping needs char and should_be ("keep", the replacement, or "" to remove)')
    report = textclean.mapping_report(char, should_be, details, "research MCP", snippet=snippet,
                                      mapped_to=mapped_to)
    count = textclean.record_report(get_log_dir() / "char_reports.jsonl", report, details)
    logged = record_agent_issue(
        title=report["improvement"]["text"],
        details=f"{details.strip()}\n\nAgent title: {title}",
        category="char_mapping",
        document_path=document_path,
        context={**(context or {}), "char_mapping": {k: report[k] for k in ("codepoint", "current", "should_be")},
                 "report_count": count},
    )
    result = {
        **logged,
        "char_mapping": {k: report[k] for k in ("codepoint", "name", "current", "should_be", "title_prefix")},
        "report_count": count,
    }
    if config.issues_mode() == "agent":
        return {**result, "improvement": report["improvement"], "file_with": config.issues_file_with(textclean.FILE_WITH)}
    notes = "\n".join(line for line in report["improvement"]["notes"].splitlines() if not line.startswith("Fix: "))
    return {**result, **_github_issue(report["improvement"]["text"], notes)}


@server.tool(annotations=OPEN_WORLD_WRITE, structured_output=True)
def add_zotero_item_by_doi(
    doi: str,
    path: str | None = None,
    collection: str | None = None,
) -> AddZoteroItemResult:
    """Add a reference to the local Zotero library by its DOI."""
    if path is not None:
        validate_path_in_root(path)
    return add_item_by_doi(doi=doi, path=path, collection=collection)


@server.tool(annotations=OPEN_WORLD_WRITE, structured_output=True)
def update_zotero_item(
    key: str,
    fields: dict[str, Any],
    dry_run: bool = False,
    clean_text: bool = True,
    sync_timeout_seconds: float = 60.0,
) -> UpdateZoteroItemResult:
    """Change an existing Zotero item's metadata (the way to fix a reference).

    `fields` maps Zotero field names to new values: strings for ordinary fields
    (`title`, `proceedingsTitle`, `conferenceName`, `volume`, `pages`, `extra`...;
    "" clears one), full lists for `creators` and `tags`. Names are checked
    against the item's type; `itemType`, `key`, `version` and collections are
    never changed. Writes through the Zotero web API, refusing if the item
    changed elsewhere since it was read, then waits until the local Zotero has
    synced (an error after `sync_timeout_seconds`, since Word renders from it).
    Returns the field diff and `previous`: pass `previous` back as `fields`
    with `clean_text=False` to undo exactly. `dry_run=True` returns the diff
    without writing. Text is cleaned like document writes (U+2011 hyphens
    become "-") unless `clean_text=False`. Then refresh the document's
    citations so the bibliography picks the change up.
    """
    return update_item(
        key, fields, dry_run=dry_run, clean_text=clean_text,
        sync_timeout_seconds=sync_timeout_seconds,
    )


def build_parser(prog: str | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=prog,
        description="Research Word and Zotero MCP server.",
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "streamable-http", "http"),
        default="stdio",
        help="Transport protocol (default: stdio).",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host for HTTP transport (default: 127.0.0.1).",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for HTTP transport (default: 8000).",
    )
    parser.add_argument(
        "--root-dir",
        type=Path,
        default=None,
        help="Root directory to sandbox file operations.",
    )
    return parser


def main(argv: list[str] | None = None, *, prog: str | None = None) -> None:
    """Run the server (stdio by default). ``argv`` defaults to ``sys.argv[1:]``."""
    parser = build_parser(prog)
    args = parser.parse_args(argv)
    if args.root_dir:
        set_root_dir(args.root_dir)
    if args.transport == "stdio":
        server.run(transport="stdio")
    elif args.transport == "streamable-http":
        server.run(transport="streamable-http", host=args.host, port=args.port)
    elif args.transport == "http":
        server.run(transport="http", host=args.host, port=args.port)


if __name__ == "__main__":
    main()
