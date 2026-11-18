"""Thin MCP 2.0 boundary over the tested research_mcp core."""

from __future__ import annotations

import argparse
import functools
import inspect
from dataclasses import asdict
import os
from pathlib import Path
from typing import Any, Literal

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError, UnexpectedToolError
from mcp.types import ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field, ValidationError

from . import __version__
from .audit import audit_docx, is_file_locked
from .bookmarks import inventory, release_to_version
from .citation_ops import insert_citation_to_version
from .cluster_ops import edit_cluster_to_version
from .refresh_ops import refresh_to_version
from .zotero_prefs import set_style_to_version
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
    DeleteCitationOperation,
    DeleteParagraphOperation,
    StyleTextOperation,
    FormatParagraphOperation,
    FIND_MAX,
    SetTrackChangesOperation,
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
from .legacy_doc import (
    convert_in_worker,
    docx_for_reading,
    is_legacy_doc,
    read_refusal,
    resolve_convert_target,
    write_refusal,
)
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
    EditFeedbackResult,
    RefreshCitationsResult,
    FormatTablesResult,
    InsertCitationResult,
    InspectDocumentResult,
    ListBookmarksResult,
    ListCollectionsResult,
    MCPMergeSpec,
    MCPTableFormatSpec,
    ExportPdfResult,
    RenderDocumentResult,
    ReportIssueResult,
    RevertDocumentResult,
    SearchZoteroResult,
    SetCitationStyleResult,
    SetProjectCollectionResult,
    UndoDocumentResult,
    UpdateZoteroItemResult,
)
from .styling import apply_journal_style
from .rendering import render_docx_pages, resolve_pdf_target
from .tables import MergeGroup, TableFormatSpec, format_to_version
from .word import staged_document
from .word_tracked import TrackedFormatTablesResult
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
   - Use the edit_document op format_paragraph for a whole paragraph's font, size, indents, spacing or list numbering (`like_paragraph=M` matches paragraph M); it follows the tracking switch. A `format_note` on a fill of a blank paragraph names the call to make.
   - Use style_text for exact phrase bold/italic. Preview the selected text, then apply with preview_only=False under existing user authorization. Never use a whole-paragraph font assignment for a phrase request.
   - Use `apply_document_style` to enforce target journal typography (e.g. `profile="radiology_ai"` for 11pt Times New Roman, 1-inch margins, double spacing, configured paragraph roles, and stripped page numbers).
   - Use `render_document_pages` to export high-resolution PNG page images (clean print layout without comment markup balloon gutters; markup=True shows redlines and comment balloons) to visually verify tables, figures, and formatting.
   - Use `export_pdf` to save a PDF of the document (same clean print layout) for the user to check or hand in: next to the .docx as `<name>.pdf` by default, or at `output`. It never overwrites an existing file and never changes the .docx. Tracked changes appear accepted (no markup) and comments are not shown.
6. When citing or searching references:
   - Use `search_zotero` or `add_zotero_item_by_doi`.
   - A wrong reference (title, venue, volume, pages...) is fixed in Zotero with `update_zotero_item`, then the document is refreshed; never edit the reference-list text (a refresh overwrites it).
   - To change which references an existing citation cites (add, remove, reorder), call `edit_citation_cluster` (pick the cluster by `citation_id` or its rendered `current_text` in `paragraph_index`); it works on documents with tracked changes and refreshes through Zotero.
   - To remove a whole citation (e.g. a stray one in a figure legend), use the `edit_document` op `delete_citation` (same selectors); it is a tracked deletion by default and does not refresh Zotero: run `refresh_citations` after the deletion is accepted (the result lists the references that will drop out).
   - To re-render every citation and the bibliography (after fixing items in Zotero, or to check for hand edits), call `refresh_citations` (`preview=True` first writes nothing); it works on documents with tracked changes. A bibliography entry it lists as changed that you did not change in Zotero was probably edited by hand: fix the item with `update_zotero_item` and refresh; never edit the reference-list text.
   - To change the document's citation style (e.g. to american-medical-association for a journal), call `set_citation_style` (`preview=True` first writes nothing): it sets the style in the document's Zotero preferences and re-renders everything like `refresh_citations`; the style must be installed in Zotero.
   - Call `insert_zotero_citation`: `item_donor` is optional in python core and defaults to the bundled donor document.
   - Verify Zotero Desktop is running (port 23119).
   - If `audit_document` reports `citation_readiness.status == "virgin_document"`, open the document in Word, insert one citation manually, save, and close Word. Then rerun `audit_document`, confirm readiness, and retry with the fresh `sha256`. If status is `structurally_invalid`, repair the package before opening it in Word or attempting citation operations.
7. If you encounter any unexpected issues, formatting quirks, Word lock contention, or agent friction during execution, call `report_issue` to log a detailed report to logs/ISSUES.md for developer improvement.
8. Text you write is cleaned of invisible and unusual characters (results list them under `cleaned`; pass `keep` to hold one). Equations, fields and Zotero citations are never touched. Sweep an existing document only when asked (`sweep_characters`, tracked). If a character was mapped wrongly, call `report_issue(title=..., details=<why>, category="char_mapping", char=..., should_be=...)` and follow what it returns (give the user its `issue_url`, or follow its `file_with` steps when it has them).
9. Tracked changes: `set_track_changes(on, author)` (an edit_document op, or edit_document's `track_changes=true`) turns on research-mcp's own tracking switch, kept in the document (not Word's Track Changes setting; inspect/audit show it as `track_changes`). While it is on, replace_text, replace_paragraph and insert_text are made as tracked changes by that author (an op's own `author` overrides it), delete_citation follows it, insert_paragraph adds each new paragraph as a tracked insertion (w:ins on its runs and mark; with tracking off on a revised document it inserts plainly), replace_table_row and replace_section record their changed, added and removed lines or paragraphs as tracked changes (table rows added or removed by replace_section(rows=...) as Word's tracked row insertion/deletion: w:ins/w:del in w:trPr, every run and paragraph mark marked; with tracking off on a revised document they edit plainly by direct XML), style_text (tool or op) records bold/italic as a tracked formatting change (w:rPrChange; text in the author's own pending insertion is formatted in place), format_paragraph records its paragraph and run formatting as w:pPrChange/w:rPrChange, apply_document_style and format_tables run with Word's own Track Changes and their revisions are relabelled to the switch author (what Word can't track, such as cell merges, gets one comment each, listed under `untracked`), insert_table writes the caption and table as one tracked insertion (merges get one comment each), and ops that cannot be tracked yet (clean_bookmarks) refuse unless given `tracked: false`; insert_zotero_citation is made as one tracked insertion (w:ins) of the whole citation. An op's `tracked` overrides the switch for that op. Results say `tracking` on/off.
10. To answer a comment in plain words use disposition `reply`; the other dispositions add their label.
11. Edit feedback: every edit you make is recorded as a proposal, and when a later call sees the document the user's outcome is recorded (accepted, rejected, reworded, or removed with its paragraph; tracked or plain edits). Per-document results then carry one line, `edit_feedback` ("Since your last edit: 2 accepted, 1 rejected, 1 reworded (3 pending). Call edit_feedback for details."), only when something is new. Call `edit_feedback(path)` before editing a document again: it lists this document's new net outcomes (first proposal → what the user kept, `rounds` when you edited a span several times), newest first, at most 10; pending proposals are counted, not listed. Learn from rewordings and rejections; do not re-propose rejected text. Only on need: `scope="all", query=...` searches every document; `record_id=..., detail="rounds"` shows each round of one item. A `feedback_warning` means feedback was not recorded; the tool itself succeeded.
"""

server = MCPServer(
    "research-mcp",
    title="Research Word and Zotero MCP",
    description="Safety-enforcing, versioned Word formatting and Zotero operations.",
    version=__version__,
    instructions=SERVER_INSTRUCTIONS,
)

_orig_call_tool = server._tool_manager.call_tool

from .stale_exit import ActivityTracker  # noqa: E402

# Calls in flight and the time of the last start/finish, read by --exit-when-stale's idle check.
TOOL_ACTIVITY = ActivityTracker()


async def _logged_call_tool(
    name: str,
    arguments: dict[str, Any],
    context: Any,
    convert_result: bool = False,
) -> Any:
    start = time.perf_counter()
    error = None
    result = None
    TOOL_ACTIVITY.start()
    try:
        result = await _orig_call_tool(
            name, arguments, context, convert_result=convert_result
        )
        return result
    except UnexpectedToolError as exc:
        # mcp >= 2.1 sends the client only "Error executing tool <name>" for an
        # exception that is not a ToolError. Refusals (RefusedUnsafe and the rest)
        # are the message the model must read, so the tool's own text is kept.
        cause = exc.__cause__
        if cause is None or isinstance(cause, ToolError):
            error = exc
            raise
        error = ToolError(f"Error executing tool {name}: {cause}")
        raise error from cause
    except Exception as exc:
        error = exc
        raise
    finally:
        TOOL_ACTIVITY.finish()
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
        "delete_citation",
        "delete_paragraph",
        "style_text",
        "format_paragraph",
        "set_track_changes",
    ]
    paragraph_index: int = 1
    find: str = Field(default="", max_length=FIND_MAX, description="Exact visible text to replace (any length up to 5000; longer than 255 is matched by its start and checked in full). delete_citation: same as current_text.")
    replace: str = ""
    match_case: bool = True
    occurrence: int = 1
    expected_count: int = 1
    position: Literal["before", "after"] = "before"
    anchor: str = Field(default="", max_length=FIND_MAX, description="Exact visible anchor (up to 5000 characters).")
    after_paragraph_index: int = 0
    text: str = ""
    style: str | None = None
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    descriptive_columns: list[int] = Field(default_factory=list)
    merges: list[dict[str, Any]] = Field(default_factory=list)
    caption: str | None = Field(default=None, description="insert_table: one-line caption paragraph written before the table (e.g. 'Table 1. ...').")
    caption_style: str | None = Field(default=None, description="insert_table: paragraph style of the caption (name or styleId); default 'Caption'.")
    comment_id: str = Field(default="", description="reply/resolve/reopen/delete_comment: the thread's comment_id from inspect_document (durable id); legacy:N or N also works when unambiguous.")
    disposition: Literal["addressed", "rejected", "no_change", "needs_information", "reply", "accepted"] = Field(default="addressed", description="reply_comment: addressed (needs depends_on: the op_id of the edit in this batch, or addressed_by), rejected, no_change, needs_information (written as 'Addressed: ...', 'Rejected: ...', 'No change: ...', 'Needs information: ...'), or reply (reply_text as given with no prefix; surrounding whitespace is trimmed, as for every reply). 'accepted' is an alias of addressed.")
    addressed_by: list[str] = Field(default_factory=list, description="reply_comment (addressed or reply): evidence from earlier calls, each a revision id from list_revisions (still in the document) or an undo_token / document_history entry id of an edit of this document still in effect.")
    reply_text: str = ""
    resolve: bool = Field(default=False, description="reply_comment: also mark the thread resolved in the same write.")
    include_replies: bool = Field(default=True, description="delete_comment: true deletes the thread with its replies; false deletes only the given reply.")
    author: str = Field(default="", description="Author. Comments default to RESEARCH_MCP_AUTHOR from the config, else Word's user name, else \"Author\"; tracked edits default to the tracking switch's author, else Claude.")
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
    through_paragraph_index: int | None = Field(default=None, ge=1, description="accept/reject/list_revisions, delete_paragraph and format_paragraph: last paragraph of the range starting at paragraph_index.")
    limit: int = Field(default=200, ge=1, le=2000, description="list_revisions: most revisions to return.")
    keep: list[str] = Field(default_factory=list, description='Write ops and sweep_characters: characters to leave as written (each a character or "U+XXXX"); everything else written is cleaned with textclean and listed in the result\'s `cleaned`.')
    preview_only: bool = Field(default=False, description="sweep_characters: report what would change; write nothing.")
    citation_id: str = Field(default="", description="delete_citation: the citation's citationID (from its field JSON); or give current_text.")
    current_text: str = Field(default="", description='delete_citation: the citation\'s rendered text in paragraph_index, e.g. "(12)" or "22".')
    tracked: bool | None = Field(default=None, description="Text ops, insert_table, delete_citation, delete_paragraph, style_text and format_paragraph: true makes this edit a tracked change, false a plain one. Default: the document's tracking switch (set_track_changes); delete_citation and delete_paragraph default to tracked when the switch was never set.")
    bold: bool | None = Field(default=None, description="style_text: true sets bold on the phrase, false removes it; omitted: unchanged.")
    italic: bool | None = Field(default=None, description="style_text: true sets italic on the phrase, false removes it; omitted: unchanged.")
    allow_whole_paragraph: bool = Field(default=False, description="style_text: needed when the phrase is the whole paragraph.")
    font: str | None = Field(default=None, description="format_paragraph: font name for every run and the paragraph mark (ascii, hAnsi and cs).")
    size: float | None = Field(default=None, description="format_paragraph: font size in points (steps of 0.5).")
    left_indent: float | None = Field(default=None, description="format_paragraph: left indent in points.")
    right_indent: float | None = Field(default=None, description="format_paragraph: right indent in points.")
    first_line_indent: float | None = Field(default=None, description="format_paragraph: first-line indent in points; negative is a hanging indent.")
    space_before: float | None = Field(default=None, ge=0, description="format_paragraph: space before in points.")
    space_after: float | None = Field(default=None, ge=0, description="format_paragraph: space after in points.")
    remove_numbering: bool = Field(default=False, description="format_paragraph: take the paragraph out of its list (w:numPr removed; numId 0 when its style numbers it).")
    like_paragraph: int | None = Field(default=None, ge=1, description="format_paragraph: copy this paragraph's indents, spacing, numbering and main font and size; explicit fields win.")
    on: bool | None = Field(default=None, description="set_track_changes: true turns tracking on for this document (later ops in this batch and later calls), false turns it off.")


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
    "delete_citation": DeleteCitationOperation,
    "delete_paragraph": DeleteParagraphOperation,
    "style_text": StyleTextOperation,
    "format_paragraph": FormatParagraphOperation,
    "set_track_changes": SetTrackChangesOperation,
}
# Ops that take the tracking switch's per-op overrides (`tracked`, `author`).
_SWITCH_OPS = ("replace_text", "replace_paragraph", "insert_text", "insert_paragraph", "insert_table",
               "replace_table_row", "replace_section", "delete_paragraph", "style_text", "format_paragraph")
_BASE_FIELDS = {"op_id", "type", "depends_on", "group"}


# Ops whose model makes a required choice optional fields (one of text/find; bold and/or italic): what the
# caller must give, as validation errors and the op docs state it.
_CONTRACT_OVERRIDES: dict[str, tuple[tuple[str, ...], tuple[str, ...]]] = {
    "style_text": (("paragraph_index", "text (or find, another name for it; one of them)", "bold and/or italic"),
                   ("allow_whole_paragraph", "tracked", "author")),
    "format_paragraph": (("paragraph_index", "one or more of font, size, left_indent, right_indent, first_line_indent, "
                          "space_before, space_after, remove_numbering, like_paragraph"),
                         ("through_paragraph_index", "tracked", "author")),
}


def _field_contract(op_type: str) -> tuple[tuple[str, ...], tuple[str, ...]]:
    if op_type in _CONTRACT_OVERRIDES:
        return _CONTRACT_OVERRIDES[op_type]
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
    if op.type in _SWITCH_OPS:
        # Plan 864: only overrides the caller gave; otherwise the document's tracking switch decides.
        if op.tracked is not None:
            base["tracked"] = op.tracked
        if op.author:
            base["author"] = op.author
    if op.type == "set_track_changes":
        if op.on is None:
            raise RefusedUnsafe(f"operation {op.op_id!r} (set_track_changes) invalid: field 'on' is required "
                                "(true or false). set_track_changes requires: on. Optional: author")
        return SetTrackChangesOperation(**base, on=op.on, author=op.author or None)
    if op.type == "sweep_characters":
        kwargs = {**base, "preview_only": op.preview_only, "through_paragraph_index": op.through_paragraph_index}
        if "paragraph_index" in given or op.through_paragraph_index is not None:
            kwargs["paragraph_index"] = op.paragraph_index
        if "author" in given:
            kwargs["author"] = op.author
        return SweepCharactersOperation(**kwargs)
    if op.type == "delete_citation":
        # Plan 864 decision 18: `find` is taken as current_text (the rendered text, e.g. "8").
        if op.find and op.current_text:
            raise RefusedUnsafe(f"operation {op.op_id!r} (delete_citation) invalid: find and current_text both "
                                "given; find is another name for current_text, so give one")
        kwargs = {**base, "paragraph_index": op.paragraph_index,
                  "citation_id": op.citation_id or None, "current_text": op.current_text or op.find or None}
        if op.tracked is not None:
            kwargs["tracked"] = op.tracked
        if "author" in given:
            kwargs["author"] = op.author
        return DeleteCitationOperation(**kwargs)
    if op.type == "style_text":
        # Plan 864 S6: `find` is another name for text; tracked/author were taken above (_SWITCH_OPS).
        if op.find and op.text:
            raise RefusedUnsafe(f"operation {op.op_id!r} (style_text) invalid: text and find both given; find is "
                                "another name for text, so give one. style_text requires: paragraph_index, text, "
                                "bold and/or italic. Optional: allow_whole_paragraph, tracked, author")
        return StyleTextOperation(**base, paragraph_index=op.paragraph_index, text=op.text or op.find or None,
                                  bold=op.bold, italic=op.italic, allow_whole_paragraph=op.allow_whole_paragraph)
    if op.type == "format_paragraph":
        # Plan 916 S3: tracked/author were taken above (_SWITCH_OPS); bold/italic belong to style_text.
        if op.bold is not None or op.italic is not None:
            raise RefusedUnsafe(f"operation {op.op_id!r} (format_paragraph) invalid: format_paragraph does not "
                                "set bold or italic; use style_text (paragraph_index, text, bold/italic) for that")
        return FormatParagraphOperation(
            **base, paragraph_index=op.paragraph_index, through_paragraph_index=op.through_paragraph_index,
            font=op.font, size=op.size, left_indent=op.left_indent, right_indent=op.right_indent,
            first_line_indent=op.first_line_indent, space_before=op.space_before, space_after=op.space_after,
            remove_numbering=op.remove_numbering, like_paragraph=op.like_paragraph)
    if op.type == "delete_paragraph":
        # Plan 864 S3c: tracked/author were taken above (_SWITCH_OPS); the switch, else tracked by "Claude".
        return DeleteParagraphOperation(**base, paragraph_index=op.paragraph_index,
                                        through_paragraph_index=op.through_paragraph_index)
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
        if op.caption:
            kwargs["caption"] = op.caption
        if op.caption_style:
            kwargs["caption_style"] = op.caption_style
        return InsertTableOperation(**kwargs)
    if op.type == "reply_comment":
        kwargs: dict[str, Any] = {
            **base,
            "comment_id": op.comment_id,
            "disposition": "addressed" if op.disposition == "accepted" else op.disposition,
            "resolve": op.resolve,
            "author": op.author or None,
            "initials": op.initials or None,
            "addressed_by": op.addressed_by,
        }
        if op.reply_text:
            kwargs["reply_text"] = op.reply_text
        return ReplyCommentOperation(**kwargs)
    raise ValueError(f"unsupported operation type: {op.type}")


def _operation_error_text(op: MCPEditOperation, error: dict) -> str:
    """One readable line for a validation error of one operation (plan 864 decision 18).

    A whole-model validator error has an empty `loc`; it reads
    "operation 'x' (type) invalid: <validator message>. type requires: ...", never an IndexError."""

    loc = tuple(part for part in (error.get("loc") or ()) if isinstance(part, str) and part != op.type)
    message = str(error.get("msg") or "validation error")
    if message.startswith("Value error, "):
        message = message[len("Value error, "):]
    message = message.rstrip(".")
    required, optional = _field_contract(op.type)
    head = f"field {loc[-1]!r} {message}" if loc else message
    text = f"operation {op.op_id!r} ({op.type}) invalid: {head}. {op.type} requires: {', '.join(required) or '-'}"
    if optional:
        text += f". Optional: {', '.join(optional)}"
    return text


def _describe_operation_error(
    operations: list[MCPEditOperation], exc: ValidationError
) -> str:
    lines = []
    for error in exc.errors():
        loc = error.get("loc") or ()
        position = next((part for part in loc if isinstance(part, int)), None)
        if position is not None and position < len(operations):
            op = operations[position]
            inner = {**error, "loc": tuple(loc[list(loc).index(position) + 1:])}
            lines.append(_operation_error_text(op, inner))
        else:
            message = str(error.get("msg") or "validation error")
            lines.append(message[len("Value error, "):] if message.startswith("Value error, ") else message)
    return " | ".join(lines)


def _feedback(param: str = "path", *, via: str = "user"):
    """Plan 829: add ``edit_feedback`` / ``feedback_warning`` to a per-document tool's result.

    A write the tool published carries the line from its own capture (publication records); otherwise the
    document is reconciled on sight (cheap unless it has open proposals and a sha the log has not seen).
    ``via="model"`` tools (undo, revert) first settle the user's outcomes, then record what the tool itself
    undid as decided by the model. Feedback never fails the tool."""

    def wrap(fn):
        signature = inspect.signature(fn)

        @functools.wraps(fn)
        def inner(*args, **kwargs):
            from .edit_feedback_word import on_sight, tool_fields

            try:
                path = signature.bind_partial(*args, **kwargs).arguments.get(param)
            except TypeError:
                path = None
            before = {}
            if via == "model" and path:
                before = {k: v for k, v in on_sight(path, via="user").items() if v}
            with publication.capture() as records:
                result = fn(*args, **kwargs)
            fields = tool_fields(path, records, via=via) if path else {}
            for key, value in before.items():
                fields.setdefault(key, value)
            if not fields:
                return result
            if isinstance(result, BaseModel):
                return result.model_copy(update=fields)
            if isinstance(result, dict):
                return {**result, **fields}
            return result

        return inner

    return wrap


def _docx_only(param: str = "source", *, read: bool = False):
    """Plan 916: refuse a Word 97-2003 .doc (detected by its header, whatever its name) before a tool that edits,
    or parses the .docx package itself, starts. One line naming convert_document; nothing is read or written."""

    def wrap(fn):
        signature = inspect.signature(fn)

        @functools.wraps(fn)
        def inner(*args, **kwargs):
            try:
                path = signature.bind_partial(*args, **kwargs).arguments.get(param)
            except TypeError:
                path = None
            if path:
                validate_path_in_root(path)
                if is_legacy_doc(path):
                    raise RefusedUnsafe((read_refusal if read else write_refusal)(fn.__name__, path))
            return fn(*args, **kwargs)

        return inner

    return wrap


class _ConvertedFromDoc(BaseModel):
    converted_from_doc: bool | None = Field(
        default=None,
        description="Plan 916: true when the file is a Word 97-2003 .doc, read through a temporary .docx Word "
        "made from it (removed afterwards; the .doc is never written). path and sha256 are the .doc's.")


class AuditDocumentOrDocResult(AuditDocumentResult, _ConvertedFromDoc):
    model_config = ConfigDict(title="AuditDocumentResult")


class InspectDocumentOrDocResult(InspectDocumentResult, _ConvertedFromDoc):
    model_config = ConfigDict(title="InspectDocumentResult")


class RenderDocumentOrDocResult(RenderDocumentResult, _ConvertedFromDoc):
    model_config = ConfigDict(title="RenderDocumentResult")


class ExportPdfOrDocResult(ExportPdfResult, _ConvertedFromDoc):
    model_config = ConfigDict(title="ExportPdfResult")


class ConvertDocumentResult(BaseModel):
    source: str
    source_sha256: str = Field(description="The .doc's sha256, unchanged by the conversion.")
    output: str
    output_sha256: str
    size_bytes: int
    elapsed_seconds: float


@server.tool(annotations=READ_ONLY, structured_output=True)
@_feedback("path")
def audit_document(path: str) -> AuditDocumentOrDocResult:
    """Audit a DOCX package without opening or changing it in Word."""
    # A Word 97-2003 .doc is audited through a temporary .docx Word makes from it (`converted_from_doc`;
    # the .doc is never written). Not in the docstring: the benchmark's recorded schema size is frozen.
    validate_path_in_root(path)
    from .track_setting import describe as _track_changes

    with docx_for_reading(path, "audit_document") as (docx, legacy):
        result = {**audit_docx(docx).to_agent_dict(), "track_changes": _track_changes(docx)}
    return legacy.relabel(result, path_key="path", sha_key="sha256") if legacy else result


@server.tool(annotations=READ_ONLY, structured_output=True)
@_feedback("path")
def inspect_document(
    path: str,
    first_paragraph: int | None = None,
    last_paragraph: int | None = None,
    include_comments: bool = False,
    search_text: str | None = None,
) -> InspectDocumentOrDocResult:
    """Read a document's headings, search paragraphs, or read one bounded window of text. A Word 97-2003
    .doc is read through a temporary .docx Word makes from it (`converted_from_doc`; the .doc is never
    written; to edit it, convert_document first)."""
    validate_path_in_root(path)
    with docx_for_reading(path, "inspect_document") as (docx, legacy):
        result = _inspect_document(
            docx,
            first_paragraph=first_paragraph,
            last_paragraph=last_paragraph,
            include_comments=include_comments,
            search_text=search_text,
        )
    return legacy.relabel(result, path_key="path", sha_key="source_sha256") if legacy else result


@server.tool(annotations=READ_ONLY, structured_output=True)
@_docx_only("path", read=True)
@_feedback("path")
def list_bookmarks(path: str) -> ListBookmarksResult:
    """List every bookmark in a document and every field referencing it."""
    validate_path_in_root(path)
    return {"bookmarks": [asdict(item) for item in inventory(path)]}


@server.tool(annotations=READ_ONLY, structured_output=True)
@_docx_only("path", read=True)
@_feedback("path")
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
@_docx_only("source")
@_feedback("source")
def clean_bookmarks(
    source: str,
    output: str,
    exact_names: list[str],
    source_sha256: SHA256 | None = None,
    tracked: bool | None = None,
) -> CleanBookmarksResult:
    """Delete only the supplied exact bookmark names at the requested output path; output may equal source for an in-place edit.

    Refused while the document's tracking switch is on (set_track_changes) unless tracked=False."""
    validate_path_in_root(source)
    validate_path_in_root(output)
    _refuse_while_tracking(source, "clean_bookmarks", tracked)
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
@_docx_only("source")
@_feedback("source")
def format_tables(
    source: str,
    output: str,
    table_specs: list[MCPTableFormatSpec],
    source_sha256: SHA256 | None = None,
    tracked: bool | None = None,
) -> TrackedFormatTablesResult:
    """Apply the fixed journal-neutral table policy at the requested output path; output may equal source for an in-place edit.

    While the document's tracking switch is on (set_track_changes), or with tracked=True, borders and alignment
    are recorded with Word's Track Changes by the switch author; cell merges can't be tracked and get one comment
    each (listed under `untracked`). tracked=False formats plainly."""
    validate_path_in_root(source)
    validate_path_in_root(output)
    from . import word_tracked

    tracking_on, tracking_author = word_tracked.tracking_for(source, tracked)
    try:
        _refuse_tracked_revisions(source, "format_tables")
    except RefusedUnsafe as exc:
        raise RefusedUnsafe(f"{exc}{word_tracked.REVISED_HINT}") from None
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
        with staged_document(source, output) as stage, word_tracked.input_copy(stage, tracking_on) as input_copy:
            saved = run_operation(
                "tables",
                {
                    "source": str(stage),
                    "output": str(stage),
                    "specs": [asdict(spec) for spec in specs],
                    "tracked": tracking_on,
                },
            )
            tracking = {**word_tracked.off_fields(), "compared_against": saved.get("compared_against", "source")}
            if tracking_on:
                from .tables import verify_saved_tables

                tracking = word_tracked.finish(
                    stage, input_copy, author=tracking_author,
                    merges=[(spec.index, group) for spec in specs for group in spec.merges],
                    verify_accept=lambda path, reference: verify_saved_tables(
                        audit_docx(reference), path, specs, reference=reference))
            after = audit_docx(stage)
    return {
        "source": str(Path(source).resolve()),
        "source_sha256": before.sha256,
        "output": str(Path(output).resolve()),
        "output_sha256": after.sha256,
        "formatted_tables": [spec.index for spec in specs],
        "protected_structure_unchanged": True,
        **tracking,
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
@_docx_only("source")
@_feedback("source")
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
    tracked: bool | None = None,
) -> InsertCitationResult:
    """Insert one live citation cluster at the requested output path and verify it.

    `anchor` is the exact visible text the citation is inserted immediately
    before; it must occur exactly once in `paragraph_index` (a duplicate
    elsewhere in the document is not a conflict). Word resolves the anchor's
    position itself rather than using a numeric offset, because a paragraph
    holding a field has Range coordinates that do not match its visible text.

    Bibliography donor and paragraph must either both be omitted or both be
    supplied. The source is never overwritten.

    Works on documents with tracked changes: the anchor must lie outside every
    revision (refused, named), and every existing revision must survive Word's
    save and the Zotero refresh, else nothing is written. While the document's
    tracking switch is on (set_track_changes), or with tracked=True, the new
    citation is one tracked insertion (w:ins) by the switch's author, so Reject
    removes it and Accept keeps it; the bibliography Zotero updated is not
    tracked, and `warnings` says to run refresh_citations after resolving a
    rejected citation. tracked=False inserts it untracked.
    """
    from .track_setting import read_setting

    validate_path_in_root(source)
    validate_path_in_root(output)
    validate_path_in_root(item_donor)
    if bibliography_donor is not None:
        validate_path_in_root(bibliography_donor)
    setting = read_setting(source)
    use_tracking = setting.on if tracked is None else bool(tracked)
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
            tracked=use_tracking,
            author=setting.author,
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
        "tracking": "on" if use_tracking else "off",
        "tracking_author": setting.author if use_tracking else None,
        "revision_id": raw.get("revision_id"),
        "revisions_checked": raw.get("revisions_checked"),
        "revision_check": raw.get("revision_check") or {},
        "warnings": raw.get("warnings") or [],
        **publication.write_fields(published, raw["source_sha256"], raw["output_sha256"]),
    }


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
@_docx_only("source")
@_feedback("source")
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
@_docx_only("source")
@_feedback("source")
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
@_docx_only("source")
@_feedback("source")
def set_citation_style(
    source: str,
    source_sha256: SHA256,
    style: str,
    output: str,
    preview: bool = False,
    locale: str | None = None,
    max_wait_seconds: float = 300.0,
) -> SetCitationStyleResult:
    """Change the document's Zotero citation style, then re-render every citation and the bibliography.

    `style` is an installed Zotero style's id, short name ("american-medical-association") or title; a style
    not installed in the local Zotero is refused, naming the closest installed ones. `locale` (e.g. "en-GB")
    changes the citation language; omitted keeps the document's. The new style is written into the document's
    Zotero preferences (ZOTERO_PREF_n custom properties, as Zotero's Document Preferences writes it; every other
    property and part unchanged), then the document goes through refresh_citations: same guards, same result
    (citations_changed, bibliography diff, revision check), plus style_before/style_after. `preview=True` does
    all of it on a temporary copy and writes nothing. `output == source` changes it in place (undo_document
    reverts it) or `output` must not exist. Refused: a citation inside a tracked revision (as refresh is), a
    Zotero field outside the main body or stored as a simple field (the refresh would not check it), a
    note-style document, a switch between an in-text and a note style (use Zotero's Document Preferences in
    Word), a style with no bibliography for a document that has one, and a document already using that style
    (call refresh_citations).
    """
    validate_path_in_root(source)
    if not preview:
        validate_path_in_root(output)
    with publication.capture(tool="set_citation_style", op="set citation style") as published:
        raw = set_style_to_version(source, output, source_sha256=source_sha256, style=style, locale=locale,
                                   preview=preview, max_wait_seconds=max_wait_seconds)
    if preview:
        return {**raw, **publication.write_fields([], None, None)}
    return {**raw, **publication.write_fields(published, raw["source_sha256"], raw["output_sha256"])}


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
@_docx_only("source")
@_feedback("source")
def edit_document(
    source: str,
    source_sha256: SHA256,
    operations: list[MCPEditOperation],
    output: str,
    allow_partial: bool = False,
    track_changes: bool | None = None,
) -> EditBatchResult:
    """Apply up to 20 verified Word edits atomically by default.

    Tracked changes: set_track_changes (on: true/false, optional author,
    default "Claude") turns on research-mcp's tracking switch for the
    document (kept in the file; not Word's own Track Changes setting;
    track_changes=true on this call is the same as a set_track_changes op
    first). While it is on, replace_text, replace_paragraph and insert_text
    are written as tracked changes (w:del/w:ins) by the switch's author (an
    op's author overrides it), delete_citation follows it, insert_paragraph
    adds each new paragraph as a tracked insertion (w:ins on its runs and
    paragraph mark), replace_table_row and replace_section record their
    changed lines as minimal word diffs, added lines as tracked insertions
    and removed lines as tracked deletions (text in w:del, mark deleted; the
    last line of a cell Word's way: the line before's mark is deleted), and
    a table row added or removed by replace_section(rows=...) as Word's
    tracked row insertion/deletion (w:ins/w:del in w:trPr, every run and
    paragraph mark marked), style_text (paragraph_index, text, bold and/or
    italic) records the formatting as a tracked formatting change
    (w:rPrChange; in the author's own pending insertion, in place), and
    insert_table writes the caption and table as one tracked insertion (direct XML; merges are applied plainly with one comment each). tracked on an op overrides the switch for that op. A set_track_changes applies to the ops after it in
    the batch. Results carry `tracking` ("on"/"off") and `tracking_author`;
    with tracking off on a document that has tracked changes,
    `tracking_note` says so. (tracked_replace_text/_paragraph remain as
    aliases of a tracked replace_text/replace_paragraph.)
    Other revision ops: list_revisions, accept_revisions, reject_revisions
    and sweep_characters go in operations[].type.

    delete_citation: paragraph_index plus citation_id or current_text (the
    rendered text, e.g. "22"; find is taken as current_text; with neither the
    refusal lists the paragraph's citations); removes that one whole Zotero citation as a
    tracked deletion by author (default "Claude", or the switch's author;
    tracked=false, or the switch set off, removes it outright), taking one space before it along when punctuation, a space or
    the paragraph end follows ("legend 22." -> "legend."). Works on documents
    with tracked changes; Reject gives the live citation back. No Zotero
    refresh: details list orphaned_items (cited nowhere else) with their
    bibliography entries; they leave the bibliography at the next
    refresh_citations after the deletion is accepted. Refused: a citation
    inside a revision, holding a nested field, spanning paragraphs, inside a
    hyperlink, sharing a run with text outside it, with a comment/bookmark
    marker inside it, or stored as a simple field (w:fldSimple, not written by
    Zotero's Word plugin).

    delete_paragraph: paragraph_index (and through_paragraph_index for a
    range in the body or in one table cell) deletes whole paragraphs, empty
    ones included, as direct XML on any document. Tracked by default (the
    switch when set, else tracked by "Claude"): text and paragraph mark in
    w:del, indices unchanged until accepted (accepted_index_delta);
    tracked=false removes them (index_delta, later ops shifted). A paragraph
    whose text is already a tracked deletion can be deleted (the old w:del
    stays). The last line(s) of a cell go Word's way; a cell keeps one
    paragraph. Refused, named: a range across cells, a w:ins, move or
    formatting revision, a mark that is already a revision, a section break,
    a field, drawing, hyperlink, content control, bookmark/comment marker or
    footnote reference.

    format_paragraph: paragraph_index (through_paragraph_index for a range
    in the body or one table cell) with font, size (pt), left_indent,
    right_indent, first_line_indent (pt; negative = hanging), space_before,
    space_after (pt), remove_numbering, and/or like_paragraph (copy that
    paragraph's indents, spacing, numbering and main font/size; explicit
    fields win). Font and size go on every run and the paragraph mark;
    deleted text and field results keep theirs. Direct XML on any document;
    tracked (switch or tracked: true) as w:pPrChange and w:rPrChange by the
    author. Not bold/italic (style_text). Refused, named: another author's
    formatting change, a nested-table paragraph, a range across cells,
    nothing to change. When replace_paragraph/replace_text fills an empty
    paragraph that is numbered or whose size is 2 pt or more off the body
    text, the op result's format_note names the format_paragraph call that
    matches it; nothing is changed by itself.

    Obtain source_sha256 from inspect_document. Each operation runs in an
    isolated version. Any failure discards the batch unless allow_partial=True.
    The result gives per-operation status and the exact source/hash for retries,
    plus undo_token / sha256_before / sha256_after: undo_document(path) restores
    the state before an in-place batch in one call.

    Each type reads only its own fields; the rest of the envelope is ignored.
    replace_text: paragraph_index, find (up to 5000 characters), replace. find may be a whole line
    spanning bold/italic runs: only the characters that differ are rewritten,
    unchanged characters keep their formatting, and new text takes the
    formatting of the character it replaces or extends. Inside caps/small-caps
    runs, find matches in any case ("WORK EXPERIENCE" finds stored
    "WorK Experience") and the stored case is kept.
    replace_paragraph: paragraph_index, text (the whole new paragraph; same
    minimal-diff rule).
    replace_table_row: match (one cell's text, or one line of it, unique in
    the document), cells (new text per cell, "\n" between paragraphs of a
    cell; cell count must stay the same); optional table_index, tracked,
    author. While tracking is on, or on a document with tracked changes, it
    is direct XML (Word never opens): tracked while on, plain when off (a
    line added after a cell's last line is then a plain new paragraph).
    replace_section: heading (as displayed) and rows (cell texts per row, for
    a section laid out as table rows) or paragraphs (for body text); optional
    until_heading. Rows/paragraphs may be added or removed inside the section
    only; everything outside is verified unchanged, cell counts are kept and
    date cells must still read as dates. tracked/author and the direct-XML
    path as for replace_table_row; a row added or removed is a tracked row
    insertion/deletion as Word writes it while tracking, and a plain direct-XML
    copy/removal with tracking off on a document with tracked changes (refused
    for a row with a tracked revision, merged cells, a nested table or
    non-text content, and for a table's only row).
    group (any type): ops sharing a group, a table row or a section are
    all-or-nothing together under allow_partial (a row named by text an
    earlier op wrote joins that row's group; a row/section op whose target is
    found nowhere refuses the partial publish); the result's `groups` lists
    each group as applied/rolled_back/failed, and `undo_call` is the one call
    that restores the pre-batch state.
    insert_text: paragraph_index,
    anchor, position, text. insert_paragraph: after_paragraph_index, text,
    style optional ("Heading 3", "Heading3" and "heading 3" all resolve to
    the document's style). Two paths. While tracking is on, or on a document
    with tracked changes, it is direct XML (Word never opens): each line of
    text ("\\n") is one new paragraph, tracked while on and plain when off;
    each copies the anchor's paragraph properties (minus section break and
    property-change marks) and the anchor mark's run formatting, or takes
    the given style with no run formatting; the result is checked against
    the input document (place, properties, formatting, text; Reject gives
    the input back). On a revision-free document with tracking off it goes
    through Word: one paragraph only ("\\n" refused), not inside a table;
    Word splits the anchor's mark (its paragraph formatting carries over)
    and the new paragraph gets the given style, else Normal; text, style and
    tables are verified after Word's save. insert_table: after_paragraph_index,
    headers, rows, descriptive_columns (all required); optional caption, caption_style, merges, tracked, author.
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
    disposition (addressed needs depends_on = the op_id of the edit in this
    batch, or addressed_by = a revision id from list_revisions or an
    undo_token of an earlier edit still in effect; rejected, no_change,
    needs_information are written with their label, "No change: ...";
    reply writes reply_text with no prefix), reply_text (surrounding
    whitespace trimmed); resolve=true also
    resolves the thread. A resolved thread must be reopened before a plain
    reply. resolve_comment / reopen_comment: comment_id (thread and replies).
    delete_comment: comment_id; include_replies (default true) deletes the
    thread with its replies; include_replies=false with a reply's id deletes
    only that reply. Comment ops never change manuscript prose.

    Tracked text edits (replace_text, replace_paragraph and insert_text with
    tracking on; direct XML, never through Word, so other authors'
    revisions stay byte-identical): find is matched once in the accepted
    text, replace_paragraph's text is the whole accepted paragraph.
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
    does. A tracked row (a w:ins/w:del in the row's trPr) is removed whole
    when its insertion is rejected or its deletion accepted, as Word does,
    if everything in it goes with the same selection (result rows_removed;
    its paragraphs and row-end marker shift later indices); otherwise refused,
    naming the row and the revision or text that keeps it (select it too),
    as is a row holding a bookmark or comment marker, a nested table, or a
    table's last row. Refused: table-cell revisions, rejecting section or
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
    On a document with tracked revisions and tracking off, replace_text,
    replace_paragraph and insert_text are made as untracked direct XML and
    refused where they touch a revised run; the other Word-based edits are
    refused.

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
            raise RefusedUnsafe(_operation_error_text(op, exc.errors()[0])) from exc
    if track_changes is not None:
        # Plan 864: edit_document(track_changes=...) is a set_track_changes op ahead of the others.
        taken = {op.op_id for op in strict_ops}
        switch_id = next(name for n in range(1, len(taken) + 2)
                         if (name := "track_changes" if n == 1 else f"track_changes_{n}") not in taken)
        strict_ops.insert(0, SetTrackChangesOperation(op_id=switch_id, type="set_track_changes", on=track_changes))
        # Validation errors name operations by position in strict_ops; keep the caller's list aligned.
        operations = [MCPEditOperation(op_id=switch_id, type="set_track_changes", on=track_changes), *operations]

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


def _refuse_while_tracking(source: str, tool: str, tracked: bool | None) -> None:
    """Plan 864: a tool that cannot record its change as a tracked change refuses while the document's
    tracking switch is on (or tracked=True is asked), unless the call says tracked=False."""

    if tracked is False:
        return
    from .track_setting import not_trackable, read_setting

    if tracked or read_setting(source).on:
        raise not_trackable(tool)


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
@_docx_only("source")
@_feedback("source")
def apply_document_style(
    source: str,
    output: str,
    source_sha256: SHA256,
    profile: Literal["radiology_ai", "editing_compact", "compact_single", "standard_double"] = "radiology_ai",
    emphasis_policy: Literal["preserve", "headings_and_summary"] = "preserve",
    preview_only: bool = False,
    tracked: bool | None = None,
) -> dict[str, Any]:
    """Apply journal typography and layout policies at the requested output path.

    Use preview_only=True to inspect role decisions and effects without writing.
    Layout preserves existing bold/italic by default. Only the explicit
    headings_and_summary policy adds whole-paragraph bold to those roles.
    For phrase emphasis use style_text, not this document-wide tool.
    Profile settings are typography presets, not full journal compliance.
    While the document's tracking switch is on (set_track_changes), or with
    tracked=True, the change is recorded with Word's Track Changes by the
    switch author; what Word can't track gets one comment each (listed under
    `untracked`). tracked=False styles plainly.
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
    from . import word_tracked

    tracking_on, tracking_author = word_tracked.tracking_for(source, tracked)
    try:
        _refuse_tracked_revisions(source, "apply_document_style")
    except RefusedUnsafe as exc:
        raise RefusedUnsafe(f"{exc}{word_tracked.REVISED_HINT}") from None
    with publication.capture(tool="apply_document_style", op=f"profile {profile}") as published:
        result = apply_journal_style(source, output, profile_name=profile, emphasis_policy=emphasis_policy, source_sha256=source_sha256,
                                     tracked=tracking_on, author=tracking_author)
    return {**result, **publication.write_fields(published, result["source_sha256"], result["output_sha256"])}


@server.tool(annotations=READ_ONLY, structured_output=True)
@_docx_only("path", read=True)
@_feedback("path")
def audit_document_style(path: str, profile: str | None = None, first_paragraph: int = 1, last_paragraph: int | None = None, search_text: str | None = None) -> dict[str, Any]:
    """Inspect effective emphasis spans and flag suspicious styling without changing Word.

    Choose a profile to check saved typography; findings are guardrails, not a
    declaration of journal compliance. Use paragraph indices with style_text.
    """
    validate_path_in_root(path)
    from .style_audit import audit_document_style as audit_style
    return audit_style(path, profile=profile, first_paragraph=first_paragraph, last_paragraph=last_paragraph, search_text=search_text)


@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
@_docx_only("source")
@_feedback("source")
def style_text(source: str, source_sha256: SHA256, paragraph_index: int, text: str | None = None, bold: bool | None = None, italic: bool | None = None, output: str | None = None, preview_only: bool = True, allow_whole_paragraph: bool = False, tracked: bool | None = None, author: str | None = None, find: str | None = None) -> dict[str, Any]:
    """Preview or apply bold/italic to one exact phrase, preserving emphasis elsewhere.

    Defaults to a read-only preview. After checking the selected text, use
    preview_only=False to apply an already-authorized change. Text must match
    once in the paragraph's accepted text (`find` is another name for `text`;
    give one). Whole-paragraph selection requires
    allow_whole_paragraph=True. Omitted properties stay unchanged.
    Field-overlapping selections are refused.
    Tracked (tracked=True, or the document's tracking switch when tracked is
    not given): each run gets the new formatting plus a tracked formatting
    change (w:rPrChange) by `author` (else the switch's author, else
    "Claude"); text inside the author's own pending insertion is formatted in
    place; a run already carrying the author's own formatting change keeps it
    (split as Word does when the phrase covers part of it); another author's
    formatting change is refused, and so is untracked styling of text that
    carries any tracked formatting change. Tracked, or on a document with tracked changes: direct XML (Word
    never opens; every other revision unchanged). Otherwise Word formats it.
    The preview says `engine`, `tracked` and `author`.
    Output defaults to source with staged verification; undo_document undoes it.
    """
    validate_path_in_root(source)
    if output is not None:
        validate_path_in_root(output)
    if text is not None and find is not None:
        raise RefusedUnsafe("give text or find, not both; find is another name for text")
    if text is None:
        text = find
    if not text:
        raise RefusedUnsafe("give text (the phrase to style); find is another name for it")
    from .style_tools import style_text as apply_inline
    with publication.capture(tool="style_text", op=f"style paragraph {paragraph_index}") as published:
        result = apply_inline(source, source_sha256, paragraph_index, text, bold=bold, italic=italic, output=output, preview_only=preview_only, allow_whole_paragraph=allow_whole_paragraph, tracked=tracked, author=author)
    if preview_only:
        return result
    return {**result, **publication.write_fields(published, result.get("source_sha256"), result.get("output_sha256"))}


# Not read-only: it writes PNGs into ``output_dir`` (overwriting page-N.png there), so it is left out of
# read-only modes such as `research-mcp chatgpt --read-only`.
@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
@_feedback("path")
def render_document_pages(
    path: str,
    output_dir: str | None = None,
    pages: list[int] | None = None,
    dpi: int = 200,
    markup: bool = False,
) -> RenderDocumentOrDocResult:
    """Render pages of a .docx file into print-quality PNG images.

    Default (markup=False): Word's clean print layout (no redlines, no comment
    balloons), rasterized at `dpi`. markup=True exports the document with its
    markup (Word's "All Markup": tracked insertions and deletions as redlines,
    comments and formatting changes in balloons in the margin), so a review
    round can be checked by eye; if Word will not set that view, `warnings`
    says the markup may be missing. The default output folder is
    `<stem>_rendered_<sha8>` (`<stem>_rendered_markup_<sha8>` with markup).
    A Word 97-2003 .doc is rendered through a temporary .docx Word makes from
    it (`converted_from_doc`; the folder is beside the .doc, named by its sha).
    """
    validate_path_in_root(path)
    if output_dir is not None:
        validate_path_in_root(output_dir)
    from .operation_process import run_operation
    with docx_for_reading(path, "render_document_pages") as (docx, legacy):
        if legacy is None:
            return run_operation("render", {"docx_path": str(docx), "output_dir": output_dir, "pages": pages,
                                            "dpi": dpi, "markup": bool(markup)}, timeout_seconds=180)
        # A .doc: Word renders the temporary .docx into the private folder; the pages are copied beside the
        # .doc only after its hash is checked against the bytes Word converted.
        if output_dir is None:
            kind = "rendered_markup" if markup else "rendered"
            output_dir = str(legacy.path.parent / f"{legacy.path.stem}_{kind}_{legacy.sha256[:8]}")
            validate_path_in_root(output_dir)
        staged = run_operation("render", {"docx_path": str(docx), "output_dir": str(docx.parent / "pages"),
                                          "pages": pages, "dpi": dpi, "markup": bool(markup)}, timeout_seconds=180)
        result = legacy.publish_pages(staged, Path(output_dir).resolve())
    return legacy.relabel(result, path_key="source")


# Not read-only: it writes a new PDF (never over an existing file; the .docx is untouched), so it is left
# out of read-only modes such as `research-mcp chatgpt --read-only`.
@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
@_feedback("path")
def export_pdf(path: str, output: str | None = None) -> ExportPdfOrDocResult:
    """Save a PDF of a .docx: Word's clean print layout (no comment balloons), vector, heading bookmarks.

    `output` defaults to `<document name>.pdf` beside the .docx; another path must end in .pdf and its
    folder must exist. An existing file is never overwritten (the call refuses; pick another `output`).
    The .docx is opened read-only and left unchanged (`source_sha256`). Tracked changes appear accepted
    (final text, no markup) and comments are not shown; `warnings` say so, and say when the document is
    open in Word (the PDF is the saved file, without unsaved edits). `page_count` is counted in the PDF.
    A Word 97-2003 .doc is exported through a temporary .docx Word makes from it (`converted_from_doc`;
    the PDF goes beside the .doc).
    """
    validate_path_in_root(path)
    if output is not None:
        validate_path_in_root(output)
    _, target = resolve_pdf_target(path, output)  # refuse early, before starting Word
    validate_path_in_root(target)
    from .operation_process import run_operation
    with docx_for_reading(path, "export_pdf") as (docx, legacy):
        if legacy is None:
            return run_operation("export_pdf", {"docx_path": str(docx), "output": output}, timeout_seconds=180)
        # A .doc: Word exports the temporary .docx into the private folder; the PDF is published beside the
        # .doc only after its hash is checked against the bytes Word converted, right before the rename.
        staged = docx.with_suffix(".pdf")
        result = run_operation("export_pdf", {"docx_path": str(docx), "output": str(staged)}, timeout_seconds=180)
        legacy.publish_file(staged, target)
    result = legacy.relabel({**result, "output": str(target), "size_bytes": target.stat().st_size},
                            path_key="source", sha_key="source_sha256")
    if is_file_locked(legacy.path)[0]:
        result["warnings"] = [*result.get("warnings", []),
                              "the document is open in Word: the PDF shows the saved file, without unsaved changes"]
    return result


# Writes a new .docx beside the .doc (never over an existing file); the .doc is untouched and no undo entry is
# made, since nothing that existed changes.
@server.tool(annotations=VERSIONED_WRITE, structured_output=True)
def convert_document(path: str, output: str | None = None) -> ConvertDocumentResult:
    """Save a Word 97-2003 .doc as a new .docx that every tool can edit.

    Word opens the .doc read-only and saves `<name>.docx` beside it, or at `output` (must end in .docx,
    folder must exist). An existing file is never overwritten (the call refuses; pick another `output`), and
    the .doc is left unchanged (`source_sha256`). A .doc named .docx needs `output`. Password-protected
    documents and documents with macros are refused. Returns the new file's path and sha256; edit that file.
    """
    validate_path_in_root(path)
    if output is not None:
        validate_path_in_root(output)
    _, target = resolve_convert_target(path, output)  # refuse early, before starting Word
    validate_path_in_root(target)
    return convert_in_worker(path, target)


@server.tool(annotations=UNDO_WRITE, structured_output=True)
@_docx_only("path")
@_feedback("path", via="model")
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
@_docx_only("path")
@_feedback("path", via="model")
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
@_docx_only("path")
@_feedback("path")
def checkpoint_document(path: str, label: str) -> CheckpointDocumentResult:
    """Store the document's current state under a label in the undo store.

    undo_document(path, to=label) returns to it later. The document is not
    changed. Stored on this PC only, outside OneDrive; retention limits apply.
    """
    validate_path_in_root(path)
    return publication.checkpoint_document(path, label)


@server.tool(annotations=READ_ONLY, structured_output=True)
@_feedback("path")
def document_history(path: str) -> DocumentHistoryResult:
    """List a document's recorded edits, checkpoints, undos and outside changes, newest first.

    Each entry id is an undo_token. Also lists pruned entries (dropped, with the
    reason), the retention limits, and whether the document changed outside the
    server since the last record.
    """
    validate_path_in_root(path)
    return publication.document_history(path)


@server.tool(annotations=READ_ONLY, structured_output=True)
def edit_feedback(
    path: str | None = None,
    scope: Literal["document", "all"] = "document",
    query: str | None = None,
    paragraph_index: int | None = None,
    search_text: str | None = None,
    outcome: Literal["accepted", "rejected", "reworded", "removed"] | None = None,
    limit: int = 10,
    record_id: str | None = None,
    detail: Literal["net", "rounds"] = "net",
) -> EditFeedbackResult:
    """What the user did with your earlier edits: accepted, rejected, reworded, or removed.

    Default (path): this document's outcomes you have not seen yet, newest first, at most `limit` (10),
    one line each (<= 300 chars): first proposal -> what the user kept, with `rounds` when you edited
    the same span several times. Pending proposals are only counted. Filters (paragraph_index,
    search_text, outcome) list every matching outcome instead. Only when needed: scope="all" with
    `query` searches every document's log; record_id + detail="rounds" shows each round of one item.
    Reads the document only to settle outcomes; writes only the feedback log (seen marks), outside
    OneDrive (RESEARCH_MCP_FEEDBACK_DIR, default ~/research_mcp/edit_feedback).
    """
    from . import edit_feedback as fb
    from .edit_feedback_word import doc_log, feedback_root, on_sight

    try:
        root = feedback_root()
    except Exception as exc:  # noqa: BLE001 - a refused or unreadable folder is a warning, never a tool error
        return EditFeedbackResult(path=path, scope=scope, feedback_warning=f"{exc}"[:300] or type(exc).__name__)
    limit = max(1, min(int(limit), 50))
    result = EditFeedbackResult(path=path, scope=scope, folder=str(root))

    def store_failed(exc: Exception) -> EditFeedbackResult:
        # the feedback store (log read, lock, seen marks, folder) failing is a warning, never a tool error
        note = f"edit feedback store unavailable: {type(exc).__name__}: {exc}"[:300]
        result.feedback_warning = " ".join(x for x in (result.feedback_warning, note) if x)
        return result

    if path is not None:
        validate_path_in_root(path)
        fields = on_sight(path)
        result.edit_feedback = fields.get("edit_feedback")
        result.feedback_warning = fields.get("feedback_warning")
    if record_id:
        try:
            log = doc_log(path) if path is not None and scope != "all" else None
            found = fb.rounds(root, record_id, log)
        except Exception as exc:  # noqa: BLE001
            return store_failed(exc)
        if found is None:
            raise RefusedUnsafe(f"no edit feedback record {record_id!r}" + (" for this document" if log else ""))
        result.record_id = record_id
        result.items = [found["item"]]
        if detail == "rounds":
            result.rounds = found["rounds"]
        return result
    if scope == "all":
        try:
            found = fb.search_all(root, query=query or search_text, outcome=outcome, limit=limit)
        except Exception as exc:  # noqa: BLE001
            return store_failed(exc)
        result.items, result.counts, result.more = found["items"], found["counts"], found["more"]
        return result
    if path is None:
        raise RefusedUnsafe('edit_feedback needs path, or scope="all" (with query) to search every document')
    try:
        log = doc_log(path)
        result.md_path = str(log.md) if log.md.exists() else None
        if not log.exists():
            result.note = "No model edits recorded for this document yet."
            return result
        found = fb.view(log, label=Path(path).name, paragraph_index=paragraph_index,
                        search_text=search_text or query, outcome=outcome, limit=limit)
    except Exception as exc:  # noqa: BLE001
        return store_failed(exc)
    result.items, result.counts = found["items"], found["counts"]
    result.pending, result.more = found["pending"], found["more"]
    result.note = found.get("note")
    if found.get("warning"):
        result.feedback_warning = " ".join(x for x in (result.feedback_warning, found["warning"]) if x)
    return result


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
    create_collection: bool = False,
) -> AddZoteroItemResult:
    """Add a reference to the local Zotero library by its DOI.

    Metadata comes from CrossRef; arXiv DOIs (10.48550/arXiv.*) come from
    DataCite and are added as preprints with the arXiv PDF attached. Files into
    `collection` (or the `.zotero-collection` marker found from `path`); a
    missing collection is refused unless `create_collection=True`, and a name
    close to an existing one is refused even then.
    """
    if path is not None:
        validate_path_in_root(path)
    return add_item_by_doi(
        doi=doi, path=path, collection=collection, create_collection=create_collection
    )


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


from . import stale_exit as _stale_exit  # noqa: E402


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
    parser.add_argument(
        "--exit-when-stale",
        action="store_true",
        help=(
            "HTTP only: exit with code 75 once idle when this server's git checkout gets new commits on master "
            "that change src/, pyproject.toml or uv.lock (and those files are committed), so a supervisor loop "
            "restarts it on the new code. Ignored under stdio."
        ),
    )
    parser.add_argument(
        "--stale-idle-seconds",
        type=float,
        default=_stale_exit.DEFAULT_IDLE_SECONDS,
        help=(
            "With --exit-when-stale: how long no tool call may have started or finished (and --idle-log not "
            f"changed) before exiting (default: {_stale_exit.DEFAULT_IDLE_SECONDS})."
        ),
    )
    parser.add_argument(
        "--idle-log",
        type=Path,
        default=None,
        help="With --exit-when-stale: also wait until this file (e.g. the supervisor's log) is that old.",
    )
    return parser


def _refresh_word_worker_hook():
    """``before_exit`` for --exit-when-stale: when the new commits change the Word worker's code, ask the
    shared worker to shut down cleanly so the next call starts it on the new code. A busy worker declines the
    exit for this round (retried at the next check). The worker's code is word_worker.py and every package
    module it imports, read from the files at start and again at the check."""
    from .errors import IntegrationBusy
    from .word_service import shutdown_word_worker

    worker = Path(__file__).resolve().with_name("word_worker.py")
    code_at_start = _stale_exit.import_closure(worker)

    def before_exit(repo: Path, changed: list[str]) -> bool:
        code = code_at_start | _stale_exit.import_closure(worker)
        hit = sorted(f for f in changed if (repo / f).resolve() in code)
        if not hit:
            return True
        try:
            shutdown_word_worker()
        except IntegrationBusy as exc:
            _stale_exit.log_line(f"Word worker busy ({exc}); exit retried at the next check")
            return False
        _stale_exit.log_line(f"Word worker asked to shut down (its code changed: {', '.join(hit)})")
        return True

    return before_exit


def start_stale_exit(args: argparse.Namespace) -> "_stale_exit.StaleExitWatcher | None":
    """Start the --exit-when-stale watcher for an HTTP server; None when the flag is off or under stdio."""
    if not args.exit_when_stale:
        return None
    if args.transport == "stdio":
        _stale_exit.log_line("--exit-when-stale is ignored under stdio (the client would lose its server)")
        return None
    from .checkpoints import home

    return _stale_exit.StaleExitWatcher(
        __file__,
        tracker=TOOL_ACTIVITY,
        state_dir=home,
        idle_seconds=args.stale_idle_seconds,
        idle_log=args.idle_log,
        before_exit=_refresh_word_worker_hook(),
    ).start()


def main(argv: list[str] | None = None, *, prog: str | None = None) -> None:
    """Run the server (stdio by default). ``argv`` defaults to ``sys.argv[1:]``."""
    parser = build_parser(prog)
    args = parser.parse_args(argv)
    if args.root_dir:
        set_root_dir(args.root_dir)
    start_stale_exit(args)
    if args.transport == "stdio":
        server.run(transport="stdio")
    elif args.transport == "streamable-http":
        server.run(transport="streamable-http", host=args.host, port=args.port)
    elif args.transport == "http":
        server.run(transport="http", host=args.host, port=args.port)


if __name__ == "__main__":
    main()
