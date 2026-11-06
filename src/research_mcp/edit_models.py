"""Data models for batch document editing requests, plans, and results."""

from __future__ import annotations

import re
from typing import Annotated, Any, Literal, Union
from pydantic import BaseModel, ConfigDict, Field, StringConstraints, model_validator


SHA256 = Annotated[str, StringConstraints(pattern=r"^[0-9a-fA-F]{64}$")]
DEFAULT_MAX_OPERATIONS = 20
SAFE_MAX_PARAGRAPHS = 100


GROUP_HELP = (
    "Operations sharing a group are all-or-nothing together under allow_partial. Operations on the same "
    "table row or inside the same replace_section are grouped automatically."
)


KEEP_HELP = (
    'Characters to leave as written (each a character or "U+XXXX"). New text is otherwise cleaned with '
    "textclean (non-breaking/odd spaces, zero-width and control characters, ...); the result's `cleaned` "
    "lists what changed."
)


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class ReplaceTextOperation(StrictModel):
    op_id: str
    type: Literal["replace_text"]
    paragraph_index: int = Field(ge=1)
    find: str = Field(min_length=1, max_length=255)
    replace: str
    match_case: bool = True
    occurrence: int = Field(default=1, ge=1)
    expected_count: int = Field(default=1, ge=1)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class InsertTextOperation(StrictModel):
    op_id: str
    type: Literal["insert_text"]
    paragraph_index: int = Field(ge=1)
    position: Literal["before", "after"]
    anchor: str = Field(min_length=1, max_length=255)
    text: str = Field(min_length=1)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class InsertParagraphOperation(StrictModel):
    op_id: str
    type: Literal["insert_paragraph"]
    after_paragraph_index: int = Field(ge=0)
    text: str = Field(min_length=1)
    style: str | None = None
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class InsertTableOperation(StrictModel):
    op_id: str
    type: Literal["insert_table"]
    after_paragraph_index: int = Field(ge=0)
    headers: list[str] = Field(min_length=1)
    rows: list[list[str]] = Field(min_length=1)
    descriptive_columns: list[int] = Field(min_length=1)
    merges: list[dict[str, Any]] = Field(default_factory=list)
    style: str | None = None
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="after")
    def validate_table(self) -> InsertTableOperation:
        expected = len(self.headers)
        for row in self.rows:
            if len(row) != expected:
                raise ValueError(f"every row must have the same width as headers ({expected})")
        for col in self.descriptive_columns:
            if col < 1 or col > expected:
                raise ValueError(f"descriptive column {col} outside table width {expected}")
        for m in self.merges:
            col = m.get("column", 1) if isinstance(m, dict) else getattr(m, "column", 1)
            start_row = m.get("start_row", 2) if isinstance(m, dict) else getattr(m, "start_row", 2)
            end_row = m.get("end_row", 2) if isinstance(m, dict) else getattr(m, "end_row", 2)
            label = m.get("label", "") if isinstance(m, dict) else getattr(m, "label", "")
            for r in range(start_row, end_row + 1):
                row_idx = r - 2
                if 0 <= row_idx < len(self.rows) and 0 <= col - 1 < len(self.rows[row_idx]):
                    if self.rows[row_idx][col - 1] != label:
                        raise ValueError(f"cells in merge must exactly match label {label!r}")
        return self


AUTHOR_HELP = ("Comment author shown in Word (default: RESEARCH_MCP_AUTHOR from the config, else Word's own "
               "user name, else \"Author\").")
INITIALS_HELP = ("Initials shown in Word (default: the configured or Word initials for the default author, else "
                 "the author's initials).")
COMMENT_ID_HELP = ("The thread's comment_id from inspect_document (durable id, e.g. 0A5777F8); a legacy id "
                   "(legacy:5 or 5) is accepted when it names exactly one comment.")


class AddCommentOperation(StrictModel):
    op_id: str
    type: Literal["add_comment"]
    paragraph_index: int = Field(ge=1)
    anchor: str = Field(min_length=1, max_length=255)
    text: str = Field(min_length=1)
    author: str | None = Field(default=None, min_length=1, description=AUTHOR_HELP)
    initials: str | None = Field(default=None, description=INITIALS_HELP)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


CommentDisposition = Literal["addressed", "rejected", "no_change", "needs_information"]


class ReplyCommentOperation(StrictModel):
    op_id: str
    type: Literal["reply_comment"]
    comment_id: str = Field(min_length=1, description=COMMENT_ID_HELP)
    disposition: CommentDisposition = "addressed"
    reply_text: str = Field(min_length=1)
    resolve: bool = Field(default=False, description="Also mark the thread resolved (done) in the same write.")
    author: str | None = Field(default=None, min_length=1, description=AUTHOR_HELP)
    initials: str | None = Field(default=None, description=INITIALS_HELP)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class ResolveCommentOperation(StrictModel):
    """Mark a comment thread resolved (Word's Resolve): the thread and its replies get done=1."""

    op_id: str
    type: Literal["resolve_comment"]
    comment_id: str = Field(min_length=1, description=COMMENT_ID_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class ReopenCommentOperation(StrictModel):
    """Reopen a resolved comment thread (Word's Reopen): the thread and its replies get done=0."""

    op_id: str
    type: Literal["reopen_comment"]
    comment_id: str = Field(min_length=1, description=COMMENT_ID_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class DeleteCommentOperation(StrictModel):
    """Delete a comment thread (its markers and every part entry); include_replies=false deletes one reply."""

    op_id: str
    type: Literal["delete_comment"]
    comment_id: str = Field(min_length=1, description=COMMENT_ID_HELP)
    include_replies: bool = Field(
        default=True,
        description="true (default): a thread id deletes the thread with all its replies. false: delete only "
                    "the given reply (a thread id with replies is refused).")
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class ReplaceParagraphOperation(StrictModel):
    """Make one paragraph read exactly `text`; only differing characters change (D8)."""

    op_id: str
    type: Literal["replace_paragraph"]
    paragraph_index: int = Field(ge=1)
    text: str
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class ReplaceTableRowOperation(StrictModel):
    """Rewrite one table row located by any cell's text (or one line of a cell)."""

    op_id: str
    type: Literal["replace_table_row"]
    match: str = Field(min_length=1)
    cells: list[str] = Field(min_length=1, description='New text per cell; "\n" separates paragraphs in a cell.')
    table_index: int | None = Field(default=None, ge=1)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class ReplaceSectionOperation(StrictModel):
    """Rewrite the rows or paragraphs under a heading, up to the next heading of the same kind."""

    op_id: str
    type: Literal["replace_section"]
    heading: str = Field(min_length=1)
    rows: list[list[str]] | None = None
    paragraphs: list[str] | None = None
    until_heading: str | None = Field(default=None, min_length=1)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="after")
    def validate_content(self) -> ReplaceSectionOperation:
        if (self.rows is None) == (self.paragraphs is None):
            raise ValueError("give exactly one of rows (table sections) or paragraphs (body sections)")
        if self.rows is not None and (not self.rows or any(not row for row in self.rows)):
            raise ValueError("rows must be a non-empty list of non-empty cell lists")
        if self.paragraphs is not None and not self.paragraphs:
            raise ValueError("paragraphs must not be empty")
        return self


REVISION_AUTHOR_HELP = 'Author recorded on the new w:ins/w:del; default "Claude".'
SELECTOR_HELP = (
    "Revision selectors; every one given must match. ids: w:id values. author: exact author name. "
    "paragraph_index..through_paragraph_index: a layout paragraph range. heading (+ until_heading): a section."
)


class TrackedReplaceTextOperation(StrictModel):
    """Replace `find` (once, in the accepted text) as tracked changes: w:del of the old words, w:ins of the new."""

    op_id: str
    type: Literal["tracked_replace_text"]
    paragraph_index: int = Field(ge=1)
    find: str = Field(min_length=1, max_length=255)
    replace: str
    author: str = Field(default="Claude", min_length=1, description=REVISION_AUTHOR_HELP)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class TrackedReplaceParagraphOperation(StrictModel):
    """Make one paragraph's accepted text read `text`; the minimal diff is recorded as tracked changes."""

    op_id: str
    type: Literal["tracked_replace_paragraph"]
    paragraph_index: int = Field(ge=1)
    text: str
    author: str = Field(default="Claude", min_length=1, description=REVISION_AUTHOR_HELP)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class _RevisionSelection(StrictModel):
    op_id: str
    ids: list[int] | None = Field(default=None, min_length=1, description=SELECTOR_HELP)
    author: str | None = Field(default=None, min_length=1)
    paragraph_index: int | None = Field(default=None, ge=1)
    through_paragraph_index: int | None = Field(default=None, ge=1)
    heading: str | None = Field(default=None, min_length=1)
    until_heading: str | None = Field(default=None, min_length=1)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="after")
    def validate_selection(self):
        if self.until_heading is not None and self.heading is None:
            raise ValueError("until_heading needs heading")
        return self


class AcceptRevisionsOperation(_RevisionSelection):
    """Accept the selected existing revisions (at least one selector)."""

    type: Literal["accept_revisions"]

    @model_validator(mode="after")
    def validate_selector(self) -> AcceptRevisionsOperation:
        if not (self.ids or self.author or self.paragraph_index or self.through_paragraph_index or self.heading):
            raise ValueError("give ids, author, a paragraph range or a heading")
        return self


class RejectRevisionsOperation(_RevisionSelection):
    """Reject the selected existing revisions (at least one selector)."""

    type: Literal["reject_revisions"]

    @model_validator(mode="after")
    def validate_selector(self) -> RejectRevisionsOperation:
        if not (self.ids or self.author or self.paragraph_index or self.through_paragraph_index or self.heading):
            raise ValueError("give ids, author, a paragraph range or a heading")
        return self


class ListRevisionsOperation(_RevisionSelection):
    """Read-only: list revisions (id, type, author, date, paragraph, text); no selector lists all."""

    type: Literal["list_revisions"]
    limit: int = Field(default=200, ge=1, le=2000)


class SweepCharactersOperation(StrictModel):
    """Clean unusual characters already in the text as tracked changes (w:del/w:ins by `author`).

    Skips OMML equations, every field (code and result, Zotero citations included), text inside existing
    tracked insertions/deletions, and paragraphs holding drawings; the result counts what was skipped."""

    op_id: str
    type: Literal["sweep_characters"]
    paragraph_index: int | None = Field(default=None, ge=1, description="First paragraph; default the first.")
    through_paragraph_index: int | None = Field(
        default=None, ge=1, description="Last paragraph; default paragraph_index when given, else the last.")
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    author: str = Field(default="Claude", min_length=1, description=REVISION_AUTHOR_HELP)
    preview_only: bool = Field(default=False, description="Report what would change; write nothing.")
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="after")
    def validate_range(self) -> SweepCharactersOperation:
        if self.through_paragraph_index is not None and self.paragraph_index is not None \
                and self.through_paragraph_index < self.paragraph_index:
            raise ValueError("through_paragraph_index must not be before paragraph_index")
        return self


REVISION_OPERATION_TYPES = ("tracked_replace_text", "tracked_replace_paragraph", "accept_revisions",
                            "reject_revisions", "list_revisions", "sweep_characters")
READ_ONLY_OPERATION_TYPES = ("list_revisions",)
COMMENT_OPERATION_TYPES = ("add_comment", "reply_comment", "resolve_comment", "reopen_comment",
                           "delete_comment")


EditOperation = Annotated[
    Union[
        ReplaceTextOperation,
        InsertTextOperation,
        InsertParagraphOperation,
        InsertTableOperation,
        ReplyCommentOperation,
        AddCommentOperation,
        ResolveCommentOperation,
        ReopenCommentOperation,
        DeleteCommentOperation,
        ReplaceParagraphOperation,
        ReplaceTableRowOperation,
        ReplaceSectionOperation,
        TrackedReplaceTextOperation,
        TrackedReplaceParagraphOperation,
        AcceptRevisionsOperation,
        RejectRevisionsOperation,
        ListRevisionsOperation,
        SweepCharactersOperation,
    ],
    Field(discriminator="type"),
]


class EditDocumentRequest(StrictModel):
    source: str
    source_sha256: SHA256
    output: str | None = None
    allow_partial: bool = False
    operations: list[EditOperation] = Field(min_length=1, max_length=DEFAULT_MAX_OPERATIONS)

    @model_validator(mode="after")
    def validate_batch(self) -> EditDocumentRequest:
        op_ids = [op.op_id for op in self.operations]
        if len(op_ids) != len(set(op_ids)):
            raise ValueError("op_id values must be unique within a batch")

        all_ids = set(op_ids)
        seen_ids: set[str] = set()
        graph: dict[str, list[str]] = {}

        for op in self.operations:
            if op.type == "reply_comment" and getattr(op, "disposition", "") == "addressed":
                if not op.depends_on:
                    raise ValueError(
                        f"{op.op_id}: an 'addressed' reply requires at least one dependency (depends_on: the "
                        "op_id of the edit that addresses the comment, in this batch). If no text changed, "
                        "use disposition 'no_change' (the point stands as written) or 'rejected' (the "
                        "suggestion is declined); 'needs_information' asks the reviewer for more.")
            for dep in op.depends_on:
                if dep not in all_ids:
                    raise ValueError(f"unknown dependencies: {dep!r}")
                if dep not in seen_ids:
                    raise ValueError(f"dependency {dep!r} must appear earlier in operations batch than {op.op_id!r}")
            graph[op.op_id] = op.depends_on
            seen_ids.add(op.op_id)

        visited: set[str] = set()
        visiting: set[str] = set()

        def visit(op_id: str) -> None:
            if op_id in visiting:
                raise ValueError(f"cycle detected in dependencies involving {op_id!r}")
            if op_id in visited:
                return
            visiting.add(op_id)
            for dep in graph.get(op_id, []):
                visit(dep)
            visiting.remove(op_id)
            visited.add(op_id)

        for op_id in graph:
            visit(op_id)
        return self


class OperationError(StrictModel):
    code: str
    message: str
    retryable: bool = False
    context: dict[str, Any] = Field(default_factory=dict)


class OperationResult(StrictModel):
    op_id: str
    type: str
    status: Literal["applied", "failed", "skipped_dependency", "rolled_back"]
    summary: str = ""
    details: dict[str, Any] = Field(default_factory=dict)
    error: OperationError | None = None
    failed_dependencies: list[str] = Field(default_factory=list)
    forensic_output: str | None = None


class RetryInfo(StrictModel):
    source: str
    source_sha256: SHA256
    requires_reinspection: bool


class ProtectedResult(StrictModel):
    fields_preserved: bool
    comments_preserved: bool
    drawings_preserved: bool
    media_preserved: bool
    unrelated_tables_preserved: bool


class GroupResult(StrictModel):
    group: str
    op_ids: list[str]
    status: Literal["applied", "rolled_back", "failed", "not_applied"]
    reason: str = ""


class EditBatchResult(StrictModel):
    status: Literal["complete", "partial", "no_changes", "refused"]
    source: str
    source_sha256: SHA256
    output: str | None = None
    output_sha256: SHA256 | None = None
    operations: list[OperationResult]
    counts: dict[str, int]
    retry: RetryInfo | None = None
    protected: ProtectedResult | None = None
    message: str = ""
    summary_diffs: list[str] = Field(default_factory=list)
    undo_token: str | None = Field(default=None, description="Pass to undo_document to restore the state before this in-place batch; null when nothing was published in place.")
    sha256_before: SHA256 | None = None
    sha256_after: SHA256 | None = None
    publish_warning: str | None = None
    groups: list[GroupResult] = Field(default_factory=list, description="All-or-nothing operation groups and what happened to each.")
    consistency_check: str | None = None
    undo_call: str | None = Field(default=None, description="The one call that restores the state before this in-place batch.")
    cleaned: list[dict[str, Any]] = Field(
        default_factory=list,
        description="Characters textclean changed in the text this call wrote or swept, merged across operations.")
