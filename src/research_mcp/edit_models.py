"""Data models for batch document editing requests, plans, and results."""

from __future__ import annotations

import re
from typing import Annotated, Any, Literal, Union
from pydantic import BaseModel, ConfigDict, Field, StringConstraints, model_validator


SHA256 = Annotated[str, StringConstraints(pattern=r"^[0-9a-fA-F]{64}$")]
DEFAULT_MAX_OPERATIONS = 20
SAFE_MAX_PARAGRAPHS = 100
# Plan 864 decision 14: find/anchor strings up to this long (Word Find itself takes 255; longer ones are
# found by their first 255 characters and checked in full).
FIND_MAX = 5000


GROUP_HELP = (
    "Operations sharing a group are all-or-nothing together under allow_partial. Operations on the same "
    "table row or inside the same replace_section are grouped automatically."
)


KEEP_HELP = (
    'Characters to leave as written (each a character or "U+XXXX"). New text is otherwise cleaned with '
    "textclean (non-breaking/odd spaces, zero-width and control characters, ...); the result's `cleaned` "
    "lists what changed."
)


TRACKED_HELP = (
    "Make this edit as a tracked change (true) or plainly (false). Default: the document's tracking switch "
    "(set_track_changes); off when it was never set."
)
OP_AUTHOR_HELP = ("Author of the tracked change when this op is tracked; default: the switch's author "
                  "(set_track_changes), else \"Claude\".")


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class ReplaceTextOperation(StrictModel):
    op_id: str
    type: Literal["replace_text"]
    paragraph_index: int = Field(ge=1)
    find: str = Field(min_length=1, max_length=FIND_MAX)
    replace: str
    match_case: bool = True
    occurrence: int = Field(default=1, ge=1)
    expected_count: int = Field(default=1, ge=1)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    tracked: bool | None = Field(default=None, description=TRACKED_HELP)
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class InsertTextOperation(StrictModel):
    op_id: str
    type: Literal["insert_text"]
    paragraph_index: int = Field(ge=1)
    position: Literal["before", "after"]
    anchor: str = Field(min_length=1, max_length=FIND_MAX)
    text: str = Field(min_length=1)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    tracked: bool | None = Field(default=None, description=TRACKED_HELP)
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


class InsertParagraphOperation(StrictModel):
    op_id: str
    type: Literal["insert_paragraph"]
    after_paragraph_index: int = Field(ge=0)
    text: str = Field(min_length=1)
    style: str | None = None
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    tracked: bool | None = Field(default=None, description=TRACKED_HELP)
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
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
    caption: str | None = Field(default=None, min_length=1, description=(
        "One-line caption paragraph written before the table (e.g. 'Table 1. ...')."))
    caption_style: str | None = Field(default=None, min_length=1, description=(
        "Paragraph style of the caption (name or styleId); default 'Caption'."))
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    tracked: bool | None = Field(default=None, description=TRACKED_HELP)
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
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
    anchor: str = Field(min_length=1, max_length=FIND_MAX)
    text: str = Field(min_length=1)
    author: str | None = Field(default=None, min_length=1, description=AUTHOR_HELP)
    initials: str | None = Field(default=None, description=INITIALS_HELP)
    keep: list[str] = Field(default_factory=list, description=KEEP_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


CommentDisposition = Literal["addressed", "rejected", "no_change", "needs_information", "reply"]

ADDRESSED_BY_HELP = (
    "addressed / reply: evidence from earlier calls that the comment was answered, each a revision id from "
    "list_revisions (must still be in the document) or an undo_token / history entry id from document_history "
    "of an edit still in effect. 'addressed' needs this or depends_on (an edit in this batch).")


class ReplyCommentOperation(StrictModel):
    op_id: str
    type: Literal["reply_comment"]
    comment_id: str = Field(min_length=1, description=COMMENT_ID_HELP)
    disposition: CommentDisposition = "addressed"
    reply_text: str = Field(min_length=1)
    addressed_by: list[str] = Field(default_factory=list, description=ADDRESSED_BY_HELP)
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
    tracked: bool | None = Field(default=None, description=TRACKED_HELP)
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
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
    tracked: bool | None = Field(default=None, description=TRACKED_HELP)
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
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
    tracked: bool | None = Field(default=None, description=TRACKED_HELP)
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
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
    find: str = Field(min_length=1, max_length=FIND_MAX)
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


class DeleteCitationOperation(StrictModel):
    """Remove one whole Zotero citation field, as a tracked deletion by default (plan 831 S4)."""

    op_id: str
    type: Literal["delete_citation"]
    paragraph_index: int = Field(ge=1)
    citation_id: str | None = Field(default=None, min_length=1,
                                    description="The citation's citationID (field JSON); or give current_text.")
    current_text: str | None = Field(default=None, min_length=1,
                                     description='The citation\'s rendered text in that paragraph, e.g. "(12)".')
    author: str = Field(default="Claude", min_length=1, description=REVISION_AUTHOR_HELP)
    tracked: bool = Field(default=True, description=(
        "false removes the field outright instead of a tracked w:del. Default: the document's tracking switch "
        "when it was set (set_track_changes), else true."))
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="after")
    def one_selector(self) -> DeleteCitationOperation:
        # Neither selector refuses when the op runs, listing the paragraph's citations (plan 864 decision 18).
        if self.citation_id is not None and self.current_text is not None:
            raise ValueError("give citation_id or current_text, not both")
        return self


class DeleteParagraphOperation(StrictModel):
    """Delete whole paragraphs (plan 864 S3c): one paragraph, or a consecutive range in the body or in one
    table cell. A paragraph whose text is already a tracked deletion can be deleted too (its mark is)."""

    op_id: str
    type: Literal["delete_paragraph"]
    paragraph_index: int = Field(ge=1)
    through_paragraph_index: int | None = Field(
        default=None, ge=1, description="Last paragraph of the range (same container); default paragraph_index.")
    tracked: bool | None = Field(default=None, description=(
        "true: a tracked deletion (text and paragraph mark in w:del); false: the paragraphs are removed. Default: "
        "the document's tracking switch when it was set (set_track_changes), else true."))
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="after")
    def validate_range(self) -> DeleteParagraphOperation:
        if self.through_paragraph_index is not None and self.through_paragraph_index < self.paragraph_index:
            raise ValueError("through_paragraph_index must not be before paragraph_index")
        return self


class StyleTextOperation(StrictModel):
    """Bold/italic on one exact phrase of a paragraph (plan 864 S6). Tracked: each run of the phrase gets the new
    w:b/w:i and a w:rPrChange holding its old formatting; untracked: the new formatting only. Direct XML when
    tracked or on a document with revisions; otherwise through Word."""

    op_id: str
    type: Literal["style_text"]
    paragraph_index: int = Field(ge=1)
    text: str | None = Field(default=None, min_length=1, max_length=FIND_MAX, description=(
        "The phrase as displayed (accepted text); must occur once in the paragraph. find is another name."))
    find: str | None = Field(default=None, min_length=1, max_length=FIND_MAX, description="Another name for text.")
    bold: bool | None = Field(default=None, description="true sets bold, false removes it; omitted: unchanged.")
    italic: bool | None = Field(default=None, description="true sets italic, false removes it; omitted: unchanged.")
    allow_whole_paragraph: bool = Field(default=False, description="Needed when the phrase is the whole paragraph.")
    tracked: bool | None = Field(default=None, description=(
        "true: a tracked formatting change (w:rPrChange); false: plain. Default: the document's tracking switch "
        "(set_track_changes), else plain."))
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="after")
    def one_text(self) -> StyleTextOperation:
        if self.text is not None and self.find is not None:
            raise ValueError("give text or find, not both; find is another name for text")
        if self.text is None and self.find is None:
            raise ValueError("give text (the phrase to style)")
        if self.text is None:
            self.text, self.find = self.find, None
        if self.bold is None and self.italic is None:
            raise ValueError("give bold and/or italic; omitted properties stay unchanged")
        return self


class FormatParagraphOperation(StrictModel):
    """Font, size, indents, spacing and list numbering of whole paragraphs (plan 916 S3), as direct XML: one
    paragraph or a range in one container. Tracked: a w:pPrChange per paragraph and a w:rPrChange on each
    changed run and on the paragraph mark; untracked: the new formatting only. Runs of deleted text and of
    fields keep their formatting."""

    op_id: str
    type: Literal["format_paragraph"]
    paragraph_index: int = Field(ge=1)
    through_paragraph_index: int | None = Field(
        default=None, ge=1, description="Last paragraph of the range (same container); default paragraph_index.")
    font: str | None = Field(default=None, min_length=1, max_length=64, description="Font name for every run "
                             "and the paragraph mark.")
    size: float | None = Field(default=None, description="Font size in pt (steps of 0.5).")
    left_indent: float | None = Field(default=None, description="Left indent in pt.")
    right_indent: float | None = Field(default=None, description="Right indent in pt.")
    first_line_indent: float | None = Field(default=None, description="First-line indent in pt; negative = hanging.")
    space_before: float | None = Field(default=None, ge=0, description="Space before in pt.")
    space_after: float | None = Field(default=None, ge=0, description="Space after in pt.")
    remove_numbering: bool = Field(default=False, description="Remove list numbering (numId 0 when the style "
                                   "numbers the paragraph).")
    like_paragraph: int | None = Field(default=None, ge=1, description=(
        "Copy that paragraph's indents, spacing, numbering and its main font and size; explicit fields win."))
    tracked: bool | None = Field(default=None, description=(
        "true: tracked formatting changes (w:pPrChange, w:rPrChange); false: plain. Default: the document's "
        "tracking switch (set_track_changes), else plain."))
    author: str | None = Field(default=None, min_length=1, description=OP_AUTHOR_HELP)
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)

    @model_validator(mode="before")
    @classmethod
    def no_bold_italic(cls, data: Any) -> Any:
        if isinstance(data, dict):
            given = [name for name in ("bold", "italic") if data.get(name) is not None]
            if given:
                raise ValueError(f"format_paragraph does not set {' or '.join(given)}; use style_text "
                                 "(paragraph_index, text, bold/italic) for that")
        return data

    @model_validator(mode="after")
    def validate_range(self) -> FormatParagraphOperation:
        if self.through_paragraph_index is not None and self.through_paragraph_index < self.paragraph_index:
            raise ValueError("through_paragraph_index must not be before paragraph_index")
        if not (self.remove_numbering or self.like_paragraph is not None or any(
                getattr(self, name) is not None for name in (
                    "font", "size", "left_indent", "right_indent", "first_line_indent", "space_before",
                    "space_after"))):
            raise ValueError("give at least one of font, size, left_indent, right_indent, first_line_indent, "
                             "space_before, space_after, remove_numbering, like_paragraph")
        return self


class SetTrackChangesOperation(StrictModel):
    """Plan 864: turn research-mcp's tracking switch on or off (kept in the document; not Word's own switch).

    While on, replace_text, replace_paragraph and insert_text are made as tracked changes by the author, and
    ops that cannot be tracked yet refuse unless they say tracked: false. Applies to later ops in the batch."""

    op_id: str
    type: Literal["set_track_changes"]
    on: bool
    author: str | None = Field(default=None, min_length=1, description=(
        'Author of the tracked changes made while on; default: the author already set, else "Claude".'))
    depends_on: list[str] = Field(default_factory=list)
    group: str | None = Field(default=None, min_length=1, description=GROUP_HELP)


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
        DeleteCitationOperation,
        DeleteParagraphOperation,
        StyleTextOperation,
        FormatParagraphOperation,
        SetTrackChangesOperation,
    ],
    Field(discriminator="type"),
]


class EditDocumentRequest(StrictModel):
    source: str
    source_sha256: SHA256
    output: str | None = None
    allow_partial: bool = False
    # One more than the cap so edit_document(track_changes=...)'s prepended set_track_changes fits beside 20
    # edits; validate_batch keeps every other batch at DEFAULT_MAX_OPERATIONS.
    operations: list[EditOperation] = Field(min_length=1, max_length=DEFAULT_MAX_OPERATIONS + 1)

    @model_validator(mode="after")
    def validate_batch(self) -> EditDocumentRequest:
        switch_first = self.operations[0].type == "set_track_changes"
        if len(self.operations) - int(switch_first) > DEFAULT_MAX_OPERATIONS:
            raise ValueError(f"at most {DEFAULT_MAX_OPERATIONS} operations per batch "
                             "(besides a leading set_track_changes); split the edits into several calls")
        op_ids = [op.op_id for op in self.operations]
        if len(op_ids) != len(set(op_ids)):
            raise ValueError("op_id values must be unique within a batch")

        all_ids = set(op_ids)
        seen_ids: set[str] = set()
        graph: dict[str, list[str]] = {}

        for op in self.operations:
            if op.type == "reply_comment" and getattr(op, "disposition", "") == "addressed":
                if not op.depends_on and not op.addressed_by:
                    raise ValueError(
                        f"{op.op_id}: an 'addressed' reply requires evidence of the edit that addresses the "
                        "comment: depends_on (its op_id in this batch) or addressed_by (a revision id from "
                        "list_revisions, or an undo_token from an earlier edit of this document). If no text "
                        "changed, use disposition 'no_change' (the point stands as written) or 'rejected' "
                        "(the suggestion is declined); 'needs_information' asks the reviewer for more; "
                        "'reply' writes reply_text with no prefix.")
            if op.type == "reply_comment" and op.addressed_by and op.disposition not in ("addressed", "reply"):
                raise ValueError(f"{op.op_id}: addressed_by is evidence for an 'addressed' or 'reply' reply, "
                                 f"not {op.disposition!r}")
            if op.type == "reply_comment" and any(not item.strip() for item in op.addressed_by):
                raise ValueError(f"{op.op_id}: addressed_by items must not be empty")
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
    size_growth: dict[str, Any] | None = Field(
        default=None,
        description="Set only when the write made the file much larger (>=256 KB and >=20%): sizes, stored "
                    "bytes by kind (fonts, media, other), fonts Word newly embedded, whether image bytes "
                    "changed, and a note with the cause and remedy.")
    groups: list[GroupResult] = Field(default_factory=list, description="All-or-nothing operation groups and what happened to each.")
    consistency_check: str | None = None
    edit_feedback: str | None = Field(default=None, description="Plan 829: one line when the user decided earlier model edits since the model last heard; call edit_feedback for details.")
    feedback_warning: str | None = Field(default=None, description="Edit feedback could not be recorded; the edit itself succeeded.")
    undo_call: str | None = Field(default=None, description="The one call that restores the state before this in-place batch.")
    cleaned: list[dict[str, Any]] = Field(
        default_factory=list,
        description="Characters textclean changed in the text this call wrote or swept, merged across operations.")
    tracking: Literal["on", "off"] | None = Field(
        default=None, description="Plan 864: the document's tracking switch after this batch (set_track_changes).")
    tracking_author: str | None = Field(default=None, description="The switch's author while tracking is on.")
    tracking_note: str | None = Field(
        default=None, description="Set when tracking is off and the document has tracked changes.")
