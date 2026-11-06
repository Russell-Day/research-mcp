"""Pydantic return types for MCP tools exposed to LLM agents."""

from __future__ import annotations

from typing import Any
from pydantic import BaseModel, ConfigDict, Field


class MCPMergeSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    column: int = Field(ge=1, description="1-based column index to merge.")
    start_row: int = Field(ge=2, description="1-based start row index (must be data row >= 2).")
    end_row: int = Field(ge=2, description="1-based end row index (must be >= start_row).")
    label: str = Field(description="Expected uniform label across the merged cells.")


class MCPTableFormatSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    index: int = Field(ge=1, description="1-based table index in document order.")
    descriptive_columns: list[int] = Field(
        min_length=1,
        description="1-based column indices containing row headers or descriptive text.",
    )
    merges: list[MCPMergeSpec] = Field(
        default_factory=list,
        description="Optional vertical merge specifications for repeated grouping labels.",
    )


class AuditDocumentResult(BaseModel):
    path: str
    sha256: str
    body_text_sha256: str
    prose_text_sha256: str
    paragraph_count: int
    paragraph_count_basis: str = "raw OOXML paragraph elements (w:p); not Word COM paragraph numbering"
    table_count: int
    tables: list[dict[str, Any]]
    bookmark_names: list[str]
    bookmark_references: dict[str, Any]
    field_count: int
    zotero_item_fields: int
    zotero_bibliography_fields: int
    zotero_pref_properties: int
    revision_count: int
    comment_count: int
    live_comment_count: int
    orphan_comment_ids: list[str]
    drawing_count: int
    media_count: int
    live_media_count: int
    dangling_media_targets: list[str]
    structural_errors: list[str] = []
    is_locked: bool = False
    lock_reason: str | None = None
    citation_readiness: dict[str, Any] | None = None


class InspectDocumentResult(BaseModel):
    path: str
    source_sha256: str
    is_locked: bool = False
    lock_reason: str | None = None
    paragraph_count: int
    paragraph_count_basis: str = "Word-compatible 1-based Paragraphs collection, including table-row end markers"
    table_paragraph_count: int
    field_paragraph_count: int
    inexact_text_paragraph_count: int
    paragraph_index_matches_word: bool
    outline: list[dict[str, Any]] | None = None
    outline_empty: bool | None = None
    first_paragraph: int | None = None
    last_paragraph: int | None = None
    paragraphs: list[dict[str, Any]] | None = None
    window_truncated: bool | None = None
    returned_paragraph_count: int | None = None
    search_query: str | None = None
    match_count: int | None = None
    matches: list[dict[str, Any]] | None = None
    truncated: bool | None = None
    search_truncated: bool | None = None
    comments: list[dict[str, Any]] | None = None
    comment_thread_count: int | None = None
    unusual_characters: dict[str, Any] | None = Field(
        default=None,
        description="Characters edit_document's sweep_characters would clean: {total, by_codepoint: "
        "{'U+00A0': {name, action, count}}, skipped_total, skipped}; math, fields and existing revisions excluded.")


class ListBookmarksResult(BaseModel):
    bookmarks: list[dict[str, Any]]


class CleanBookmarksResult(BaseModel):
    source: str
    source_sha256: str
    output: str
    output_sha256: str
    removed: list[str]
    retained: list[str]
    protected_structure_unchanged: bool
    undo_token: str | None = None
    sha256_before: str | None = None
    sha256_after: str | None = None
    publish_warning: str | None = None


class FormatTablesResult(BaseModel):
    source: str
    source_sha256: str
    output: str
    output_sha256: str
    formatted_tables: list[int]
    protected_structure_unchanged: bool
    undo_token: str | None = None
    sha256_before: str | None = None
    sha256_after: str | None = None
    publish_warning: str | None = None


class InsertCitationResult(BaseModel):
    source: str
    source_sha256: str
    output: str
    output_sha256: str
    item_donor: str
    paragraph_index: int
    anchor: str
    item_keys: list[str]
    rendered_citation_text: str
    zotero_field_count: int
    undo_token: str | None = None
    sha256_before: str | None = None
    sha256_after: str | None = None
    publish_warning: str | None = None


class EditCitationClusterResult(BaseModel):
    success: bool
    source: str
    source_sha256: str
    output: str
    output_sha256: str
    paragraph_index: int
    citation_id: str
    item_keys_before: list[str]
    item_keys_after: list[str]
    rendered_citation_before: str
    rendered_citation_after: str
    bibliography_before: list[str]
    bibliography_after: list[str]
    revisions_checked: int
    revision_check: dict[str, Any] = Field(
        default_factory=dict,
        description="How every tracked revision compared before the edit and after Word's save: "
        "level byte_identical or serialization_only (rsid/attribute order/proofing markers only).",
    )
    warnings: list[str] = Field(default_factory=list)
    undo_token: str | None = None
    sha256_before: str | None = None
    sha256_after: str | None = None
    publish_warning: str | None = None


class BibliographyEntryChange(BaseModel):
    item_key: str | None = Field(None, description="Zotero item of the entry, when the citation order fixes it "
                                 "(numbered citation-sequence styles); null otherwise.")
    before_index: int
    after_index: int
    before: str
    after: str


class BibliographyEntryDropped(BaseModel):
    before_index: int
    text: str


class BibliographyEntryAdded(BaseModel):
    after_index: int
    item_key: str | None = None
    text: str


class CitationChange(BaseModel):
    citation_id: str
    item_keys: list[str]
    before: str
    after: str


class RefreshCitationsResult(BaseModel):
    success: bool
    preview: bool = Field(description="True: the refresh ran on a temporary copy that was discarded; nothing written.")
    source: str
    source_sha256: str
    output: str | None = None
    output_sha256: str | None = None
    citations_refreshed: int
    citations_changed: list[CitationChange] = Field(default_factory=list)
    bibliography_changed: list[BibliographyEntryChange] = Field(
        default_factory=list,
        description="Entries whose text the refresh changed. One you did not change in Zotero was probably "
        "edited by hand: fix the item in Zotero (update_zotero_item) and refresh; never edit the list text.")
    bibliography_dropped: list[BibliographyEntryDropped] = Field(default_factory=list)
    bibliography_added: list[BibliographyEntryAdded] = Field(default_factory=list)
    bibliography_unchanged_count: int
    bibliography_renumbered_count: int = 0
    bibliography_before: list[str]
    bibliography_after: list[str]
    custom_entries: int = Field(0, description="Custom entries in the ZOTERO_BIBL field, left as Zotero renders them.")
    revisions_checked: int
    revision_check: dict[str, Any] = Field(default_factory=dict)
    warnings: list[str] = Field(default_factory=list)
    undo_token: str | None = None
    sha256_before: str | None = None
    sha256_after: str | None = None
    publish_warning: str | None = None


class UndoDocumentResult(BaseModel):
    path: str
    restored: bool
    undo_token: str | None = Field(
        default=None,
        description="Id of this undo in the document's history; undo_document(path) again reverses it.",
    )
    restored_sha256: str
    sha256_before: str
    sha256_after: str
    publish_warning: str | None = None
    sha256: str
    undid: dict[str, Any] | None = None
    forced: bool = False
    source: str = Field(default="checkpoint", description="checkpoint (undo store) or bak (legacy .docx.bak fallback).")
    backup_path: str | None = None
    output: str | None = None
    message: str


class RevertDocumentResult(UndoDocumentResult):
    """revert_document is an alias of undo_document(path)."""


class CheckpointDocumentResult(BaseModel):
    path: str
    checkpoint: str
    label: str
    sha256: str
    changed_outside_since_last_record: bool
    message: str


class DocumentHistoryResult(BaseModel):
    path: str
    sha256: str
    entries: list[dict[str, Any]]
    dropped: list[dict[str, Any]]
    limits: dict[str, Any]
    changed_outside_since_last_record: bool
    legacy_backup: str | None = None
    store: str
    recovery_pending: list[dict[str, Any]] = []


class SearchZoteroResult(BaseModel):
    candidates: list[dict[str, Any]]
    collection_marker_path: str | None = None


class ListCollectionsResult(BaseModel):
    collections: list[dict[str, Any]]


class SetProjectCollectionResult(BaseModel):
    success: bool
    collection: str
    collection_key: str
    marker_path: str
    previous_value: str | None = None
    replaced: bool = False


class AddZoteroItemResult(BaseModel):
    success: bool
    created: bool
    item_key: str
    doi: str
    title: str
    first_author: str
    abstract: str
    pdf_attached: bool
    pdf_status: str
    visible_locally: bool
    collection: str | None = None
    collection_marker_path: str | None = None


class ZoteroFieldChange(BaseModel):
    field: str
    old: Any
    new: Any


class UpdateZoteroItemResult(BaseModel):
    success: bool
    item_key: str
    item_type: str
    title: str
    changed: bool
    dry_run: bool
    changes: list[ZoteroFieldChange]
    previous: dict[str, Any]
    cleaned: list[dict[str, Any]] = Field(default_factory=list)
    version_before: int
    version_after: int | None = None
    synced_locally: bool | None = None
    sync_seconds: float | None = None


class ReportIssueResult(BaseModel):
    ok: bool
    issue_id: str
    logged_to: str
    message: str
    # plan 551: category char_mapping only
    char_mapping: dict[str, Any] | None = None
    improvement: dict[str, Any] | None = None
    file_with: str | None = None
    report_count: int | None = None
    # default (GitHub) mode: a prefilled new-issue link and the same text
    issue_url: str | None = None
    issue_title: str | None = None
    issue_body: str | None = None


class ApplyStyleResult(BaseModel):
    source: str
    source_sha256: str
    output: str
    output_sha256: str
    profile: str
    paragraphs_styled: int
    margins: dict[str, Any]
    typography: dict[str, Any]
    protected_structure_unchanged: bool
    undo_token: str | None = None
    sha256_before: str | None = None
    sha256_after: str | None = None
    publish_warning: str | None = None


class RenderDocumentResult(BaseModel):
    source: str
    output_directory: str
    total_rendered: int
    dpi: int
    pages: list[dict[str, Any]]
