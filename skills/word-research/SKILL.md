---
name: word-research
description: Edit Microsoft Word (.docx) research documents and cite from Zotero safely with the research-mcp tools - read before every write, pin the document's sha256, edit in place with one-call undo, track changes and comments, and fix references in Zotero then refresh. Use when asked to revise, comment on, restyle or add citations to a Word document.
---

# Word research documents with research-mcp

These tools edit `.docx` files on disk (not the Word window) and cite through the Zotero desktop app.
Follow the loop below for every change.

## 1. Read before you write

- Call `inspect_document(path)` (or `audit_document(path)`) first. Note the `sha256` and `is_locked`.
- Find a passage with `inspect_document(path, search_text="...")`; it returns paragraph indices, styles
  and the matching text, so you never page through a long document.
- If `is_locked` is true the file is open in Word. Ask the user to save and close it, then inspect again.

## 2. Pin the version you read

Pass the `sha256` you just read as `source_sha256` on every write (`edit_document`, `format_tables`,
`clean_bookmarks`, `insert_zotero_citation`, `apply_document_style`, `style_text`). If the file changed
since your read, the write is refused: inspect again and redo the edit on the new text. Each write
returns `sha256_after`; use it for the next write.

## 3. Edit in place, undo in one call

- Edit in place by default: `output` equal to `path`. The previous state goes into the undo store first.
- Save a new file (`output="paper_v2.docx"`) only when the user asks for a new version. For a big rewrite,
  ask before creating one.
- `undo_document(path)` reverses the last edit or batch; `undo_document(path, undo_token=...)` goes back
  to before a given edit. Before a risky series, `checkpoint_document(path, label)`; return with
  `undo_document(path, to=label)`. `document_history(path)` lists everything.
- Undo refuses when someone changed the file outside the tools. Tell the user; `output=<new path>`
  writes the earlier state as a separate file without touching theirs.

## 3b. Learn from what the user kept

- Results carry an `edit_feedback` line when the user has accepted, rejected or reworded your earlier edits
  since you last looked. Before editing that document again, call `edit_feedback(path)`: up to 10 new
  outcomes, newest first, each "first proposal -> what the user kept". Follow their rewordings; do not
  propose rejected text again.
- More only when needed: `scope="all", query="..."` searches every document; `record_id=..., detail="rounds"`
  shows each round of one span.

## 4. Make the edit

- Batches in `edit_document` are all-or-nothing. Group the operations that belong together (one entry's
  title, dates and description) instead of using `allow_partial`.
- `replace_text` and `replace_paragraph` change only the characters that differ, so bold and italic runs
  survive a whole-line rewrite. `replace_table_row` and `replace_section` rewrite a table row or a
  section by its visible text.
- Track changes when the user wants to review them: tracked edits, then accept or reject by revision.
  Read existing ones with `list_revisions`.
- Comments: add, reply to and resolve Word comments through `edit_document`. Quote the exact text the
  comment is about.
- Bold or italic one phrase with `style_text` (preview first). Never restyle a whole paragraph for a
  phrase. `audit_document_style` shows emphasis and typography before you change it.
- Check layout with `render_document_pages` (PNG pages) after table, figure or style changes.
- Save a PDF to look at or hand in with `export_pdf` (beside the .docx as `<name>.pdf`, or `output`). It refuses
  an existing file rather than overwrite it; tracked changes appear accepted and comments are hidden.

## 5. Citations go through Zotero

- Find references with `search_zotero`; add a missing one with `add_zotero_item_by_doi`.
- Insert with `insert_zotero_citation`. It writes Zotero's own field, so the Zotero Word plugin keeps
  working on the document.
- Change what one citation cites with `edit_citation_cluster`.
- A wrong reference (title, journal, pages...) is fixed in Zotero with `update_zotero_item`, then
  `refresh_citations` re-renders the document. Never edit the reference-list text by hand: the next
  refresh overwrites it. Run `refresh_citations(preview=True)` first to see what would change.
- If `audit_document` reports `citation_readiness.status == "virgin_document"`, the document has never
  had a Zotero citation: ask the user to insert one citation in Word with the Zotero plugin, save and
  close, then inspect again.
- Zotero desktop must be running with the local API on (Settings > Advanced, "Allow other applications
  on this computer to communicate with Zotero").

## 6. Text cleaning and problems

- Text you write is cleaned of invisible and unusual characters; results list them under `cleaned`.
  Equations, fields and citations are never touched.
- When something behaves unexpectedly, call `report_issue` and give the user the issue link it returns.
