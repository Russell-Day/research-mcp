"""High-fidelity visual rendering of Word documents into print-grade PNGs."""

from __future__ import annotations

import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import time
from typing import Any, Sequence
import uuid

from .audit import audit_docx, file_sha256
from .errors import RefusedUnsafe
from .word import IsolatedWordDocument

# Word's own deadline for one render or export session. The default 30 s is short for a long manuscript;
# this stays under the 180 s the server gives the whole worker process (run_operation).
WORD_EXPORT_TIMEOUT_SECONDS = 150.0


def find_pdftoppm() -> str:
    """Locate pdftoppm executable from system PATH or common MiKTeX install paths."""
    found = shutil.which("pdftoppm")
    if found:
        return found

    candidates = [
        Path(os.path.expandvars(r"%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64\pdftoppm.exe")),
        Path(r"C:\Program Files\MiKTeX\miktex\bin\x64\pdftoppm.exe"),
        Path(r"C:\Program Files (x86)\MiKTeX\miktex\bin\x64\pdftoppm.exe"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)

    raise RefusedUnsafe(
        "pdftoppm executable not found. Ensure MiKTeX or Poppler is installed and in PATH."
    )


# ExportAsFixedFormat Item: wdExportDocumentContent (clean print layout) and wdExportDocumentWithMarkup.
EXPORT_CONTENT = 0
EXPORT_WITH_MARKUP = 7


def _show_all_markup(document: Any) -> bool:
    """Set the document's window to Word's "All Markup" view, so a markup export shows redlines and balloons
    whatever view the file was saved in: final view with markup (wdRevisionsViewFinal 0), all markup
    (wdRevisionsMarkupAll 2), revisions and comments shown, balloons in the margin (wdBalloonRevisions 0).
    Best effort: False when any of it cannot be set (the export then uses Word's current view and the render
    warns)."""
    try:
        view = document.Windows(1).View
        view.ShowRevisionsAndComments = True
        view.RevisionsView = 0
        # Word 2013+ keeps the markup level in RevisionsFilter; a filter left at Simple or No Markup would hide
        # changes, so failing to set it counts as failing to set the view.
        view.RevisionsFilter.Markup = 2
        view.RevisionsFilter.View = 0
        view.MarkupMode = 0
        return True
    except Exception:  # noqa: BLE001 - COM may refuse on a hidden read-only window; fakes have no window
        return False


MARKUP_VIEW_WARNING = (
    "Word did not let the render set the All Markup view, so the pages use the view the file was saved in: "
    "redlines or comment balloons may be missing. Open the file in Word, choose Review > All Markup, save, and "
    "render again")


def _export_pdf(document: Any, pdf_path: Path, *, markup: bool = False) -> bool:
    """Export a vector PDF of the whole document (shared by page renders and export_pdf).

    wdExportFormatPDF = 17, wdExportOptimizeForPrint = 0, wdExportAllDocument = 0,
    wdExportDocumentContent = 0 (no comment balloon gutter; the default), or with `markup`
    wdExportDocumentWithMarkup = 7 (redlines and comment balloons), wdExportCreateHeadingBookmarks = 1.
    Returns False only for a markup export whose All Markup view could not be set (the caller warns).
    """
    view_set = _show_all_markup(document) if markup else True
    document.ExportAsFixedFormat(
        OutputFileName=str(pdf_path),
        ExportFormat=17,
        OpenAfterExport=False,
        OptimizeFor=0,
        Range=0,
        Item=EXPORT_WITH_MARKUP if markup else EXPORT_CONTENT,
        IncludeDocProps=True,
        KeepIRM=True,
        CreateBookmarks=1,
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=False,
    )
    return view_set


def render_docx_pages(
    docx_path: str | Path,
    output_dir: str | Path | None = None,
    pages: Sequence[int] | None = None,
    dpi: int = 200,
    *,
    visible: bool = False,
    markup: bool = False,
) -> dict[str, Any]:
    """Render pages of a .docx file into PNG images.

    Uses Word COM ExportAsFixedFormat with Item=0 (wdExportDocumentContent)
    to suppress comment markup gutters and render true margin-to-margin print layout,
    or with `markup` Item=7 (wdExportDocumentWithMarkup: redlines and comment balloons),
    followed by pdftoppm rasterization at specified DPI.
    """
    resolved_docx = Path(docx_path).resolve()
    if not resolved_docx.is_file():
        raise FileNotFoundError(f"document not found: {resolved_docx}")

    pdftoppm_bin = find_pdftoppm()

    # Determine output directory
    if output_dir is None:
        file_sha = audit_docx(resolved_docx).sha256[:8]
        # The markup render gets its own folder, so it never overwrites the clean pages of the same file.
        kind = "rendered_markup" if markup else "rendered"
        out_path = resolved_docx.parent / f"{resolved_docx.stem}_{kind}_{file_sha}"
    else:
        out_path = Path(output_dir).resolve()
    out_path.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_path = Path(tmp_dir) / f"{resolved_docx.stem}.pdf"
        # Render into an isolated directory.  Reusing ``out_path`` as the
        # pdftoppm destination makes a shorter later render discover stale
        # page images left by an earlier, longer render.
        render_prefix = Path(tmp_dir) / "rendered-page"

        # Export clean vector PDF via Word COM
        with IsolatedWordDocument(
            resolved_docx, visible=visible, read_only=True, timeout_seconds=WORD_EXPORT_TIMEOUT_SECONDS
        ) as session:
            if markup:
                view_set = _export_pdf(session.document, pdf_path, markup=True)
            else:
                view_set = _export_pdf(session.document, pdf_path)

        if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
            raise RuntimeError(f"failed to export PDF from {resolved_docx}")

        # Construct pdftoppm command
        prefix = render_prefix
        cmd = [pdftoppm_bin, "-png", "-r", str(dpi)]
        if pages:
            cmd.extend(["-f", str(min(pages)), "-l", str(max(pages))])
        cmd.extend([str(pdf_path), str(prefix)])

        proc = subprocess.run(
            cmd, capture_output=True, text=True, check=False,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            timeout=120,
        )
        if proc.returncode != 0:
            raise RuntimeError(f"pdftoppm failed ({proc.returncode}): {proc.stderr}")

        # Copy only this invocation's generated files.  Existing files in the
        # requested output directory are intentionally preserved, but are not
        # included in the response.
        generated_files = sorted(Path(tmp_dir).glob("rendered-page-*.png"))
        rendered_files: list[tuple[int, Path]] = []
        for generated in generated_files:
            try:
                page_num = int(generated.stem.rsplit("-", 1)[-1])
            except ValueError:
                continue
            destination = out_path / f"page-{page_num}.png"
            shutil.copy2(generated, destination)
            rendered_files.append((page_num, destination))

    results: list[dict[str, Any]] = []
    for page_num, p in rendered_files:
        if pages and page_num not in pages:
            continue
        results.append({
            "page_number": page_num,
            "image_path": str(p),
            "size_bytes": p.stat().st_size,
        })

    return {
        "source": str(resolved_docx),
        "output_directory": str(out_path),
        "total_rendered": len(results),
        "dpi": dpi,
        "markup": bool(markup),
        "pages": results,
        **({} if view_set else {"warnings": [MARKUP_VIEW_WARNING]}),
    }


def resolve_pdf_target(docx_path: str | Path, output: str | Path | None = None) -> tuple[Path, Path]:
    """Return (source .docx, PDF target) for export_pdf, refusing a target it must not write.

    The default target is ``<stem>.pdf`` beside the document. An existing file is never overwritten.
    """
    resolved_docx = Path(docx_path).resolve()
    if not resolved_docx.is_file():
        raise FileNotFoundError(f"document not found: {resolved_docx}")
    # An existing path is refused as given, before resolving: a dangling symlink must not be followed.
    if output is not None and os.path.lexists(output):
        raise RefusedUnsafe(
            f"export_pdf: {str(output)!r} already exists and is never overwritten; "
            "pass another output path, or move or delete that file first"
        )
    target = resolved_docx.with_suffix(".pdf") if output is None else Path(output).resolve()
    if target.suffix.lower() != ".pdf":
        raise RefusedUnsafe(f"export_pdf: output must end in .pdf, got {str(target)!r}")
    if not target.parent.is_dir():
        raise RefusedUnsafe(f"export_pdf: output folder does not exist: {str(target.parent)!r}")
    if os.path.lexists(target):
        raise RefusedUnsafe(
            f"export_pdf: {str(target)!r} already exists and is never overwritten; "
            "pass another output path, or move or delete that file first"
        )
    return resolved_docx, target


_PDF_PAGE_OBJECT = re.compile(rb"/Type\s*/Page(?![A-Za-z])")


def pdf_page_count(pdf_path: Path) -> int | None:
    """Pages in a PDF Word exported: its page objects (Word writes them uncompressed); None if none found.

    Word's own page statistic counts the markup view of a document with tracked changes, not the exported
    final view (a 21-page export of a 1303-revision manuscript reported 22), so the PDF is counted.
    """
    return len(_PDF_PAGE_OBJECT.findall(pdf_path.read_bytes())) or None


def _publish_new_file(temporary: Path, target: Path) -> None:
    """Move ``temporary`` to ``target`` only if ``target`` does not exist (no overwrite, no race)."""
    try:
        if os.name == "nt":
            os.rename(temporary, target)  # MoveFileEx without REPLACE_EXISTING: refuses an existing target
        else:
            os.link(temporary, target)  # POSIX rename would overwrite; link refuses an existing target
            os.unlink(temporary)
    except FileExistsError:
        raise RefusedUnsafe(
            f"export_pdf: {str(target)!r} appeared while Word was exporting and was left untouched; "
            "nothing was written"
        ) from None


def export_docx_pdf(
    docx_path: str | Path,
    output: str | Path | None = None,
    *,
    visible: bool = False,
) -> dict[str, Any]:
    """Export a .docx to a new PDF (Word's clean print layout); never overwrites, never changes the source."""
    from .editing import revision_total

    started = time.monotonic()
    resolved_docx, target = resolve_pdf_target(docx_path, output)
    audit = audit_docx(resolved_docx)
    revision_count = revision_total(resolved_docx)
    temporary = target.parent / f".~research-mcp-export-{uuid.uuid4().hex}.pdf"
    try:
        with IsolatedWordDocument(
            resolved_docx, visible=visible, read_only=True, timeout_seconds=WORD_EXPORT_TIMEOUT_SECONDS
        ) as session:
            _export_pdf(session.document, temporary)
        if not temporary.is_file() or temporary.stat().st_size == 0:
            raise RuntimeError(f"Word did not export a PDF from {resolved_docx}")
        page_count = pdf_page_count(temporary)
        sha_after = file_sha256(resolved_docx)
        if sha_after != audit.sha256:
            raise RefusedUnsafe(
                f"export_pdf: {str(resolved_docx)!r} changed on disk during the export "
                f"(sha256 {audit.sha256} -> {sha_after}); nothing was written. Export again."
            )
        _publish_new_file(temporary, target)
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass

    warnings: list[str] = []
    if revision_count:
        warnings.append(
            f"the document has {revision_count} tracked revision elements; the PDF shows the final text with "
            "every change accepted and no markup (deleted text left out, inserted text as plain text)"
        )
    if audit.comment_count:
        warnings.append(f"the document has {audit.comment_count} comments; the PDF does not show them")
    if audit.is_locked:
        warnings.append(
            "the document is open in Word: the PDF shows the saved file, without unsaved changes"
        )
    return {
        "source": str(resolved_docx),
        "source_sha256": audit.sha256,
        "output": str(target),
        "size_bytes": target.stat().st_size,
        "page_count": page_count,
        "revision_count": revision_count,
        "comment_count": audit.comment_count,
        "warnings": warnings,
        "elapsed_seconds": round(time.monotonic() - started, 2),
    }
