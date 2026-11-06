"""High-fidelity visual rendering of Word documents into print-grade PNGs."""

from __future__ import annotations

import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import Any, Sequence

from .audit import audit_docx
from .errors import RefusedUnsafe
from .word import IsolatedWordDocument


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


def render_docx_pages(
    docx_path: str | Path,
    output_dir: str | Path | None = None,
    pages: Sequence[int] | None = None,
    dpi: int = 200,
    *,
    visible: bool = False,
) -> dict[str, Any]:
    """Render pages of a .docx file into clean PNG images.

    Uses Word COM ExportAsFixedFormat with Item=0 (wdExportDocumentContent)
    to suppress comment markup gutters and render true margin-to-margin print layout,
    followed by pdftoppm rasterization at specified DPI.
    """
    resolved_docx = Path(docx_path).resolve()
    if not resolved_docx.is_file():
        raise FileNotFoundError(f"document not found: {resolved_docx}")

    pdftoppm_bin = find_pdftoppm()

    # Determine output directory
    if output_dir is None:
        file_sha = audit_docx(resolved_docx).sha256[:8]
        out_path = resolved_docx.parent / f"{resolved_docx.stem}_rendered_{file_sha}"
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
        with IsolatedWordDocument(resolved_docx, visible=visible, read_only=True) as session:
            doc = session.document
            # wdExportFormatPDF = 17, wdExportOptimizeForPrint = 0, wdExportAllDocument = 0
            # wdExportDocumentContent = 0 (suppresses comment markup balloon gutter)
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                Item=0,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )

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
        "pages": results,
    }
