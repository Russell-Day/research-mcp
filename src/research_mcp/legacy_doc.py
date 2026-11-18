"""Word 97-2003 .doc files (plan 916 S1): detect them, read them through a temporary .docx, convert them.

A .doc is an OLE2 compound file. It is detected by its header, never by its extension: a .docx package
named ``.doc`` is read as the .docx it is, and a .doc renamed to ``.docx`` is still a .doc. (A .doc can
hold an embedded zip, so ``zipfile`` may open it and fail later with "There is no item named
'word/document.xml'"; the header check comes first.)

Reading tools (inspect, audit, render, export) convert the .doc with Word into a private temporary .docx
and read that; the .doc is opened read-only and never written. Write tools refuse a .doc with one line
that names ``convert_document``, which saves ``<stem>.docx`` beside it (never over an existing file).

This module imports only the standard library and ``errors`` at load time, so ``audit`` and
``paragraphs`` can use the detection and messages; Word is imported where it is used.
"""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
import os
from pathlib import Path
import shutil
import tempfile
import time
from typing import Any, Iterator
import uuid

from .errors import IntegrationRefused, RefusedUnsafe, VerificationFailed

OLE2_HEADER = bytes.fromhex("D0CF11E0A1B11AE1")
ZIP_HEADER = b"PK\x03\x04"

# Word fills a password prompt with this instead of showing it: a protected file fails to open with a
# password error (refused below), an unprotected one ignores it.
NO_PASSWORD = "research-mcp-no-password"
WORD_CONVERT_TIMEOUT_SECONDS = 120.0
WORKER_TIMEOUT_SECONDS = 180.0
TEMP_PREFIX = ".~research-mcp-convert-"


def is_legacy_doc(path: str | Path) -> bool:
    """True when the file starts with the OLE2 compound-file header (a Word 97-2003 .doc), whatever its name."""
    try:
        with open(path, "rb") as handle:
            return handle.read(len(OLE2_HEADER)) == OLE2_HEADER
    except OSError:
        return False


def _beside_name(path: str | Path) -> str:
    return f"{Path(path).stem}.docx"


def write_refusal(tool: str, path: str | Path) -> str:
    """The one line every write tool gives for a .doc (decision 1A)."""
    p = Path(path)
    if p.suffix.lower() == ".docx":
        return (f"{tool} can't edit {p.name}: it is a Word 97-2003 .doc file named .docx; call "
                "convert_document(path, output=<new .docx path>) and edit that.")
    return f"{tool} can't edit a .doc; call convert_document(path) to make {_beside_name(p)} beside it and edit that."


def read_refusal(tool: str, path: str | Path) -> str:
    """For a read tool that parses the .docx package itself (not one of the four that convert first)."""
    p = Path(path)
    return (f"{tool} reads .docx files and {p.name} is a Word 97-2003 .doc; call convert_document(path) to make "
            f"a .docx and read that (inspect_document, audit_document, render_document_pages and export_pdf "
            "read a .doc as is).")


def not_a_package(path: str | Path) -> str:
    """For low-level readers (audit, paragraphs) handed a .doc directly."""
    p = Path(path)
    return (f"{p.name} is a Word 97-2003 .doc file, not a .docx package; call convert_document(path) to make "
            f"{_beside_name(p)} beside it.")


def password_refusal(path: str | Path) -> str:
    return (f"{Path(path).name} is password-protected; Word can't open it without the password. Open it in Word, "
            "remove the password (File > Info > Protect Document) and save it, then try again.")


def macro_refusal(path: str | Path) -> str:
    return (f"{Path(path).name} contains macros (a VBA project); research-mcp does not read or convert documents "
            "with macros. Open it in Word and save it as a .docx (Word drops the macros), then use that file.")


# --- target and publication ---------------------------------------------------------------------------


def resolve_convert_target(source: str | Path, output: str | Path | None = None) -> tuple[Path, Path]:
    """Return (source, target .docx) for convert_document, refusing anything it must not write."""
    src = Path(source).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"document not found: {src}")
    if not is_legacy_doc(src):
        # The header decides: a .docx package named .doc is already a .docx.
        with open(src, "rb") as handle:
            is_package = handle.read(len(ZIP_HEADER)) == ZIP_HEADER
        if is_package:
            raise RefusedUnsafe(f"convert_document: {src.name} is already a .docx package; read and edit it directly")
        raise RefusedUnsafe(f"convert_document: {src.name} is not a Word 97-2003 .doc file (no OLE2 header)")
    # An existing path is refused as given, before resolving: a dangling symlink must not be followed.
    if output is not None and os.path.lexists(output):
        raise RefusedUnsafe(
            f"convert_document: {str(output)!r} already exists and is never overwritten; "
            "pass another output path, or move or delete that file first")
    if output is None:
        target = src.with_suffix(".docx")
        if target == src:
            raise RefusedUnsafe(
                f"convert_document: {src.name} is a Word 97-2003 .doc file named .docx; pass "
                "output=<new .docx path> (the file itself is never overwritten)")
    else:
        target = Path(output).resolve()
    if target.suffix.lower() != ".docx":
        raise RefusedUnsafe(f"convert_document: output must end in .docx, got {str(target)!r}")
    if not target.parent.is_dir():
        raise RefusedUnsafe(f"convert_document: output folder does not exist: {str(target.parent)!r}")
    if os.path.lexists(target):
        raise RefusedUnsafe(
            f"convert_document: {str(target)!r} already exists and is never overwritten; "
            "pass another output path, or move or delete that file first")
    return src, target


def _publish_new_file(temporary: Path, target: Path, tool: str = "convert_document") -> None:
    """Move ``temporary`` to ``target`` only if ``target`` does not exist (no overwrite, no race)."""
    try:
        if os.name == "nt":
            os.rename(temporary, target)  # MoveFileEx without REPLACE_EXISTING: refuses an existing target
        else:
            os.link(temporary, target)  # POSIX rename would overwrite; link refuses an existing target
            os.unlink(temporary)
    except FileExistsError:
        raise RefusedUnsafe(
            f"{tool}: {str(target)!r} appeared while Word was working and was left untouched; "
            "nothing was written") from None


# --- Word -----------------------------------------------------------------------------------------------


def _is_password_error(exc: BaseException) -> bool:
    return "password" in str(exc).lower()


def _legacy_session_class():
    """IsolatedWordDocument without the OOXML preflight (a .doc is no zip), opened read-only, never prompting."""
    from .processes import process_creation_time
    from .word import IsolatedWordDocument
    from .word_com import co_initialize, create_word_application, word_pid

    class LegacyWordDocument(IsolatedWordDocument):
        def __enter__(self):
            try:
                co_initialize()
                self._com_initialized = True
                self.word = create_word_application(visible=self.visible)
                self.pid = word_pid(self.word)
                if self.pid:
                    self.process_created = process_creation_time(self.pid)
                self._start_watchdog()
                self.document = self.word.Documents.Open(
                    FileName=str(self.path),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    PasswordDocument=NO_PASSWORD,
                    WritePasswordDocument=NO_PASSWORD,
                    Visible=self.visible,
                    OpenAndRepair=False,
                    NoEncodingDialog=True,
                )
                self._raise_if_timed_out()
                return self
            except Exception as exc:
                self._force_cleanup()
                if self._timed_out.is_set():
                    raise IntegrationRefused(
                        f"Word operation exceeded timeout_seconds={self.timeout_seconds:g}; the owned Word "
                        "process was terminated and nothing was written") from exc
                if _is_password_error(exc):
                    raise RefusedUnsafe(password_refusal(self.path)) from exc
                raise IntegrationRefused(f"Word could not open {self.path}: {exc}. Nothing was written.") from exc

    return LegacyWordDocument


def save_as_docx(document: Any, source: Path, target: Path) -> None:
    """Save the open document as a .docx (Word's default format) at ``target``; refuse one with macros."""
    from .constants import WD_SAVE_FORMAT_DOCUMENT_DEFAULT

    if bool(getattr(document, "HasVBProject", False)):
        raise RefusedUnsafe(macro_refusal(source))
    document.SaveAs2(FileName=str(target), FileFormat=WD_SAVE_FORMAT_DOCUMENT_DEFAULT, AddToRecentFiles=False)


def convert_to_docx(source: str | Path, output: str | Path | None = None, *, visible: bool = False,
                    temporary: str | Path | None = None) -> dict[str, Any]:
    """Word saves a new .docx from ``source`` (opened read-only); never overwrites, never changes the source.

    Runs inside the operation worker (``run_operation("convert_doc", ...)``). Word writes a hidden temporary
    in the target folder (``temporary``, named by the caller so it can remove it if the worker is stopped),
    which is checked (a .docx package, the source unchanged) and then published with a no-replace rename.
    ``source_sha256`` in the result is the hash of the bytes Word converted: taken before Word opened the
    file and found again after the save.
    """
    from .audit import document_structure_errors, file_sha256

    started = time.monotonic()
    src, target = resolve_convert_target(source, output)
    sha_before = file_sha256(src)
    temporary = Path(temporary) if temporary is not None else _temporary_beside(target)
    if temporary.parent != target.parent or not temporary.name.startswith(TEMP_PREFIX):
        raise RefusedUnsafe(f"convert_document: the temporary {str(temporary)!r} must be a {TEMP_PREFIX}* file "
                            "in the target folder")
    try:
        with _legacy_session_class()(src, visible=visible, read_only=True,
                                     timeout_seconds=WORD_CONVERT_TIMEOUT_SECONDS) as session:
            save_as_docx(session.document, src, temporary)
        if not temporary.is_file() or temporary.stat().st_size == 0:
            raise VerificationFailed(f"Word did not save a .docx from {src}")
        problems = document_structure_errors(temporary)
        if problems:
            raise VerificationFailed(
                f"Word's .docx of {src.name} failed the package check: {'; '.join(problems[:5])}; nothing was written")
        sha_after = file_sha256(src)
        if sha_after != sha_before:
            raise RefusedUnsafe(
                f"convert_document: {src.name} changed on disk during the conversion (sha256 {sha_before} -> "
                f"{sha_after}); nothing was written. Try again.")
        _publish_new_file(temporary, target)
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
    return {
        "source": str(src),
        "source_sha256": sha_before,
        "output": str(target),
        "output_sha256": file_sha256(target),
        "size_bytes": target.stat().st_size,
        "elapsed_seconds": round(time.monotonic() - started, 3),
    }


# --- server-side helpers --------------------------------------------------------------------------------


def _temporary_beside(target: Path) -> Path:
    return target.parent / f"{TEMP_PREFIX}{uuid.uuid4().hex}.docx"


def convert_in_worker(source: str | Path, target: Path) -> dict[str, Any]:
    """Run the conversion in the operation worker. The worker's hidden temporary is named here and removed
    here too, so a worker stopped at its deadline leaves nothing behind (unless Word still holds it)."""
    from .operation_process import run_operation

    temporary = _temporary_beside(Path(target))
    try:
        return run_operation("convert_doc", {"source": str(source), "output": str(target),
                                             "temporary": str(temporary)},
                             timeout_seconds=WORKER_TIMEOUT_SECONDS)
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


@dataclass
class LegacyRead:
    """A .doc read through a temporary .docx: what the result reports instead of the temporary's identity.

    ``sha256`` is the hash of the bytes Word converted (the worker's own), so a save between the call and
    the conversion is never labelled with the older hash; ``check_unchanged`` refuses when the .doc no longer
    has those bytes."""

    path: Path
    sha256: str
    tool: str = "read"
    checked_at_publish: bool = False

    def check_unchanged(self) -> None:
        from .audit import file_sha256

        now = file_sha256(self.path)
        if now != self.sha256:
            raise RefusedUnsafe(
                f"{self.tool}: {self.path.name} changed on disk while it was being read (sha256 {self.sha256} -> "
                f"{now}); nothing was returned or written. Try again.")

    def publish_file(self, produced: Path, target: Path) -> None:
        """Copy ``produced`` (made from the temporary .docx) to ``target`` beside the .doc: staged as a hidden
        file in the target folder, the .doc's hash checked right before a no-replace rename."""
        staged = target.parent / f".~research-mcp-{uuid.uuid4().hex}{target.suffix}"
        try:
            shutil.copy2(produced, staged)
            self.check_unchanged()
            self.checked_at_publish = True
            _publish_new_file(staged, target, self.tool)
        finally:
            try:
                staged.unlink(missing_ok=True)
            except OSError:
                pass

    def publish_pages(self, result: dict[str, Any], output_dir: Path) -> dict[str, Any]:
        """Copy rendered pages from the private folder to ``output_dir`` (as render does for a .docx: an
        existing page-N.png is replaced), after checking the .doc's hash."""
        self.check_unchanged()
        self.checked_at_publish = True
        output_dir.mkdir(parents=True, exist_ok=True)
        pages = []
        for page in result.get("pages", []):
            destination = output_dir / Path(page["image_path"]).name
            shutil.copy2(page["image_path"], destination)
            pages.append({**page, "image_path": str(destination)})
        return {**result, "output_directory": str(output_dir), "pages": pages}

    def relabel(self, result: dict[str, Any], *, path_key: str, sha_key: str | None = None) -> dict[str, Any]:
        """The result read from the temporary .docx, with the .doc's path, sha256 and lock state."""
        from .audit import is_file_locked

        out = {**result, path_key: str(self.path), "converted_from_doc": True}
        if sha_key:
            out[sha_key] = self.sha256
        if "is_locked" in result:
            out["is_locked"], out["lock_reason"] = is_file_locked(self.path)
        return out


@contextmanager
def docx_for_reading(path: str | Path, tool: str = "read") -> Iterator[tuple[Path, LegacyRead | None]]:
    """Yield (a .docx to read, LegacyRead or None). A .doc is converted into a private temporary folder,
    removed afterwards; anything else is yielded as given. When the read finishes, a .doc that no longer has
    the bytes Word converted is refused (unless the caller already checked it right before publishing)."""
    p = Path(path)
    if not is_legacy_doc(p):
        yield p, None
        return
    resolved = p.resolve(strict=True)
    with tempfile.TemporaryDirectory(prefix="research-mcp-doc-", ignore_cleanup_errors=True) as folder:
        temporary = Path(folder) / _beside_name(resolved)
        converted = convert_in_worker(resolved, temporary)
        legacy = LegacyRead(path=resolved, sha256=converted["source_sha256"], tool=tool)
        yield temporary, legacy
        if not legacy.checked_at_publish:
            legacy.check_unchanged()
