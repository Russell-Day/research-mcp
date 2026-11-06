"""Isolated Microsoft Word COM process lifecycle and error handling."""

from __future__ import annotations

from contextlib import contextmanager
import logging
import json
import os
from pathlib import Path
import shutil
import tempfile
import threading
import time
import math
from typing import Iterator

from .constants import (
    RPC_E_DISCONNECTED,
    RPC_S_CALLPENDING,
    WD_DO_NOT_SAVE_CHANGES,
    WD_SAVE_FORMAT_DOCUMENT_DEFAULT,
)
from .errors import IntegrationRefused, RefusedUnsafe, VerificationFailed
from .audit import document_structure_errors, file_sha256, is_file_locked
from .publication import note_versioned, publish_inplace
from .processes import process_creation_time, process_identity_matches
from .word_com import (
    co_initialize,
    co_uninitialize,
    com_error_code,
    create_word_application,
    kill_process,
    word_pid,
)

logger = logging.getLogger(__name__)


@contextmanager
def staged_document(source: str | Path, output: str | Path) -> Iterator[Path]:
    """Provide a verified working copy for either versioned or in-place mutations.

    If output != source:
        Enforces that output does not already exist. Copies source to output.
        Yields output. On unhandled failure, cleans up the incomplete output.

    If output == source:
        Enforces atomic in-place mutation: creates a temporary staging copy,
        yields staging path for Word COM mutations and verifications.
        Only on clean exit from the context block does publish_inplace store the
        verified pre-edit state in the checkpoint store (outside OneDrive; no
        .bak sidecar) and atomically swap (os.replace) the staged copy over
        source. The publication (undo_token, sha256_before, sha256_after) is
        reported to any active ``publication.capture()``.
        If source is locked by Word or another process, raises RefusedUnsafe.
        On any verification failure or error, staging is cleanly discarded
        and source is never modified.
    """
    source_path = Path(source).resolve(strict=True)
    output_path = Path(output).resolve()
    if source_path.suffix.lower() != ".docx" or output_path.suffix.lower() != ".docx":
        raise RefusedUnsafe("source and output must both be .docx files")

    is_inplace = (source_path == output_path)
    if is_inplace:
        locked, reason = is_file_locked(source_path)
        if locked:
            raise RefusedUnsafe(f"Cannot stage {source_path.name}: {reason}. Save and close it, then inspect again.")
        expected_sha = file_sha256(source_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="research-mcp-stage-", dir=output_path.parent) as temp_dir:
            temp_stage = Path(temp_dir) / source_path.name
            shutil.copy2(source_path, temp_stage)
            yield temp_stage
            publish_inplace(temp_stage, source_path, expected_sha256=expected_sha)
    else:
        if output_path.exists():
            raise RefusedUnsafe(f"refusing to overwrite existing output: {output_path}")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        source_sha = file_sha256(source_path)
        shutil.copy2(source_path, output_path)
        try:
            yield output_path
        except Exception:
            if output_path.exists():
                try:
                    output_path.unlink()
                except OSError:
                    pass
            raise
        note_versioned(source_sha, output_path)


def create_versioned_copy(source: str | Path, output: str | Path) -> tuple[Path, Path]:
    """Create a new version output file copied from source, refusing if output exists."""
    source_path = Path(source).resolve(strict=True)
    output_path = Path(output).resolve()
    if output_path.exists():
        raise RefusedUnsafe(f"refusing to overwrite existing output: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_path, output_path)
    return source_path, output_path


class IsolatedWordDocument:
    """Run Word COM operations with aggressive cleanup and zombie prevention."""

    def __init__(
        self,
        path: str | Path,
        *,
        visible: bool = False,
        read_only: bool = False,
        timeout_seconds: float = 30.0,
    ) -> None:
        self.path = Path(path).resolve(strict=True)
        self.visible = visible
        self.read_only = read_only
        self.timeout_seconds = timeout_seconds
        self.word = None
        self.document = None
        self.pid = None
        self.process_created = None
        self._com_initialized = False
        self._timed_out = threading.Event()
        self._watchdog_stop = threading.Event()
        self._watchdog_thread: threading.Thread | None = None
        self._owned_word_file = os.environ.get("RESEARCH_MCP_OWNED_WORD_FILE")

        if not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
            raise ValueError("timeout_seconds must be a finite positive number")

    def _start_watchdog(self) -> None:
        """Start a deadline watcher after we have an owned Word identity.

        COM calls can block the calling thread indefinitely.  A daemon thread
        can still enforce the process deadline while that thread is blocked.
        It only terminates a process whose PID *and* creation time still match
        the instance captured for this context, so a recycled PID or the
        user's Word process cannot be terminated accidentally.
        """
        if not self.pid or self.process_created is None:
            raise IntegrationRefused(
                "cannot enforce Word deadline because the owned Word PID and creation identity "
                "were unavailable after create_word_application(); the application creation "
                "call itself remains outside the bounded portion of this lifecycle"
            )

        self._register_owned_word()

        def watch() -> None:
            if self._watchdog_stop.wait(self.timeout_seconds):
                return
            self._timed_out.set()
            if process_identity_matches(self.pid, self.process_created):
                kill_process(self.pid)
                logger.warning("Word deadline expired; terminated owned Word PID %s", self.pid)
            else:
                logger.warning(
                    "Word deadline expired, but owned process identity no longer matched; "
                    "refusing to terminate PID %s",
                    self.pid,
                )

        self._watchdog_thread = threading.Thread(
            target=watch,
            name="research-mcp-word-deadline",
            daemon=True,
        )
        self._watchdog_thread.start()

    def _register_owned_word(self) -> None:
        """Publish the owned process identity for an outer worker supervisor."""
        if not self._owned_word_file:
            return
        target = Path(self._owned_word_file).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        temp = target.with_name(f".{target.name}.{os.getpid()}.tmp")
        temp.write_text(
            json.dumps({"pid": self.pid, "created": self.process_created}, separators=(",", ":")),
            encoding="utf-8",
        )
        os.replace(temp, target)

    def _unregister_owned_word(self, *, only_if_gone: bool = False) -> None:
        """Remove only our registration; preserve a newer owner's record.

        On timeout, keep the record while the process is still observable so an
        outer supervisor can retry cleanup after this worker is stopped.
        """
        if not self._owned_word_file:
            return
        target = Path(self._owned_word_file).resolve()
        try:
            payload = json.loads(target.read_text(encoding="utf-8"))
            if payload.get("pid") == self.pid and payload.get("created") == self.process_created:
                if not only_if_gone or not process_identity_matches(self.pid, self.process_created):
                    target.unlink()
        except (FileNotFoundError, OSError, ValueError, TypeError):
            pass

    def _stop_watchdog(self) -> None:
        self._watchdog_stop.set()
        thread = self._watchdog_thread
        if thread and thread is not threading.current_thread():
            thread.join(timeout=0.2)
        self._watchdog_thread = None

    def _raise_if_timed_out(self) -> None:
        if self._timed_out.is_set():
            raise IntegrationRefused(
                f"Word operation exceeded timeout_seconds={self.timeout_seconds:g}; "
                "the owned Word process was terminated and no staged output may be published"
            )

    def __enter__(self) -> IsolatedWordDocument:
        problems = document_structure_errors(self.path)
        if problems:
            raise RefusedUnsafe(
                "DOCX preflight failed before starting Word: "
                + "; ".join(problems[:5])
                + ". Repair the package and inspect again; retrying Word will not fix it."
            )
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
                ReadOnly=self.read_only,
                AddToRecentFiles=False,
                Visible=self.visible,
                ConfirmConversions=False,
                OpenAndRepair=False,
                NoEncodingDialog=True,
            )
            if not self.read_only and bool(getattr(self.document, "ReadOnly", False)):
                raise RefusedUnsafe(
                    "Word opened the document read-only although a writable session was requested; "
                    "refusing to continue because edits could not be saved"
                )
            self._raise_if_timed_out()
            return self
        except Exception as exc:
            self._force_cleanup()
            if self._timed_out.is_set():
                raise IntegrationRefused(
                    f"Word operation exceeded timeout_seconds={self.timeout_seconds:g}; "
                    "the owned Word process was terminated and no staged output may be published"
                ) from exc
            raise IntegrationRefused(
                f"Word could not open {self.path}: {exc}. "
                "No edits were made in this session. Stop Word retries and inspect the document or integration."
            ) from exc

    def save(self) -> None:
        self._raise_if_timed_out()
        if not self.document:
            raise RefusedUnsafe("no open Word document to save")
        self.document.Save()
        self._raise_if_timed_out()

    def save_as(self, target_path: str | Path) -> Path:
        self._raise_if_timed_out()
        if not self.document:
            raise RefusedUnsafe("no open Word document to save")
        resolved = Path(target_path).resolve()
        resolved.parent.mkdir(parents=True, exist_ok=True)
        self.document.SaveAs2(
            FileName=str(resolved),
            FileFormat=WD_SAVE_FORMAT_DOCUMENT_DEFAULT,
            AddToRecentFiles=False,
        )
        self._raise_if_timed_out()
        return resolved

    def _force_cleanup(self) -> None:
        # Once the watchdog fires, Close/Quit may themselves block on the
        # disconnected COM server.  The watchdog already terminated the owned
        # process, so release Python references and uninitialize COM directly.
        if self._timed_out.is_set():
            self.document = None
            self.word = None
            self._stop_watchdog()
            self._unregister_owned_word(only_if_gone=True)
            if self._com_initialized:
                co_uninitialize()
                self._com_initialized = False
            return

        if self.document:
            try:
                self.document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception as exc:
                if com_error_code(exc) not in (RPC_E_DISCONNECTED, RPC_S_CALLPENDING):
                    logger.debug("ignoring document.Close error during force cleanup: %s", exc)
            self.document = None

        # The deadline may fire while document.Close is returning.  Do not
        # begin another COM call after that point.
        if self._timed_out.is_set():
            self.word = None
            self._stop_watchdog()
            self._unregister_owned_word(only_if_gone=True)
            if self._com_initialized:
                co_uninitialize()
                self._com_initialized = False
            return

        if self.word:
            try:
                self.word.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception as exc:
                if com_error_code(exc) not in (RPC_E_DISCONNECTED, RPC_S_CALLPENDING):
                    logger.debug("ignoring word.Quit error during force cleanup: %s", exc)
            self.word = None

        self._stop_watchdog()

        if self.pid and process_identity_matches(self.pid, self.process_created):
            start = time.monotonic()
            while time.monotonic() - start < 1.0:
                if not process_identity_matches(self.pid, self.process_created):
                    break
                time.sleep(0.05)
            if process_identity_matches(self.pid, self.process_created):
                kill_process(self.pid)

        self._unregister_owned_word(only_if_gone=True)
        if self._com_initialized:
            co_uninitialize()
            self._com_initialized = False

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self._force_cleanup()
        if self._timed_out.is_set():
            raise IntegrationRefused(
                f"Word operation exceeded timeout_seconds={self.timeout_seconds:g}; "
                "the owned Word process was terminated and no staged output may be published"
            ) from exc_val
