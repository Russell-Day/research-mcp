"""Persistent, project-owned Word process for serialized Zotero integration."""

from __future__ import annotations

import json
from multiprocessing.connection import Listener
import os
from pathlib import Path
import random
import re
import string
import sys
import time
import traceback

from .constants import WD_ALERTS_NONE, WD_DO_NOT_SAVE_CHANGES, WD_FIELD_ADDIN
from .audit import document_structure_errors
from .processes import process_creation_time
from .word_com import ensure_default_desktop
from .word_service import refresh_state_path, worker_state_path



def _record_progress(**fields) -> None:
    """Publish refresh progress where a timed-out supervisor can still read it.

    Best effort by design: a diagnostic that can fail the operation it is
    diagnosing is worse than no diagnostic.
    """

    try:
        path = refresh_state_path()
        payload = {"pid": os.getpid(), "wall_clock": time.time(), **fields}
        temporary = path.with_suffix(".tmp")
        temporary.write_text(json.dumps(payload), encoding="utf-8")
        temporary.replace(path)
    except Exception:
        pass


def _citation_id(field) -> str | None:
    code = str(field.Code.Text).strip()
    start = code.find("{")
    if start < 0:
        return None
    try:
        return json.loads(code[start:]).get("citationID")
    except ValueError:
        return None  # Zotero is rewriting this field's code right now


def _item_fields(document) -> list:
    return [field for field in document.Fields if int(field.Type) == WD_FIELD_ADDIN and str(field.Code.Text).strip().startswith("ADDIN ZOTERO_ITEM")]


def _bibliography_fields(document) -> list:
    return [field for field in document.Fields if int(field.Type) == WD_FIELD_ADDIN and str(field.Code.Text).strip().startswith("ADDIN ZOTERO_BIBL")]


def _bibliographies_complete(results: list[str]) -> bool:
    """An empty rendered bibliography is legal; only the placeholder is pending."""

    return all(value != "{Bibliography}" for value in results)


def _zotero_preferences(document) -> tuple[list, list[str], str]:
    """Return ordered split properties and their combined document data."""

    properties = []
    for index in range(1, int(document.CustomDocumentProperties.Count) + 1):
        prop = document.CustomDocumentProperties(index)
        match = re.fullmatch(r"ZOTERO_PREF_(\d+)", str(prop.Name))
        if match:
            properties.append((int(match.group(1)), prop))
    properties.sort(key=lambda item: item[0])
    if not properties:
        raise RuntimeError("document has no ZOTERO_PREF_* properties")
    ordered = [prop for _, prop in properties]
    parts = [str(prop.Value) for prop in ordered]
    return ordered, parts, "".join(parts)


def _write_zotero_preferences(properties: list, old_parts: list[str], value: str) -> None:
    offset = 0
    for position, prop in enumerate(properties):
        length = len(old_parts[position]) if position < len(properties) - 1 else len(value) - offset
        prop.Value = value[offset : offset + length]
        offset += length


def _mark_refresh_pending(document) -> str:
    """Set a harmless marker that Zotero replaces in setDocumentData at completion.

    Zotero parses ``zotero-version`` as informational metadata and always
    serializes its current version immediately after a successful command.  A
    marker here therefore distinguishes a completed idempotent refresh from
    fields that merely looked complete before the asynchronous command began.
    """

    properties, parts, combined = _zotero_preferences(document)
    match = re.search(r'zotero-version="([^"]+)"', combined)
    if match is None:
        raise RuntimeError("Zotero document data has no zotero-version attribute")
    old_value = match.group(1)
    marker = "".join(random.SystemRandom().choices(string.ascii_letters, k=len(old_value)))
    while marker == old_value:
        marker = "".join(random.SystemRandom().choices(string.ascii_letters, k=len(old_value)))
    updated = combined[: match.start(1)] + marker + combined[match.end(1) :]
    _write_zotero_preferences(properties, parts, updated)
    if marker not in _zotero_preferences(document)[2]:
        raise RuntimeError("refresh completion marker did not read back")
    return marker


def _refresh_completed(document, marker: str) -> bool:
    return marker not in _zotero_preferences(document)[2]


def _same_path(candidate: str, target: Path) -> bool:
    candidate_normalized = str(candidate).replace("\\", "/").lower()
    target_normalized = str(target).replace("\\", "/").lower()
    return candidate_normalized == target_normalized or candidate_normalized.endswith("/" + target.name.lower())


def _close_open_copies(word, path: Path) -> int:
    closed = 0
    for index in range(int(word.Documents.Count), 0, -1):
        try:
            candidate = word.Documents(index)
            if not _same_path(str(candidate.FullName), path):
                continue
            candidate.Close(WD_DO_NOT_SAVE_CHANGES)
            closed += 1
        except Exception:
            continue
    return closed


def _word_pid(word) -> int:
    import win32process

    control = word.Documents.Add()
    try:
        return int(win32process.GetWindowThreadProcessId(int(control.ActiveWindow.Hwnd))[1])
    finally:
        control.Close(WD_DO_NOT_SAVE_CHANGES)


def _configure_word_application(word) -> None:
    """Configure the persistent Word instance for unattended operation."""

    word.Visible = False
    word.DisplayAlerts = WD_ALERTS_NONE
    word.ScreenUpdating = False
    word.AutomationSecurity = 3


def _refresh(word, path: Path, citation_ids: tuple[str, ...], timeout: float) -> dict:
    document = None
    started = time.perf_counter()
    last_state: dict = {"note": "refresh loop never executed"}
    try:
        problems = document_structure_errors(path)
        if problems:
            raise RuntimeError(
                "DOCX preflight failed before Word opened the refresh target: "
                + "; ".join(problems[:5])
            )
        document = word.Documents.Open(
            FileName=str(path),
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False,
            ConfirmConversions=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        if bool(document.ReadOnly):
            raise RuntimeError(f"Word opened the refresh target read-only: {path}")
        document.Activate()
        completion_marker = _mark_refresh_pending(document)
        # Measured: Run returns in ~0.02s and Zotero renders asynchronously,
        # so the polling loop below is what actually bounds the refresh. The
        # 36-citation manuscript completes in ~7s. Every prior timeout was the
        # post-refresh unwind, not the refresh. Keep these records: they are
        # what distinguished the two.
        _record_progress(
            phase="zotero_refresh_invoked",
            path=str(path),
            citation_count=len(citation_ids),
            loop_timeout=timeout,
        )
        run_started = time.perf_counter()
        document.Application.Run("ZoteroRefresh")
        run_seconds = time.perf_counter() - run_started
        _record_progress(
            phase="zotero_refresh_returned",
            path=str(path),
            citation_count=len(citation_ids),
            run_seconds=run_seconds,
        )
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            by_id = {_citation_id(field): field for field in _item_fields(document)}
            missing = [citation_id for citation_id in citation_ids if citation_id not in by_id]
            command_complete = _refresh_completed(document, completion_marker)
            if missing and command_complete:
                raise RuntimeError(f"citation fields disappeared: {missing}")
            # While Zotero is still refreshing it rewrites each field's code, so a
            # poll can catch a field mid-rewrite (seen live with 34 placeholders).
            results = {
                citation_id: str(by_id[citation_id].Result.Text).strip() if citation_id in by_id else ""
                for citation_id in citation_ids
            }
            citations_complete = all(value and value != "{Citation}" for value in results.values())
            bibliography_results = [str(field.Result.Text).strip() for field in _bibliography_fields(document)]
            bibliographies_complete = _bibliographies_complete(bibliography_results)
            if citations_complete and bibliographies_complete and command_complete:
                # setDocumentData precedes Zotero's final cleanup/activate/complete
                # sequence. Keep the document alive through that short tail.
                time.sleep(0.5)
                document.Save()
                _record_progress(
                    phase="complete",
                    path=str(path),
                    run_seconds=run_seconds,
                    total_seconds=time.perf_counter() - started,
                )
                return {
                    "rendered": results,
                    "bibliography_results": bibliography_results,
                    "seconds": time.perf_counter() - started,
                    "zotero_refresh_seconds": run_seconds,
                }
            last_state = {
                "citations_complete": citations_complete,
                "bibliographies_complete": bibliographies_complete,
                "command_complete": command_complete,
                "unrendered_citations": [k for k, v in results.items() if not v or v == "{Citation}"][:5],
                "bibliography_results": [b[:40] for b in bibliography_results],
            }
            _record_progress(phase="awaiting_predicates", path=str(path), **last_state)
            time.sleep(0.5)
        # Report WHICH predicate failed. A bare timeout here cost a full
        # debugging cycle: Zotero had finished and saved the document itself,
        # while one predicate never flipped, so a completed refresh was reported
        # as a failure with no way to tell the two apart.
        _record_progress(
            phase="loop_timeout", path=str(path), run_seconds=run_seconds, **last_state
        )
        raise RuntimeError(
            f"Zotero refresh did not complete within {timeout}s; last state: {last_state}"
        )
    finally:
        # Cleanup MUST complete before the reply. Deferring it until after the
        # response was tried and reverted: the caller immediately reopened the
        # file for the next refresh and got
        # `PermissionError: [Errno 13] Permission denied` because Word still
        # held the handle. Cleanup was never the cause of the old hangs -- that
        # was a pipe-buffer deadlock in the supervisor -- so there is nothing
        # to gain here and a real file-lock race to lose.
        try:
            if document is not None:
                document.Close(WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            pass
        _record_progress(phase="cleanup_started", path=str(path))
        closed = _close_open_copies(word, path)
        _record_progress(phase="cleanup_finished", path=str(path), closed=closed)


def main(argv: list[str] | None = None) -> int:
    args = argv if argv is not None else sys.argv[1:]
    if len(args) != 2:
        raise SystemExit("usage: python -m research_mcp.word_worker PIPE AUTHKEY_HEX")
    pipe, authkey_hex = args
    import pythoncom
    import win32com.client

    ensure_default_desktop()
    pythoncom.CoInitialize()
    listener = None
    word = None
    state_file = worker_state_path()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        _configure_word_application(word)
        word_pid = _word_pid(word)
        listener = Listener(pipe, family="AF_PIPE", authkey=bytes.fromhex(authkey_hex))
        state_file.write_text(json.dumps({
            "pid": os.getpid(), "process_created": process_creation_time(os.getpid()),
            "word_pid": word_pid, "word_process_created": process_creation_time(word_pid),
            "pipe": pipe, "authkey": authkey_hex, "started": time.time(),
        }), encoding="utf-8")
        stopping = False
        while not stopping:
            connection = listener.accept()
            try:
                request = connection.recv()
                command = request.get("command") if isinstance(request, dict) else None
                if command == "ping":
                    documents = []
                    for index in range(1, int(word.Documents.Count) + 1):
                        try:
                            documents.append(str(word.Documents(index).FullName))
                        except Exception:
                            documents.append("<unreadable>")
                    response = {
                        "ok": True,
                        "pid": os.getpid(),
                        "word_pid": word_pid,
                        "document_count": int(word.Documents.Count),
                        "documents": documents,
                    }
                    try:
                        response["user_name"] = str(word.UserName or "")
                        response["user_initials"] = str(word.UserInitials or "")
                    except Exception:
                        pass
                elif command == "refresh":
                    response = {"ok": True, **_refresh(word, Path(request["path"]).resolve(strict=True), tuple(request.get("citation_ids", ())), float(request.get("timeout", 60.0)))}
                elif command == "shutdown":
                    if int(word.Documents.Count):
                        raise RuntimeError("refusing worker shutdown while documents remain open")
                    response = {"ok": True, "stopping": True}
                    stopping = True
                else:
                    raise ValueError(f"unknown Word worker command: {command!r}")
            except Exception as exc:
                response = {"ok": False, "error": str(exc), "traceback": traceback.format_exc()}
            try:
                connection.send(response)
            finally:
                connection.close()
    finally:
        state_file.unlink(missing_ok=True)
        if listener is not None:
            listener.close()
        if word is not None:
            try:
                for index in range(int(word.Documents.Count), 0, -1):
                    word.Documents(index).Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
            try:
                word.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
