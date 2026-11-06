"""Publish a staged document over its source, keeping the pre-edit state in the undo store.

Plan 358: the rolling ``<name>.docx.bak`` sidecar is gone. Before the swap, the verified pre-edit bytes
go to the checkpoint store (``checkpoints.py``, under %LOCALAPPDATA%\\research_mcp, outside OneDrive) and
the write is journalled; after the swap the post-edit sha is recorded in the document's history. The
history entry id is the ``undo_token``. ``undo_document`` (and its alias ``revert_document``),
``checkpoint_document`` and ``document_history`` live here as well: stale guard, undo of undo, named checkpoints, ``output=`` restores.

Hash checks detect edits made during staging. They are not an OS-wide lock against arbitrary external
writers between the last check and replacement.
"""

from __future__ import annotations

from contextlib import contextmanager
import json
import os
from pathlib import Path
import shutil
import tempfile
import threading
import uuid

from . import checkpoints as ck
from .audit import is_file_locked
from .checkpoints import Refusal, file_sha256

# ---------------------------------------------------------------- what a tool call published
#
# Writers reach publish_inplace through staged_document or edit_batch, several calls below the MCP tool.
# A tool wraps its work in ``capture(tool=...)``; every publication in that thread is appended to the
# captured list, so the tool can return undo_token / sha256_before / sha256_after without threading them
# through each writer module.

_CAP = threading.local()


@contextmanager
def capture(tool: str | None = None, op: str | None = None):
    frames = getattr(_CAP, "frames", None)
    if frames is None:
        frames = _CAP.frames = []
    frame = {"tool": tool, "op": op, "records": []}
    frames.append(frame)
    try:
        yield frame["records"]
    finally:
        frames.remove(frame)


def _current(key: str):
    for frame in reversed(getattr(_CAP, "frames", None) or []):
        if frame.get(key):
            return frame[key]
    return None


def note_write(record: dict) -> dict:
    for frame in getattr(_CAP, "frames", None) or []:
        frame["records"].append(record)
    return record


def note_versioned(source_sha256: str, output: Path) -> dict:
    """A versioned write (output != source): the source is untouched, so there is nothing to undo."""
    return note_write({"path": str(Path(output).resolve()), "undo_token": None, "versioned": True,
                       "sha256_before": source_sha256.upper(), "sha256_after": file_sha256(output).upper()})


def write_fields(records: list, sha256_before: str | None = None, sha256_after: str | None = None) -> dict:
    """undo_token / sha256_before / sha256_after / publish_warning for a tool result: the last publication,
    else the given pair."""
    if records:
        rec = records[-1]
        return {"undo_token": rec["undo_token"], "sha256_before": rec["sha256_before"],
                "sha256_after": rec["sha256_after"], "publish_warning": rec.get("publish_warning")}
    return {"undo_token": None, "sha256_before": sha256_before.upper() if sha256_before else None,
            "sha256_after": sha256_after.upper() if sha256_after else None, "publish_warning": None}


# ---------------------------------------------------------------- publish

def check_source(target: Path, expected: str) -> None:
    locked, reason = is_file_locked(target)
    if locked:
        raise Refusal("doc_open", f"Cannot replace {target.name}: {reason}. Save and close it, then inspect again.")
    actual = file_sha256(target)
    if actual != expected.lower():
        raise Refusal("wrong_sha", f"Source changed during editing: {target.name}. The external changes were "
                      "preserved; inspect again before retrying.",
                      {"expected": expected.upper(), "actual": actual.upper()})


_REPLACEFILE_IGNORE_MERGE_ERRORS = 0x2
_REPLACEFILE_IGNORE_ACL_ERRORS = 0x4
_ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 = 1177


def _replace_file(replacement: Path, replaced: Path, displaced: Path) -> None:
    """Windows ReplaceFileW: put ``replacement`` at ``replaced`` and move the file actually there to ``displaced``
    in one call, so a save that lands after the last hash check is kept, not discarded. All three paths must be
    on one volume. Raises OSError (``winerror`` set) on failure."""
    ck.require_windows()
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    fn = kernel32.ReplaceFileW
    fn.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPVOID,
                   wintypes.LPVOID]
    fn.restype = wintypes.BOOL
    if not fn(str(replaced), str(replacement), str(displaced),
              _REPLACEFILE_IGNORE_MERGE_ERRORS | _REPLACEFILE_IGNORE_ACL_ERRORS, None, None):
        err = ctypes.get_last_error()
        raise OSError(None, ctypes.FormatError(err).strip(), str(replaced), err)


def _swap(staged: Path, target: Path, *, expected_sha256: str, state: dict, displaced: Path) -> str:
    """Replace ``target`` with ``staged`` if ``target`` still has ``expected_sha256``. Returns the new sha.

    The file the swap displaces goes to ``displaced`` (see _settle). ``state["replaced"]`` turns true as soon
    as the swap onto ``target`` succeeded; a caller must trust it, not a later hash, to know whether the
    document changed. ``state["half"]`` marks a swap that stopped with ``target`` missing and its earlier
    state at ``displaced`` that could not be put back."""
    state["replaced"] = False
    want = file_sha256(staged)
    try:
        check_source(target, expected_sha256)
        try:
            _replace_file(staged, target, displaced)
            state["replaced"] = True
        except OSError as exc:
            if (getattr(exc, "winerror", None) == _ERROR_UNABLE_TO_MOVE_REPLACEMENT_2
                    or (not target.exists() and displaced.exists())):
                try:
                    os.rename(displaced, target)  # the swap stopped half-way: put the original back
                except OSError:
                    state["half"] = True
            raise
    except OSError as exc:
        if state["replaced"]:
            raise _unverified(target, want, type(exc).__name__) from exc
        raise Refusal("publish_failed", f"Could not safely publish {target.name}: {exc}. Original document was not "
                      "replaced. Check file access and free disk space, then inspect again.") from exc
    try:
        got = file_sha256(target)
    except OSError as exc:
        raise _unverified(target, want, f"{type(exc).__name__} reading it back") from exc
    if got != want:
        raise _unverified(target, want, f"read back sha {got.upper()}")
    return got


def _unverified(target: Path, want: str, why: str) -> Refusal:
    return Refusal("publish_failed", f"{target.name} was replaced, but the check after the swap failed ({why}).",
                   {"published": True, "verified": False, "expected": want.upper()}, changed_nothing=False)


def _settle(target: Path, txid: str, displaced: Path, cur_sha: str) -> str | None:
    """After the swap: store the displaced file's state, verify it and record it in the journal, then delete the
    file (ck.settle_displaced). When it is not the checked ``cur_sha`` (a save landed between the last check and
    the swap) it becomes this write's ``sha_before``, so undo brings it back, and the outside change is recorded
    in history. Returns a warning, or None when nothing raced. When storing fails the file stays where it is,
    named in the journal, and the refusal names that existing path."""
    if not displaced.is_file():
        return None
    try:
        x = ck.settle_displaced(txid, displaced)
    except (OSError, Refusal) as exc:
        ck.mark_replaced(txid)
        kept = displaced.is_file()
        where = (f"It is kept unchanged at {displaced}; the next call on this document stores it and records the "
                 "change once the store accepts it." if kept else "It could not be found afterwards.")
        raise Refusal("publish_failed", f"{target.name} was replaced, but the state the swap displaced could not be "
                      f"stored ({type(exc).__name__}: {exc}). {where} Inspect both before the next edit.",
                      {"published": True, "displaced": str(displaced) if kept else None, "undo_token": txid},
                      changed_nothing=False) from exc
    if x == cur_sha:
        return None
    return (f"{target.name} was saved outside the server (sha {x.upper()}) after the last check and just before "
            f"this write replaced it (checked sha {cur_sha.upper()}). That saved state was kept: undo_document(path) "
            "(this write's undo_token) restores it, and document_history lists the outside change. Inspect before "
            "the next edit.")


def _publish(staged: Path, target: Path, cur_sha: str, new_sha: str, txid: str) -> list[str]:
    """Publish ``staged`` over ``target``; returns warnings. Before the swap a failure drops the journal (nothing
    changed); after it the journal is kept (marked replaced) and the failure is reported as a published,
    unverified write. The displaced file goes to a dot-named file beside the document (same volume, outside any
    temporary folder), journalled before the swap so recover() finds it after a crash or interruption."""
    displaced = ck.displaced_path(target, txid)
    ck.note_displaced(txid, displaced)
    state: dict = {}
    try:
        _swap(staged, target, expected_sha256=cur_sha, state=state, displaced=displaced)
    except BaseException as exc:
        if state.get("half"):
            _settle(target, txid, displaced, "0" * 64)  # stores the original and keeps the journal
            raise Refusal("publish_failed", f"Windows stopped the swap half-way: {target.name} is missing and could not "
                          "be put back. Its earlier state is kept in the undo store; once a file is back at the path, "
                          f"undo_document(path, undo_token='{txid}', force=True) restores it.",
                          {"published": True, "undo_token": txid}, changed_nothing=False) from exc
        if not state.get("replaced"):
            ck.abort(txid)
            raise
        warning = _settle(target, txid, displaced, cur_sha)
        ck.mark_replaced(txid)
        if not isinstance(exc, Exception):
            raise
        raise Refusal("publish_failed", f"{target.name} was replaced with the new state (sha {new_sha[:12].upper()}), "
                      f"but the check after the swap failed ({type(exc).__name__}). The journal keeps the write: the "
                      f"next call on this document records it, and undo_document(path, undo_token='{txid}') restores "
                      "the earlier state. Inspect the document again before the next edit."
                      + (f" {warning}" if warning else ""),
                      {"published": True, "verified": False, "sha256_after": new_sha.upper(), "undo_token": txid},
                      changed_nothing=False) from exc
    warning = _settle(target, txid, displaced, cur_sha)
    return [warning] if warning else []


def _commit(target: Path, txid: str, sha_after: str, warnings: list | None = None) -> dict:
    """Record the published write. A failure here is reported as a write that happened."""
    try:
        return ck.commit(target, txid, warnings)
    except Exception as exc:  # noqa: BLE001
        raise Refusal("publish_failed", f"{target.name} was written (sha {sha_after[:12].upper()}) but its history "
                      f"entry could not be recorded ({type(exc).__name__}). The journal keeps it: the next call on "
                      f"this document records it, and undo_document(path, undo_token='{txid}') then restores the "
                      "earlier state.", {"published": True, "sha256_after": sha_after.upper(), "undo_token": txid},
                      changed_nothing=False) from exc


def publish_inplace(staged: Path, target: Path, *, expected_sha256: str, tool: str | None = None,
                    op: str | None = None) -> dict:
    """Replace ``target`` with ``staged`` if ``target`` still has ``expected_sha256``.

    The pre-edit state is stored and verified in the checkpoint store first (refused ``store_full`` when it
    cannot fit, ``wrong_sha`` when the document changed, ``doc_open`` when Word holds it, ``recovery_pending``
    while an earlier write on it could not be recorded yet). Returns
    {"path", "undo_token", "sha256_before", "sha256_after"}; ``undo_token`` is None when the staged bytes
    equal the source (nothing was written).
    """
    staged, target = Path(staged), Path(target).resolve()
    expected = expected_sha256.lower()
    tool = tool or _current("tool") or "write"
    op = op or _current("op")
    try:
        with ck.doc_lock(target):
            ck.recover(target)
            check_source(target, expected)
            sha_after = file_sha256(staged)
            if sha_after == expected:
                return note_write({"path": str(target), "undo_token": None, "sha256_before": expected.upper(),
                                   "sha256_after": expected.upper(), "publish_warning": None})
            txid = ck.begin(target, {"kind": "edit", "tool": tool, "op": op, "sha_before": expected,
                                     "sha_after": sha_after, "stored": True}, {expected: target})
            warnings = _publish(staged, target, expected, sha_after, txid)
            entry = _commit(target, txid, sha_after, warnings)
    except OSError as exc:
        raise Refusal("publish_failed", f"Could not safely publish {target.name}: {exc}. Original document was not "
                      "replaced. Check file access and free disk space, then inspect again.") from exc
    return note_write({"path": str(target), "undo_token": entry["id"], "sha256_before": entry["sha_before"].upper(),
                       "sha256_after": sha_after.upper(), "publish_warning": " ".join(warnings) or None})


# ---------------------------------------------------------------- checkpoint / undo / history

def _doc_path(path, *, must_exist: bool = True) -> Path:
    if path is None or str(path).strip() == "":
        raise Refusal("bad_request", "path is required")
    p = Path(path).resolve()
    if p.suffix.lower() != ".docx":
        raise Refusal("bad_request", f"{p.name} is not a .docx (only .docx documents are supported)")
    store = ck.home()
    if p == store or store in p.parents:
        raise Refusal("outside_root", "paths inside the server's own store are not documents to edit")
    if must_exist and not p.is_file():
        raise Refusal("bad_request", f"{p} does not exist")
    return p


def _refuse_locked(p: Path, action: str) -> None:
    locked, reason = is_file_locked(p)
    if locked:
        raise Refusal("doc_open", f"Cannot {action}: {p.name} is open or locked ({reason}). Save and close it in "
                      "Word, then retry. Nothing was written.", {"lock_reason": reason})


def _recover_quietly(p: Path) -> list[dict]:
    """Recover what can be recovered; returns the journalled writes still unresolved (none when a writer holds
    the document: it records its own entry)."""
    try:
        with ck.doc_lock(p, timeout=2.0):
            return ck.recover_report(p)[1]
    except Refusal:
        return []  # a writer holds the document, or this is not Windows Python


def checkpoint_document(path, label: str) -> dict:
    """Store the document's current state under ``label``; ``undo_document(path, to=label)`` restores it."""
    p = _doc_path(path)
    if not label or not str(label).strip():
        raise Refusal("bad_request", "checkpoint needs a label (undo_document(path, to=label) restores it)")
    with ck.doc_lock(p):
        ck.recover(p)
        sha = file_sha256(p)
        prev = ck.last(ck.load(p))
        ext = prev is not None and prev["sha_after"] != sha
        e = ck.record(p, {"kind": "checkpoint", "tool": "checkpoint_document", "label": label, "sha_before": sha,
                          "sha_after": sha, "stored": True}, {sha: p})
    return {"path": str(p), "checkpoint": e["id"], "label": label, "sha256": sha.upper(),
            "changed_outside_since_last_record": ext,
            "message": f"Checkpoint '{label}' stored for {p.name} (sha {sha[:12].upper()}); "
                       f"undo_document(path, to='{label}') restores it."
                       + (" The document had changed outside the server since the last record; noted in history."
                          if ext else "")}


def write_new(src: Path, out: Path, sha: str) -> None:
    """Copy ``src`` to a new file ``out``: a verified sibling temp file, then an atomic no-clobber rename (a
    file that appears at ``out`` meanwhile is never overwritten). Nothing is left behind on failure."""
    sha = sha.lower()
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_name(f".~research-mcp-{uuid.uuid4().hex[:8]}.docx.tmp")
    try:
        shutil.copyfile(src, tmp)
        if file_sha256(tmp) != sha:
            raise Refusal("publish_failed", "the output copy did not verify; nothing was written")
        try:
            if os.name == "nt":
                os.rename(tmp, out)  # MoveFileEx without REPLACE_EXISTING: fails if out exists
            else:
                os.link(tmp, out)  # fails if out exists
        except FileExistsError:
            raise Refusal("bad_request", f"output {out.name} appeared while it was being written; it was left "
                          "untouched and nothing was written. Give a new file name.") from None
    finally:
        tmp.unlink(missing_ok=True)
    if file_sha256(out) != sha:
        raise Refusal("publish_failed", f"output {out.name} did not verify after the rename", changed_nothing=False)


def _new_output(output) -> Path:
    out = Path(output).resolve()
    if out.suffix.lower() != ".docx":
        raise Refusal("bad_request", "output must be a .docx path")
    if out.exists():
        raise Refusal("bad_request", f"output {out.name} exists; give a new file name")
    if out == ck.home() or ck.home() in out.parents:
        raise Refusal("outside_root", "output inside the server's own store")
    return out


def _restore(p: Path, src: Path, restore: str, cur: str, entry: dict) -> dict:
    """Swap the stored state ``src`` (sha ``restore``) over ``p`` (sha ``cur``), storing ``cur`` first so this
    restore can itself be undone. Call under the document lock."""
    check_source(p, cur)
    with tempfile.TemporaryDirectory(prefix="research-mcp-undo-", dir=p.parent, ignore_cleanup_errors=True) as folder:
        tmp = Path(folder) / p.name
        shutil.copyfile(src, tmp)
        if file_sha256(tmp) != restore:
            raise Refusal("internal", "the stored state did not verify; nothing was written", {"sha": restore.upper()})
        txid = ck.begin(p, entry, {cur: p}, keep=(restore,) if ck.has_blob(restore) else ())
        warnings = _publish(tmp, p, cur, restore, txid)
    return dict(_commit(p, txid, restore, warnings), publish_warning=" ".join(warnings) or None)


def _result(p: Path, *, restored: bool, cur: str, after: str, message: str, undo_token=None, undid=None,
            forced=False, source="checkpoint", backup_path=None, output=None, restored_sha=None,
            entry: dict | None = None) -> dict:
    """A restore result. ``entry`` (the recorded undo) can carry a publish warning: a save that landed just
    before the swap, kept in the store; its sha is then the real ``sha256_before``."""
    warning = (entry or {}).get("publish_warning")
    if entry is not None:
        cur = entry["sha_before"]
    return {"path": str(p), "restored": restored, "undo_token": undo_token,
            "restored_sha256": (restored_sha or after).upper(), "sha256_before": cur.upper(),
            "sha256_after": after.upper(), "sha256": after.upper(), "undid": undid, "forced": forced,
            "source": source, "backup_path": backup_path, "output": output,
            "message": message + (f" WARNING: {warning}" if warning else ""), "publish_warning": warning}


def _stale(p: Path, cur: str, last: str | None, why: str) -> Refusal:
    return Refusal("stale", f"{p.name} {why} (current sha {cur.upper()}; last sha recorded by the MCP "
                   f"{last.upper() if last else 'none'}). Undo refused; nothing was written. Pass force=True to "
                   "restore anyway (the current state is checkpointed first, so that can be undone too) or "
                   "output=<new .docx path> to write the earlier state as a new file.",
                   {"current_sha256": cur.upper(), "last_recorded_sha256": last.upper() if last else None})


def undo_document(path, undo_token: str | None = None, to: str | None = None, force: bool = False,
                  output: str | None = None, *, tool: str = "undo_document") -> dict:
    """Restore the state before the last MCP edit (or before ``undo_token``, or checkpoint ``to``).

    Refuses (``stale``) when the document changed since the MCP last recorded it, unless ``force`` (the
    overwritten state is checkpointed first) or ``output`` (the earlier state is written as a new file and
    the document is left as it is). Refuses (``doc_open``) while Word holds the file. With no recorded edits,
    a legacy ``<name>.docx.bak`` beside the document is the fallback; it is read, never moved or deleted.
    """
    p = _doc_path(path)
    out = _new_output(output) if output is not None else None
    if out is None:
        _refuse_locked(p, "undo")
    with ck.doc_lock(p):
        ck.recover(p)
        cur = file_sha256(p)
        hist = ck.load(p)
        try:
            plan = ck.plan_undo(hist, cur, token=undo_token, to=to)
        except Refusal as r:
            if r.code == "not_found" and not undo_token and not to and ck.last(hist, ("edit", "undo")) is None:
                return _bak_fallback(p, cur, hist, force=force, out=out, tool=tool)
            raise
        e, restore = plan["entry"], plan["restore_sha"]
        what = ck.describe(e)
        if out is not None:
            write_new(ck.blob_path(restore), out, restore)
            return _result(p, restored=True, cur=cur, after=cur, restored_sha=restore, undid=what, output=str(out),
                           message=f"Wrote the state before {e['kind']} {e['id']} to {out.name}; {p.name} is unchanged.")
        if plan["stale"] and not force:
            raise _stale(p, cur, plan["last_recorded_sha"], "changed outside the server since the MCP last recorded it")
        if restore == cur:
            return _result(p, restored=False, cur=cur, after=cur, undid=what,
                           message=f"{p.name} is already in that state; nothing was written.")
        entry = _restore(p, ck.blob_path(restore), restore, cur,
                         {"kind": "undo", "tool": tool, "sha_before": cur, "sha_after": restore, "stored": True,
                          "undid": e["id"], "op": f"undo {e['kind']} {e.get('label') or e.get('op') or e['id']}"})
    label = e.get("label")
    how = (f"checkpoint '{label}'" if e["kind"] == "checkpoint" else
           f"{e['kind']} {e['id']} ({e.get('tool')}: {e.get('op')}, {e['time']})")
    return _result(p, restored=True, cur=cur, after=restore, undo_token=entry["id"], undid=what,
                   forced=bool(plan["stale"]), entry=entry,
                   message=f"Restored {p.name} to the state {'of' if e['kind'] == 'checkpoint' else 'before'} {how}; "
                           f"sha {restore[:12].upper()}. undo_document(path) again reverses this undo."
                           + (" The overwritten (externally changed) state was checkpointed first."
                              if plan["stale"] else ""))


_ENVELOPE = {"structured_content", "structuredContent", "is_error", "isError", "content", "result_type"}
_NOT_WRITTEN = {"refused", "failed", "no_changes", "preview", "error"}


def _logged_payload(res) -> dict | None:
    """The tool's own result from a traffic-log record, or None unless it is a successful write payload.

    The logger stores the MCP envelope (meta / content / structured_content / is_error / result_type) whose
    fields sit under ``structured_content``; older records hold the flat result."""
    if not isinstance(res, dict):
        return None
    if _ENVELOPE & set(res) and "source" not in res:
        if res.get("is_error") or res.get("isError"):
            return None
        sc = res.get("structured_content", res.get("structuredContent"))
        if isinstance(sc, dict) and isinstance(sc.get("result"), dict) and "source" not in sc:
            sc = sc["result"]  # a tool whose return was wrapped as {"result": ...}
        if not isinstance(sc, dict):
            for block in res.get("content") or []:
                try:
                    sc = json.loads(block["text"]) if isinstance(block, dict) and "text" in block else None
                except ValueError:
                    sc = None
                if isinstance(sc, dict):
                    break
        res = sc if isinstance(sc, dict) else None
    if res is None or res.get("is_error") or res.get("error") or str(res.get("status", "")).lower() in _NOT_WRITTEN:
        return None
    return res


def _logged_last_write(p: Path, max_bytes: int = 16 * 1024 ** 2) -> dict | None:
    """The last in-place write of ``p`` in the MCP traffic log (its tail): {"sha_before", "sha_after"}."""
    from .logger import get_traffic_log_path

    try:
        log = get_traffic_log_path()
        with open(log, "rb") as f:
            f.seek(max(0, log.stat().st_size - max_bytes))
            lines = f.read().decode("utf-8", errors="replace").splitlines()
    except OSError:
        return None
    want = ck.norm_path(p)
    for line in reversed(lines):
        if p.name not in line:
            continue
        try:
            entry = json.loads(line)
        except ValueError:
            continue
        res = _logged_payload(entry.get("result")) if entry.get("status") == "success" else None
        if res is None or not res.get("source") or not res.get("output"):
            continue
        try:
            if ck.norm_path(res["output"]) != want or ck.norm_path(res["source"]) != want:
                continue
        except (OSError, ValueError, TypeError):
            continue
        before = res.get("sha256_before") or res.get("source_sha256")
        after = res.get("sha256_after") or res.get("output_sha256")
        if isinstance(before, str) and isinstance(after, str):
            return {"sha_before": before.lower(), "sha_after": after.lower()}
    return None


def _bak_fallback(p: Path, cur: str, hist: dict, *, force: bool, out: Path | None, tool: str) -> dict:
    """No recorded edits: restore a legacy ``.docx.bak`` (from before the undo store). Call under the doc lock.

    The stale guard needs the sha the MCP recorded after its last edit; for a legacy .bak that comes from the
    traffic log (the last logged in-place write must have gone from the .bak's sha to the current sha). When
    the log has no such record the document counts as changed. The .bak is never moved or deleted."""
    bak = p.with_name(p.name + ".bak")
    if not bak.is_file():
        gone = hist.get("dropped", [])[-5:]
        raise Refusal("not_found", "nothing to undo: this document has no recorded edits in the undo store and no "
                      "legacy .bak beside it" + (" (older entries were pruned)" if gone else "")
                      + ". OneDrive version history is the outer fallback.", {"pruned": gone})
    bak_sha = file_sha256(bak)
    undid = {"kind": "legacy_bak", "backup_path": str(bak), "sha256": bak_sha.upper()}
    if out is not None:
        write_new(bak, out, bak_sha)
        return _result(p, restored=True, cur=cur, after=cur, restored_sha=bak_sha, undid=undid, source="bak",
                       backup_path=str(bak), output=str(out),
                       message=f"Wrote the legacy backup {bak.name} to {out.name}; {p.name} is unchanged.")
    logged = _logged_last_write(p)
    verified = logged is not None and logged["sha_after"] == cur and logged["sha_before"] == bak_sha
    if bak_sha == cur:
        return _result(p, restored=False, cur=cur, after=cur, undid=undid, source="bak", backup_path=str(bak),
                       message=f"{p.name} already matches {bak.name}; nothing was written.")
    if not verified and not force:
        raise _stale(p, cur, logged["sha_after"] if logged else None,
                     f"has no recorded edits, only a legacy backup {bak.name}, and the MCP log does not show that the "
                     "document is unchanged since the edit that backup preceded")
    entry = _restore(p, bak, bak_sha, cur, {"kind": "undo", "tool": tool, "sha_before": cur, "sha_after": bak_sha,
                                            "stored": True, "undid": "legacy .bak", "op": f"restore {bak.name}"})
    return _result(p, restored=True, cur=cur, after=bak_sha, undo_token=entry["id"], undid=undid,
                   forced=not verified, source="bak", backup_path=str(bak), entry=entry,
                   message=f"Restored {p.name} from the legacy backup {bak.name} (left in place); sha "
                           f"{bak_sha[:12].upper()}. The previous state was checkpointed: undo_document(path) "
                           "reverses this.")


def document_history(path) -> dict:
    p = _doc_path(path)
    stuck = _recover_quietly(p)
    hist = ck.load(p)
    cur = file_sha256(p)
    prev = ck.last(hist)
    bak = p.with_name(p.name + ".bak")
    return {"path": str(p), "sha256": cur.upper(), "entries": [ck.describe(e) for e in reversed(hist["entries"])],
            "dropped": hist.get("dropped", [])[-20:], "limits": ck.limits(),
            "changed_outside_since_last_record": bool(prev and prev["sha_after"] != cur),
            "legacy_backup": str(bak) if bak.is_file() else None,
            "store": str(ck.checkpoints_root() / ck.doc_key(p)), "recovery_pending": stuck}
