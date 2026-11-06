"""Checkpoint store and history: what one-call undo restores.

Notes:
  - each document's history sits under ``checkpoints/<doc-id>/`` next to the shared ``blobs/`` and
    ``journal/`` folders; blobs are ``<sha>.docx``; the busy refusal is ``doc_busy``;
  - the store lives under RESEARCH_MCP_HOME or %LOCALAPPDATA%\\research_mcp and is refused inside OneDrive;
  - ``Refusal`` is a ``RefusedUnsafe`` subclass, so every refusal path handles it;
  - shas are kept lower-case inside the store (``file_sha256`` is upper-case; callers compare
    case-insensitively and report upper-case);
  - the doc-id normalises WSL ``/mnt/<drive>/`` paths and ``\\\\?\\`` prefixes so a read from either side
    finds the same history;
  - limits come from RESEARCH_MCP_MAX_ENTRIES / RESEARCH_MCP_MAX_DAYS / RESEARCH_MCP_MAX_BYTES (30 / 14 / 1 GB).

Layout (under checkpoints_root(), outside OneDrive):
  blobs/<sha256>.docx         content-addressed document states, shared across documents
  <doc-id>/history.json       {"path", "entries": [...], "dropped": [...]}
  <doc-id>/.lock              OS byte-range lock on a held handle (with the in-process lock: one writer per doc)
  journal/<id>.json           a write between storing its states and recording its history entry
  .store.lock                 OS lock: one writer of blobs, journal and histories at a time

Locks are OS locks (msvcrt.locking on Windows, flock elsewhere) held on an open handle: the OS releases
them when the holder exits, so a lock is never taken over on age, and the lock files are never deleted.
Windows msvcrt locks and WSL flock locks do not see each other, so writes run on Windows Python only
(``require_windows``): elsewhere every lock refuses ``windows_only`` and read-only tools still work.

A journal or history that cannot be read or does not validate refuses the write (store_unreadable): pruning
without it could delete a state another document's undo needs.

Entry kinds:
  edit        a write tool or batch; ``sha_before`` is stored, ``undo_token`` = id
  checkpoint  an explicit snapshot with a label; its sha is stored
  undo        an undo; ``sha_before`` (the state it overwrote) is stored, so undo of undo works
  external    the document changed outside the server between two recorded states (no blob), or, stored,
              a save that landed just before an MCP swap (its state is the next entry's sha_before)

Retention: 30 entries per document, 14 days, 1 GB of blobs in total (env RESEARCH_MCP_MAX_ENTRIES,
RESEARCH_MCP_MAX_DAYS, RESEARCH_MCP_MAX_BYTES override). The byte cap is hard: any document's entries may
go to make room, and a write whose own states do not fit is refused (store_full) before it publishes.
Pruned entries are listed in "dropped" so a refused undo can say why; blobs nothing references are deleted.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import sys
import threading
import time
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from .audit import file_sha256 as _file_sha256_upper
from .errors import RefusedUnsafe

HEX64 = re.compile(r"[0-9a-f]{64}")
_LOCKS: dict[str, threading.Lock] = {}
_LOCKS_GUARD = threading.Lock()


class Refusal(RefusedUnsafe):
    """An explicit refusal: nothing was written. ``code`` is stable; ``details`` holds the measured values.
    ``changed_nothing=False`` only for a failure after a document was published (the message says so)."""

    def __init__(self, code: str, message: str, details: dict | None = None, *, changed_nothing: bool = True):
        super().__init__(message)
        self.code = code
        self.details = details or {}
        self.changed_nothing = changed_nothing


def file_sha256(path: str | Path) -> str:
    return _file_sha256_upper(path).lower()


# ---------------------------------------------------------------- paths

def home() -> Path:
    if os.environ.get("RESEARCH_MCP_HOME"):
        p = Path(os.environ["RESEARCH_MCP_HOME"])
    else:
        p = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local"))) / "research_mcp"
    return p.resolve()


def assert_outside_onedrive(path: Path) -> Path:
    if "onedrive" in str(Path(path).resolve()).lower():
        raise RuntimeError(f"refusing to keep server state inside OneDrive: {path}")
    return Path(path)


def checkpoints_root() -> Path:
    p = assert_outside_onedrive(home() / "checkpoints")
    p.mkdir(parents=True, exist_ok=True)
    return p


def limits() -> dict:
    from . import config

    return {"entries": int(config.get("RESEARCH_MCP_MAX_ENTRIES", "30")),
            "days": float(config.get("RESEARCH_MCP_MAX_DAYS", "14")),
            "bytes": int(config.get("RESEARCH_MCP_MAX_BYTES", str(1024 ** 3)))}


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


_WSL = re.compile(r"^/mnt/([a-zA-Z])(/.*)?$")


def norm_path(path: Path | str) -> str:
    """The resolved path as Windows spells it, lower-cased: one doc-id for every alias of one file.

    A WSL spelling (/mnt/c/...) is converted before resolving on Windows, where resolving it first would
    give C:\\mnt\\c\\...; elsewhere the resolved path is converted."""
    raw = str(path)
    m = _WSL.match(raw.replace("\\", "/"))
    if m and on_windows():
        raw = f"{m.group(1)}:{(m.group(2) or '/')}".replace("/", "\\")
    s = str(Path(raw).resolve())
    m = _WSL.match(s.replace("\\", "/"))
    if m:
        s = f"{m.group(1)}:{(m.group(2) or '/')}".replace("/", "\\")
    for prefix in ("\\\\?\\UNC\\", "\\\\?\\"):
        if s.startswith(prefix):
            s = ("\\\\" if prefix.endswith("UNC\\") else "") + s[len(prefix):]
    return s.lower()


def doc_key(path: Path) -> str:
    p = Path(path)
    stem = re.sub(r"[^\w.-]", "_", p.stem)[:40]
    return f"{stem}-{hashlib.sha256(norm_path(p).encode()).hexdigest()[:12]}"


def blobs_dir() -> Path:
    d = checkpoints_root() / "blobs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def doc_dir(path: Path) -> Path:
    d = checkpoints_root() / doc_key(path)
    d.mkdir(parents=True, exist_ok=True)
    return d


def blob_path(sha: str) -> Path:
    return blobs_dir() / f"{sha.lower()}.docx"


def store_blob(src: Path, sha: str | None = None) -> str:
    """Copy ``src`` into the store (once per content) and verify it. Returns its sha."""
    sha = (sha or file_sha256(src)).lower()
    dst = blob_path(sha)
    if dst.exists() and file_sha256(dst) == sha:
        return sha
    tmp = dst.with_suffix(f".{uuid.uuid4().hex[:6]}.tmp")
    try:
        shutil.copyfile(src, tmp)
        got = file_sha256(tmp)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise
    if got != sha:
        tmp.unlink(missing_ok=True)
        raise Refusal("wrong_sha", f"the stored copy of {Path(src).name} is not the expected state (the file changed "
                      "while it was being copied, or the copy failed); nothing was written. Inspect again.",
                      {"expected": sha.upper(), "actual": got.upper()})
    os.replace(tmp, dst)
    return sha


def has_blob(sha: str) -> bool:
    return bool(sha) and blob_path(sha).exists()


# ---------------------------------------------------------------- locking

_STORE_LOCK = threading.RLock()
_TL = threading.local()


def on_windows() -> bool:
    return sys.platform == "win32"


def require_windows() -> None:
    """Refuse a write (``windows_only``) unless this is Windows Python. A server under WSL would take flock
    locks that a Windows server's msvcrt locks never see: both could write one document or prune one store."""
    if not on_windows():
        raise Refusal("windows_only", f"writes need the server on Windows Python, not {sys.platform}: its document "
                      "and store locks would not exclude a Windows server (WSL flock and Windows msvcrt locks do not "
                      "interlock). Nothing was changed; read-only tools still work.", {"platform": sys.platform})


if os.name == "nt":
    import msvcrt

    def _try_lock(fd: int) -> bool:
        os.lseek(fd, 0, os.SEEK_SET)
        try:
            msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            return False

    def _unlock(fd: int) -> None:
        os.lseek(fd, 0, os.SEEK_SET)
        msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
else:
    import fcntl

    def _try_lock(fd: int) -> bool:
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
            return True
        except OSError:
            return False

    def _unlock(fd: int) -> None:
        fcntl.flock(fd, fcntl.LOCK_UN)


def _hold_file(lockfile: Path, timeout: float, message: str) -> int:
    """Lock byte 0 of ``lockfile`` through an open handle; returns the handle (pass it to ``_release``).

    The OS lock lasts until the handle is closed or its process exits, however old the file is: nothing is
    ever taken over from a live holder, and a dead holder's lock is gone with it."""
    fd = os.open(lockfile, os.O_CREAT | os.O_RDWR | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOINHERIT", 0))
    t0 = time.monotonic()
    while not _try_lock(fd):
        if time.monotonic() - t0 > timeout:
            os.close(fd)
            raise Refusal("doc_busy", message, {"lock": str(lockfile)})
        time.sleep(0.05)
    return fd


def _release(fd: int) -> None:
    try:
        _unlock(fd)
    finally:
        os.close(fd)


@contextmanager
def doc_lock(path: Path, timeout: float = 30.0):
    require_windows()
    key = doc_key(path)
    with _LOCKS_GUARD:
        lk = _LOCKS.setdefault(key, threading.Lock())
    if not lk.acquire(timeout=timeout):
        raise Refusal("doc_busy", "another call is writing this document; retry when it finishes")
    lockfile = doc_dir(path) / ".lock"
    try:
        fd = _hold_file(lockfile, timeout, "another server process is writing this document (lock held); retry later")
        try:
            yield
        finally:
            _release(fd)
    finally:
        lk.release()


@contextmanager
def store_lock(timeout: float = 30.0):
    """One writer of the shared store (blobs, journal, every document's history) across threads and processes.
    Re-entrant in a thread. Lock order: a document lock first, then the store lock."""
    require_windows()
    if not _STORE_LOCK.acquire(timeout=timeout):
        raise Refusal("doc_busy", "the checkpoint store is busy; retry when the other call finishes")
    depth = getattr(_TL, "depth", 0)
    lockfile = checkpoints_root() / ".store.lock"
    try:
        fd = None
        if depth == 0:
            lockfile.parent.mkdir(parents=True, exist_ok=True)
            fd = _hold_file(lockfile, timeout, "another server process is updating the checkpoint store (lock held); "
                                               "retry later")
        _TL.depth = depth + 1
        try:
            yield
        finally:
            _TL.depth = depth
            if fd is not None:
                _release(fd)
    finally:
        _STORE_LOCK.release()


# ---------------------------------------------------------------- history

def _hist_path(path: Path) -> Path:
    return doc_dir(path) / "history.json"


def _write_json(target: Path, obj) -> None:
    """Write durably: a sibling temp file, flushed to disk, then an atomic replace."""
    tmp = target.with_name(f"{target.name}.{uuid.uuid4().hex[:6]}.tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(json.dumps(obj, indent=1, ensure_ascii=False))
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, target)
    finally:
        tmp.unlink(missing_ok=True)


def _unreadable(p: Path, why: str) -> Refusal:
    return Refusal("store_unreadable", f"the checkpoint store file {p} cannot be read or is not valid ({why}); "
                   "nothing was written, because pruning without it could delete a state an undo needs. Check the "
                   "file's permissions (or move a damaged file aside), then retry.", {"file": str(p), "reason": why})


def _read_json(p: Path):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except OSError as exc:
        raise _unreadable(p, type(exc).__name__) from exc
    except ValueError as exc:
        raise _unreadable(p, "not valid JSON") from exc


def _sha_ok(s) -> bool:
    return isinstance(s, str) and HEX64.fullmatch(s) is not None


def _entry_ok(e, timed: bool) -> bool:
    if not (isinstance(e, dict) and isinstance(e.get("id"), str) and isinstance(e.get("kind"), str)
            and isinstance(e.get("sha_before"), str)):
        return False
    if e.get("stored") and not _sha_ok(e["sha_before"]):
        return False
    if timed or "time" in e:
        try:
            datetime.fromisoformat(e["time"])
        except (KeyError, TypeError, ValueError):
            return False
    return True


def read_history(hp: Path) -> dict:
    """A history file, validated; refuses store_unreadable when it cannot be read or checked."""
    h = _read_json(hp)
    if not (isinstance(h, dict) and isinstance(h.get("entries"), list) and all(_entry_ok(e, True) for e in h["entries"])
            and isinstance(h.get("dropped", []), list)):
        raise _unreadable(hp, "unexpected content")
    h.setdefault("dropped", [])
    return h


def read_journal(jp: Path) -> dict:
    j = _read_json(jp)
    if not (isinstance(j, dict) and isinstance(j.get("tx"), str) and isinstance(j.get("doc"), str)
            and _entry_ok(j.get("entry"), False) and j["entry"]["id"] == j["tx"]
            and isinstance(j.get("blobs"), list) and all(_sha_ok(s) for s in j["blobs"])):
        raise _unreadable(jp, "unexpected content")
    return j


def load(path: Path) -> dict:
    hp = _hist_path(path)
    if hp.exists():
        return read_history(hp)
    return {"path": str(Path(path).resolve()), "entries": [], "dropped": []}


def save(path: Path, hist: dict) -> None:
    with store_lock():
        _write_json(_hist_path(path), hist)


def last(hist: dict, kinds=None) -> dict | None:
    for e in reversed(hist["entries"]):
        if kinds is None or e["kind"] in kinds:
            return e
    return None


def new_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:10]}"


def note_external(hist: dict, current_sha: str) -> dict | None:
    """Record that the document changed outside the server since the last recorded state."""
    prev = last(hist)
    if prev is None or prev["sha_after"] == current_sha:
        return None
    e = {"id": new_id("x"), "kind": "external", "time": now_iso(), "sha_before": prev["sha_after"],
         "sha_after": current_sha, "stored": False}
    hist["entries"].append(e)
    return e


# ---------------------------------------------------------------- journal (write-ahead record of a publication)
#
# A write that publishes a document first stores the states it needs and writes journal/<id>.json (the history
# entry to add, the blobs it keeps), then publishes, then commits (history entry added, journal deleted). A
# crash or a failed history write between publish and commit leaves the journal; recover() finishes it on
# the next call for that document. Journalled blobs are never pruned.

def journal_dir() -> Path:
    d = checkpoints_root() / "journal"
    d.mkdir(parents=True, exist_ok=True)
    return d


def pending() -> list[dict]:
    """Every journal, validated. One that cannot be read refuses (store_unreadable): its blobs are unknown."""
    return [read_journal(jp) for jp in sorted(journal_dir().glob("*.json"))]


def _size(sha: str, src: Path | None = None) -> int:
    if has_blob(sha):
        return blob_path(sha).stat().st_size
    return Path(src).stat().st_size if src is not None else 0


def begin(path: Path, entry: dict, blobs: dict, keep=()) -> str:
    """Before publishing: make room, store ``blobs`` ({sha: source file}) and journal ``entry``.

    Refuses ``store_full`` (nothing stored or written) when the states that must be kept (this write's, the
    restore source ``keep`` and other pending writes') do not fit under the byte cap."""
    entry = dict(entry)
    entry.setdefault("id", new_id({"edit": "u", "checkpoint": "c", "undo": "r"}.get(entry["kind"], "e")))
    with store_lock():
        cap = limits()["bytes"]
        need = {s.lower(): _size(s, src) for s, src in blobs.items()}
        need.update({s.lower(): _size(s) for s in keep})
        for j in pending():
            for s in j.get("blobs", []):
                need.setdefault(s, _size(s))
        if sum(need.values()) > cap:
            raise Refusal("store_full", f"the undo store cap ({cap} bytes) is smaller than the states this write must "
                          f"keep ({sum(need.values())} bytes); nothing was written. Raise RESEARCH_MCP_MAX_BYTES.",
                          {"cap_bytes": cap, "needed_bytes": sum(need.values())})
        prune(path, extra=need)
        for s, src in blobs.items():
            store_blob(src, s)
        _write_json(journal_dir() / f"{entry['id']}.json",
                    {"tx": entry["id"], "doc": str(Path(path).resolve()), "entry": entry,
                     "blobs": sorted({s.lower() for s in blobs} | {s.lower() for s in keep}),
                     "created": now_iso()})
    return entry["id"]


def commit(path: Path, txid: str, warnings: list | None = None) -> dict:
    """After publishing: add the journalled entry to the document's history, prune, drop the journal. A store
    file that cannot be read skips the pruning (a warning), never the history entry."""
    with store_lock():
        jp = journal_dir() / f"{txid}.json"
        j = read_journal(jp)
        hist = load(path)
        entry = next((e for e in hist["entries"] if e["id"] == txid), None)
        if entry is None:
            entry = dict(j["entry"])
            entry.setdefault("time", now_iso())
            raced = j.get("raced")
            if raced:  # saved outside the server between the last check and the swap: both states are kept
                note_external(hist, raced["from"])
                hist["entries"].append({"id": new_id("x"), "kind": "external", "time": now_iso(),
                                        "sha_before": raced["from"], "sha_after": raced["to"], "stored": True,
                                        "op": "saved outside the server just before an MCP write replaced it"})
            note_external(hist, entry["sha_before"])
            hist["entries"].append(entry)
        try:
            prune(path, hist, keep_ids=(txid,))
        except Refusal as r:
            if r.code != "store_unreadable":
                raise
            if warnings is not None:
                warnings.append(f"the undo store was not pruned this time: {r}")
        save(path, hist)
        jp.unlink(missing_ok=True)
    return entry


def abort(txid: str) -> None:
    """Nothing was published: drop the journal (its blobs go at the next prune unless referenced)."""
    with store_lock():
        (journal_dir() / f"{txid}.json").unlink(missing_ok=True)


def note_displaced(txid: str, displaced: Path) -> None:
    """Before the swap: where the swap moves the file it displaces, so recover() can find it after a crash."""
    with store_lock():
        jp = journal_dir() / f"{txid}.json"
        _write_json(jp, dict(read_journal(jp), displaced=str(displaced)))


def mark_raced(txid: str, displaced_sha: str) -> None:
    """After the swap: the displaced file (already stored as a blob) is the state this write overwrote. When it
    is not the checked state, the document was saved outside the server just before the swap: the entry's
    ``sha_before`` becomes that state (so undo brings it back) and commit records the outside change."""
    displaced_sha = displaced_sha.lower()
    with store_lock():
        jp = journal_dir() / f"{txid}.json"
        j = read_journal(jp)
        e = dict(j["entry"])
        if displaced_sha != e["sha_before"]:
            j["raced"] = {"from": e["sha_before"], "to": displaced_sha}
            e["sha_before"] = displaced_sha
            j["blobs"] = sorted(set(j["blobs"]) | {displaced_sha})
        _write_json(jp, dict(j, entry=e, replaced=True))


def displaced_path(target: Path, txid: str) -> Path:
    """Where a swap puts the file it displaces: a dot-named file beside the document (not hidden; same volume, as
    ReplaceFileW needs), never inside a temporary folder that is cleaned up automatically. It is named in the
    journal and deleted only after its state is stored, verified and recorded (settle_displaced). A name that is
    already taken gets a suffix, so an earlier leftover is never overwritten."""
    d = Path(target).with_name(f".~research-mcp-displaced-{txid}.docx")
    while d.exists():
        d = d.with_name(f".~research-mcp-displaced-{txid}-{uuid.uuid4().hex[:6]}.docx")
    return d


def settle_displaced(txid: str, displaced: Path) -> str:
    """Store the file a swap displaced, verify the stored copy, record it in the journal (mark_raced), then
    delete it. Returns its sha. Any failure raises with the file left where it is and the journal still naming
    it, so a later recover() can finish."""
    x = file_sha256(displaced)
    store_blob(displaced, x)
    got = file_sha256(blob_path(x))
    if got != x:
        raise Refusal("wrong_sha", f"the stored copy of {Path(displaced).name} did not verify; it was kept in place",
                      {"expected": x.upper(), "actual": got.upper()})
    mark_raced(txid, x)
    try:
        Path(displaced).unlink(missing_ok=True)
    except OSError:
        pass  # its state is stored and recorded; the file is only a leftover copy now
    return x


def mark_replaced(txid: str) -> None:
    """The document was replaced: recovery must record this journal, never drop it (best effort)."""
    try:
        with store_lock():
            jp = journal_dir() / f"{txid}.json"
            _write_json(jp, dict(read_journal(jp), replaced=True))
    except (OSError, Refusal):
        pass


def record(path: Path, entry: dict, blobs: dict) -> dict:
    """An entry with nothing to publish (a checkpoint): store, journal and commit in one step."""
    return commit(path, begin(path, entry, blobs))


def recover_report(path: Path) -> tuple[list[dict], list[dict]]:
    """Finish journalled writes on this document that never reached its history. Call under the doc lock.
    Returns (recorded entries, unresolved journals); an unresolved one keeps its journal and displaced file."""
    p = Path(path).resolve()
    done, stuck = [], []
    with store_lock():
        for j in pending():
            if norm_path(j.get("doc")) != norm_path(p):
                continue
            d = Path(j["displaced"]) if j.get("displaced") else None
            if d is not None and d.is_file():  # the swap ran but the displaced file was not settled
                try:
                    if not p.exists() and not j.get("replaced"):
                        os.rename(d, p)  # the swap stopped half-way: put the original back; nothing was published
                        abort(j["tx"])
                        continue
                    settle_displaced(j["tx"], d)
                    j = read_journal(journal_dir() / f"{j['tx']}.json")
                except (OSError, Refusal) as exc:  # keep the journal and the displaced file for a later call
                    stuck.append({"tx": j["tx"], "displaced": str(d), "reason": f"{type(exc).__name__}: {exc}"})
                    continue
            e = j["entry"]
            try:
                cur = file_sha256(p) if p.exists() else None
            except OSError as exc:  # cannot tell; keep the journal for a later call
                stuck.append({"tx": j["tx"], "displaced": None, "reason": f"{type(exc).__name__}: {exc}"})
                continue
            if cur == e["sha_before"] and not j.get("replaced"):
                abort(j["tx"])  # never published
                continue
            e["recovered"] = ("recorded from the journal after an interrupted write" if cur == e["sha_after"] else
                              "recorded from the journal; the document changed again since, so publication is "
                              "unverified")
            _write_json(journal_dir() / f"{j['tx']}.json", dict(j, entry=e))
            got = commit(p, j["tx"])
            done.append(got)
    return done, stuck


def recover(path: Path) -> list[dict]:
    """recover_report, refusing (``recovery_pending``) while a journalled write on this document is unresolved:
    a new edit, undo or checkpoint would be recorded before it and put undo out of order."""
    done, stuck = recover_report(path)
    if stuck:
        p = Path(path)
        kept = [s["displaced"] for s in stuck if s["displaced"]]
        raise Refusal("recovery_pending", f"{p.name} has an earlier MCP write that could not be recorded yet "
                      f"({stuck[0]['reason']})."
                      + (f" The state it overwrote is kept unchanged at {', '.join(kept)}." if kept else "")
                      + " Edits, undo and checkpoints on this document are refused until it is recorded, so undo "
                      "stays in order; nothing was written. Free space in the undo store (or raise "
                      "RESEARCH_MCP_MAX_BYTES) and check file access, then retry: every call on this document retries "
                      "the recovery. Other documents are not affected; document_history(path) shows the pending write.",
                      {"pending": stuck})
    return done


def recover_all() -> list[dict]:
    """Recover every journalled document whose lock is free (server start)."""
    out = []
    for doc in sorted({j.get("doc") for j in pending() if j.get("doc")}):
        try:
            with doc_lock(Path(doc), timeout=0.5):
                out += recover(Path(doc))
        except Refusal:
            continue
    return out


def append(path: Path, hist: dict, entry: dict) -> dict:
    """Add an entry with nothing to store (kept for callers that only record)."""
    entry.setdefault("id", new_id({"edit": "u", "checkpoint": "c", "undo": "r"}.get(entry["kind"], "e")))
    entry.setdefault("time", now_iso())
    with store_lock():
        hist["entries"].append(entry)
        prune(path, hist, keep_ids=(entry["id"],))
        save(path, hist)
    return entry


# ---------------------------------------------------------------- retention

def _drop(hist: dict, e: dict, reason: str) -> None:
    hist["entries"].remove(e)
    hist.setdefault("dropped", []).append({"id": e["id"], "kind": e["kind"], "time": e["time"], "label": e.get("label"),
                                           "reason": reason, "pruned": now_iso()})
    hist["dropped"] = hist["dropped"][-200:]


def all_histories() -> list[Path]:
    root = checkpoints_root()
    return sorted(root.glob("*/history.json")) if root.exists() else []


def referenced(hists) -> set[str]:
    out = set()
    for h in hists:
        for e in h["entries"]:
            if e.get("stored"):
                out.add(e["sha_before"])
    return out


def prune(path: Path, hist: dict | None = None, keep_ids=(), extra: dict | None = None) -> list[str]:
    """Apply the per-document limits to this document and the global byte cap to every document, under the
    store lock.

    ``hist`` (this document's history, saved by the caller) defaults to the one on disk (then saved here).
    Other documents' histories are re-read and saved inside the lock. Never dropped: ``keep_ids`` and entries
    whose state a pending (journalled) write keeps. ``extra`` ({sha: bytes}) counts states about to be stored."""
    lim, notes = limits(), []
    with store_lock():
        own = hist is None
        if own:
            hist = load(path)
        keep_ids = set(keep_ids)
        cutoff = time.time() - lim["days"] * 86400
        for e in list(hist["entries"]):
            if e["id"] not in keep_ids and datetime.fromisoformat(e["time"]).timestamp() < cutoff:
                _drop(hist, e, f"older than {lim['days']:g} days")
        while len(hist["entries"]) > lim["entries"]:
            e = next((x for x in hist["entries"] if x["id"] not in keep_ids), None)
            if e is None:
                break
            _drop(hist, e, f"more than {lim['entries']} entries")
        me = str(_hist_path(path).resolve())
        hists = [(None, hist)]
        for hp in all_histories():
            if str(hp.resolve()) != me:
                hists.append((hp, read_history(hp)))  # an unreadable one refuses: its states would be unknown
        protected = {s for j in pending() for s in j.get("blobs", [])}
        extra = {s.lower(): n for s, n in (extra or {}).items()}

        def size():
            shas = referenced([h for _, h in hists]) | protected
            return sum(_size(s) for s in shas) + sum(n for s, n in extra.items() if s not in shas)

        total, touched = size(), set()
        while total > lim["bytes"]:
            cands = [(e["time"], i, e) for i, (_, h) in enumerate(hists) for e in h["entries"]
                     if e.get("stored") and e["id"] not in keep_ids and e["sha_before"] not in protected
                     and e["sha_before"] not in extra]
            if not cands:
                break
            _, i, e = min(cands, key=lambda c: c[0])
            _drop(hists[i][1], e, f"checkpoint store above {lim['bytes']} bytes")
            touched.add(i)
            notes.append(f"pruned {e['id']} ({e['kind']}, {e['time']})")
            total = size()
        for i in touched:
            if hists[i][0] is not None:  # another document's history
                _write_json(hists[i][0], hists[i][1])
        if own:
            _write_json(_hist_path(path), hist)
        keep = referenced([h for _, h in hists]) | protected | set(extra)
        for b in blobs_dir().glob("*.docx"):
            if b.stem not in keep:
                b.unlink(missing_ok=True)
    return notes


# ---------------------------------------------------------------- undo planning

def plan_undo(hist: dict, current_sha: str, *, token: str | None = None, to: str | None = None) -> dict:
    """Which state an undo restores. Returns {"entry", "restore_sha", "stale"}; refuses when nothing fits."""
    current_sha = current_sha.lower()
    if token and to:
        raise Refusal("bad_request", "give undo_token or to (a checkpoint label), not both")
    if to:
        e = next((x for x in reversed(hist["entries"]) if x["kind"] == "checkpoint" and x.get("label") == to), None)
        if e is None:
            gone = [d for d in hist.get("dropped", []) if d.get("label") == to]
            raise Refusal("not_found", f"no checkpoint labelled {to!r}" + (" (it was pruned)" if gone else ""),
                          {"labels": [x.get("label") for x in hist["entries"] if x["kind"] == "checkpoint"],
                           "pruned": gone})
        restore = e["sha_before"]
    elif token:
        e = next((x for x in hist["entries"] if x["id"] == token), None)
        if e is None:
            gone = [d for d in hist.get("dropped", []) if d["id"] == token]
            raise Refusal("not_found", f"undo_token {token!r} is not in this document's history" +
                          (f" (pruned: {gone[0]['reason']})" if gone else ""), {"pruned": gone})
        if e["kind"] == "external" and not e.get("stored"):
            raise Refusal("bad_request", "that entry records an outside change; nothing of it is stored")
        restore = e["sha_before"]
    else:
        e = last(hist, ("edit", "undo"))
        if e is None:
            raise Refusal("not_found", "nothing to undo: this document has no recorded edits")
        restore = e["sha_before"]
    if not has_blob(restore):
        raise Refusal("not_found", "the stored state for that entry is missing from the checkpoint store",
                      {"sha": restore})
    prev = last(hist)
    stale = prev is not None and prev["sha_after"] != current_sha
    return {"entry": e, "restore_sha": restore, "stale": stale,
            "last_recorded_sha": prev["sha_after"] if prev else None}


def describe(e: dict) -> dict:
    keys = ("id", "kind", "time", "tool", "op", "label", "summary", "sha_before", "sha_after", "undid", "recovered")
    out = {k: e[k] for k in keys if e.get(k) is not None}
    for k in ("sha_before", "sha_after"):
        if k in out:
            out[k] = out[k].upper()
    return out
