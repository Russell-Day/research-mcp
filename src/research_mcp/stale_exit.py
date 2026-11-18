"""Let a long-running server exit by itself, once idle, when its own git checkout has new runtime commits.

A supervisor (a shell loop, a service manager) that restarts the process on exit then loads the new code, so
fixes go live without anyone remembering to restart the server.

The process exits with ``EXIT_CODE`` (75, EX_TEMPFAIL) when ALL of these hold, checked every
``DEFAULT_INTERVAL_SECONDS`` by a daemon thread:

a. the repository holding the server package (found from the package's ``__file__``) has ``refs/heads/<branch>``
   at a commit other than the one recorded at start, and ``git diff --name-only <start> <now> -- RUNTIME_PATHS``
   is non-empty (commits that touch only docs or tests do not restart the server);
b. the working tree's runtime files equal that commit: ``git diff --name-only <now> -- RUNTIME_PATHS`` and the
   untracked files under them are both empty, so uncommitted work is never loaded (the working tree is compared
   with the commit rather than read through ``git status``, because a commit made from a private index leaves the
   shared index behind and ``git status`` would then report every committed file as changed);
c. the server is idle: no call in flight and none started or finished in the last ``idle_seconds``
   (``ActivityTracker``, fed by the server's call wrapper), and, when ``idle_log`` is given, that file was last
   modified at least ``idle_seconds`` ago.

Then an optional ``before_exit(changed_files)`` callback runs (it may refuse, and the exit is retried at the next
check). After it, git is read once more: if the branch moved, a runtime file differs from the commit or a new
untracked one appeared while the hook ran, or git fails, the process stays up and the next check decides again on
whatever the branch holds then. A last look at the tracker under its lock follows (a call that started meanwhile
keeps the server up). Then one line goes to stderr, ``{at, from, to, files, idle_s}`` is appended to
``<state_dir>/restarts.jsonl`` and the process ends with ``os._exit(EXIT_CODE)``, which also ends it while other
non-daemon threads (a web server's) are still running. A commit or edit landing in the milliseconds between that
last git read and the new process importing its code can still be loaded unchecked; no check inside this process
can close that window.

If the start commit cannot be read when the watcher is made (git timed out or failed, but exists), the code the
process loaded is unknown. The watcher then never adopts a later commit as its start (a commit landing in that gap
would go unnoticed). At the first check git answers: if the last line of ``restarts.jsonl`` was written before
this process started and its ``to`` is the branch's commit now, that commit is what this process loaded (the
branch only moves forward), so it becomes the start; otherwise every runtime file at the branch counts as changed
(``from`` is null in the record) and the process restarts once, when clean and idle like any restart, to load a
known commit. So git timing out at every boot never restarts the server more than once.

Every git call uses ``--no-optional-locks`` and a timeout; a timeout or an error means "not stale this round".
This module is standard library only and names nothing server-specific, so it can be copied unchanged into
other servers.
"""

from __future__ import annotations

import ast
from dataclasses import dataclass, field
import datetime as _dt
import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import time
from typing import Callable, Iterable, Sequence

EXIT_CODE = 75  # EX_TEMPFAIL: a planned restart, not a crash
DEFAULT_IDLE_SECONDS = 300
DEFAULT_INTERVAL_SECONDS = 60
GIT_TIMEOUT_SECONDS = 10
RUNTIME_PATHS: tuple[str, ...] = ("src", "pyproject.toml", "uv.lock")
DEFAULT_BRANCH = "master"
RESTARTS_FILE = "restarts.jsonl"
LOG_PREFIX = "[stale_exit]"


# --------------------------------------------------------------------------- activity


class ActivityTracker:
    """Counts calls in flight and remembers when the last one started or finished. Thread-safe."""

    def __init__(self, clock: Callable[[], float] = time.time) -> None:
        self._clock = clock
        self._lock = threading.Lock()
        self._in_flight = 0
        self._last_activity = clock()

    def start(self) -> None:
        with self._lock:
            self._in_flight += 1
            self._last_activity = self._clock()

    def finish(self) -> None:
        with self._lock:
            self._in_flight = max(0, self._in_flight - 1)
            self._last_activity = self._clock()

    @property
    def in_flight(self) -> int:
        with self._lock:
            return self._in_flight

    def idle_seconds(self, now: float | None = None) -> float | None:
        """Seconds since the last call started or finished; None while a call is in flight."""
        with self._lock:
            return self._idle_locked(now)

    def _idle_locked(self, now: float | None) -> float | None:
        if self._in_flight:
            return None
        return max(0.0, (self._clock() if now is None else now) - self._last_activity)


# --------------------------------------------------------------------------- git


def find_git_root(start: str | os.PathLike[str]) -> Path | None:
    """The nearest directory at or above ``start`` that holds ``.git`` (a directory, or a file for a worktree)."""
    p = Path(start).resolve()
    for d in (p, *p.parents):
        if (d / ".git").exists():
            return d
    return None


class GitError(Exception):
    pass


class GitMissing(GitError):
    pass


def run_git(
    repo: Path,
    args: Sequence[str],
    *,
    git_command: Sequence[str] = ("git",),
    timeout: float = GIT_TIMEOUT_SECONDS,
) -> str:
    env = dict(os.environ)
    env["GIT_OPTIONAL_LOCKS"] = "0"
    env["GIT_TERMINAL_PROMPT"] = "0"
    env.pop("GIT_INDEX_FILE", None)
    env.pop("GIT_DIR", None)
    env.pop("GIT_WORK_TREE", None)
    cmd = [*git_command, "--no-optional-locks", *args]
    try:
        proc = subprocess.run(
            cmd,
            cwd=str(repo),
            env=env,
            stdin=subprocess.DEVNULL,
            capture_output=True,
            timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except FileNotFoundError as exc:
        raise GitMissing(f"git not found: {exc}") from exc
    except subprocess.TimeoutExpired as exc:
        raise GitError(f"git {' '.join(args[:2])} timed out after {timeout:g} s") from exc
    except OSError as exc:
        raise GitError(f"git {' '.join(args[:2])} failed: {exc}") from exc
    if proc.returncode != 0:
        err = proc.stderr.decode("utf-8", "replace").strip().splitlines()
        raise GitError(f"git {' '.join(args[:2])} exited {proc.returncode}: {err[-1] if err else ''}")
    return proc.stdout.decode("utf-8", "replace")


def _lines(text: str) -> list[str]:
    return [line for line in text.splitlines() if line.strip()]


# --------------------------------------------------------------------------- import closure


def _package_chain(module_file: Path) -> tuple[Path, list[str]]:
    """(import root, dotted name parts) of a module file inside a regular package tree."""
    module_file = module_file.resolve()
    parts = [] if module_file.name == "__init__.py" else [module_file.stem]
    d = module_file.parent
    while (d / "__init__.py").exists():
        parts.insert(0, d.name)
        d = d.parent
    return d, parts


def _module_file(root: Path, parts: Sequence[str]) -> Path | None:
    if not parts:
        return None
    base = root.joinpath(*parts)
    if (base / "__init__.py").is_file():
        return base / "__init__.py"
    if base.with_suffix(".py").is_file():
        return base.with_suffix(".py")
    return None


def import_closure(entry: str | os.PathLike[str]) -> set[Path]:
    """Every source file of the entry module's own top-level package that running the entry can import.

    Static and conservative: every ``import``/``from ... import`` anywhere in a file (function bodies included)
    counts, as do the ``__init__.py`` files of each package on the way. Imports of other packages are ignored.
    """
    entry_path = Path(entry).resolve()
    root, entry_parts = _package_chain(entry_path)
    top = entry_parts[0] if entry_parts else None
    seen: set[Path] = set()
    todo: list[Path] = [entry_path]

    def add_module(parts: Sequence[str]) -> bool:
        if not parts or parts[0] != top:
            return False
        found = False
        for i in range(1, len(parts) + 1):  # the packages on the way, then the module itself
            f = _module_file(root, parts[:i])
            if f is not None:
                todo.append(f)
                found = True
        return found

    while todo:
        f = todo.pop().resolve()
        if f in seen or not f.is_file():
            continue
        seen.add(f)
        _, parts = _package_chain(f)
        package = parts if f.name == "__init__.py" else parts[:-1]
        add_module(parts)  # the packages it sits in (their __init__.py run first)
        try:
            tree = ast.parse(f.read_text(encoding="utf-8"), filename=str(f))
        except (OSError, SyntaxError, ValueError):
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    add_module(alias.name.split("."))
            elif isinstance(node, ast.ImportFrom):
                if node.level:
                    keep = len(package) - (node.level - 1)
                    if keep <= 0:
                        continue
                    base = list(package[:keep])
                else:
                    base = []
                mod = base + (node.module.split(".") if node.module else [])
                add_module(mod)
                for alias in node.names:  # `from pkg import submodule`
                    if alias.name != "*":
                        f2 = _module_file(root, mod + [alias.name])
                        if f2 is not None:
                            todo.append(f2)
    return seen


# --------------------------------------------------------------------------- watcher


@dataclass
class Decision:
    stale: bool
    reason: str
    start: str | None = None
    head: str | None = None
    files: list[str] = field(default_factory=list)
    idle_s: float | None = None


def log_line(message: str) -> None:
    """One timestamped, prefixed line to stderr (the default ``log`` of the watcher)."""
    stamp = _dt.datetime.now().astimezone().isoformat(timespec="seconds")
    try:
        print(f"{stamp} {LOG_PREFIX} {message}", file=sys.stderr, flush=True)
    except Exception:
        pass


def _default_exit(code: int) -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.flush()
        except Exception:
            pass
    os._exit(code)


class StaleExitWatcher:
    """Checks for new runtime commits and exits the process once idle. See the module docstring."""

    def __init__(
        self,
        package_file: str | os.PathLike[str],
        *,
        tracker: ActivityTracker,
        state_dir: str | os.PathLike[str] | Callable[[], str | os.PathLike[str]],
        idle_seconds: float = DEFAULT_IDLE_SECONDS,
        interval_seconds: float = DEFAULT_INTERVAL_SECONDS,
        idle_log: str | os.PathLike[str] | None = None,
        before_exit: Callable[[Path, list[str]], bool] | None = None,
        branch: str = DEFAULT_BRANCH,
        runtime_paths: Iterable[str] = RUNTIME_PATHS,
        git_command: Sequence[str] = ("git",),
        git_timeout: float = GIT_TIMEOUT_SECONDS,
        clock: Callable[[], float] = time.time,
        log: Callable[[str], None] = log_line,
        exit_fn: Callable[[int], None] = _default_exit,
    ) -> None:
        self.tracker = tracker
        self.state_dir = state_dir
        self.idle_seconds = float(idle_seconds)
        self.interval_seconds = float(interval_seconds)
        self.idle_log = Path(idle_log) if idle_log else None
        self.before_exit = before_exit
        self.branch = branch
        self.runtime_paths = tuple(runtime_paths)
        self.git_command = tuple(git_command)
        self.git_timeout = git_timeout
        self.clock = clock
        self.log = log
        self.exit_fn = exit_fn
        self.repo = find_git_root(Path(package_file).resolve().parent)
        self.enabled = self.repo is not None
        self.start_commit: str | None = None
        self.start_unknown = False  # the start commit could not be read: the loaded code is unknown
        self._made_at = time.time()  # wall clock, comparable with the ``at`` of restarts.jsonl
        self._said: set[tuple[str, str | None]] = set()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        if self.repo is None:
            self.log(f"off: {Path(package_file).resolve().parent} is not inside a git repository")
            return
        self._read_start()

    # -- helpers

    def _git(self, *args: str) -> str:
        assert self.repo is not None
        return run_git(self.repo, args, git_command=self.git_command, timeout=self.git_timeout)

    def _say_once(self, key: str, head: str | None, message: str) -> None:
        if (key, head) not in self._said:
            self._said.add((key, head))
            self.log(message)

    def _branch_commit(self) -> str:
        return self._git("rev-parse", "--verify", "--quiet", f"refs/heads/{self.branch}^{{commit}}").strip()

    def _read_start(self) -> None:
        try:
            self.start_commit = self._branch_commit()
        except GitMissing as exc:
            self.enabled = False
            self.log(f"off: {exc}")
            return
        except GitError as exc:
            self.start_unknown = True
            self.log(
                f"start commit not read ({exc}); the loaded code is unknown, so once git answers the server "
                f"restarts once (when clean and idle) to load a known {self.branch} commit"
            )
            return
        self.log(
            f"watching {self.repo} {self.branch} at {self.start_commit[:12]}; exits {EXIT_CODE} once idle "
            f"{self.idle_seconds:g} s after a commit that changes {', '.join(self.runtime_paths)}"
        )

    def _restarted_into(self, head: str) -> bool:
        """True when the last restart was recorded before this process started and loaded ``head``."""
        try:
            lines = (self._state_dir() / RESTARTS_FILE).read_text(encoding="utf-8").splitlines()
            last = json.loads(next(line for line in reversed(lines) if line.strip()))
            return last.get("to") == head and _dt.datetime.fromisoformat(last["at"]).timestamp() <= self._made_at
        except Exception:  # no file, an empty or a broken line: nothing known
            return False

    def _changed(self, start: str | None, head: str) -> list[str]:
        """Runtime files that differ between ``start`` and ``head``; every runtime file at ``head`` when the start
        commit is unknown."""
        if start is None:
            return _lines(self._git("ls-tree", "-r", "--name-only", head, "--", *self.runtime_paths))
        return _lines(self._git("diff", "--name-only", "--no-renames", start, head, "--", *self.runtime_paths))

    def _dirty(self, head: str) -> list[str]:
        """Runtime files in the working tree that differ from ``head``, plus untracked ones under the runtime paths."""
        dirty = _lines(self._git("diff", "--name-only", "--no-renames", head, "--", *self.runtime_paths))
        dirty += _lines(self._git("ls-files", "--others", "--exclude-standard", "--", *self.runtime_paths))
        return dirty

    def _say_dirty(self, head: str, dirty: list[str]) -> None:
        shown = ", ".join(sorted(set(dirty))[:5])
        self._say_once(
            "dirty", head,
            f"runtime files differ from {head[:12]} in the working tree ({shown}); waiting until they are committed",
        )

    def _state_dir(self) -> Path:
        d = self.state_dir() if callable(self.state_dir) else self.state_dir
        return Path(d)

    # -- the check

    def check(self) -> Decision:
        """One round: is the process stale and idle? Never raises."""
        if not self.enabled:
            return Decision(False, "off")
        start = self.start_commit  # None only when start_unknown: then every runtime file counts as changed
        try:
            head = self._branch_commit()
            if start is None and self._restarted_into(head):
                self.start_commit = start = head
                self.start_unknown = False
                self.log(
                    f"watching {self.repo} {self.branch} from {head[:12]}: the start commit was not read, but the "
                    f"last restart before this process started loaded it and {self.branch} has not moved since"
                )
            if head == start:
                return Decision(False, "no new commit", start, head)
            changed = self._changed(start, head)
            if not changed:
                self._say_once("docs", head, f"{self.branch} moved to {head[:12]}; no runtime file changed, staying")
                return Decision(False, "no runtime file changed", start, head)
            dirty = self._dirty(head)
        except GitError as exc:
            self._say_once("git", str(exc), f"git check skipped this round: {exc}")
            return Decision(False, f"git error: {exc}", start)
        if start is None:
            self._say_once("unknown", head, f"start commit was never read; restarting once to load {head[:12]}")
        else:
            self._say_once(
                "new", head, f"{self.branch} moved {start[:12]}..{head[:12]} ({len(changed)} runtime files changed)"
            )
        if dirty:
            self._say_dirty(head, dirty)
            return Decision(False, "runtime files not committed", start, head, changed)
        now = self.clock()
        idle = self.tracker.idle_seconds(now)
        if idle is None:
            return Decision(False, f"{self.tracker.in_flight} call(s) in flight", start, head, changed)
        if self.idle_log is not None:
            try:
                idle = min(idle, max(0.0, now - self.idle_log.stat().st_mtime))
            except OSError:
                pass  # a missing log says nothing about activity
        if idle < self.idle_seconds:
            return Decision(False, f"idle {idle:.0f} s of {self.idle_seconds:g} s", start, head, changed, idle)
        return Decision(True, "stale and idle", start, head, changed, idle)

    def check_and_exit(self) -> Decision:
        decision = self.check()
        if not decision.stale:
            return decision
        if self.before_exit is not None:
            try:
                ok = self.before_exit(self.repo, list(decision.files))
            except Exception as exc:  # a failing hook must not take the server down
                ok = False
                self._say_once("hook", decision.head, f"before-exit hook failed ({exc!r}); retrying next check")
            if not ok:
                return Decision(False, "before-exit hook declined", decision.start, decision.head, decision.files)
        deferred = self._recheck(decision)
        if deferred is not None:
            return deferred
        # Last look under the tracker's lock: a call that started while the hook ran keeps the server up, and
        # no call can start between this look and the exit.
        with self.tracker._lock:
            idle = self.tracker._idle_locked(self.clock())
            if idle is None or idle < self.idle_seconds:
                return Decision(False, "a call arrived during the before-exit hook", decision.start, decision.head,
                                decision.files)
            record = {
                "at": _dt.datetime.now().astimezone().isoformat(timespec="seconds"),
                "from": decision.start,
                "to": decision.head,
                "files": decision.files,
                "idle_s": round(decision.idle_s or 0.0, 1),
            }
            try:
                d = self._state_dir()
                d.mkdir(parents=True, exist_ok=True)
                with open(d / RESTARTS_FILE, "a", encoding="utf-8") as fh:
                    fh.write(json.dumps(record) + "\n")
            except Exception as exc:
                self.log(f"could not record the restart in {RESTARTS_FILE}: {exc}")
            self.log(
                f"exiting {EXIT_CODE} to load {self.branch} {decision.head[:12] if decision.head else '?'} "
                f"(from {decision.start[:12] if decision.start else 'an unknown start'}; "
                f"{len(decision.files)} runtime files; "
                f"idle {record['idle_s']:g} s)"
            )
            self.exit_fn(EXIT_CODE)
        return decision

    def _recheck(self, decision: Decision) -> Decision | None:
        """Git once more, after the hook and outside the tracker's lock (git may take up to its timeout, and
        holding the lock that long would stall every call): None when the tree is still the one ``decision``
        checked, else the decision to stay up this round."""
        assert decision.head is not None
        try:
            head = self._branch_commit()
            dirty = self._dirty(head) if head == decision.head else []
        except GitError as exc:
            self._say_once("recheck-git", str(exc), f"git check before exiting failed ({exc}); staying up this round")
            return Decision(False, f"git error before exiting: {exc}", decision.start, decision.head, decision.files)
        if head != decision.head:
            self._say_once(
                "moved", head,
                f"{self.branch} moved to {head[:12]} while getting ready to exit; checking it at the next round",
            )
            return Decision(False, f"branch moved to {head[:12]} before exiting", decision.start, decision.head,
                            decision.files)
        if dirty:
            self._say_dirty(head, dirty)
            return Decision(False, "runtime files not committed", decision.start, head, decision.files)
        return None

    # -- the thread

    def _run(self) -> None:
        while not self._stop.wait(self.interval_seconds):
            try:
                self.check_and_exit()
            except Exception as exc:  # never let the watcher die silently
                self._say_once("loop", repr(exc), f"check failed: {exc!r}")

    def start(self) -> StaleExitWatcher:
        if self.enabled and self._thread is None:
            self._thread = threading.Thread(target=self._run, name="stale-exit", daemon=True)
            self._thread.start()
        return self

    def stop(self) -> None:
        self._stop.set()
