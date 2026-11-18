"""--exit-when-stale (stale_exit.py): a server exits, once idle, when its checkout gets new runtime commits.

Every repository here is a throwaway one made in the test's tmp dir, never this checkout.
"""

from __future__ import annotations

import argparse
import asyncio
import datetime
import json
import os
from pathlib import Path
import subprocess
import sys
import textwrap
import time

import pytest

from research_mcp import stale_exit
from research_mcp.stale_exit import (
    DEFAULT_IDLE_SECONDS,
    EXIT_CODE,
    RESTARTS_FILE,
    RUNTIME_PATHS,
    ActivityTracker,
    StaleExitWatcher,
    import_closure,
)

IDLE = 300.0


def git(repo: Path, *args: str, env: dict | None = None) -> str:
    full_env = {**os.environ, **(env or {})}
    return subprocess.run(
        ["git", "-c", "user.name=Test", "-c", "user.email=test@example.invalid", "-c", "core.autocrlf=false", *args],
        cwd=repo, check=True, capture_output=True, text=True, env=full_env,
    ).stdout


def write(repo: Path, rel: str, text: str) -> None:
    p = repo / rel
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8")


def commit(repo: Path, message: str, *rels: str) -> str:
    git(repo, "add", "--", *rels)
    git(repo, "commit", "-q", "-m", message)
    return git(repo, "rev-parse", "HEAD").strip()


@pytest.fixture
def repo(tmp_path: Path) -> Path:
    r = tmp_path / "checkout"
    r.mkdir()
    git(r, "init", "-q", "-b", "master")
    write(r, "src/pkg/__init__.py", "")
    write(r, "src/pkg/server.py", "VALUE = 1\n")
    write(r, "pyproject.toml", "[project]\nname = 'pkg'\n")
    write(r, "docs/notes.md", "notes\n")
    write(r, "README.md", "readme\n")
    commit(r, "initial", ".")
    return r


class Clock:
    def __init__(self) -> None:
        self.t = time.time()

    def __call__(self) -> float:
        return self.t


class Harness:
    def __init__(self, repo: Path, tmp_path: Path, **kw) -> None:
        self.clock = Clock()
        self.tracker = ActivityTracker(self.clock)
        self.logs: list[str] = []
        self.exits: list[int] = []
        self.state = tmp_path / "state"
        self.watcher = StaleExitWatcher(
            repo / "src" / "pkg" / "server.py",
            tracker=self.tracker,
            state_dir=self.state,
            idle_seconds=IDLE,
            clock=self.clock,
            log=self.logs.append,
            exit_fn=self.exits.append,
            **kw,
        )

    def go_idle(self, seconds: float = IDLE + 1) -> None:
        self.clock.t += seconds

    def restarts(self) -> list[dict]:
        f = self.state / RESTARTS_FILE
        if not f.exists():
            return []
        return [json.loads(line) for line in f.read_text(encoding="utf-8").splitlines()]


@pytest.fixture
def h(repo: Path, tmp_path: Path) -> Harness:
    return Harness(repo, tmp_path)


# ---------------------------------------------------------------- (a) new runtime commits


def test_src_commit_exits_75_and_records_the_restart(repo, h):
    start = h.watcher.start_commit
    write(repo, "src/pkg/server.py", "VALUE = 2\n")
    head = commit(repo, "fix", "src/pkg/server.py")
    h.go_idle()

    decision = h.watcher.check_and_exit()

    assert decision.stale
    assert h.exits == [EXIT_CODE] == [75]
    (record,) = h.restarts()
    assert record["from"] == start and record["to"] == head
    assert record["files"] == ["src/pkg/server.py"]
    assert record["idle_s"] >= IDLE
    assert set(record) == {"at", "from", "to", "files", "idle_s"}
    assert any("exiting 75" in line for line in h.logs)


def test_pyproject_and_lock_commits_count_as_runtime(repo, h):
    write(repo, "uv.lock", "lock\n")
    commit(repo, "lock", "uv.lock")
    h.go_idle()
    assert h.watcher.check_and_exit().stale
    assert h.exits == [EXIT_CODE]
    assert h.restarts()[0]["files"] == ["uv.lock"]


def test_docs_only_commit_keeps_running(repo, h):
    write(repo, "docs/notes.md", "more notes\n")
    write(repo, "README.md", "new readme\n")
    commit(repo, "docs", "docs/notes.md", "README.md")
    h.go_idle()

    decision = h.watcher.check_and_exit()

    assert not decision.stale and decision.reason == "no runtime file changed"
    assert h.exits == [] and h.restarts() == []
    h.watcher.check_and_exit()
    assert sum("no runtime file changed" in line for line in h.logs) == 1  # once per head


def test_docs_commit_then_src_commit_lists_only_runtime_files(repo, h):
    write(repo, "docs/notes.md", "x\n")
    commit(repo, "docs", "docs/notes.md")
    write(repo, "src/pkg/new_mod.py", "X = 1\n")
    commit(repo, "feature", "src/pkg/new_mod.py")
    h.go_idle()
    assert h.watcher.check_and_exit().stale
    assert h.restarts()[0]["files"] == ["src/pkg/new_mod.py"]


def test_no_new_commit_never_exits(repo, h):
    h.go_idle(10 * IDLE)
    assert h.watcher.check_and_exit().reason == "no new commit"
    assert h.exits == []


def test_commit_on_another_branch_is_ignored(repo, h):
    git(repo, "checkout", "-q", "-b", "side")
    write(repo, "src/pkg/server.py", "VALUE = 3\n")
    commit(repo, "side work", "src/pkg/server.py")
    git(repo, "checkout", "-q", "master")
    h.go_idle()
    assert not h.watcher.check_and_exit().stale
    assert h.exits == []


# ---------------------------------------------------------------- (b) uncommitted runtime files


def test_dirty_src_file_keeps_running_until_committed(repo, h):
    write(repo, "src/pkg/server.py", "VALUE = 2\n")
    commit(repo, "fix", "src/pkg/server.py")
    write(repo, "src/pkg/server.py", "VALUE = 'work in progress'\n")  # another stage's uncommitted edit
    h.go_idle()

    for _ in range(3):
        decision = h.watcher.check_and_exit()
        assert not decision.stale and decision.reason == "runtime files not committed"
    assert h.exits == []
    assert sum("waiting until they are committed" in line for line in h.logs) == 1  # once per head

    commit(repo, "wip done", "src/pkg/server.py")
    assert h.watcher.check_and_exit().stale
    assert h.exits == [EXIT_CODE]


def test_untracked_src_file_keeps_running(repo, h):
    write(repo, "src/pkg/server.py", "VALUE = 2\n")
    commit(repo, "fix", "src/pkg/server.py")
    write(repo, "src/pkg/half_written.py", "X = 1\n")
    h.go_idle()
    assert h.watcher.check_and_exit().reason == "runtime files not committed"
    assert h.exits == []


def test_dirty_docs_file_does_not_block(repo, h):
    write(repo, "src/pkg/server.py", "VALUE = 2\n")
    commit(repo, "fix", "src/pkg/server.py")
    write(repo, "docs/notes.md", "editing notes\n")
    h.go_idle()
    assert h.watcher.check_and_exit().stale


def test_commit_from_private_index_is_not_mistaken_for_dirty(repo, h, tmp_path):
    """A commit made with its own GIT_INDEX_FILE leaves the shared index behind, so `git status` shows the
    committed file as changed although the working tree equals the commit."""
    write(repo, "src/pkg/server.py", "VALUE = 2\n")
    base = git(repo, "rev-parse", "HEAD").strip()
    env = {"GIT_INDEX_FILE": str(tmp_path / "private-index")}
    git(repo, "read-tree", base, env=env)
    git(repo, "add", "--", "src/pkg/server.py", env=env)
    tree = git(repo, "write-tree", env=env).strip()
    new = git(repo, "commit-tree", tree, "-p", base, "-m", "fix", env=env).strip()
    git(repo, "update-ref", "refs/heads/master", new, base)
    assert git(repo, "status", "--porcelain", "--", "src").strip()  # the shared index is behind
    h.go_idle()

    assert h.watcher.check_and_exit().stale
    assert h.restarts()[0]["to"] == new


# ---------------------------------------------------------------- (c) idle


def src_commit(repo: Path) -> str:
    """A runtime commit landing after the watcher recorded its start commit."""
    write(repo, "src/pkg/server.py", "VALUE = 2\n")
    return commit(repo, "fix", "src/pkg/server.py")


def test_call_in_flight_past_the_window_waits(repo, h):
    src_commit(repo)
    h.tracker.start()
    h.go_idle(5 * IDLE)  # the call has run far longer than the window

    decision = h.watcher.check_and_exit()
    assert not decision.stale and "in flight" in decision.reason
    assert h.exits == []

    h.tracker.finish()
    h.go_idle(IDLE - 5)  # finishing counts as activity
    assert not h.watcher.check_and_exit().stale
    h.go_idle(10)
    assert h.watcher.check_and_exit().stale
    assert h.exits == [EXIT_CODE]


def test_recent_call_waits_for_the_window(repo, h):
    src_commit(repo)
    h.go_idle(IDLE * 2)
    h.tracker.start()
    h.tracker.finish()
    h.go_idle(IDLE / 2)
    assert "idle" in h.watcher.check_and_exit().reason
    assert h.exits == []


def test_idle_log_touched_recently_waits(repo, tmp_path):
    log = tmp_path / "server.log"
    log.write_text("GET /mcp\n", encoding="utf-8")
    hh = Harness(repo, tmp_path, idle_log=log)
    src_commit(repo)
    hh.go_idle(IDLE * 2)  # no tool call for a long time ...
    os.utime(log, (hh.clock.t - 30, hh.clock.t - 30))  # ... but the log moved 30 s ago

    assert not hh.watcher.check_and_exit().stale
    assert hh.exits == []

    os.utime(log, (hh.clock.t - IDLE - 1, hh.clock.t - IDLE - 1))
    assert hh.watcher.check_and_exit().stale
    assert hh.exits == [EXIT_CODE]


def test_missing_idle_log_does_not_block(repo, tmp_path):
    hh = Harness(repo, tmp_path, idle_log=tmp_path / "absent.log")
    src_commit(repo)
    hh.go_idle()
    assert hh.watcher.check_and_exit().stale


def test_call_arriving_during_the_hook_keeps_the_server_up(repo, tmp_path):
    holder: dict = {}

    def hook(repo_root, files):
        holder["h"].tracker.start()  # a call arrives while the hook runs
        return True

    hh = Harness(repo, tmp_path, before_exit=hook)
    src_commit(repo)
    holder["h"] = hh
    hh.go_idle()
    assert not hh.watcher.check_and_exit().stale
    assert hh.exits == [] and hh.restarts() == []


# ---------------------------------------------------------------- git read again right before exiting


def test_commit_landing_during_the_hook_defers_to_the_next_round(repo, tmp_path):
    """Input: a runtime commit is made while the before-exit hook runs.
    Expected: no exit this round; the next round exits to the newer commit (never to the one checked earlier)."""
    landed: list[str] = []

    def hook(repo_root, files):
        if not landed:
            write(repo, "src/pkg/server.py", "VALUE = 3\n")
            landed.append(commit(repo, "second fix", "src/pkg/server.py"))
        return True

    hh = Harness(repo, tmp_path, before_exit=hook)
    first = src_commit(repo)
    hh.go_idle()

    decision = hh.watcher.check_and_exit()
    assert not decision.stale and decision.reason == f"branch moved to {landed[0][:12]} before exiting"
    assert hh.exits == [] and hh.restarts() == []
    assert decision.head == first != landed[0]  # the deferral describes the tree it checked
    assert sum("while getting ready to exit" in line for line in hh.logs) == 1

    assert hh.watcher.check_and_exit().stale
    assert hh.exits == [EXIT_CODE]
    assert hh.restarts()[0]["to"] == landed[0]


def test_uncommitted_pyproject_edit_during_the_hook_defers_until_committed(repo, tmp_path):
    """Sibling: the change arriving during the hook is uncommitted and in another runtime path than src/."""
    edited: list[bool] = []

    def hook(repo_root, files):
        if not edited:
            write(repo, "pyproject.toml", "[project]\nname = 'pkg'\nversion = '2'\n")
            edited.append(True)
        return True

    hh = Harness(repo, tmp_path, before_exit=hook)
    src_commit(repo)
    hh.go_idle()

    decision = hh.watcher.check_and_exit()
    assert not decision.stale and decision.reason == "runtime files not committed"
    assert hh.exits == []
    assert any("pyproject.toml" in line and "waiting until they are committed" in line for line in hh.logs)
    assert hh.watcher.check_and_exit().reason == "runtime files not committed"  # still dirty: still up

    head = commit(repo, "version", "pyproject.toml")
    assert hh.watcher.check_and_exit().stale
    assert hh.restarts()[0]["to"] == head
    assert sorted(hh.restarts()[0]["files"]) == ["pyproject.toml", "src/pkg/server.py"]


def test_untracked_src_file_appearing_during_the_hook_defers(repo, tmp_path):
    def hook(repo_root, files):
        write(repo, "src/pkg/new_module.py", "X = 1\n")
        return True

    hh = Harness(repo, tmp_path, before_exit=hook)
    src_commit(repo)
    hh.go_idle()
    assert hh.watcher.check_and_exit().reason == "runtime files not committed"
    assert hh.exits == [] and hh.restarts() == []


def test_docs_edit_during_the_hook_does_not_defer(repo, tmp_path):
    def hook(repo_root, files):
        write(repo, "docs/notes.md", "edited while the hook ran\n")
        return True

    hh = Harness(repo, tmp_path, before_exit=hook)
    src_commit(repo)
    hh.go_idle()
    assert hh.watcher.check_and_exit().stale
    assert hh.exits == [EXIT_CODE]


def test_git_timing_out_after_the_hook_keeps_the_server_up(repo, tmp_path):
    holder: dict = {}

    def hook(repo_root, files):
        w = holder["h"].watcher
        w.git_command = (sys.executable, "-c", "import time; time.sleep(30)")
        w.git_timeout = 0.5
        return True

    hh = Harness(repo, tmp_path, before_exit=hook)
    holder["h"] = hh
    src_commit(repo)
    hh.go_idle()

    decision = hh.watcher.check_and_exit()
    assert not decision.stale and "timed out" in decision.reason
    assert hh.exits == [] and hh.restarts() == []

    hh.watcher.before_exit = None
    hh.watcher.git_command = ("git",)
    hh.watcher.git_timeout = stale_exit.GIT_TIMEOUT_SECONDS
    assert hh.watcher.check_and_exit().stale  # the next round decides again
    assert hh.exits == [EXIT_CODE]


# ---------------------------------------------------------------- before_exit hook


def test_before_exit_gets_the_changed_files_and_can_defer(repo, tmp_path):
    calls: list[tuple[Path, list[str]]] = []
    answers = [False, True]

    def hook(repo_root, files):
        calls.append((repo_root, files))
        return answers.pop(0)

    hh = Harness(repo, tmp_path, before_exit=hook)
    src_commit(repo)
    hh.go_idle()
    assert hh.watcher.check_and_exit().reason == "before-exit hook declined"
    assert hh.exits == []
    hh.go_idle(60)
    assert hh.watcher.check_and_exit().stale  # retried at the next check
    assert hh.exits == [EXIT_CODE]
    assert calls == [(repo.resolve(), ["src/pkg/server.py"])] * 2


def test_before_exit_raising_keeps_the_server_up(repo, tmp_path):
    def hook(repo_root, files):
        raise RuntimeError("boom")

    hh = Harness(repo, tmp_path, before_exit=hook)
    src_commit(repo)
    hh.go_idle()
    assert not hh.watcher.check_and_exit().stale
    assert hh.exits == []
    assert any("boom" in line for line in hh.logs)


# ---------------------------------------------------------------- off / errors


def test_no_git_repo_turns_the_feature_off(tmp_path):
    pkg = tmp_path / "plain" / "src" / "pkg"
    pkg.mkdir(parents=True)
    (pkg / "server.py").write_text("", encoding="utf-8")
    logs: list[str] = []
    exits: list[int] = []
    w = StaleExitWatcher(pkg / "server.py", tracker=ActivityTracker(), state_dir=tmp_path,
                         log=logs.append, exit_fn=exits.append, clock=lambda: 1e12)
    assert not w.enabled
    assert len(logs) == 1 and logs[0].startswith("off:")
    assert w.check_and_exit().reason == "off"
    assert w.start()._thread is None
    assert exits == []


def test_git_missing_turns_the_feature_off(repo, tmp_path):
    hh = Harness(repo, tmp_path, git_command=(str(tmp_path / "no-such-git.exe"),))
    src_commit(repo)
    assert not hh.watcher.enabled
    assert any(line.startswith("off: git not found") for line in hh.logs)


def test_git_timeout_is_not_stale_this_round(repo, tmp_path):
    hh = Harness(repo, tmp_path)
    src_commit(repo)
    hh.go_idle()
    hh.watcher.git_command = (sys.executable, "-c", "import time; time.sleep(30)")
    hh.watcher.git_timeout = 0.5

    decision = hh.watcher.check_and_exit()

    assert not decision.stale and "timed out" in decision.reason
    assert hh.exits == []
    hh.watcher.git_command = ("git",)
    hh.watcher.git_timeout = stale_exit.GIT_TIMEOUT_SECONDS
    assert hh.watcher.check_and_exit().stale  # the next round checks again


SLOW_GIT = (sys.executable, "-c", "import time; time.sleep(30)")


def slow_start_harness(repo: Path, tmp_path: Path, name: str = "first") -> Harness:
    """A server whose start commit could not be read: git timed out when the watcher was made."""
    hh = Harness(repo, tmp_path / name, git_command=SLOW_GIT, git_timeout=0.5)
    assert hh.watcher.enabled and hh.watcher.start_commit is None and hh.watcher.start_unknown
    hh.watcher.git_command = ("git",)
    hh.watcher.git_timeout = stale_exit.GIT_TIMEOUT_SECONDS
    return hh


def test_commit_while_the_start_was_unread_still_restarts(repo, tmp_path):
    """Input: the start read times out, a src commit lands, git recovers.
    Expected: exit; the record has from = null and lists the runtime files (the commit is not adopted as start)."""
    hh = slow_start_harness(repo, tmp_path)
    head = src_commit(repo)
    hh.go_idle()

    decision = hh.watcher.check_and_exit()

    assert decision.stale and hh.exits == [EXIT_CODE]
    (record,) = hh.restarts()
    assert record["from"] is None and record["to"] == head
    assert sorted(record["files"]) == ["pyproject.toml", "src/pkg/__init__.py", "src/pkg/server.py"]
    assert any(f"start commit was never read; restarting once to load {head[:12]}" in line for line in hh.logs)
    assert any("an unknown start" in line for line in hh.logs)


def test_unread_start_restarts_once_even_without_a_new_commit_then_stays(repo, tmp_path):
    """Sibling: nothing lands while the start is unread. The loaded code is still unknown, so it restarts once;
    the new process (start read fine) at the same commit never restarts again."""
    hh = slow_start_harness(repo, tmp_path)
    head = git(repo, "rev-parse", "HEAD").strip()
    hh.go_idle()
    assert hh.watcher.check_and_exit().stale
    assert hh.restarts()[0]["to"] == head

    after = Harness(repo, tmp_path / "second")
    assert after.watcher.start_commit == head
    after.go_idle()
    for _ in range(3):
        assert after.watcher.check_and_exit().reason == "no new commit"
    assert after.exits == []


def test_git_slow_at_every_boot_restarts_only_once(repo, tmp_path):
    """Input: the start read times out at two boots in a row, nothing new is committed.
    Expected: the first process restarts once; the second, started after that restart's record, takes the
    recorded commit as its start and stays up."""
    state = tmp_path / "shared_state"
    first = slow_start_harness(repo, tmp_path, "first")
    first.state = first.watcher.state_dir = state
    head = git(repo, "rev-parse", "HEAD").strip()
    first.go_idle()
    assert first.watcher.check_and_exit().stale

    second = slow_start_harness(repo, tmp_path, "second")
    second.state = second.watcher.state_dir = state
    second.go_idle()
    for _ in range(3):
        assert second.watcher.check_and_exit().reason == "no new commit"
    assert second.exits == []
    assert second.watcher.start_commit == head and not second.watcher.start_unknown
    assert len(second.restarts()) == 1

    head2 = src_commit(repo)  # and it still notices the next commit
    assert second.watcher.check_and_exit().stale
    assert second.restarts()[-1]["from"] == head and second.restarts()[-1]["to"] == head2


def test_unread_start_ignores_a_restart_record_it_cannot_rely_on(repo, tmp_path):
    """Sibling: the last record names an older commit, or was written after this process started (another
    process sharing the state dir): the loaded code stays unknown and the process restarts once."""
    older = git(repo, "rev-parse", "HEAD").strip()
    hh = slow_start_harness(repo, tmp_path)
    hh.state.mkdir(parents=True)
    old_at = "2020-01-01T00:00:00+00:00"
    (hh.state / RESTARTS_FILE).write_text(json.dumps({"at": old_at, "from": None, "to": older}) + "\n",
                                          encoding="utf-8")
    head = src_commit(repo)
    hh.go_idle()
    assert hh.watcher.check_and_exit().stale and hh.restarts()[-1]["to"] == head

    later = slow_start_harness(repo, tmp_path, "later")
    later.state.mkdir(parents=True)
    after_start = datetime.datetime.fromtimestamp(later.watcher._made_at + 5).astimezone().isoformat(timespec="seconds")
    (later.state / RESTARTS_FILE).write_text(json.dumps({"at": after_start, "from": older, "to": head}) + "\n",
                                             encoding="utf-8")
    later.go_idle()
    assert later.watcher.check_and_exit().stale
    assert later.restarts()[-1]["from"] is None


def test_unread_start_still_waits_for_uncommitted_runtime_files(repo, tmp_path):
    hh = slow_start_harness(repo, tmp_path)
    write(repo, "src/pkg/server.py", "VALUE = 'work in progress'\n")
    hh.go_idle()
    assert hh.watcher.check_and_exit().reason == "runtime files not committed"
    assert hh.exits == []


def test_git_error_is_not_stale(repo, tmp_path):
    hh = Harness(repo, tmp_path, branch="no-such-branch")
    src_commit(repo)
    hh.go_idle()
    assert hh.watcher.start_commit is None
    assert not hh.watcher.check_and_exit().stale
    assert hh.exits == []


# ---------------------------------------------------------------- the real process


def test_process_really_exits_75_with_other_threads_running(repo, tmp_path):
    """Like uvicorn, a non-daemon thread keeps the process alive; the watcher still ends it with 75."""
    state = tmp_path / "proc_state"
    script = tmp_path / "fake_server.py"
    script.write_text(textwrap.dedent(f"""
        import importlib.util, sys, threading, time
        spec = importlib.util.spec_from_file_location("stale_exit", {str(stale_exit.__file__)!r})
        mod = importlib.util.module_from_spec(spec); sys.modules["stale_exit"] = mod; spec.loader.exec_module(mod)
        threading.Thread(target=lambda: time.sleep(120), daemon=False).start()  # the web server
        w = mod.StaleExitWatcher({str(repo / "src" / "pkg" / "server.py")!r}, tracker=mod.ActivityTracker(),
                                 state_dir={str(state)!r}, idle_seconds=0, interval_seconds=0.2).start()
        print("ready", w.start_commit, flush=True)
        time.sleep(120)
    """), encoding="utf-8")
    proc = subprocess.Popen([sys.executable, str(script)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    try:
        first = proc.stdout.readline()
        if not first.startswith("ready "):
            proc.wait(timeout=30)
            pytest.fail(f"fake server did not start: {first!r} {proc.stderr.read()}")
        write(repo, "src/pkg/server.py", "VALUE = 2\n")
        head = commit(repo, "fix", "src/pkg/server.py")
        code = proc.wait(timeout=60)
    finally:
        if proc.poll() is None:
            proc.kill()
    stderr = proc.stderr.read()
    assert code == EXIT_CODE, stderr
    assert "exiting 75" in stderr
    (line,) = (state / RESTARTS_FILE).read_text(encoding="utf-8").splitlines()
    assert json.loads(line)["to"] == head


# ---------------------------------------------------------------- import closure


def test_import_closure_follows_package_imports_only(tmp_path):
    root = tmp_path / "src"
    write(root, "pkg/__init__.py", "from .base import B\n")
    write(root, "pkg/base.py", "import json\nB = 1\n")
    write(root, "pkg/worker.py", "import os\nfrom .helpers import h\nfrom . import sibling\n"
                                 "def late():\n    import pkg.sub.deep\n")
    write(root, "pkg/helpers.py", "from pkg.errors import E\n")
    write(root, "pkg/errors.py", "E = 1\n")
    write(root, "pkg/sibling.py", "")
    write(root, "pkg/sub/__init__.py", "")
    write(root, "pkg/sub/deep.py", "from ..errors import E\n")
    write(root, "pkg/server.py", "from .worker import late\n")  # imports the worker, not imported by it
    write(root, "pkg/unused.py", "")

    names = {p.relative_to(root.resolve()).as_posix() for p in import_closure(root / "pkg" / "worker.py")}

    assert names == {
        "pkg/__init__.py", "pkg/base.py", "pkg/worker.py", "pkg/helpers.py", "pkg/errors.py",
        "pkg/sibling.py", "pkg/sub/__init__.py", "pkg/sub/deep.py",
    }


def test_word_worker_closure_of_this_package():
    import research_mcp

    pkg = Path(research_mcp.__file__).resolve().parent
    names = {p.name for p in import_closure(pkg / "word_worker.py")}
    assert {"word_worker.py", "word_service.py", "errors.py", "__init__.py", "audit.py"} <= names
    assert "server.py" not in names and "stale_exit.py" not in names


# ---------------------------------------------------------------- server wiring


def _args(**kw) -> argparse.Namespace:
    from research_mcp.server import build_parser

    argv = []
    for k, v in kw.items():
        flag = "--" + k.replace("_", "-")
        argv += [flag] if v is True else [flag, str(v)]
    return build_parser().parse_args(argv)


def test_server_flags_default_to_the_module_constants():
    args = _args()
    assert args.exit_when_stale is False
    assert args.stale_idle_seconds == DEFAULT_IDLE_SECONDS == 300
    assert args.idle_log is None
    assert RUNTIME_PATHS == ("src", "pyproject.toml", "uv.lock")


def test_stdio_ignores_exit_when_stale(capsys):
    from research_mcp.server import start_stale_exit

    assert start_stale_exit(_args(transport="stdio", exit_when_stale=True)) is None
    assert "ignored under stdio" in capsys.readouterr().err


def test_flag_off_starts_nothing():
    from research_mcp.server import start_stale_exit

    assert start_stale_exit(_args(transport="streamable-http")) is None


def test_http_with_flag_starts_the_watcher(monkeypatch, tmp_path):
    from research_mcp import server

    created: list[dict] = []

    class FakeWatcher:
        def __init__(self, package_file, **kw):
            created.append({"package_file": package_file, **kw})

        def start(self):
            return self

    monkeypatch.setattr(server._stale_exit, "StaleExitWatcher", FakeWatcher)
    log = tmp_path / "host.log"
    w = server.start_stale_exit(_args(transport="streamable-http", exit_when_stale=True,
                                      stale_idle_seconds=60, idle_log=log))
    assert isinstance(w, FakeWatcher)
    (kw,) = created
    assert kw["tracker"] is server.TOOL_ACTIVITY
    assert kw["idle_seconds"] == 60 and kw["idle_log"] == log
    assert Path(kw["package_file"]).name == "server.py"
    assert callable(kw["before_exit"])
    assert Path(kw["state_dir"]()) == Path(os.environ["RESEARCH_MCP_HOME"]).resolve()


def test_tool_calls_are_counted_in_flight(monkeypatch):
    from research_mcp import server

    seen: list[int] = []

    async def fake_call(name, arguments, context, convert_result=False):
        seen.append(server.TOOL_ACTIVITY.in_flight)
        return "ok"

    async def failing_call(name, arguments, context, convert_result=False):
        seen.append(server.TOOL_ACTIVITY.in_flight)
        raise ValueError("bad")

    monkeypatch.setattr(server, "log_tool_call", lambda **kw: None)
    before = server.TOOL_ACTIVITY.in_flight
    monkeypatch.setattr(server, "_orig_call_tool", fake_call)
    assert asyncio.run(server._logged_call_tool("t", {}, None)) == "ok"
    monkeypatch.setattr(server, "_orig_call_tool", failing_call)
    with pytest.raises(ValueError):
        asyncio.run(server._logged_call_tool("t", {}, None))
    assert seen == [before + 1, before + 1]
    assert server.TOOL_ACTIVITY.in_flight == before
    assert server.TOOL_ACTIVITY.idle_seconds() < 5


class _Busy(Exception):
    pass


def _worker_hook(monkeypatch, behaviour):
    from research_mcp import errors, server, word_service

    calls: list[int] = []

    def fake_shutdown():
        calls.append(1)
        if behaviour == "busy":
            raise errors.IntegrationBusy("worker busy")
        return {"ok": True}

    monkeypatch.setattr(word_service, "shutdown_word_worker", fake_shutdown)
    return server._refresh_word_worker_hook(), calls, Path(server.__file__).resolve().parents[2]


def test_word_worker_hook_shuts_the_worker_down_when_its_code_changed(monkeypatch):
    hook, calls, repo_root = _worker_hook(monkeypatch, "ok")
    assert hook(repo_root, ["src/research_mcp/word_service.py", "README.md"]) is True
    assert calls == [1]


def test_word_worker_hook_leaves_the_worker_for_other_changes(monkeypatch):
    hook, calls, repo_root = _worker_hook(monkeypatch, "ok")
    assert hook(repo_root, ["src/research_mcp/server.py", "pyproject.toml"]) is True
    assert calls == []


def test_word_worker_hook_defers_the_exit_while_the_worker_is_busy(monkeypatch):
    hook, calls, repo_root = _worker_hook(monkeypatch, "busy")
    assert hook(repo_root, ["src/research_mcp/word_worker.py"]) is False
    assert calls == [1]
