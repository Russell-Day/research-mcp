"""Settings: the process environment first, then ``%APPDATA%\\research-mcp\\config.env``.

The config file holds simple ``KEY=value`` lines (``#`` starts a comment, values may be quoted), so a
Zotero key or the comment author lives in one place instead of every MCP client's launch config.
``RESEARCH_MCP_CONFIG`` points at another file (tests, a scratch setup). The file is parsed once and cached;
``reload()`` drops the cache. The file may be UTF-8 (with or without a BOM) or UTF-16 with a BOM (what
Windows PowerShell's ``>`` writes); an unreadable file counts as missing.

Known keys: ``ZOTERO_API_KEY``, ``ZOTERO_LIBRARY_ID``, ``ZOTERO_LIBRARY_TYPE``, ``ZOTERO_MCP_CONTACT_EMAIL``,
``RESEARCH_MCP_AUTHOR``, ``RESEARCH_MCP_AUTHOR_INITIALS``, ``RESEARCH_MCP_ISSUES`` (``github``, the default, or
``agent``), ``RESEARCH_MCP_ISSUES_FILE_WITH`` (the filing steps returned in ``agent`` mode), plus the technical
``RESEARCH_MCP_*`` toggles.
"""

from __future__ import annotations

import os
import re
import threading
from pathlib import Path

_LOCK = threading.Lock()
_FILE_VALUES: dict[str, str] | None = None

DEFAULT_AUTHOR = "Author"
WORD_NAME_TIMEOUT = 3.0


def config_path() -> Path:
    """Where the config file is (it need not exist)."""

    override = os.environ.get("RESEARCH_MCP_CONFIG", "").strip()
    if override:
        return Path(override)
    appdata = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    return Path(appdata) / "research-mcp" / "config.env"


def parse(text: str) -> dict[str, str]:
    """``KEY=value`` lines: blank and ``#`` lines skipped, an optional ``export``, a quoted value is the text
    between its quotes, and an unquoted value ends at `` #`` (a comment after whitespace)."""

    values: dict[str, str] = {}
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = re.sub(r"^export\s+", "", key.strip())
        value = value.strip()
        if value and value[0] in "\"'" and value.find(value[0], 1) > 0:
            value = value[1:value.find(value[0], 1)]
        else:
            value = re.split(r"\s+#", value, maxsplit=1)[0].rstrip()
        if key:
            values[key] = value
    return values


def decode(data: bytes) -> str:
    """UTF-16 when the bytes start with its BOM, else UTF-8 (BOM optional); undecodable bytes become U+FFFD."""

    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16", errors="replace")
    return data.decode("utf-8-sig", errors="replace")


def _file_values() -> dict[str, str]:
    global _FILE_VALUES
    with _LOCK:
        if _FILE_VALUES is None:
            try:
                _FILE_VALUES = parse(decode(config_path().read_bytes()))
            except OSError:  # missing, a directory, no permission: no file settings
                _FILE_VALUES = {}
        return _FILE_VALUES


def get(name: str, default: str | None = None) -> str | None:
    """The environment's value, else the config file's, else ``default``. Empty values count as unset."""

    value = os.environ.get(name, "").strip()
    if value:
        return value
    value = _file_values().get(name, "").strip()
    return value if value else default


def reload() -> None:
    """Forget the parsed file (and the cached comment author) so the next read sees changes."""

    global _FILE_VALUES, _WORD_USER
    with _LOCK:
        _FILE_VALUES = None
        _WORD_USER = None


# --- Comment author -----------------------------------------------------------------------------------
# RESEARCH_MCP_AUTHOR (env or config) -> Word's own user name (Application.UserName, read through the
# project's Word worker when it is already running, else the Office user-info registry key Word reads it
# from) -> "Author". Never starts Word just to ask, and never waits on it longer than WORD_NAME_TIMEOUT:
# the worker answers one request at a time, so during a long Zotero refresh even its pipe handshake blocks.
# A name that is found is cached for the process.

_WORD_USER: tuple[str, str] | None = None


def _word_user_from_worker() -> tuple[str, str] | None:
    from .word_service import _request_from_state, read_worker_state, worker_state_is_live

    state = read_worker_state()
    if not worker_state_is_live(state):
        return None
    response = _request_from_state(state, {"command": "ping"}, timeout=WORD_NAME_TIMEOUT)
    name = str(response.get("user_name") or "").strip()
    return (name, str(response.get("user_initials") or "").strip()) if name else None


def _word_user_from_registry() -> tuple[str, str] | None:
    import winreg

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\Common\UserInfo") as key:
        def value(name: str) -> str:
            try:
                return str(winreg.QueryValueEx(key, name)[0] or "").strip()
            except OSError:
                return ""
        name = value("UserName")
        return (name, value("UserInitials")) if name else None


def _bounded(source, timeout: float):
    """``source()`` in a daemon thread; None when it raises or is still running after ``timeout`` seconds."""

    box: dict = {}

    def run() -> None:
        try:
            box["value"] = source()
        except Exception:
            box["value"] = None

    thread = threading.Thread(target=run, name="research-mcp-word-user", daemon=True)
    thread.start()
    thread.join(timeout)
    return box.get("value")


def _word_user() -> tuple[str, str] | None:
    """Word's user name and initials, or None when Word can't tell (not installed, not set, not Windows)."""

    for source in (_word_user_from_worker, _word_user_from_registry):
        found = _bounded(source, WORD_NAME_TIMEOUT)
        if found:
            return found
    return None


def word_user() -> tuple[str, str] | None:
    global _WORD_USER
    if _WORD_USER is None:
        _WORD_USER = _word_user()
    return _WORD_USER


def initials_for(name: str) -> str:
    return "".join(word[0] for word in name.split() if word[:1].isalpha()).upper()[:9]


def comment_author() -> tuple[str, str]:
    """The default comment author and initials."""

    name = get("RESEARCH_MCP_AUTHOR")
    if name:
        return name, get("RESEARCH_MCP_AUTHOR_INITIALS") or initials_for(name)
    found = word_user()
    if found:
        return found[0], found[1] or initials_for(found[0])
    return DEFAULT_AUTHOR, initials_for(DEFAULT_AUTHOR)


def default_comment_author() -> str:
    return comment_author()[0]


def issues_mode() -> str:
    """``agent``: reports come back as an improvement payload plus ``file_with`` steps for the agent's own
    issue tool (``RESEARCH_MCP_ISSUES_FILE_WITH`` overrides the generic steps). Anything else: a GitHub
    issue link."""

    return (get("RESEARCH_MCP_ISSUES", "github") or "github").strip().lower()


def issues_file_with(default: str) -> str:
    return (get("RESEARCH_MCP_ISSUES_FILE_WITH") or "").strip() or default
