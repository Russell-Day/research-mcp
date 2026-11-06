"""``research-mcp setup`` and ``research-mcp doctor``.

setup: check this PC (Windows, Python, Word, Zotero), optionally store a Zotero write key and the comment
author in the config file (``config.config_path()``), then add one ``research`` server entry to each MCP
client the user picks: Claude Desktop, Claude Code, Codex, Cursor, VS Code. Existing client files are
backed up and merged (other servers and settings are kept); a file that can't be parsed is refused, never
replaced. Keys are never written into client configs: the server reads them from the config file.

doctor: the same checks, which clients have a research entry, and where the config file is. Exit code 1
when a required check fails.

Everything that touches the machine (paths, PATH lookups, HTTP, the Word registry check, subprocesses) is a
parameter, so tests run against temp dirs and fakes.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
import datetime as _dt
import getpass
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import sysconfig
import tomllib
from typing import Callable, Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from . import config

SERVER_NAME = "research"
PACKAGE = "research-mcp"
ZOTERO_LOCAL_URL = "http://127.0.0.1:23119"
ZOTERO_KEYS_PAGE = "https://www.zotero.org/settings/keys"
ZOTERO_DOWNLOAD = "https://www.zotero.org/download/"
ZOTERO_SETTING = (
    "In Zotero: Edit > Settings > Advanced, tick 'Allow other applications on this computer to "
    "communicate with Zotero'"
)
UV_INSTALL = 'powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"'
CLIENTS = ("claude-desktop", "claude-code", "codex", "cursor", "vscode")
CLIENT_LABELS = {
    "claude-desktop": "Claude Desktop",
    "claude-code": "Claude Code",
    "codex": "Codex",
    "cursor": "Cursor",
    "vscode": "VS Code",
}
CODEX_STARTUP_TIMEOUT = 60  # seconds; uvx's first run downloads the package, Codex's default is 10


class SetupRefused(Exception):
    """A client file can't be merged safely; nothing was written to it."""


# --- Where things live -----------------------------------------------------------------------------------


@dataclass(frozen=True)
class Home:
    appdata: Path
    localappdata: Path
    userprofile: Path
    codex_home: Path
    config_file: Path

    @classmethod
    def from_env(cls, env: dict[str, str] | None = None) -> "Home":
        env = dict(os.environ if env is None else env)
        profile = Path(env.get("USERPROFILE") or env.get("HOME") or Path.home())
        appdata = Path(env.get("APPDATA") or profile / "AppData" / "Roaming")
        local = Path(env.get("LOCALAPPDATA") or profile / "AppData" / "Local")
        codex = Path(env.get("CODEX_HOME") or profile / ".codex")
        override = (env.get("RESEARCH_MCP_CONFIG") or "").strip()
        cfg = Path(override) if override else appdata / "research-mcp" / "config.env"
        return cls(appdata, local, profile, codex, cfg)


@dataclass(frozen=True)
class Launch:
    command: str
    args: list[str]
    why: str


def _ran_from_uv_cache(prefix: str) -> bool:
    """uvx runs tools from an environment in uv's cache (``...\\uv\\cache\\archive-v0\\...``)."""

    parts = [part.lower() for part in Path(prefix).parts]
    return any(parts[i] == "uv" and parts[i + 1] == "cache" for i in range(len(parts) - 1))


def detect_launch(mode: str = "auto", *, which: Callable[[str], str | None] = shutil.which,
                  prefix: str | None = None, scripts_dir: str | None = None) -> Launch:
    """How MCP clients should start the server.

    ``uvx``: ``uvx research-mcp``. Setup was run with uvx, whose environment lives in uv's cache and can be
    pruned, so the client asks uvx again each time (it reuses the cached copy; ``uvx research-mcp@latest``
    or ``uv cache clean`` picks up a new release). The absolute path of ``uvx.exe`` is used when known,
    because apps started from the Start menu may not see a PATH change made after they started.
    ``path``: the ``research-mcp`` executable of the Python environment setup runs in (``pip install``,
    ``uv tool install``, a dev checkout). That environment is permanent, so the client runs it directly.
    """

    prefix = sys.prefix if prefix is None else prefix
    if mode == "auto":
        mode = "uvx" if _ran_from_uv_cache(prefix) else "path"
    if mode == "uvx":
        uvx = which("uvx")
        return Launch(uvx or "uvx", [PACKAGE],
                      "uvx runs the published package and keeps it cached; setup itself ran through uvx")
    scripts = Path(scripts_dir or sysconfig.get_path("scripts"))
    for name in (f"{PACKAGE}.exe", PACKAGE):
        exe = scripts / name
        if exe.is_file():
            return Launch(str(exe), [], "the research-mcp program of the Python environment setup ran in")
    return Launch(sys.executable, ["-m", "research_mcp"], "the Python that ran setup, as a module")


# --- Client config formats -------------------------------------------------------------------------------


@dataclass(frozen=True)
class Target:
    client: str
    label: str
    path: Path | None  # None: Claude Code through its own CLI
    fmt: str  # json:mcpServers | json:claude-code | json:servers | toml:codex | cli:claude
    detected: bool
    entry_file: Path | None = None  # cli:claude: the file the CLI keeps user-scope servers in (~/.claude.json)


def _entry(fmt: str, launch: Launch) -> dict:
    if fmt == "json:servers":  # VS Code
        return {"type": "stdio", "command": launch.command, "args": list(launch.args)}
    if fmt == "json:claude-code":
        return {"type": "stdio", "command": launch.command, "args": list(launch.args), "env": {}}
    if fmt == "toml:codex":
        return {"command": launch.command, "args": list(launch.args),
                "startup_timeout_sec": CODEX_STARTUP_TIMEOUT}
    return {"command": launch.command, "args": list(launch.args)}


def _toml_str(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)  # JSON's escapes are valid TOML basic-string escapes


def codex_table(launch: Launch) -> str:
    args = ", ".join(_toml_str(arg) for arg in launch.args)
    return (
        f"[mcp_servers.{SERVER_NAME}]\n"
        f"command = {_toml_str(launch.command)}\n"
        f"args = [{args}]\n"
        f"startup_timeout_sec = {CODEX_STARTUP_TIMEOUT}\n"
    )


def claude_cli_command(launch: Launch, claude: str = "claude") -> list[str]:
    return [claude, "mcp", "add", "--scope", "user", SERVER_NAME, "--", launch.command, *launch.args]


def snippet(fmt: str, launch: Launch) -> str:
    """What a user would paste by hand for this format."""

    if fmt == "toml:codex":
        return codex_table(launch)
    if fmt == "cli:claude":
        return subprocess.list2cmdline(claude_cli_command(launch))
    top = "servers" if fmt == "json:servers" else "mcpServers"
    return json.dumps({top: {SERVER_NAME: _entry(fmt, launch)}}, indent=2, ensure_ascii=False)


def _read_text(path: Path) -> str | None:
    try:
        data = path.read_bytes()
    except FileNotFoundError:
        return None
    return config.decode(data)


def merge_json(text: str | None, top_key: str, entry: dict, *, where: str = "file") -> tuple[str, bool]:
    """``text`` with ``{top_key: {"research": entry}}`` merged in; every other key is kept.
    Returns (new text, changed). Raises SetupRefused on text that isn't a JSON object."""

    if text is None or not text.strip():
        data: dict = {}
    else:
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise SetupRefused(
                f"{where} is not valid JSON ({exc.msg} at line {exc.lineno}, column {exc.colno}); "
                "it may contain comments. Fix it or add the entry by hand."
            ) from exc
        if not isinstance(data, dict):
            raise SetupRefused(f"{where} does not hold a JSON object; add the entry by hand.")
    servers = data.get(top_key)
    if servers is None:
        servers = data[top_key] = {}
    if not isinstance(servers, dict):
        raise SetupRefused(f'"{top_key}" in {where} is not an object; add the entry by hand.')
    if servers.get(SERVER_NAME) == entry:
        return text or "", False
    servers[SERVER_NAME] = entry
    return json.dumps(data, indent=2, ensure_ascii=False) + "\n", True


_TOML_HEADER = re.compile(r"^\s*\[\[?\s*(?P<name>[^\]]+?)\s*\]\]?\s*(?:#.*)?$")


def _toml_key_parts(name: str) -> list[str]:
    parts = re.findall(r'"(?:[^"\\]|\\.)*"|\'[^\']*\'|[^.\s]+', name)
    return [part[1:-1] if part[:1] in "\"'" else part for part in parts]


def merge_codex_toml(text: str | None, launch: Launch, *, where: str = "file") -> tuple[str, bool]:
    """Replace (in place) or append the ``[mcp_servers.research]`` table; every other line stays as it was."""

    text = text or ""
    try:
        data = tomllib.loads(text)
    except tomllib.TOMLDecodeError as exc:
        raise SetupRefused(f"{where} is not valid TOML ({exc}); fix it or add the entry by hand.") from exc
    wanted = _entry("toml:codex", launch)
    if (data.get("mcp_servers") or {}).get(SERVER_NAME) == wanted:
        return text, False

    nl = "\r\n" if "\r\n" in text else "\n"  # keep the file's own line endings
    not_a_table = (f"{where} defines mcp_servers.{SERVER_NAME} in a form setup can't edit (an inline table or "
                   "dotted keys); add the entry by hand.")
    lines = text.splitlines(keepends=True)
    start = end = None
    for index, line in enumerate(lines):
        match = _TOML_HEADER.match(line)
        if not match:
            continue
        parts = _toml_key_parts(match.group("name"))
        ours = parts[:2] == ["mcp_servers", SERVER_NAME]
        if start is None and ours:
            start = index
        elif start is not None and end is None and not ours:
            end = index
    if start is None and SERVER_NAME in (data.get("mcp_servers") or {}):
        raise SetupRefused(not_a_table)
    table = codex_table(launch).replace("\n", nl)
    if start is None:
        prefix = text if not text or text.endswith("\n") else text + nl
        gap = nl if prefix.strip() and not prefix.endswith(nl + nl) else ""
        new = prefix + gap + table
    else:
        end = len(lines) if end is None else end
        # Comments and blank lines just above the next header belong to that header.
        while end > start + 1 and (not lines[end - 1].strip() or lines[end - 1].lstrip().startswith("#")):
            end -= 1
        tail = lines[end:]
        new = "".join(lines[:start]) + table + (nl if tail else "") + "".join(tail)
    try:
        check = tomllib.loads(new)
    except tomllib.TOMLDecodeError as exc:  # part of the entry is set outside its [table] (dotted keys)
        raise SetupRefused(not_a_table) from exc
    if (check.get("mcp_servers") or {}).get(SERVER_NAME) != wanted:
        raise SetupRefused(not_a_table)
    return new, True


# --- Finding clients ---------------------------------------------------------------------------------------


def find_targets(home: Home, *, which: Callable[[str], str | None] = shutil.which) -> list[Target]:
    targets: list[Target] = []

    desktop_seen = (
        (home.appdata / "Claude").is_dir()
        or (home.localappdata / "AnthropicClaude").is_dir()
        or any(home.localappdata.glob("Packages/Claude_*"))
    )
    targets.append(Target("claude-desktop", "Claude Desktop",
                          home.appdata / "Claude" / "claude_desktop_config.json", "json:mcpServers",
                          desktop_seen))
    # The Microsoft Store (MSIX) build keeps its settings in the package's own copy of AppData.
    for package in sorted(home.localappdata.glob("Packages/Claude_*")):
        roaming = package / "LocalCache" / "Roaming" / "Claude"
        if roaming.is_dir():
            targets.append(Target("claude-desktop", "Claude Desktop (Store app)",
                                  roaming / "claude_desktop_config.json", "json:mcpServers", True))

    claude = which("claude")
    claude_json = home.userprofile / ".claude.json"
    if claude:
        targets.append(Target("claude-code", "Claude Code", None, "cli:claude", True, entry_file=claude_json))
    else:
        targets.append(Target("claude-code", "Claude Code", claude_json, "json:claude-code",
                              claude_json.exists() or (home.userprofile / ".claude").is_dir()))

    targets.append(Target("codex", "Codex", home.codex_home / "config.toml", "toml:codex",
                          home.codex_home.is_dir() or bool(which("codex"))))
    targets.append(Target("cursor", "Cursor", home.userprofile / ".cursor" / "mcp.json", "json:mcpServers",
                          (home.userprofile / ".cursor").is_dir() or bool(which("cursor"))))
    targets.append(Target("vscode", "VS Code", home.appdata / "Code" / "User" / "mcp.json", "json:servers",
                          (home.appdata / "Code" / "User").is_dir() or bool(which("code"))))
    insiders = home.appdata / "Code - Insiders" / "User"
    if insiders.is_dir():
        targets.append(Target("vscode", "VS Code Insiders", insiders / "mcp.json", "json:servers", True))
    return targets


# --- Planning and writing client entries -----------------------------------------------------------------


@dataclass
class Plan:
    target: Target
    action: str  # create | update | unchanged | refused | run
    new_text: str | None = None
    message: str = ""
    command: list[str] = field(default_factory=list)
    dropped: list[str] = field(default_factory=list)  # settings of the old entry the new one leaves out


def _old_entry(text: str | None, fmt: str) -> dict:
    try:
        if fmt == "toml:codex":
            data = tomllib.loads(text or "")
            servers = data.get("mcp_servers") or {}
        else:
            data = json.loads(text or "{}")
            servers = data.get("servers" if fmt == "json:servers" else "mcpServers") or {}
        old = servers.get(SERVER_NAME) if isinstance(servers, dict) else None
    except (ValueError, AttributeError):
        return {}
    return old if isinstance(old, dict) else {}


def plan_target(target: Target, launch: Launch, *, which: Callable[[str], str | None] = shutil.which) -> Plan:
    if target.fmt == "cli:claude":
        return Plan(target, "run", command=claude_cli_command(launch, which("claude") or "claude"))
    assert target.path is not None
    where = str(target.path)
    text = _read_text(target.path)
    try:
        if target.fmt == "toml:codex":
            new, changed = merge_codex_toml(text, launch, where=where)
        else:
            top = "servers" if target.fmt == "json:servers" else "mcpServers"
            new, changed = merge_json(text, top, _entry(target.fmt, launch), where=where)
    except SetupRefused as exc:
        return Plan(target, "refused", message=str(exc))
    if not changed:
        return Plan(target, "unchanged", message="already has this research entry")
    dropped = sorted(set(_old_entry(text, target.fmt)) - set(_entry(target.fmt, launch)))
    return Plan(target, "create" if text is None else "update", new_text=new, dropped=dropped)


def backup_path(path: Path, now: _dt.datetime | None = None) -> Path:
    plain = path.with_name(path.name + ".bak")
    if not plain.exists():
        return plain
    stamp = (now or _dt.datetime.now()).strftime("%Y%m%d-%H%M%S")
    candidate = path.with_name(f"{path.name}.{stamp}.bak")
    counter = 1
    while candidate.exists():
        candidate = path.with_name(f"{path.name}.{stamp}-{counter}.bak")
        counter += 1
    return candidate


def _atomic_write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".research-mcp.tmp")
    tmp.write_text(text, encoding="utf-8", newline="")
    os.replace(tmp, path)


Runner = Callable[[list[str]], "subprocess.CompletedProcess[str]"]


def _run(command: list[str]) -> "subprocess.CompletedProcess[str]":
    return subprocess.run(command, capture_output=True, text=True, timeout=120)


def apply_plan(plan: Plan, *, run: Runner = _run) -> str:
    """Carry out one plan; returns a one-line result."""

    if plan.action == "run":
        # The CLI rewrites ~/.claude.json; keep a copy first, as for every other client file.
        entry_file = plan.target.entry_file
        backup = None
        if entry_file is not None and entry_file.exists():
            backup = backup_path(entry_file)
            shutil.copy2(entry_file, backup)
        result = run(plan.command)
        output = (result.stdout or "") + (result.stderr or "")
        if result.returncode != 0 and "already exists" in output.lower():
            return _replace_claude_entry(plan, run, backup)
        if result.returncode != 0:
            raise SetupRefused(f"`claude mcp add` failed: {output.strip()[:300]}")
        return "added with `claude mcp add --scope user`"
    if plan.action in ("unchanged", "refused"):
        return plan.message
    assert plan.target.path is not None and plan.new_text is not None
    backup = None
    if plan.target.path.exists():
        backup = backup_path(plan.target.path)
        shutil.copy2(plan.target.path, backup)
    _atomic_write(plan.target.path, plan.new_text)
    return f"wrote {plan.target.path}" + (f" (backup: {backup.name})" if backup else "")


def _claude_user_entry(path: Path | None) -> dict | None:
    """The user-scope ``research`` entry in ~/.claude.json, or None when it can't be read."""
    text = _read_text(path) if path else None
    try:
        servers = (json.loads(text) if text else {}).get("mcpServers")
    except (ValueError, AttributeError):
        return None
    entry = servers.get(SERVER_NAME) if isinstance(servers, dict) else None
    return entry if isinstance(entry, dict) else None


def _replace_claude_entry(plan: Plan, run: Runner, backup: Path | None = None) -> str:
    """Claude Code already has a ``research`` server: keep a copy, remove it, add ours, and put the old one
    back if the new add fails. Nothing is removed when the old entry can't be read first."""

    claude, path = plan.command[0], plan.target.entry_file
    old = _claude_user_entry(path)
    if old is None or path is None:
        raise SetupRefused(
            f"Claude Code already has a `{SERVER_NAME}` server that setup could not read from {path}, so it "
            f"was left as it is. Replace it by hand: `claude mcp remove --scope user {SERVER_NAME}`, then "
            "run setup again."
        )
    if backup is None:
        backup = backup_path(path)
        shutil.copy2(path, backup)

    def attempt(command: list[str]) -> tuple[int, str]:
        # After the remove, a timeout or a missing CLI must still reach the rollback below.
        try:
            result = run(command)
        except (subprocess.SubprocessError, OSError) as exc:
            return 1, f"{type(exc).__name__}: {exc}"
        return result.returncode, (result.stdout or "") + (result.stderr or "")

    run([claude, "mcp", "remove", "--scope", "user", SERVER_NAME])
    code, output = attempt(plan.command)
    if code == 0:
        return f"replaced the old entry with `claude mcp add --scope user` (backup: {backup.name})"
    output = output.strip()[:300]
    restored, _ = attempt([claude, "mcp", "add-json", "--scope", "user", SERVER_NAME, json.dumps(old)])
    if restored == 0 and _claude_user_entry(path) == old:
        how = "your previous entry was put back"
    else:
        try:
            new, _ = merge_json(_read_text(path), "mcpServers", old, where=str(path))
            _atomic_write(path, new)
            how = f"your previous entry was written back into {path}"
        except (SetupRefused, OSError):
            how = f"putting your previous entry back failed; it is in {backup}"
    raise SetupRefused(f"`claude mcp add` failed after removing the old entry ({output}); {how}.")


# --- config.env ----------------------------------------------------------------------------------------


def _env_value(value: str) -> str:
    if value and not re.search(r"\s|#|['\"]", value):
        return value
    return f"'{value}'" if '"' in value else f'"{value}"'


def merge_env_text(text: str | None, updates: dict[str, str]) -> str:
    """Set each key of ``updates`` (replacing its line in place, or appending); other lines are kept."""

    text = text or ""
    newline = "\r\n" if "\r\n" in text else "\n"
    lines = text.splitlines()
    done: set[str] = set()
    for index, line in enumerate(lines):
        match = re.match(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=", line)
        if match and match.group(1) in updates and match.group(1) not in done:
            lines[index] = f"{match.group(1)}={_env_value(updates[match.group(1)])}"
            done.add(match.group(1))
    if not text:
        lines.append("# research-mcp settings (see `research-mcp doctor`). Keep this file private.")
    for key, value in updates.items():
        if key not in done:
            lines.append(f"{key}={_env_value(value)}")
    return newline.join(lines) + newline


def write_config_env(path: Path, updates: dict[str, str]) -> None:
    text = _read_text(path)
    _atomic_write(path, merge_env_text(text, updates))
    config.reload()


# --- Checks -----------------------------------------------------------------------------------------------


@dataclass
class Check:
    name: str
    status: str  # ok | fail | warn | info
    detail: str
    fix: str = ""


Fetch = Callable[[str, dict[str, str], float], tuple[int, bytes]]


def _http_get(url: str, headers: dict[str, str], timeout: float) -> tuple[int, bytes]:
    """(status, body); status 0 when nothing answered."""

    try:
        with urlopen(Request(url, headers=headers), timeout=timeout) as response:
            return response.status, response.read()
    except HTTPError as exc:
        return exc.code, exc.read() if exc.fp else b""
    except (URLError, OSError, TimeoutError):
        return 0, b""


def _word_registered() -> bool | None:
    """Whether the COM class ``Word.Application`` is registered; None off Windows. Never starts Word."""

    if sys.platform != "win32":
        return None
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except OSError:
        return False


def zotero_key_info(key: str, *, fetch: Fetch = _http_get) -> dict:
    """The Zotero web API's description of a key (one read-only call). Raises SetupRefused when the key
    is rejected; returns {} when Zotero's servers can't be reached."""

    headers = {"Zotero-API-Key": key, "Zotero-API-Version": "3", "User-Agent": f"{PACKAGE}-setup"}
    status, body = fetch("https://api.zotero.org/keys/current", headers, 15.0)
    if status == 0:
        return {}
    if status in (403, 404):
        raise SetupRefused("Zotero did not accept this key (check that you copied all of it)")
    if status != 200:
        raise SetupRefused(f"Zotero answered HTTP {status} when checking the key")
    try:
        info = json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise SetupRefused("Zotero's answer about the key was not JSON") from exc
    return info if isinstance(info, dict) else {}


def _key_summary(info: dict) -> tuple[str, bool]:
    user = (info.get("access") or {}).get("user") or {}
    return str(info.get("userID") or ""), bool(user.get("write"))


def run_checks(home: Home, *, fetch: Fetch = _http_get, word_registered: Callable[[], bool | None] = _word_registered,
               platform: str | None = None, version: tuple[int, int] | None = None,
               online: bool = True) -> list[Check]:
    platform = sys.platform if platform is None else platform
    version = tuple(sys.version_info[:2]) if version is None else version
    checks: list[Check] = []

    checks.append(Check("Windows", "ok", platform) if platform == "win32" else Check(
        "Windows", "fail", f"this is {platform}",
        "research-mcp drives Microsoft Word through COM, which only exists on Windows."))

    py = f"{version[0]}.{version[1]}"
    checks.append(Check("Python", "ok", py) if (3, 12) <= version < (3, 15) else Check(
        "Python", "fail", f"Python {py}; research-mcp needs 3.12, 3.13 or 3.14",
        "Run it through uv, which picks a matching Python: uvx research-mcp setup"))

    word = word_registered()
    if word:
        checks.append(Check("Microsoft Word", "ok", "installed (Word.Application is registered)"))
    elif word is None:
        checks.append(Check("Microsoft Word", "fail", "can't check off Windows", "Use a Windows PC with Word."))
    else:
        checks.append(Check("Microsoft Word", "fail", "Word.Application is not registered",
                            "Install the desktop Microsoft Word (Microsoft 365 or Office 2016+) for this user."))

    status, _ = fetch(f"{ZOTERO_LOCAL_URL}/api/users/0/items/top?limit=1",
                      {"Zotero-Allowed-Request": "1", "Accept": "application/json"}, 3.0)
    if status == 200:
        checks.append(Check("Zotero", "ok", "running, local API on"))
    elif status == 0:
        checks.append(Check("Zotero", "fail", "Zotero is not running (or not installed)",
                            f"Install Zotero 7 ({ZOTERO_DOWNLOAD}), start it, and keep it open while you work. "
                            + ZOTERO_SETTING + "."))
    elif status == 403:
        checks.append(Check("Zotero", "fail", "Zotero is running but its local API is off", ZOTERO_SETTING + "."))
    else:
        checks.append(Check("Zotero", "fail", f"Zotero's local API answered HTTP {status}",
                            "Update Zotero to version 7 or later. " + ZOTERO_SETTING + "."))

    values = config.parse(_read_text(home.config_file) or "")
    key = (os.environ.get("ZOTERO_API_KEY") or values.get("ZOTERO_API_KEY") or "").strip()
    library = (os.environ.get("ZOTERO_LIBRARY_ID") or values.get("ZOTERO_LIBRARY_ID") or "").strip()
    if not key:
        checks.append(Check("Zotero write key", "info",
                            "not set (optional: only needed to add or change Zotero items)",
                            "research-mcp setup"))
    elif not online:
        checks.append(Check("Zotero write key", "info", "set (not checked)"))
    else:
        try:
            info = zotero_key_info(key, fetch=fetch)
        except SetupRefused as exc:
            checks.append(Check("Zotero write key", "fail", str(exc),
                                f"Make a new key at {ZOTERO_KEYS_PAGE} and run research-mcp setup"))
        else:
            if not info:
                checks.append(Check("Zotero write key", "warn", "set, but api.zotero.org could not be reached"))
            else:
                user_id, can_write = _key_summary(info)
                problems = []
                if not can_write:
                    problems.append("the key can't write to your library (tick 'Allow write access')")
                if not library:
                    problems.append("ZOTERO_LIBRARY_ID is not set")
                elif user_id and library != user_id and (values.get("ZOTERO_LIBRARY_TYPE") or "user") == "user":
                    problems.append(f"ZOTERO_LIBRARY_ID is {library} but the key belongs to user {user_id}")
                checks.append(Check("Zotero write key", "fail", "; ".join(problems),
                                    f"Fix it at {ZOTERO_KEYS_PAGE}, then run research-mcp setup")
                              if problems else Check("Zotero write key", "ok", f"valid for user {user_id}"))
    return checks


def _print_checks(checks: Iterable[Check], out) -> None:
    tags = {"ok": "OK  ", "fail": "FAIL", "warn": "WARN", "info": "--  "}
    for check in checks:
        print(f"  [{tags[check.status]}] {check.name}: {check.detail}", file=out)
        if check.fix and check.status in ("fail", "warn"):
            print(f"         fix: {check.fix}", file=out)


def client_entries(home: Home) -> list[tuple[str, str]]:
    """(client label, where) for each client config that has a research-mcp server entry."""

    found: list[tuple[str, str]] = []
    for target in find_targets(home, which=lambda name: None):  # file route, so ~/.claude.json is read too
        text = _read_text(target.path) if target.path else None
        if not text:
            continue
        try:
            data = tomllib.loads(text) if target.fmt == "toml:codex" else json.loads(text)
        except (tomllib.TOMLDecodeError, json.JSONDecodeError):
            found.append((target.label, f"{target.path} (could not be parsed)"))
            continue
        key = {"toml:codex": "mcp_servers", "json:servers": "servers"}.get(target.fmt, "mcpServers")
        names = _research_names(data.get(key) if isinstance(data, dict) else None)
        if names:
            found.append((target.label, f"{target.path} ({', '.join(names)})"))
    return found


def _research_names(servers) -> list[str]:
    if not isinstance(servers, dict):
        return []
    names = []
    for name, entry in servers.items():
        blob = json.dumps(entry, ensure_ascii=False).lower() if isinstance(entry, dict) else ""
        if name == SERVER_NAME or "research-mcp" in blob or "research_mcp" in blob:
            names.append(name)
    return names


# --- Commands ------------------------------------------------------------------------------------------


def _setup_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{PACKAGE} setup",
        description="Check this PC, store an optional Zotero write key, and add research-mcp to your MCP clients.",
    )
    parser.add_argument("--yes", "-y", action="store_true",
                        help="no questions: add the server to every detected client (or each --client)")
    parser.add_argument("--client", action="append", choices=CLIENTS, default=[],
                        help="add the server to this client (repeatable); skips detection for it")
    parser.add_argument("--print", dest="print_only", action="store_true",
                        help="print the config snippet for every client and change nothing")
    parser.add_argument("--dry-run", action="store_true", help="show what would change and change nothing")
    parser.add_argument("--zotero-key-from-env", action="store_true",
                        help="store ZOTERO_API_KEY (and ZOTERO_LIBRARY_ID) from the environment")
    parser.add_argument("--author", help="comment author name for Word comments (default: Word's user name)")
    parser.add_argument("--launcher", choices=("auto", "uvx", "path"), default="auto",
                        help="how clients start the server (default: uvx when setup runs under uvx)")
    return parser


def _ask(prompt: str, input_fn: Callable[[str], str]) -> str:
    try:
        return input_fn(prompt).strip()
    except EOFError:
        return ""


def _yes(prompt: str, input_fn: Callable[[str], str]) -> bool:
    return _ask(prompt + " [y/N] ", input_fn).lower() in ("y", "yes")


@dataclass
class Io:
    out: object = None
    input: Callable[[str], str] = input
    secret: Callable[[str], str] = getpass.getpass
    interactive: bool | None = None

    def __post_init__(self) -> None:
        if self.out is None:
            self.out = sys.stdout


def _say(io: Io, text: str = "") -> None:
    print(text, file=io.out)


def setup_main(argv: list[str] | None = None, *, home: Home | None = None, io: Io | None = None,
               which: Callable[[str], str | None] = shutil.which, fetch: Fetch = _http_get,
               word_registered: Callable[[], bool | None] = _word_registered, run: Runner = _run,
               launch: Launch | None = None) -> int:
    args = _setup_parser().parse_args(argv)
    home = home or Home.from_env()
    io = io or Io()
    _quiet_encoding_errors()
    launch = launch or detect_launch(args.launcher, which=which)

    if args.print_only:
        _say(io, f"research-mcp server entry for each client (starts: {_shown(launch)})")
        _say(io, f"  why this command: {launch.why}")
        _say(io, "  Keys don't go here; they live in " + str(home.config_file) + ".")
        for target in find_targets(home, which=which):
            where = str(target.path) if target.path else "run in a terminal"
            _say(io, f"\n# {target.label}: {where}")
            _say(io, snippet(target.fmt, launch))
            if target.fmt == "cli:claude":
                _say(io, f"# (or add to {home.userprofile / '.claude.json'}:)")
                _say(io, snippet("json:claude-code", launch))
        return 0

    interactive = io.interactive if io.interactive is not None else sys.stdin.isatty()
    interactive = interactive and not args.yes and not args.dry_run
    failed = False

    _say(io, "Checking this PC")
    checks = run_checks(home, fetch=fetch, word_registered=word_registered)
    _print_checks(checks, io.out)
    failed |= any(check.status == "fail" and check.name != "Zotero write key" for check in checks)

    # Zotero write key and comment author -> config file
    updates: dict[str, str] = {}
    existing = config.parse(_read_text(home.config_file) or "")
    key = library = ""
    if args.zotero_key_from_env:
        key = (os.environ.get("ZOTERO_API_KEY") or "").strip()
        library = (os.environ.get("ZOTERO_LIBRARY_ID") or "").strip()
        if not key:
            _say(io, "\n--zotero-key-from-env: ZOTERO_API_KEY is not set in the environment; skipped.")
    elif interactive:
        _say(io, "\nZotero write access (optional)")
        _say(io, "  Searching and citing use the Zotero app on this PC and need no key. A key is only needed")
        _say(io, "  for tools that add or change items in your library (add by DOI, fix a reference).")
        _say(io, f"  Make one at {ZOTERO_KEYS_PAGE} ('Create new private key', tick 'Allow write access').")
        _say(io, "  Your user ID is shown on the same page ('Your user ID for use in API calls').")
        question = "  Replace the saved key?" if existing.get("ZOTERO_API_KEY") else "  Add a key now?"
        if _yes(question, io.input):
            key = io.secret("  Zotero API key (hidden): ").strip()
    if key:
        try:
            info = zotero_key_info(key, fetch=fetch)
        except SetupRefused as exc:
            _say(io, f"  The key was not saved: {exc}.")
            failed = True
            key = ""
        else:
            user_id, can_write = _key_summary(info)
            if not info:
                _say(io, "  Could not reach api.zotero.org to check the key; saving it anyway.")
            elif not can_write:
                _say(io, "  Note: this key is read-only; tick 'Allow write access' for it on the keys page.")
            if not library and interactive:
                hint = f" [{user_id}]" if user_id else ""
                for _attempt in range(3):
                    library = _ask(f"  Your Zotero user ID (a number){hint}: ", io.input) or user_id
                    if not library or library.isdigit():
                        break
                    _say(io, f"  {library!r} is not a user ID; it is the number on {ZOTERO_KEYS_PAGE}.")
            library = library or user_id
            if library and not library.isdigit():
                _say(io, f"  The key was not saved: the user ID must be a number (see {ZOTERO_KEYS_PAGE}).")
                failed = True
                key = ""
            elif user_id and library and library != user_id:
                _say(io, f"  Note: this key belongs to user {user_id}, not {library}.")
            elif not library:
                _say(io, f"  The key was not saved: no user ID (see {ZOTERO_KEYS_PAGE}).")
                failed = True
                key = ""
    if key:
        updates.update({"ZOTERO_API_KEY": key, "ZOTERO_LIBRARY_ID": library})
        if not existing.get("ZOTERO_LIBRARY_TYPE"):
            updates["ZOTERO_LIBRARY_TYPE"] = "user"

    author = (args.author or "").strip()
    if not author and interactive:
        word_name = (config.word_user() or ("", ""))[0]
        current = existing.get("RESEARCH_MCP_AUTHOR") or ""
        _say(io, "\nComment author (optional)")
        _say(io, "  Word comments added through research-mcp carry an author name. By default it is the")
        _say(io, "  user name Word shows in File > Options > General" + (f" ({word_name})." if word_name else "."))
        if current:
            _say(io, f"  Saved now: {current}.")
        author = _ask("  Author name (Enter keeps the default): ", io.input)
    if author:
        updates["RESEARCH_MCP_AUTHOR"] = author

    if updates:
        shown_keys = ", ".join(updates)  # names only; never the values of keys
        if args.dry_run:
            _say(io, f"\nWould save to {home.config_file}: {shown_keys}")
        else:
            write_config_env(home.config_file, updates)
            _say(io, f"\nSaved to {home.config_file}: {shown_keys}")

    # Client configs
    targets = find_targets(home, which=which)
    wanted = set(args.client)
    if wanted:
        chosen = [t for t in targets if t.client in wanted and (t.detected or t.label == CLIENT_LABELS[t.client])]
    else:
        chosen = [t for t in targets if t.detected]
    _say(io, f"\nMCP clients (server entry: {_shown(launch)})")
    _say(io, f"  why this command: {launch.why}")
    if not chosen:
        _say(io, "  No MCP clients found. `research-mcp setup --print` shows the entry to add by hand.")
    elif not (interactive or args.yes or wanted or args.dry_run):
        _say(io, "  Found: " + ", ".join(t.label for t in chosen))
        _say(io, "  Not changed (no terminal to ask in). Run again with --yes, or --client <name>.")
        chosen = []
    written = []
    for target in chosen:
        if interactive and not wanted and not _yes(f"  Add research-mcp to {target.label}?", io.input):
            continue
        plan = plan_target(target, launch, which=which)
        where = str(target.path) if target.path else "claude CLI"
        if plan.action == "refused":
            _say(io, f"  {target.label}: NOT changed. {plan.message}")
            _say(io, "  Entry to add by hand:\n" + _indent(snippet(target.fmt, launch), 4))
            failed = True
            continue
        if args.dry_run:
            if plan.action == "run":
                _say(io, f"  {target.label}: would run: {subprocess.list2cmdline(plan.command)}")
            elif plan.action == "unchanged":
                _say(io, f"  {target.label}: {where} already has this entry; no change")
            else:
                verb = "create" if plan.action == "create" else "back up and update"
                _say(io, f"  {target.label}: would {verb} {where} with:")
                _say(io, _indent(snippet(target.fmt, launch), 4))
                if plan.dropped:
                    _say(io, f"    (replaces the existing research entry; its {', '.join(plan.dropped)} "
                             "setting(s) would be dropped, kept in the .bak)")
            continue
        try:
            _say(io, f"  {target.label}: {apply_plan(plan, run=run)}")
            written.append(target.label)
            if plan.dropped:
                _say(io, f"    dropped from the old research entry: {', '.join(plan.dropped)} (kept in the .bak)")
        except (SetupRefused, OSError, subprocess.SubprocessError) as exc:
            _say(io, f"  {target.label}: NOT changed. {exc}")
            failed = True

    zotero = next((check for check in checks if check.name == "Zotero"), None)
    if zotero is not None and zotero.status == "fail":
        _say(io, f"\n[FAIL] Zotero: {zotero.detail}")
        if written:
            _say(io, f"  The client configs were written ({', '.join(written)}); Zotero only has to be"
                     " started (with its local API on), no need to run setup again.")
        if zotero.fix:
            _say(io, f"  fix: {zotero.fix}")
    if not args.dry_run:
        _say(io, "\nNext: restart your MCP client (quit it fully, then open it again) and ask it to")
        _say(io, "'list the research tools'. `research-mcp doctor` checks the install later.")
        _say(io, "ChatGPT has no local setup: run `research-mcp chatgpt` and paste the URL it prints.")
    return 1 if failed else 0


def doctor_main(argv: list[str] | None = None, *, home: Home | None = None, out=None,
                fetch: Fetch = _http_get, word_registered: Callable[[], bool | None] = _word_registered) -> int:
    parser = argparse.ArgumentParser(prog=f"{PACKAGE} doctor",
                                     description="Check a research-mcp install. Exit code 1 when a required check fails.")
    parser.add_argument("--offline", action="store_true", help="don't check the Zotero key with api.zotero.org")
    args = parser.parse_args(argv)
    home = home or Home.from_env()
    out = out or sys.stdout
    _quiet_encoding_errors()
    from . import __version__

    print(f"research-mcp {__version__} (Python {sys.version.split()[0]}, {sys.executable})", file=out)
    checks = run_checks(home, fetch=fetch, word_registered=word_registered, online=not args.offline)
    _print_checks(checks, out)
    values = config.parse(_read_text(home.config_file) or "")
    state = "exists" if home.config_file.exists() else "not created yet (optional)"
    print(f"\nConfig file: {home.config_file} ({state})", file=out)
    if values:
        print("  keys set: " + ", ".join(sorted(values)), file=out)  # names only
    entries = client_entries(home)
    print("\nClients with the server entry (named `research`):", file=out)
    for label, where in entries:
        print(f"  {label}: {where}", file=out)
    if not entries:
        print("  none (run `research-mcp setup`)", file=out)
    return 1 if any(check.status == "fail" for check in checks) else 0


def _shown(launch: Launch) -> str:
    return subprocess.list2cmdline([launch.command, *launch.args])


def _indent(text: str, spaces: int) -> str:
    pad = " " * spaces
    return "\n".join(pad + line for line in text.splitlines())


def _quiet_encoding_errors() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
