"""research-mcp setup / doctor and the command dispatcher.

Every test runs against a scratch home (APPDATA, LOCALAPPDATA, USERPROFILE, CODEX_HOME under tmp_path),
fakes the network and the Word registry check, and never runs a real `claude` command.
"""

from __future__ import annotations

import io
import json
import tomllib
from pathlib import Path

import pytest

from research_mcp import __main__ as dispatcher
from research_mcp import setup_cmd as sc

KEY = "FAKEKEY0123456789abcdefSECRET"
USER_ID = "1234567"
LAUNCH = sc.Launch("uvx", ["research-mcp"], "test")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    for name in ("ZOTERO_API_KEY", "ZOTERO_LIBRARY_ID", "ZOTERO_LIBRARY_TYPE"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setattr(sc.config, "word_user", lambda: None)  # don't ask a running Word worker or the registry


@pytest.fixture
def home(tmp_path) -> sc.Home:
    profile = tmp_path / "profile"
    h = sc.Home(
        appdata=profile / "AppData" / "Roaming",
        localappdata=profile / "AppData" / "Local",
        userprofile=profile,
        codex_home=profile / ".codex",
        config_file=profile / "AppData" / "Roaming" / "research-mcp" / "config.env",
    )
    for d in (h.appdata, h.localappdata, h.userprofile):
        d.mkdir(parents=True, exist_ok=True)
    return h


def all_clients(home: sc.Home) -> None:
    """Make every client look installed (its config folder exists)."""

    (home.appdata / "Claude").mkdir(parents=True, exist_ok=True)
    (home.userprofile / ".claude").mkdir(parents=True, exist_ok=True)
    home.codex_home.mkdir(parents=True, exist_ok=True)
    (home.userprofile / ".cursor").mkdir(parents=True, exist_ok=True)
    (home.appdata / "Code" / "User").mkdir(parents=True, exist_ok=True)


def no_which(name: str) -> None:
    return None


def fake_fetch(*, zotero_local: int = 200, key_status: int = 200, user_id: str = USER_ID, write: bool = True):
    calls: list[tuple[str, dict]] = []

    def fetch(url: str, headers: dict[str, str], timeout: float) -> tuple[int, bytes]:
        calls.append((url, dict(headers)))
        if url.startswith(sc.ZOTERO_LOCAL_URL):
            return zotero_local, b"[]"
        if url == "https://api.zotero.org/keys/current":
            if key_status != 200:
                return key_status, b"Forbidden"
            body = {"key": headers.get("Zotero-API-Key"), "userID": int(user_id),
                    "access": {"user": {"library": True, "write": write}}}
            return 200, json.dumps(body).encode()
        raise AssertionError(f"unexpected request {url}")

    fetch.calls = calls  # type: ignore[attr-defined]
    return fetch


def run_setup(home, argv, *, inputs=(), secret="", interactive=False, fetch=None, word=True, which=no_which,
              run=None):
    out = io.StringIO()
    answers = list(inputs)

    def ask(prompt: str) -> str:
        if not answers:
            raise EOFError
        return answers.pop(0)

    commands: list[list[str]] = []

    def fake_run(cmd):
        commands.append(cmd)
        import subprocess

        return subprocess.CompletedProcess(cmd, 0, "Added stdio MCP server research", "")

    code = sc.setup_main(
        argv, home=home, io=sc.Io(out=out, input=ask, secret=lambda prompt: secret, interactive=interactive),
        which=which, fetch=fetch or fake_fetch(), word_registered=lambda: word, run=run or fake_run,
        launch=LAUNCH,
    )
    return code, out.getvalue(), commands


def client_files(home: sc.Home) -> dict[str, Path]:
    return {
        "claude-desktop": home.appdata / "Claude" / "claude_desktop_config.json",
        "claude-code": home.userprofile / ".claude.json",
        "codex": home.codex_home / "config.toml",
        "cursor": home.userprofile / ".cursor" / "mcp.json",
        "vscode": home.appdata / "Code" / "User" / "mcp.json",
    }


# --- merging per format ------------------------------------------------------------------------------------


def test_merge_json_fresh_file():
    text, changed = sc.merge_json(None, "mcpServers", {"command": "uvx", "args": ["research-mcp"]})
    assert changed
    assert json.loads(text) == {"mcpServers": {"research": {"command": "uvx", "args": ["research-mcp"]}}}


def test_merge_json_keeps_other_servers_and_keys():
    before = json.dumps({"theme": "dark", "mcpServers": {"other": {"command": "x"}}})
    text, changed = sc.merge_json(before, "mcpServers", {"command": "uvx", "args": ["research-mcp"]})
    data = json.loads(text)
    assert changed
    assert data["theme"] == "dark"
    assert data["mcpServers"]["other"] == {"command": "x"}
    assert data["mcpServers"]["research"]["command"] == "uvx"


def test_merge_json_same_entry_is_unchanged():
    entry = {"command": "uvx", "args": ["research-mcp"]}
    first, _ = sc.merge_json(None, "mcpServers", entry)
    again, changed = sc.merge_json(first, "mcpServers", entry)
    assert not changed and again == first


@pytest.mark.parametrize("bad", ['{"mcpServers": {', "// comment\n{}", "[1, 2]", '{"mcpServers": []}'])
def test_merge_json_refuses_what_it_cannot_parse(bad):
    with pytest.raises(sc.SetupRefused):
        sc.merge_json(bad, "mcpServers", {"command": "uvx"})


def test_vscode_and_claude_code_entries_have_their_own_shape():
    vscode = json.loads(sc.snippet("json:servers", LAUNCH))
    assert vscode == {"servers": {"research": {"type": "stdio", "command": "uvx", "args": ["research-mcp"]}}}
    code = json.loads(sc.snippet("json:claude-code", LAUNCH))
    assert code["mcpServers"]["research"]["type"] == "stdio"


CODEX_BEFORE = """model = "o4"

[mcp_servers.other]
command = "other"

[mcp_servers.research]
command = "old"
args = ["--stale"]

[mcp_servers.research.env]
STALE = "1"

# notes about the profile below
[profiles.fast]
model = "mini"
"""


def test_codex_toml_replaces_only_the_research_table():
    text, changed = sc.merge_codex_toml(CODEX_BEFORE, LAUNCH)
    data = tomllib.loads(text)
    assert changed
    assert data["model"] == "o4"
    assert data["mcp_servers"]["other"] == {"command": "other"}
    assert data["mcp_servers"]["research"] == {"command": "uvx", "args": ["research-mcp"],
                                               "startup_timeout_sec": sc.CODEX_STARTUP_TIMEOUT}
    assert data["profiles"]["fast"] == {"model": "mini"}
    assert "# notes about the profile below\n[profiles.fast]" in text  # the comment stays with its table
    assert "STALE" not in text and "old" not in text
    again, changed_again = sc.merge_codex_toml(text, LAUNCH)
    assert not changed_again and again == text


def test_codex_toml_appends_to_a_file_without_the_table():
    before = 'model = "o4"\n\n[mcp_servers.other]\ncommand = "other"\n'
    text, changed = sc.merge_codex_toml(before, LAUNCH)
    assert changed and text.startswith(before)
    assert tomllib.loads(text)["mcp_servers"]["research"]["command"] == "uvx"


def test_codex_toml_refuses_invalid_and_inline_tables():
    with pytest.raises(sc.SetupRefused):
        sc.merge_codex_toml("[mcp_servers\n", LAUNCH)
    with pytest.raises(sc.SetupRefused):
        sc.merge_codex_toml('mcp_servers = { research = { command = "old" } }\n', LAUNCH)


def test_codex_toml_with_windows_paths_round_trips():
    launch = sc.Launch(r"C:\Users\someone\.local\bin\uvx.exe", ["research-mcp"], "test")
    text, _ = sc.merge_codex_toml("", launch)
    assert tomllib.loads(text)["mcp_servers"]["research"]["command"] == launch.command


# --- config.env -----------------------------------------------------------------------------------------------


def test_merge_env_text_replaces_in_place_and_keeps_other_lines():
    before = "# mine\r\nZOTERO_API_KEY=old\r\nOTHER=1\r\n"
    after = sc.merge_env_text(before, {"ZOTERO_API_KEY": "new", "RESEARCH_MCP_AUTHOR": "A. Writer"})
    assert after == '# mine\r\nZOTERO_API_KEY=new\r\nOTHER=1\r\nRESEARCH_MCP_AUTHOR="A. Writer"\r\n'


def test_merge_env_text_new_file_has_a_header():
    after = sc.merge_env_text(None, {"ZOTERO_LIBRARY_ID": "42"})
    assert after.splitlines()[0].startswith("#") and after.endswith("ZOTERO_LIBRARY_ID=42\n")


# --- backups ------------------------------------------------------------------------------------------------


def test_backup_path_is_timestamped_when_bak_exists(tmp_path):
    path = tmp_path / "mcp.json"
    assert sc.backup_path(path).name == "mcp.json.bak"
    (tmp_path / "mcp.json.bak").write_text("{}")
    import datetime as dt

    stamped = sc.backup_path(path, dt.datetime(2026, 1, 2, 3, 4, 5))
    assert stamped.name == "mcp.json.20260102-030405.bak"


# --- setup end to end -------------------------------------------------------------------------------------------


def test_setup_yes_writes_every_client_and_backs_up_existing(home):
    all_clients(home)
    files = client_files(home)
    files["cursor"].write_text(json.dumps({"mcpServers": {"other": {"command": "x"}}}), encoding="utf-8")
    files["codex"].write_text('[mcp_servers.other]\ncommand = "x"\n', encoding="utf-8")

    code, out, commands = run_setup(home, ["--yes"])
    assert code == 0, out
    assert commands == []  # no claude on PATH: Claude Code goes through ~/.claude.json
    for client, path in files.items():
        assert path.exists(), client
    assert json.loads(files["claude-desktop"].read_text())["mcpServers"]["research"]["args"] == ["research-mcp"]
    assert json.loads(files["claude-code"].read_text())["mcpServers"]["research"]["type"] == "stdio"
    assert json.loads(files["vscode"].read_text())["servers"]["research"]["type"] == "stdio"
    cursor = json.loads(files["cursor"].read_text())["mcpServers"]
    assert set(cursor) == {"other", "research"}
    codex = tomllib.loads(files["codex"].read_text())["mcp_servers"]
    assert set(codex) == {"other", "research"}
    assert (files["cursor"].parent / "mcp.json.bak").exists()
    assert (files["codex"].parent / "config.toml.bak").exists()
    assert not (files["vscode"].parent / "mcp.json.bak").exists()  # new file: nothing to back up

    # Running again changes nothing and makes no second backup.
    code, out, _ = run_setup(home, ["--yes"])
    assert code == 0 and "already has this research entry" in out
    assert not list(files["cursor"].parent.glob("mcp.json.2*.bak"))


def test_setup_refuses_malformed_file_and_leaves_it_alone(home):
    all_clients(home)
    cursor = client_files(home)["cursor"]
    cursor.write_text('{"mcpServers": { // my comment\n}}', encoding="utf-8")
    before = cursor.read_bytes()
    code, out, _ = run_setup(home, ["--client", "cursor"])
    assert code == 1
    assert "NOT changed" in out and '"research"' in out  # the by-hand snippet is shown
    assert cursor.read_bytes() == before
    assert not list(cursor.parent.glob("*.bak"))


def test_setup_client_flag_limits_and_creates(home):
    code, out, _ = run_setup(home, ["--client", "vscode"])  # VS Code not detected: --client still adds it
    files = client_files(home)
    assert code == 0, out
    assert files["vscode"].exists()
    assert not files["cursor"].exists() and not files["claude-desktop"].exists()


def test_setup_claude_code_uses_the_cli_when_present(home):
    which = lambda name: r"C:\bin\claude.exe" if name == "claude" else None  # noqa: E731
    code, out, commands = run_setup(home, ["--client", "claude-code"], which=which)
    assert code == 0, out
    assert commands == [[r"C:\bin\claude.exe", "mcp", "add", "--scope", "user", "research", "--",
                         "uvx", "research-mcp"]]
    assert not client_files(home)["claude-code"].exists()


OLD_CLAUDE_ENTRY = {"type": "stdio", "command": "python", "args": ["-m", "old_research"],
                    "env": {"OLD_TOKEN": "s3cret-old-token"}}


def fake_claude_cli(home, *, fail_second_add=False, fail_add_json=False):
    """A stand-in for the `claude` CLI that keeps user-scope servers in ~/.claude.json like the real one."""
    import subprocess

    path = client_files(home)["claude-code"]
    seen: list[list[str]] = []

    def load() -> dict:
        return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}

    def save(data: dict) -> None:
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def run(cmd):
        seen.append(cmd)
        data = load()
        servers = data.setdefault("mcpServers", {})
        verb = cmd[2]
        if verb == "add":
            if "research" in servers:
                return subprocess.CompletedProcess(cmd, 1, "", "MCP server research already exists in user config")
            if fail_second_add and len([c for c in seen if c[2] == "add"]) > 1:
                return subprocess.CompletedProcess(cmd, 1, "", "error: something went wrong")
            servers["research"] = {"type": "stdio", "command": cmd[7], "args": cmd[8:], "env": {}}
        elif verb == "remove":
            servers.pop("research", None)
        elif verb == "add-json":
            if fail_add_json:
                return subprocess.CompletedProcess(cmd, 1, "", "error: add-json is broken")
            servers[cmd[5]] = json.loads(cmd[6])
        save(data)
        return subprocess.CompletedProcess(cmd, 0, "", "")

    run.seen = seen  # type: ignore[attr-defined]
    run.load = load  # type: ignore[attr-defined]
    path.write_text(json.dumps({"numStartups": 3, "mcpServers": {"research": OLD_CLAUDE_ENTRY}}), encoding="utf-8")
    return run


def CLAUDE_ON_PATH(name):  # noqa: N802
    return "claude" if name == "claude" else None


def test_setup_claude_cli_first_add_backs_up_the_existing_file(home):
    run = fake_claude_cli(home)
    path = client_files(home)["claude-code"]
    path.write_text(json.dumps({"numStartups": 3, "mcpServers": {"other": OLD_CLAUDE_ENTRY}}), encoding="utf-8")
    before = path.read_bytes()
    code, out, _ = run_setup(home, ["--client", "claude-code"], which=CLAUDE_ON_PATH, run=run)
    assert code == 0, out
    assert [cmd[2] for cmd in run.seen] == ["add"]
    assert path.with_name(".claude.json.bak").read_bytes() == before
    assert len(list(path.parent.glob(".claude.json*.bak"))) == 1


def test_setup_claude_cli_already_exists_is_replaced(home):
    run = fake_claude_cli(home)
    code, out, _ = run_setup(home, ["--client", "claude-code"], which=CLAUDE_ON_PATH, run=run)
    assert code == 0, out
    assert [cmd[2] for cmd in run.seen] == ["add", "remove", "add"]
    assert run.load()["mcpServers"]["research"]["command"] == "uvx"
    backup = client_files(home)["claude-code"].with_name(".claude.json.bak")
    assert json.loads(backup.read_text(encoding="utf-8"))["mcpServers"]["research"] == OLD_CLAUDE_ENTRY


def test_setup_claude_cli_failed_readd_puts_the_old_entry_back(home):
    run = fake_claude_cli(home, fail_second_add=True)
    code, out, _ = run_setup(home, ["--client", "claude-code"], which=CLAUDE_ON_PATH, run=run)
    assert code == 1, out
    assert [cmd[2] for cmd in run.seen] == ["add", "remove", "add", "add-json"]
    assert run.load()["mcpServers"]["research"] == OLD_CLAUDE_ENTRY
    assert run.load()["numStartups"] == 3
    assert "NOT changed" in out and "something went wrong" in out and "previous entry was put back" in out
    assert "s3cret-old-token" not in out


def test_setup_claude_cli_failed_readd_and_add_json_writes_the_old_entry_back(home):
    run = fake_claude_cli(home, fail_second_add=True, fail_add_json=True)
    code, out, _ = run_setup(home, ["--client", "claude-code"], which=CLAUDE_ON_PATH, run=run)
    assert code == 1, out
    assert run.load()["mcpServers"]["research"] == OLD_CLAUDE_ENTRY
    assert run.load()["numStartups"] == 3
    assert "written back into" in out and "s3cret-old-token" not in out


def test_setup_claude_cli_readd_timeout_still_puts_the_old_entry_back(home):
    import subprocess

    inner = fake_claude_cli(home)
    adds = []

    def run(cmd):
        if cmd[2] == "add":
            adds.append(cmd)
            if len(adds) == 2:  # the re-add after the remove hangs
                inner.seen.append(cmd)
                raise subprocess.TimeoutExpired(cmd, 120)
        return inner(cmd)

    code, out, _ = run_setup(home, ["--client", "claude-code"], which=CLAUDE_ON_PATH, run=run)
    assert code == 1, out
    assert [cmd[2] for cmd in inner.seen] == ["add", "remove", "add", "add-json"]
    assert inner.load()["mcpServers"]["research"] == OLD_CLAUDE_ENTRY
    assert "TimeoutExpired" in out and "previous entry was put back" in out
    assert "s3cret-old-token" not in out


def test_setup_claude_cli_unreadable_old_entry_is_left_alone(home):
    run = fake_claude_cli(home)
    path = client_files(home)["claude-code"]
    path.write_text("{ not json", encoding="utf-8")
    run_existing = run

    def run(cmd):  # the CLI still says it exists, but the file can't be read
        import subprocess

        run_existing.seen.append(cmd)
        return subprocess.CompletedProcess(cmd, 1, "", "MCP server research already exists in user config")

    code, out, _ = run_setup(home, ["--client", "claude-code"], which=CLAUDE_ON_PATH, run=run)
    assert code == 1, out
    assert [cmd[2] for cmd in run_existing.seen] == ["add"]
    assert path.read_text(encoding="utf-8") == "{ not json"
    assert "could not read" in out and "claude mcp remove --scope user research" in out


def test_setup_dry_run_writes_nothing(home):
    all_clients(home)
    code, out, commands = run_setup(home, ["--dry-run", "--author", "A. Writer"])
    assert code == 0, out
    assert "would create" in out and "Would save to" in out
    assert commands == []
    assert not any(path.exists() for path in client_files(home).values())
    assert not home.config_file.exists()


def test_setup_without_terminal_changes_nothing(home):
    all_clients(home)
    code, out, _ = run_setup(home, [], interactive=False)
    assert code == 0
    assert "Not changed (no terminal to ask in)" in out
    assert not any(path.exists() for path in client_files(home).values())


def test_setup_interactive_asks_per_client(home):
    all_clients(home)
    # key? no; author: Enter; then Claude Desktop yes, Claude Code no, Codex no, Cursor yes, VS Code no.
    code, out, _ = run_setup(home, [], interactive=True, inputs=["n", "", "y", "n", "n", "y", "n"])
    assert code == 0, out
    files = client_files(home)
    assert files["claude-desktop"].exists() and files["cursor"].exists()
    assert not files["codex"].exists() and not files["vscode"].exists() and not files["claude-code"].exists()


def test_print_shows_every_client_and_changes_nothing(home):
    all_clients(home)
    code, out, commands = run_setup(home, ["--print"])
    assert code == 0 and commands == []
    for label in ("Claude Desktop", "Claude Code", "Codex", "Cursor", "VS Code"):
        assert f"# {label}" in out
    assert "[mcp_servers.research]" in out and '"servers"' in out
    assert not any(path.exists() for path in client_files(home).values())


# --- the Zotero key ------------------------------------------------------------------------------------------


def test_key_from_env_goes_to_config_env_only(home, monkeypatch):
    all_clients(home)
    home.config_file.parent.mkdir(parents=True)
    home.config_file.write_text("RESEARCH_MCP_AUTHOR=Someone\n", encoding="utf-8")
    monkeypatch.setenv("ZOTERO_API_KEY", KEY)
    fetch = fake_fetch()
    code, out, _ = run_setup(home, ["--yes", "--zotero-key-from-env"], fetch=fetch)
    assert code == 0, out
    assert KEY not in out
    saved = home.config_file.read_text(encoding="utf-8")
    assert f"ZOTERO_API_KEY={KEY}" in saved and f"ZOTERO_LIBRARY_ID={USER_ID}" in saved
    assert "RESEARCH_MCP_AUTHOR=Someone" in saved
    for path in client_files(home).values():
        assert KEY not in path.read_text(encoding="utf-8")
    key_calls = [url for url, _ in fetch.calls if "api.zotero.org" in url]
    assert key_calls and set(key_calls) == {"https://api.zotero.org/keys/current"}  # read-only calls only


def test_interactive_key_is_hidden_and_validated(home):
    code, out, _ = run_setup(home, ["--client", "cursor"], interactive=True, secret=KEY,
                             inputs=["y", "", ""])  # add key: y; user id: Enter (from the key); author: Enter
    assert code == 0, out
    assert KEY not in out
    assert f"ZOTERO_API_KEY={KEY}" in home.config_file.read_text(encoding="utf-8")
    assert KEY not in client_files(home)["cursor"].read_text(encoding="utf-8")


def test_rejected_key_is_not_saved(home, monkeypatch):
    monkeypatch.setenv("ZOTERO_API_KEY", KEY)
    code, out, _ = run_setup(home, ["--yes", "--zotero-key-from-env"], fetch=fake_fetch(key_status=403))
    assert code == 1
    assert "was not saved" in out and KEY not in out
    assert not home.config_file.exists()


# --- checks and doctor --------------------------------------------------------------------------------------


def run_doctor(home, argv=(), **kwargs):
    out = io.StringIO()
    code = sc.doctor_main(list(argv), home=home, out=out, **kwargs)
    return code, out.getvalue()


def test_doctor_ok_without_key_and_lists_clients(home, monkeypatch):
    monkeypatch.setattr(sc.sys, "platform", "win32")
    monkeypatch.setattr(sc.sys, "version_info", (3, 13, 0, "final", 0))
    all_clients(home)
    run_setup(home, ["--client", "cursor"])
    code, out = run_doctor(home, fetch=fake_fetch(), word_registered=lambda: True)
    assert code == 0, out
    assert "[OK  ] Microsoft Word" in out and "[--  ] Zotero write key" in out
    assert "Cursor:" in out and "(research)" in out


def test_doctor_fails_without_word(home, monkeypatch):
    monkeypatch.setattr(sc.sys, "platform", "win32")
    monkeypatch.setattr(sc.sys, "version_info", (3, 13, 0, "final", 0))
    code, out = run_doctor(home, fetch=fake_fetch(), word_registered=lambda: False)
    assert code == 1
    assert "[FAIL] Microsoft Word" in out and "fix:" in out


def test_doctor_fails_when_zotero_is_not_running(home, monkeypatch):
    monkeypatch.setattr(sc.sys, "platform", "win32")
    monkeypatch.setattr(sc.sys, "version_info", (3, 12, 1, "final", 0))
    code, out = run_doctor(home, fetch=fake_fetch(zotero_local=0), word_registered=lambda: True)
    assert code == 1
    assert "[FAIL] Zotero" in out and "start it" in out and "Allow other applications on this computer" in out


def test_doctor_fails_when_zotero_local_api_is_off(home, monkeypatch):
    monkeypatch.setattr(sc.sys, "platform", "win32")
    monkeypatch.setattr(sc.sys, "version_info", (3, 13, 0, "final", 0))
    code, out = run_doctor(home, fetch=fake_fetch(zotero_local=403), word_registered=lambda: True)
    assert code == 1
    assert "[FAIL] Zotero" in out and "local API is off" in out


def test_setup_writes_configs_but_fails_when_zotero_is_not_running(home):
    all_clients(home)
    code, out, _ = run_setup(home, ["--yes"], fetch=fake_fetch(zotero_local=0))
    assert code == 1, out
    assert all(path.exists() for name, path in client_files(home).items() if name != "claude-code")
    assert "[FAIL] Zotero: Zotero is not running" in out
    assert "The client configs were written" in out and "Zotero only has to be started" in out


def test_doctor_fails_on_rejected_key_and_never_prints_it(home, monkeypatch):
    monkeypatch.setattr(sc.sys, "platform", "win32")
    monkeypatch.setattr(sc.sys, "version_info", (3, 13, 0, "final", 0))
    home.config_file.parent.mkdir(parents=True)
    home.config_file.write_text(f"ZOTERO_API_KEY={KEY}\nZOTERO_LIBRARY_ID={USER_ID}\n", encoding="utf-8")
    code, out = run_doctor(home, fetch=fake_fetch(key_status=403), word_registered=lambda: True)
    assert code == 1
    assert KEY not in out and "keys set: ZOTERO_API_KEY, ZOTERO_LIBRARY_ID" in out
    code, out = run_doctor(home, ["--offline"], fetch=fake_fetch(), word_registered=lambda: True)
    assert code == 0 and KEY not in out


def test_doctor_flags_a_library_id_that_does_not_match_the_key(home, monkeypatch):
    monkeypatch.setattr(sc.sys, "platform", "win32")
    monkeypatch.setattr(sc.sys, "version_info", (3, 13, 0, "final", 0))
    home.config_file.parent.mkdir(parents=True)
    home.config_file.write_text(f"ZOTERO_API_KEY={KEY}\nZOTERO_LIBRARY_ID=999\n", encoding="utf-8")
    code, out = run_doctor(home, fetch=fake_fetch(), word_registered=lambda: True)
    assert code == 1 and f"belongs to user {USER_ID}" in out


def test_checks_fail_off_windows_and_on_old_python(home):
    checks = sc.run_checks(home, fetch=fake_fetch(), word_registered=lambda: None, platform="linux",
                           version=(3, 11))
    status = {check.name: check.status for check in checks}
    assert status["Windows"] == "fail" and status["Python"] == "fail" and status["Microsoft Word"] == "fail"


# --- finding clients and choosing the launch command ---------------------------------------------------------


def test_find_targets_includes_store_app_config(home):
    roaming = home.localappdata / "Packages" / "Claude_abc123" / "LocalCache" / "Roaming" / "Claude"
    roaming.mkdir(parents=True)
    labels = {t.label: t for t in sc.find_targets(home, which=no_which)}
    assert labels["Claude Desktop"].detected
    assert labels["Claude Desktop (Store app)"].path == roaming / "claude_desktop_config.json"


def test_detect_launch_modes(tmp_path):
    cache_prefix = str(tmp_path / "uv" / "cache" / "archive-v0" / "abc")
    uvx = sc.detect_launch("auto", which=lambda n: r"C:\bin\uvx.exe" if n == "uvx" else None,
                           prefix=cache_prefix)
    assert (uvx.command, uvx.args) == (r"C:\bin\uvx.exe", ["research-mcp"])
    scripts = tmp_path / "Scripts"
    scripts.mkdir()
    (scripts / "research-mcp.exe").write_bytes(b"")
    path = sc.detect_launch("auto", which=no_which, prefix=str(tmp_path / "venv"), scripts_dir=str(scripts))
    assert (path.command, path.args) == (str(scripts / "research-mcp.exe"), [])
    fallback = sc.detect_launch("path", which=no_which, prefix=str(tmp_path), scripts_dir=str(tmp_path / "none"))
    assert fallback.args == ["-m", "research_mcp"]


# --- the research-mcp command ----------------------------------------------------------------------------


def test_dispatcher_version_and_unknown(capsys):
    assert dispatcher.main(["--version"]) == 0
    assert "research-mcp" in capsys.readouterr().out
    assert dispatcher.main(["no-such-command"]) == 2


@pytest.mark.parametrize("argv, expected", [
    ([], []),
    (["serve"], []),
    (["serve", "--transport", "stdio"], ["--transport", "stdio"]),
    (["--transport", "stdio"], ["--transport", "stdio"]),
])
def test_dispatcher_runs_the_server(monkeypatch, argv, expected):
    seen = {}

    def fake_serve(args, *, prog=None):
        seen["args"] = args

    import research_mcp.server as server

    monkeypatch.setattr(server, "main", fake_serve)
    assert dispatcher.main(argv) == 0
    assert seen["args"] == expected


def test_dispatcher_routes_setup_doctor_and_rmcp(monkeypatch):
    import research_mcp.cli as cli

    seen = []
    monkeypatch.setattr(sc, "setup_main", lambda args: seen.append(("setup", args)) or 0)
    monkeypatch.setattr(sc, "doctor_main", lambda args: seen.append(("doctor", args)) or 0)
    monkeypatch.setattr(cli, "main", lambda args: seen.append(("rmcp", args)) or 0)
    command = sorted(dispatcher._rmcp_commands(cli.build_parser()))[0]
    assert dispatcher.main(["setup", "--print"]) == 0
    assert dispatcher.main(["doctor"]) == 0
    assert dispatcher.main([command]) == 0
    assert seen == [("setup", ["--print"]), ("doctor", []), ("rmcp", [command])]


# --- review round 1 ------------------------------------------------------------------------------------------


def test_python_314_passes_the_check(home):
    checks = sc.run_checks(home, fetch=fake_fetch(), word_registered=lambda: True, platform="win32",
                           version=(3, 14))
    assert {check.name: check.status for check in checks}["Python"] == "ok"
    checks = sc.run_checks(home, fetch=fake_fetch(), word_registered=lambda: True, platform="win32",
                           version=(3, 15))
    assert {check.name: check.status for check in checks}["Python"] == "fail"


def test_codex_toml_keeps_crlf_line_endings():
    before = 'model = "o4"\r\n\r\n[mcp_servers.other]\r\ncommand = "other"\r\n'
    text, changed = sc.merge_codex_toml(before, LAUNCH)
    assert changed and text.startswith(before)
    assert "\n" not in text.replace("\r\n", "")  # no bare LF anywhere
    assert "\r\n\r\n\r\n" not in text
    replaced, _ = sc.merge_codex_toml(text.replace('"research-mcp"', '"old"'), LAUNCH)
    assert "\n" not in replaced.replace("\r\n", "")
    assert tomllib.loads(replaced)["mcp_servers"]["research"]["args"] == ["research-mcp"]


def test_codex_toml_append_after_blank_line_adds_one_gap():
    text, _ = sc.merge_codex_toml('model = "o4"\n\n', LAUNCH)
    assert "\n\n\n" not in text


@pytest.mark.parametrize("before", [
    'mcp_servers = { research = { command = "old" } }\n',
    '[mcp_servers]\nresearch = { command = "old" }\n',
    'mcp_servers.research.command = "old"\n',
    '[mcp_servers]\nresearch.command = "old"\n',
    '[mcp_servers]\nresearch.command = "old"\n\n[mcp_servers.research.env]\nA = "1"\n',
])
def test_codex_toml_inline_or_dotted_entry_gets_the_clear_refusal(before):
    with pytest.raises(sc.SetupRefused, match="inline table or dotted keys"):
        sc.merge_codex_toml(before, LAUNCH)


def test_dry_run_and_apply_name_dropped_settings(home):
    home.codex_home.mkdir(parents=True)
    path = client_files(home)["codex"]
    path.write_text('[mcp_servers.research]\ncommand = "old"\n\n[mcp_servers.research.env]\nA = "1"\n',
                    encoding="utf-8")
    code, out, _ = run_setup(home, ["--dry-run", "--client", "codex"])
    assert code == 0, out
    assert "its env setting(s) would be dropped, kept in the .bak" in out
    code, out, _ = run_setup(home, ["--yes", "--client", "codex"])
    assert code == 0, out
    assert "dropped from the old research entry: env" in out
    assert "env" not in tomllib.loads(path.read_text(encoding="utf-8"))["mcp_servers"]["research"]
    assert "[mcp_servers.research.env]" in path.with_name("config.toml.bak").read_text(encoding="utf-8")


def test_non_numeric_user_id_is_asked_again_then_accepted(home):
    code, out, _ = run_setup(home, ["--client", "cursor"], interactive=True, secret=KEY,
                             inputs=["y", "Dr X", USER_ID, ""])  # key: y; bad id; good id; author: Enter
    assert code == 0, out
    assert "is not a user ID" in out
    assert f"ZOTERO_LIBRARY_ID={USER_ID}" in home.config_file.read_text(encoding="utf-8")


def test_non_numeric_user_id_from_env_is_refused(home, monkeypatch):
    monkeypatch.setenv("ZOTERO_API_KEY", KEY)
    monkeypatch.setenv("ZOTERO_LIBRARY_ID", "Dr X")
    code, out, _ = run_setup(home, ["--yes", "--zotero-key-from-env"])
    assert code == 1
    assert "user ID must be a number" in out and KEY not in out
    assert not home.config_file.exists()


def test_main_module_runs_as_a_plain_script():
    """The bundle's entry_point is the file itself, so it must not depend on relative imports."""
    import subprocess
    import sys

    script = Path(dispatcher.__file__)
    result = subprocess.run([sys.executable, str(script), "--version"], capture_output=True, text=True,
                            timeout=120, cwd=script.parent)
    assert result.returncode == 0, result.stderr
    assert result.stdout.startswith("research-mcp ")
