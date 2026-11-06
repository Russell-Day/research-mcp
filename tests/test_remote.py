"""``research-mcp chatgpt`` (remote.py): secret-path routing, token storage, the cloudflared download and
URL parsing, ``--url``, read-only filtering, and clean shutdown with a stand-in tunnel process.

Nothing here downloads anything, opens a real tunnel, or starts Word. HTTP tests run a real uvicorn server on
127.0.0.1 (a free port) around a small test MCPServer, and talk to it with the MCP client.
"""

from __future__ import annotations

import asyncio
import hashlib
import io
import json
from pathlib import Path
import subprocess
import sys
import textwrap
import time
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pytest
from mcp.client import Client
from mcp.server import MCPServer
from mcp.types import ToolAnnotations

from research_mcp import __main__ as dispatcher
from research_mcp import config
from research_mcp import remote

TOKEN = "T0k3n-For_Tests-only-0123456789abcdefghijkl"

SAMPLE_OUTPUT = """\
2026-11-06T07:23:00Z INF Thank you for trying Cloudflare Tunnel. ... https://www.cloudflare.com/website-terms/ ...
2026-11-06T07:23:00Z INF Requesting new quick Tunnel on trycloudflare.com...
2026-11-06T07:23:03Z INF +--------------------------------------------------------------------------------------------+
2026-11-06T07:23:03Z INF |  Your quick Tunnel has been created! Visit it at (it may take some time to be reachable):  |
2026-11-06T07:23:03Z INF |  https://seven-random-words-here.trycloudflare.com                                       |
2026-11-06T07:23:03Z INF +--------------------------------------------------------------------------------------------+
2026-11-06T07:23:03Z INF Version 2026.10.3 (Checksum 1d356dac)
2026-11-06T07:23:03Z INF Starting metrics server on 127.0.0.1:20242/metrics
2026-11-06T07:23:04Z INF Registered tunnel connection connIndex=0 connection=de98 event=0 ip=198.41.192.67 location=ord15 protocol=quic
"""
SAMPLE_HOST = "https://seven-random-words-here.trycloudflare.com"


def _test_server() -> MCPServer:
    srv = MCPServer("remote-test")

    @srv.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def look(text: str) -> str:
        """A read-only tool."""
        return f"saw {text}"

    @srv.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))
    def change(text: str) -> str:
        """A write tool."""
        return f"changed {text}"

    @srv.tool()
    def unannotated() -> str:
        """No annotations: counts as a write."""
        return "x"

    return srv


@pytest.fixture(autouse=True)
def keep_server_tools():
    """``--read-only`` removes tools from the server object (fine for a one-shot process); put them back."""
    from research_mcp.server import server

    tools = dict(server._tool_manager._tools)
    yield
    server._tool_manager._tools.clear()
    server._tool_manager._tools.update(tools)


@pytest.fixture
def running():
    """Start a SecretPathApp around a fresh test server on a free port; yields a factory."""

    started: list[remote.HttpServer] = []

    def start(*, read_only: bool = False, on_request=None):
        app = remote.build_app(TOKEN, read_only=read_only, mcp_server=_test_server(), on_request=on_request)
        port = remote.free_port()
        http = remote.HttpServer(app, port)
        http.start()
        started.append(http)
        return app, f"http://127.0.0.1:{port}"

    yield start
    for http in started:
        http.stop()


def _status(url: str, method: str = "POST") -> int:
    body = json.dumps({"jsonrpc": "2.0", "id": 1, "method": "ping"}).encode() if method == "POST" else None
    request = Request(url, data=body, method=method, headers={
        "content-type": "application/json", "accept": "application/json, text/event-stream"})
    try:
        with urlopen(request, timeout=10) as response:
            return response.status
    except HTTPError as exc:
        return exc.code


async def _session(url: str, call: tuple[str, dict] | None = None):
    async with Client(url) as client:
        tools = await client.list_tools()
        result = await client.call_tool(*call) if call else None
        return client, sorted(t.name for t in tools.tools), result


# --- Routing ---------------------------------------------------------------------------------------


def test_right_path_answers_mcp_and_lists_tools(running):
    _, base = running()
    client, names, result = asyncio.run(_session(f"{base}/mcp/{TOKEN}", ("look", {"text": "hi"})))
    assert names == ["change", "look", "unannotated"]
    assert result.is_error is False
    assert "saw hi" in json.dumps(result.model_dump(mode="json"))


def test_trailing_slash_and_query_are_accepted(running):
    _, base = running()
    assert asyncio.run(_session(f"{base}/mcp/{TOKEN}/"))[1] == ["change", "look", "unannotated"]
    assert _status(f"{base}/mcp/{TOKEN}?x=1") != 404


@pytest.mark.parametrize("path", [
    "/", "/mcp", "/mcp/", f"/mcp/{TOKEN}x", f"/mcp/{TOKEN[:-1]}", f"/mcp/{TOKEN}//", f"/mcp/{TOKEN}/x",
    f"/MCP/{TOKEN}", f"/x/mcp/{TOKEN}", "/mcp/wrong", f"/{TOKEN}", f"/mcp%2F{TOKEN}",
])
def test_every_other_path_is_404(running, path):
    app, base = running()
    assert _status(base + path) == 404
    assert _status(base + path, "GET") == 404
    assert app.requests == 0


def _scope(raw: bytes) -> dict:
    return {"type": "http", "method": "POST", "path": raw.decode("latin-1"), "raw_path": raw,
            "query_string": b"", "headers": []}


def _call(app: remote.SecretPathApp, scope: dict) -> tuple[int | None, list[dict]]:
    sent: list[dict] = []

    async def receive():
        return {"type": "http.request", "body": b"", "more_body": False}

    async def send(message):
        sent.append(message)

    asyncio.run(app(scope, receive, send))
    status = next((m["status"] for m in sent if m["type"] == "http.response.start"), None)
    return status, sent


@pytest.mark.parametrize("raw", [
    f"/mcp/{TOKEN}/../".encode(), f"/mcp/{TOKEN}/./".encode(), f"/mcp/x/../{TOKEN}".encode(),
    f"/mcp/{TOKEN}%00".encode(), f"/mcp/{TOKEN}\x00".encode(), b"", b"/mcp/",
])
def test_prefix_tricks_never_reach_the_app(raw):
    seen: list[dict] = []

    async def inner(scope, receive, send):
        seen.append(scope)

    app = remote.SecretPathApp(inner, TOKEN)
    status, _ = _call(app, _scope(raw))
    assert status == 404 and seen == [] and app.refused == 1


def test_inner_app_sees_mcp_not_the_token():
    seen: list[dict] = []

    async def inner(scope, receive, send):
        seen.append(scope)

    calls = []
    app = remote.SecretPathApp(inner, TOKEN, on_request=lambda: calls.append(1))
    _call(app, _scope(f"/mcp/{TOKEN}".encode()))
    assert seen[0]["path"] == "/mcp" and seen[0]["raw_path"] == b"/mcp"
    assert TOKEN not in json.dumps({k: v for k, v in seen[0].items() if isinstance(v, str)})
    assert calls == [1] and app.requests == 1


def test_websocket_is_refused():
    sent: list[dict] = []

    async def inner(scope, receive, send):
        raise AssertionError("must not be reached")

    async def send(message):
        sent.append(message)

    async def receive():
        return {"type": "websocket.connect"}

    app = remote.SecretPathApp(inner, TOKEN)
    asyncio.run(app({"type": "websocket", "path": f"/mcp/{TOKEN}", "raw_path": f"/mcp/{TOKEN}".encode()},
                    receive, send))
    assert sent == [{"type": "websocket.close", "code": 1008}]


def test_tunnel_host_header_is_accepted(running):
    """Requests through a tunnel carry the tunnel's host name; the server must not refuse them (421)."""
    _, base = running()
    request = Request(f"{base}/mcp/{TOKEN}", method="POST",
                      data=json.dumps({"jsonrpc": "2.0", "id": 1, "method": "ping"}).encode(),
                      headers={"content-type": "application/json", "accept": "application/json, text/event-stream",
                               "host": "seven-random-words-here.trycloudflare.com"})
    with urlopen(request, timeout=10) as response:
        assert response.status == 200
        assert response.headers["content-type"].startswith("application/json")  # no SSE through quick tunnels


def test_wire_annotations_are_camel_case(running):
    """ChatGPT decides read vs write from ``readOnlyHint`` in tools/list; check the JSON on the wire."""
    _, base = running()
    request = Request(f"{base}/mcp/{TOKEN}", method="POST",
                      data=json.dumps({"jsonrpc": "2.0", "id": 1, "method": "tools/list"}).encode(),
                      headers={"content-type": "application/json", "accept": "application/json, text/event-stream",
                               "mcp-protocol-version": "2025-06-18"})
    with urlopen(request, timeout=10) as response:
        tools = {t["name"]: t for t in json.loads(response.read())["result"]["tools"]}
    assert tools["look"]["annotations"]["readOnlyHint"] is True
    assert tools["change"]["annotations"]["readOnlyHint"] is False


# --- Read-only ---------------------------------------------------------------------------------------


def test_read_only_keeps_only_read_tools(running):
    _, base = running(read_only=True)
    assert asyncio.run(_session(f"{base}/mcp/{TOKEN}"))[1] == ["look"]
    try:
        result = asyncio.run(_session(f"{base}/mcp/{TOKEN}", ("change", {"text": "x"})))[2]
    except Exception as exc:  # an unknown tool is a protocol error
        assert "change" in str(exc) or "Unknown" in str(exc)
    else:
        assert result.is_error is True


def test_restrict_to_read_only_reports_dropped():
    srv = _test_server()
    assert sorted(remote.restrict_to_read_only(srv)) == ["change", "unannotated"]
    assert [t.name for t in srv._tool_manager.list_tools()] == ["look"]


def test_real_server_read_tools_are_annotated():
    from research_mcp.server import server

    read = {t.name for t in server._tool_manager.list_tools()
            if t.annotations is not None and t.annotations.read_only_hint is True}
    assert {"inspect_document", "audit_document"} <= read
    assert "edit_document" not in read


# --- Token ---------------------------------------------------------------------------------------


def test_token_is_created_saved_reused_and_rotated(tmp_path):
    cfg = tmp_path / "research-mcp" / "config.env"
    cfg.parent.mkdir()
    cfg.write_text("# mine\nZOTERO_LIBRARY_ID=42\nRESEARCH_MCP_AUTHOR=\"A B\"\n", encoding="utf-8")
    first, created = remote.remote_token(cfg)
    assert created and len(first) >= 43
    again, created_again = remote.remote_token(cfg)
    assert (again, created_again) == (first, False)
    rotated, created_rot = remote.remote_token(cfg, rotate=True)
    assert created_rot and rotated != first
    values = config.parse(cfg.read_text(encoding="utf-8"))
    assert values == {"ZOTERO_LIBRARY_ID": "42", "RESEARCH_MCP_AUTHOR": "A B", remote.TOKEN_KEY: rotated}
    assert cfg.read_text(encoding="utf-8").startswith("# mine\n")
    assert cfg.read_text(encoding="utf-8").count(remote.TOKEN_KEY) == 1


def test_token_ignores_environment(tmp_path, monkeypatch):
    monkeypatch.setenv(remote.TOKEN_KEY, "from-env")
    token, created = remote.remote_token(tmp_path / "config.env")
    assert created and token != "from-env"


# --- URL composition --------------------------------------------------------------------------------


@pytest.mark.parametrize("base", ["https://research.example.com", "https://research.example.com/",
                                  "  https://research.example.com  "])
def test_connector_url(base):
    assert remote.connector_url(base, TOKEN) == f"https://research.example.com/mcp/{TOKEN}"


def test_connector_url_keeps_a_base_path():
    assert remote.connector_url("https://ex.com/proxy/", "t") == "https://ex.com/proxy/mcp/t"


@pytest.mark.parametrize("bad", ["research.example.com", "ftp://x.com", "https://", "https://a b.com", "",
                                 "http://research.example.com"])
def test_connector_url_refuses_non_urls(bad):
    with pytest.raises(remote.RemoteError):
        remote.connector_url(bad, TOKEN)


# --- cloudflared --------------------------------------------------------------------------------------


def test_parse_tunnel_url_from_sample_output():
    found = [remote.parse_tunnel_url(line) for line in SAMPLE_OUTPUT.splitlines()]
    assert [f for f in found if f] == [SAMPLE_HOST]
    assert remote.parse_tunnel_url("see https://api.trycloudflare.com/tunnel") is None


def test_download_refuses_checksum_mismatch(tmp_path):
    dest = tmp_path / "bin" / "cloudflared.exe"

    def fetch(url, target):
        assert url == remote.CLOUDFLARED_URL
        Path(target).write_bytes(b"not the real binary")

    with pytest.raises(remote.RemoteError, match="SHA-256"):
        remote.download_cloudflared(dest, fetch=fetch, say=lambda _: None)
    assert list((tmp_path / "bin").iterdir()) == []


def test_download_installs_matching_file(tmp_path):
    payload = b"pretend cloudflared"
    dest = tmp_path / "bin" / "cloudflared.exe"
    remote.download_cloudflared(dest, fetch=lambda url, target: Path(target).write_bytes(payload),
                                expected=hashlib.sha256(payload).hexdigest(), say=lambda _: None)
    assert dest.read_bytes() == payload
    assert [p.name for p in dest.parent.iterdir()] == ["cloudflared.exe"]


def test_download_network_error_is_clear(tmp_path):
    def fetch(url, target):
        raise OSError("no network")

    with pytest.raises(remote.RemoteError, match="--url"):
        remote.download_cloudflared(tmp_path / "cloudflared.exe", fetch=fetch, say=lambda _: None)


def test_find_cloudflared_prefers_path_then_local_copy(tmp_path):
    assert remote.find_cloudflared(which=lambda _: r"C:\tools\cloudflared.exe") == Path(r"C:\tools\cloudflared.exe")
    local = tmp_path / "cloudflared.exe"
    local.write_bytes(b"x")
    assert remote.find_cloudflared(which=lambda _: None, folder=tmp_path) == local
    got = remote.find_cloudflared(which=lambda _: None, folder=tmp_path / "empty",
                                  download=lambda dest: dest)
    assert got == tmp_path / "empty" / "cloudflared.exe"


# --- Whole command with a stand-in tunnel ---------------------------------------------------------------

FAKE_CLOUDFLARED = textwrap.dedent("""
    import sys, time
    sample = sys.argv[1]
    for line in open(sample, encoding="utf-8"):
        print(line, end="", flush=True)
    if sys.argv[2] == "exit":
        sys.exit(3)
    while True:
        time.sleep(0.2)
""")


@pytest.fixture
def fake_tunnel(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    script = tmp_path / "fake_cloudflared.py"
    script.write_text(FAKE_CLOUDFLARED, encoding="utf-8")
    sample = tmp_path / "sample.txt"
    sample.write_text(SAMPLE_OUTPUT, encoding="utf-8")
    made: list[remote.Tunnel] = []
    commands: list[list[str]] = []

    def factory(mode: str):
        def popen(command, **kwargs):
            commands.append(command)
            return subprocess.Popen([sys.executable, str(script), str(sample), mode], **kwargs)

        def make(exe, port):
            tunnel = remote.start_tunnel(exe, port, popen=popen)
            made.append(tunnel)
            return tunnel
        return make

    return factory, made, commands


class _StopAfter:
    """Stands in for HttpServer.alive(): after n checks it raises KeyboardInterrupt, like Ctrl+C."""

    def __init__(self, server, n: int):
        self.server, self.n = server, n

    def __call__(self):
        self.n -= 1
        if self.n < 0:
            raise KeyboardInterrupt
        return self.server.thread.is_alive()


def _server_factory(stop_after: int | None):
    made: list[remote.HttpServer] = []

    def make(app, port):
        http = remote.HttpServer(app, port)
        if stop_after is not None:
            http.alive = _StopAfter(http, stop_after)
        made.append(http)
        return http
    return make, made


def _run(argv, **kwargs):
    out, err = io.StringIO(), io.StringIO()
    code = remote.chatgpt_main(argv, out=out, err=err, clipboard=lambda url: True, poll=0.05, **kwargs)
    return code, out.getvalue(), err.getvalue()


def test_ctrl_c_stops_tunnel_and_server(fake_tunnel):
    factory, made, commands = fake_tunnel
    servers, http_made = _server_factory(stop_after=3)
    code, out, err = _run(["--read-only"], find=lambda: Path("cloudflared.exe"), tunnel_factory=factory("stay"),
                          server_factory=servers)
    assert code == 0, err
    token = config.parse(Path(config.config_path()).read_text(encoding="utf-8"))[remote.TOKEN_KEY]
    assert f"{SAMPLE_HOST}/mcp/{token}" in out
    assert out.count(token) == 1  # only in the URL line
    assert "The URL is the password" in out and "No authentication" in out
    assert commands[0][:3] == ["cloudflared.exe", "tunnel", "--no-autoupdate"]
    assert commands[0][-2] == "--url" and commands[0][-1].startswith("http://127.0.0.1:")
    assert made[0].process.poll() is not None  # the tunnel process was stopped
    assert not http_made[0].thread.is_alive()


def test_tunnel_dying_exits_nonzero(fake_tunnel):
    factory, made, _ = fake_tunnel
    servers, http_made = _server_factory(stop_after=None)
    code, out, err = _run(["--read-only"], find=lambda: Path("cloudflared.exe"), tunnel_factory=factory("exit"),
                          server_factory=servers)
    assert code == 1
    assert "tunnel" in err and "run `research-mcp chatgpt` again" in err
    assert made[0].process.poll() == 3
    assert not http_made[0].thread.is_alive()


def test_tunnel_timeout_is_clear():
    class Never:
        returncode = None

        def poll(self):
            return None

    import queue

    clock = iter([0.0, 0.0, 100.0, 100.0])
    tunnel = remote.Tunnel(Never(), queue.Queue())
    with pytest.raises(remote.RemoteError, match="did not report a public address"):
        tunnel.wait_ready(timeout=10, clock=lambda: next(clock))


def test_no_tunnel_serves_locally_and_answers_mcp(tmp_path):
    servers, http_made = _server_factory(stop_after=None)
    result: dict = {}

    def check_then_stop(app, port):
        http = servers(app, port)
        original = http.alive

        def alive():
            if "names" not in result:
                token = config.parse(Path(config.config_path()).read_text(encoding="utf-8"))[remote.TOKEN_KEY]
                url = f"http://127.0.0.1:{port}/mcp/{token}"
                result["names"] = asyncio.run(_session(url))[1]
                result["wrong"] = _status(f"http://127.0.0.1:{port}/mcp/{token}x")
                raise KeyboardInterrupt
            return original()
        http.alive = alive
        return http

    # the real server's tool list, read-only: nothing here starts Word
    code, out, err = _run(["--no-tunnel", "--read-only"], server_factory=check_then_stop)
    assert code == 0, err
    assert "inspect_document" in result["names"] and "edit_document" not in result["names"]
    assert result["wrong"] == 404
    assert "http://127.0.0.1:" in out and "on this PC only" in out


def test_url_mode_prints_composed_url_and_needs_port():
    code, _, err = _run(["--url", "https://research.example.com"])
    assert code == 2 and "--port" in err
    code, _, err = _run(["--url", "https://x.example.com", "--no-tunnel", "--port", "1"])
    assert code == 2
    code, _, err = _run(["--url", "not a url", "--port", "1"])
    assert code == 1 and "--url must be" in err
    servers, _ = _server_factory(stop_after=1)
    port = remote.free_port()
    code, out, err = _run(["--url", "https://research.example.com/", "--port", str(port), "--read-only"],
                          server_factory=servers)
    assert code == 0, err
    token = config.parse(Path(config.config_path()).read_text(encoding="utf-8"))[remote.TOKEN_KEY]
    assert f"https://research.example.com/mcp/{token}" in out
    assert f"http://127.0.0.1:{port}" in out


def test_new_token_rotates(tmp_path):
    servers, _ = _server_factory(stop_after=0)
    _run(["--no-tunnel", "--read-only"], server_factory=servers)
    before = config.parse(Path(config.config_path()).read_text(encoding="utf-8"))[remote.TOKEN_KEY]
    servers, _ = _server_factory(stop_after=0)
    code, out, _ = _run(["--no-tunnel", "--read-only", "--new-token"], server_factory=servers)
    after = config.parse(Path(config.config_path()).read_text(encoding="utf-8"))[remote.TOKEN_KEY]
    assert code == 0 and after != before and before not in out and after in out
    assert "old URL no longer work" in out


def test_errors_never_show_the_token(fake_tunnel):
    factory, _, _ = fake_tunnel
    servers, _ = _server_factory(stop_after=None)

    def boom():
        token = config.parse(Path(config.config_path()).read_text(encoding="utf-8"))[remote.TOKEN_KEY]
        raise remote.RemoteError(f"failed near /mcp/{token}")

    code, out, err = _run(["--read-only"], find=boom, server_factory=servers)
    token = config.parse(Path(config.config_path()).read_text(encoding="utf-8"))[remote.TOKEN_KEY]
    assert code == 1 and token not in err and token not in out and "<token>" in err


# --- Word warm-up -------------------------------------------------------------------------------------


def test_word_warmup_runs_once_in_background(monkeypatch):
    monkeypatch.setattr(remote.sys, "platform", "win32")
    calls = []
    warm = remote.WordWarmup(True, ensure=lambda: (time.sleep(0.05), calls.append(1)))
    for _ in range(5):
        warm()
    warm.thread.join(5)
    assert calls == [1]


def test_word_warmup_off_for_read_only():
    calls = []
    warm = remote.WordWarmup(False, ensure=lambda: calls.append(1))
    warm()
    assert warm.thread is None and calls == []


def test_word_warmup_failure_is_reported_not_raised(monkeypatch, capsys):
    monkeypatch.setattr(remote.sys, "platform", "win32")

    def fail():
        raise RuntimeError("no Word")

    warm = remote.WordWarmup(True, ensure=fail)
    warm()
    warm.thread.join(5)
    assert "first Word tool call will start it" in capsys.readouterr().err


# --- Dispatcher -----------------------------------------------------------------------------------------


def test_dispatcher_knows_chatgpt(capsys):
    with pytest.raises(SystemExit) as exc:
        dispatcher.main(["chatgpt", "--help"])
    assert exc.value.code == 0
    assert "--read-only" in capsys.readouterr().out
    assert dispatcher.main(["--help"]) == 0
    assert "chatgpt" in capsys.readouterr().out


def test_clipboard_uses_clip_on_windows(monkeypatch):
    calls = []

    def run(command, **kwargs):
        calls.append((command, kwargs["input"]))
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(remote.sys, "platform", "win32")
    assert remote._copy_to_clipboard("https://x.example/mcp/t", run=run) is True
    assert calls == [(["clip"], b"https://x.example/mcp/t")]
    monkeypatch.setattr(remote.sys, "platform", "linux")
    assert remote._copy_to_clipboard("u", run=run) is False


@pytest.mark.parametrize("bad", ["https://ex.com/?q=1", "https://ex.com/#x", "https://ex.com/base?x"])
def test_connector_url_refuses_query_or_fragment(bad):
    with pytest.raises(remote.RemoteError):
        remote.connector_url(bad, TOKEN)


@pytest.mark.parametrize("value", ["-1", "65536", "70000"])
def test_port_out_of_range_is_an_argparse_error(value):
    with pytest.raises(SystemExit):
        remote._parser().parse_args(["--port", value])


def test_get_on_the_secret_path_is_405_not_a_hanging_stream():
    import asyncio

    called = []

    async def inner(scope, receive, send):
        called.append(scope)

    app = remote.SecretPathApp(inner, TOKEN)
    sent = []

    async def receive():
        return {"type": "http.request", "body": b"", "more_body": False}

    async def send(message):
        sent.append(message)

    raw = f"/mcp/{TOKEN}".encode()
    asyncio.run(app({"type": "http", "method": "GET", "path": raw.decode(), "raw_path": raw}, receive, send))
    assert sent[0]["status"] == 405 and not called
