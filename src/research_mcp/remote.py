"""``research-mcp chatgpt``: the same MCP server over HTTP for ChatGPT, reachable through a tunnel.

ChatGPT only talks to MCP servers on a public HTTPS address, so this runs the server over streamable HTTP on
``127.0.0.1:<port>`` and, by default, opens a Cloudflare quick tunnel to it (``cloudflared``, no account).

Access control is the URL itself: the server answers only on ``/mcp/<token>`` (32 random bytes, kept in
config.env as ``RESEARCH_MCP_REMOTE_TOKEN``, ``--new-token`` replaces it) and returns 404 for every other path.
The token is never logged: the tool-call log records tool names and arguments only, uvicorn's access log is
off, and the path the MCP app sees is rewritten to ``/mcp``.

Two settings follow from Cloudflare quick tunnels not carrying Server-Sent Events: responses are plain JSON
(``json_response``) and the server is stateless (no GET event stream). Host-header (DNS rebinding) checks are
off because requests arrive with the tunnel's host name; the secret path is the guard instead.

``cloudflared`` comes from PATH, else a copy downloaded once into ``%LOCALAPPDATA%\\research-mcp\\bin`` from
Cloudflare's GitHub release, pinned to one version and checked against its published SHA-256.
"""

from __future__ import annotations

import argparse
import hashlib
import hmac
import logging
import os
from pathlib import Path
import queue
import re
import secrets
import shutil
import signal
import socket
import subprocess
import sys
import threading
import time
from typing import Any, Callable, Iterable

TOKEN_KEY = "RESEARCH_MCP_REMOTE_TOKEN"
MOUNT = "/mcp"

CLOUDFLARED_VERSION = "2026.10.3"
CLOUDFLARED_ASSET = "cloudflared-windows-amd64.exe"
CLOUDFLARED_URL = (
    f"https://github.com/cloudflare/cloudflared/releases/download/{CLOUDFLARED_VERSION}/{CLOUDFLARED_ASSET}"
)
# From the release notes of that version (the same value GitHub lists as the asset's digest).
CLOUDFLARED_SHA256 = "76b2805253d0d6ffacb092a7acfbee99e7af6917835c0c5410e23f3ef5052da2"
CLOUDFLARED_INSTALL = "https://developers.cloudflare.com/cloudflare-one/networks/connectors/cloudflare-tunnel/downloads/"
TUNNEL_TIMEOUT = 60.0  # seconds to wait for the tunnel's address and its first connection

_TUNNEL_URL = re.compile(r"https://(?!api\.)[a-z0-9][a-z0-9-]*\.trycloudflare\.com")
_TUNNEL_READY = "Registered tunnel connection"

CHATGPT_STEPS = """\
In ChatGPT on the web (developer mode; per OpenAI, full MCP with write tools is for Business,
Enterprise and Edu, while Pro gets read/fetch tools only):
  1. Settings > Apps (or Apps & Connectors) > Advanced settings: turn on Developer mode
     (some accounts have it under Settings > Security).
  2. Create a connector (\"Create app\" / the + button): any name, paste the URL above as the
     MCP server URL, Authentication: No authentication, tick \"I trust this application\", Create.
  3. In a chat, pick it from the + menu (Developer mode) so ChatGPT can use the tools.
Tools that change a document or your Zotero library are marked as writes, so ChatGPT may ask before
running them (depending on its permission settings)."""


class RemoteError(Exception):
    """A clear, user-facing reason ``research-mcp chatgpt`` can't run."""


# --- Token -------------------------------------------------------------------------------------------


def new_token() -> str:
    return secrets.token_urlsafe(32)


def remote_token(config_file: Path, *, rotate: bool = False) -> tuple[str, bool]:
    """The access token from ``config_file`` (made and saved when missing, replaced with ``rotate``).

    Returns ``(token, created)``. Only the file counts, not the environment, so ``--new-token`` always takes
    effect. The file is merged (other keys kept), never replaced."""

    from . import config
    from .setup_cmd import _read_text, write_config_env

    current = config.parse(_read_text(config_file) or "").get(TOKEN_KEY, "").strip()
    if current and not rotate:
        return current, False
    token = new_token()
    write_config_env(config_file, {TOKEN_KEY: token})
    return token, True


def _port(value: str) -> int:
    port = int(value)
    if not 0 <= port <= 65535:
        raise argparse.ArgumentTypeError(f"port must be 0-65535, not {value}")
    return port


def connector_url(base: str, token: str) -> str:
    """``<base>/mcp/<token>`` for an https base address (a trailing slash on the base is fine)."""

    base = base.strip().rstrip("/")
    if "?" in base or "#" in base:
        raise RemoteError(f"--url must not contain a query or fragment: {base!r}")
    if not re.match(r"^https://[^/\s]+(/[^\s]*)?$", base):
        raise RemoteError(f"--url must be an https address like https://research.example.com, not {base!r}")
    return f"{base}{MOUNT}/{token}"


def redact(text: str, token: str) -> str:
    return text.replace(token, "<token>") if token else text


# --- ASGI app ----------------------------------------------------------------------------------------


async def _discard_body(receive, limit: int = 1 << 20) -> None:
    seen = 0
    while seen <= limit:
        message = await receive()
        if message.get("type") != "http.request":
            return
        seen += len(message.get("body", b""))
        if not message.get("more_body"):
            return


class SecretPathApp:
    """Pass requests for exactly ``/mcp/<token>`` (or with one trailing slash) to ``app`` as ``/mcp``;
    404 for everything else. Lifespan events go straight through."""

    def __init__(self, app, token: str, *, on_request: Callable[[], None] | None = None):
        self.app = app
        self._paths = (f"{MOUNT}/{token}".encode(), f"{MOUNT}/{token}/".encode())
        self.on_request = on_request
        self.requests = 0
        self.refused = 0

    def _allowed(self, scope) -> bool:
        raw = scope.get("raw_path")
        if raw is None:
            raw = scope.get("path", "").encode("utf-8", "surrogateescape")
        raw = raw.split(b"?", 1)[0]
        # compare_digest on every candidate: no early exit that would time how much of the token matched.
        return any([hmac.compare_digest(raw, path) for path in self._paths])

    async def __call__(self, scope, receive, send):
        if scope["type"] == "lifespan":
            await self.app(scope, receive, send)
            return
        if scope["type"] != "http" or not self._allowed(scope):
            self.refused += 1
            if scope["type"] == "websocket":
                await send({"type": "websocket.close", "code": 1008})
                return
            await _discard_body(receive)  # answering before the body is read can reset the connection
            await send({"type": "http.response.start", "status": 404,
                        "headers": [(b"content-type", b"text/plain; charset=utf-8")]})
            await send({"type": "http.response.body", "body": b"Not Found"})
            return
        if scope.get("method") == "GET":
            # No server-to-client stream in stateless JSON mode (quick tunnels don't carry SSE); the MCP spec
            # lets the server answer 405 here, and an open GET would only hang until shutdown.
            await _discard_body(receive)
            await send({"type": "http.response.start", "status": 405,
                        "headers": [(b"allow", b"POST, DELETE"), (b"content-type", b"text/plain; charset=utf-8")]})
            await send({"type": "http.response.body", "body": b"Method Not Allowed"})
            return
        self.requests += 1
        if self.on_request is not None:
            try:
                self.on_request()
            except Exception:
                pass
        inner = dict(scope)
        inner["path"] = MOUNT
        inner["raw_path"] = MOUNT.encode()
        await self.app(inner, receive, send)


def restrict_to_read_only(mcp_server) -> list[str]:
    """Drop every tool not annotated ``readOnlyHint=True`` from ``mcp_server``; returns the dropped names."""

    dropped = []
    for tool in list(mcp_server._tool_manager.list_tools()):
        annotations = tool.annotations
        if not (annotations is not None and annotations.read_only_hint is True):
            mcp_server.remove_tool(tool.name)
            dropped.append(tool.name)
    return dropped


def build_app(token: str, *, read_only: bool = False, mcp_server=None,
              on_request: Callable[[], None] | None = None) -> SecretPathApp:
    """The server's streamable-HTTP app behind the secret path."""

    from mcp.server.transport_security import TransportSecuritySettings

    if mcp_server is None:
        from .server import server as mcp_server
    if read_only:
        restrict_to_read_only(mcp_server)
    inner = mcp_server.streamable_http_app(
        streamable_http_path=MOUNT,
        json_response=True,
        stateless_http=True,
        transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False),
    )
    return SecretPathApp(inner, token, on_request=on_request)


# --- cloudflared -------------------------------------------------------------------------------------


def bin_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(local) / "research-mcp" / "bin"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _download(url: str, target: Path) -> None:
    from urllib.request import urlopen

    with urlopen(url, timeout=60) as response, open(target, "wb") as out:
        shutil.copyfileobj(response, out, 1 << 20)


def download_cloudflared(dest: Path, *, fetch: Callable[[str, Path], None] = _download,
                         expected: str = CLOUDFLARED_SHA256, say: Callable[[str], None] = print) -> Path:
    """Download the pinned cloudflared to ``dest``; refuse (and keep nothing) when the SHA-256 differs."""

    dest.parent.mkdir(parents=True, exist_ok=True)
    partial = dest.with_name(dest.name + ".download")
    say(f"Downloading cloudflared {CLOUDFLARED_VERSION} (about 55 MB, once) from Cloudflare's GitHub release...")
    try:
        fetch(CLOUDFLARED_URL, partial)
        actual = sha256_file(partial)
        if actual != expected:
            raise RemoteError(
                f"the downloaded cloudflared did not match its published SHA-256 (expected {expected}, got "
                f"{actual}); it was deleted and not run. Try again later, or install cloudflared yourself "
                f"({CLOUDFLARED_INSTALL}) so it is on PATH."
            )
        os.replace(partial, dest)
    except OSError as exc:
        raise RemoteError(
            f"could not download cloudflared ({exc}). Install it yourself ({CLOUDFLARED_INSTALL}) so it is on "
            "PATH, or run your own tunnel and pass its address with --url."
        ) from exc
    finally:
        partial.unlink(missing_ok=True)
    return dest


def find_cloudflared(*, which: Callable[[str], str | None] = shutil.which, folder: Path | None = None,
                     download: Callable[[Path], Path] | None = None) -> Path:
    found = which("cloudflared")
    if found:
        return Path(found)
    folder = folder or bin_dir()
    local = folder / "cloudflared.exe"
    if local.is_file():
        return local
    if sys.platform != "win32" and download is None:
        raise RemoteError(f"cloudflared is not on PATH; install it ({CLOUDFLARED_INSTALL}) or pass --url.")
    return (download or download_cloudflared)(local)


def parse_tunnel_url(line: str) -> str | None:
    match = _TUNNEL_URL.search(line)
    return match.group(0) if match else None


class Tunnel:
    """A running ``cloudflared tunnel --url`` process and its output (kept, last lines only)."""

    def __init__(self, process: subprocess.Popen, lines: "queue.Queue[str | None]"):
        self.process = process
        self.lines = lines
        self.tail: list[str] = []
        self.url: str | None = None
        self._job = None

    def wait_ready(self, timeout: float = TUNNEL_TIMEOUT, *, clock: Callable[[], float] = time.monotonic) -> str:
        """Block until the public address is known and the first connection is registered."""

        deadline = clock() + timeout
        ready = False
        while not (self.url and ready):
            left = deadline - clock()
            if left <= 0:
                raise RemoteError(self._why("did not report a public address within "
                                            f"{int(timeout)} s (blocked network or Cloudflare busy?)"))
            try:
                line = self.lines.get(timeout=min(left, 0.5))
            except queue.Empty:
                if self.process.poll() is not None:
                    raise RemoteError(self._why(f"exited with code {self.process.returncode}"))
                continue
            if line is None:
                self.process.wait(timeout=5)
                raise RemoteError(self._why(f"exited with code {self.process.returncode}"))
            self._keep(line)
            self.url = self.url or parse_tunnel_url(line)
            ready = ready or _TUNNEL_READY in line
        return self.url

    def _keep(self, line: str) -> None:
        self.tail = (self.tail + [line.rstrip()])[-15:]

    def drain(self) -> None:
        """Keep reading output in the background so cloudflared never blocks on a full pipe."""

        def run() -> None:
            while True:
                line = self.lines.get()
                if line is None:
                    return
                self._keep(line)

        threading.Thread(target=run, name="research-mcp-tunnel-drain", daemon=True).start()

    def _why(self, what: str) -> str:
        tail = "\n".join("    " + line for line in self.tail[-8:])
        return f"the Cloudflare tunnel {what}." + (f" Its last output:\n{tail}" if tail else "")

    def alive(self) -> bool:
        return self.process.poll() is None

    def stop(self, timeout: float = 5.0) -> None:
        if self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait(timeout=timeout)


def _read_lines(stream: Iterable[str], lines: "queue.Queue[str | None]") -> None:
    try:
        for line in stream:
            lines.put(line)
    except (OSError, ValueError):
        pass
    finally:
        lines.put(None)


def _kill_with_this_process(process: subprocess.Popen):
    """Windows: put ``process`` in a job that is killed when this process ends, however it ends (a closed
    window, Task Manager), so no cloudflared is left behind. Returns the job handle to keep alive."""

    if sys.platform != "win32":
        return None
    try:
        import win32api
        import win32con
        import win32job

        job = win32job.CreateJobObject(None, "")
        info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
        handle = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE, False, process.pid)
        win32job.AssignProcessToJobObject(job, handle)
        return job
    except Exception as exc:  # the normal Ctrl+C path still stops it; say so rather than skip silently
        print(f"note: cloudflared is stopped on Ctrl+C, but not if this window is killed ({exc})", file=sys.stderr)
        return None


def start_tunnel(exe: Path, port: int, *, popen: Callable[..., subprocess.Popen] = subprocess.Popen) -> Tunnel:
    # An empty config file, so a user's own ~/.cloudflared/config.yml (named-tunnel ingress rules) can't
    # redirect the quick tunnel. cloudflared logs to stderr; its own process group keeps the console's Ctrl+C
    # for us, and we stop it ourselves.
    empty = bin_dir() / "quick-tunnel.yml"
    try:
        empty.parent.mkdir(parents=True, exist_ok=True)
        if not empty.exists():
            empty.write_text("# research-mcp quick tunnel: no settings\n", encoding="utf-8")
        config_args = ["--config", str(empty)]
    except OSError:
        config_args = []
    command = [str(exe), "tunnel", "--no-autoupdate", *config_args, "--url", f"http://127.0.0.1:{port}"]
    flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", creationflags=flags)
    lines: "queue.Queue[str | None]" = queue.Queue()
    threading.Thread(target=_read_lines, args=(process.stdout, lines), name="research-mcp-tunnel-read",
                     daemon=True).start()
    tunnel = Tunnel(process, lines)
    tunnel._job = _kill_with_this_process(process)
    return tunnel


# --- HTTP server --------------------------------------------------------------------------------------


def free_port(host: str = "127.0.0.1") -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((host, 0))
        return sock.getsockname()[1]


class HttpServer:
    """uvicorn in a background thread (so the main thread can watch the tunnel and take Ctrl+C)."""

    def __init__(self, app, port: int, host: str = "127.0.0.1"):
        import uvicorn

        config = uvicorn.Config(app, host=host, port=port, log_level="warning", access_log=False,
                                lifespan="on", timeout_graceful_shutdown=3)
        self.server = uvicorn.Server(config)
        self.thread = threading.Thread(target=self.server.run, name="research-mcp-http", daemon=True)

    def start(self, timeout: float = 20.0) -> None:
        self.thread.start()
        deadline = time.monotonic() + timeout
        while not self.server.started:
            if not self.thread.is_alive():
                raise RemoteError("the HTTP server did not start (is the port already in use?)")
            if time.monotonic() > deadline:
                raise RemoteError("the HTTP server did not start in time")
            time.sleep(0.05)

    def alive(self) -> bool:
        return self.thread.is_alive()

    def stop(self, timeout: float = 10.0) -> None:
        self.server.should_exit = True
        self.thread.join(timeout)


# --- Word warm-up ------------------------------------------------------------------------------------


class WordWarmup:
    """Start the server's own hidden Word worker in the background on the first MCP request, so the first
    citation refresh doesn't spend ChatGPT's tool-call time (about a minute) waiting for Word to start.
    Never blocks a request; with --read-only (no Word-writing tools) it does nothing."""

    def __init__(self, enabled: bool, *, ensure: Callable[[], Any] | None = None):
        self.enabled = enabled and sys.platform == "win32"
        self._ensure = ensure
        self._started = False
        self._lock = threading.Lock()
        self.thread: threading.Thread | None = None

    def __call__(self) -> None:
        if not self.enabled:
            return
        with self._lock:
            if self._started:
                return
            self._started = True
        self.thread = threading.Thread(target=self._run, name="research-mcp-word-warmup", daemon=True)
        self.thread.start()

    def _run(self) -> None:
        try:
            ensure = self._ensure
            if ensure is None:
                from .word_service import ensure_word_worker as ensure
            ensure()
        except Exception as exc:  # the first Word tool call starts it (and reports any problem) instead
            print(f"note: Word did not start in the background ({type(exc).__name__}); "
                  "the first Word tool call will start it.", file=sys.stderr)


# --- Command -----------------------------------------------------------------------------------------


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="research-mcp chatgpt",
        description="Serve the research tools to ChatGPT: HTTP on this PC plus a public tunnel address.",
    )
    parser.add_argument("--url", metavar="HTTPS_BASE",
                        help="your own stable public address that forwards to --port (named Cloudflare tunnel, "
                             "ngrok, reverse proxy); no quick tunnel is started. A path in it must be stripped by "
                             "your proxy: the server itself answers on /mcp/<token>")
    parser.add_argument("--port", type=_port, default=0, help="local port, 1-65535 (default: a free one)")
    parser.add_argument("--no-tunnel", action="store_true", help="local only, for testing (http://127.0.0.1)")
    parser.add_argument("--read-only", action="store_true",
                        help="only the tools that read (inspect, audit, search Zotero); nothing on disk or in the "
                             "Zotero library can be changed")
    parser.add_argument("--new-token", action="store_true",
                        help="replace the secret part of the URL (old connector URLs stop working)")
    return parser


def _copy_to_clipboard(text: str, *, run: Callable[..., Any] = subprocess.run) -> bool:
    if sys.platform != "win32":
        return False
    try:
        result = run(["clip"], input=text.encode("ascii"), capture_output=True, timeout=10)
        return result.returncode == 0
    except (OSError, subprocess.SubprocessError, UnicodeEncodeError):
        return False


def _on_break(signum, frame):  # Ctrl+Break (and CTRL_BREAK_EVENT from scripts) stops like Ctrl+C
    raise KeyboardInterrupt


def chatgpt_main(argv: list[str] | None = None, *, out=None, err=None,
                 find: Callable[[], Path] | None = None,
                 tunnel_factory: Callable[[Path, int], Tunnel] | None = None,
                 server_factory: Callable[[Any, int], HttpServer] | None = None,
                 clipboard: Callable[[str], bool] = _copy_to_clipboard,
                 poll: float = 0.5) -> int:
    args = _parser().parse_args(argv)
    out = out or sys.stdout
    err = err or sys.stderr
    from . import config
    from .setup_cmd import _quiet_encoding_errors

    _quiet_encoding_errors()

    def say(text: str = "") -> None:
        print(text, file=out, flush=True)

    if args.url and args.no_tunnel:
        print("research-mcp chatgpt: --url and --no-tunnel don't go together", file=err)
        return 2
    if args.url and not args.port:
        print("research-mcp chatgpt: --url needs --port (the local port your tunnel forwards to)", file=err)
        return 2

    if hasattr(signal, "SIGBREAK"):
        try:
            signal.signal(signal.SIGBREAK, _on_break)
        except ValueError:  # not the main thread (tests)
            pass
    tunnel: Tunnel | None = None
    http: HttpServer | None = None
    token = ""
    try:
        token, created = remote_token(config.config_path(), rotate=args.new_token)
        if args.url:
            connector_url(args.url, token)  # a bad --url fails before anything starts
        port = args.port or free_port()
        warm = WordWarmup(enabled=not args.read_only)
        first = {"seen": False}

        def on_request() -> None:
            if not first["seen"]:
                first["seen"] = True
                say("A client connected.")
            warm()

        app = build_app(token, read_only=args.read_only, on_request=on_request)
        # The SDK logs a line per HTTP request at INFO ("Terminating session: None"); keep the window readable.
        logging.getLogger("mcp").setLevel(logging.WARNING)
        http = (server_factory or HttpServer)(app, port)
        http.start()

        if args.no_tunnel:
            url = f"http://127.0.0.1:{port}{MOUNT}/{token}"
            where = "on this PC only (--no-tunnel): ChatGPT can't reach it; for testing with a local client."
        elif args.url:
            url = connector_url(args.url, token)
            where = f"through your own address; it must forward to http://127.0.0.1:{port}."
        else:
            say("Starting a Cloudflare quick tunnel...")
            exe = (find or find_cloudflared)()
            tunnel = (tunnel_factory or start_tunnel)(exe, port)
            base = tunnel.wait_ready()
            tunnel.drain()
            url = connector_url(base, token)
            where = ("through a Cloudflare quick tunnel; a new address can take a minute to start working, and it "
                     "changes every start (use --url for a fixed one).")

        copied = clipboard(url)
        say()
        say("research-mcp is ready for ChatGPT" + (" (read-only tools)" if args.read_only else "") + ".")
        say("Connector URL" + (" (copied to the clipboard)" if copied else "") + ":")
        say(f"  {url}")
        say(f"Reachable {where}")
        if created and not args.new_token:
            say(f"(A new access token was saved to {config.config_path()}.)")
        elif args.new_token:
            say("(New access token saved; connectors with the old URL no longer work.)")
        say()
        say(CHATGPT_STEPS)
        say()
        say("The URL is the password: anyone who has it can use these tools while this window is open.")
        say("Keep this window open while you use ChatGPT. Ctrl+C stops it.")

        while True:
            if tunnel is not None and not tunnel.alive():
                print("\n" + redact(tunnel._why(f"stopped (exit code {tunnel.process.returncode})"), token)
                      + "\nChatGPT can no longer reach the tools; run `research-mcp chatgpt` again.", file=err)
                return 1
            if not http.alive():
                print("\nThe local HTTP server stopped unexpectedly.", file=err)
                return 1
            time.sleep(poll)
    except KeyboardInterrupt:
        say("\nStopping...")
        return 0
    except RemoteError as exc:
        print(f"research-mcp chatgpt: {redact(str(exc), token)}", file=err)
        return 1
    finally:
        if tunnel is not None:
            tunnel.stop()
        if http is not None:
            http.stop()
