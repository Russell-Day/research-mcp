# research-mcp

**Needs Windows, the desktop Microsoft Word, and the Zotero desktop app.**

An MCP server that lets an AI assistant (Claude, ChatGPT, Codex, Cursor, VS Code…) edit your Word
documents without breaking them, and cite from your Zotero library with live Zotero citations. Every
edit is checked before it is saved, keeps your formatting, and can be undone in one call.

## Install

### Claude Desktop (one click)

1. Download `research-mcp-<version>.mcpb` from the [latest release](https://github.com/Russell-Day/research-mcp/releases/latest).
2. Double-click it. Claude Desktop opens and asks to install it. Click **Install**.
3. Optional: fill in the Zotero write key (see [Zotero write key](#zotero-write-key)).

Claude Desktop sets up Python and everything else itself.

### Everyone else (one command)

Install [uv](https://docs.astral.sh/uv/) if you don't have it (PowerShell):

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Open a new terminal, then run:

```powershell
uvx research-mcp setup
```

Setup checks Word and Zotero, asks for an optional Zotero write key, and offers to add the server to
each MCP client it finds: Claude Desktop, Claude Code, Codex, Cursor and VS Code. It backs up every
file it changes (`<file>.bak`) and keeps your other servers. Restart the client afterwards.

Useful options:

| Command | What it does |
|---|---|
| `uvx research-mcp setup --print` | print the config entry for every client, change nothing |
| `uvx research-mcp setup --dry-run` | show what setup would change, change nothing |
| `uvx research-mcp setup --yes` | no questions: add to every client found |
| `uvx research-mcp setup --client cursor` | only this client (`claude-desktop`, `claude-code`, `codex`, `cursor`, `vscode`) |

By hand, the server entry is just `uvx research-mcp` (stdio). For example, in a `mcpServers` file:

```json
{ "mcpServers": { "research": { "command": "uvx", "args": ["research-mcp"] } } }
```

### ChatGPT

ChatGPT only reaches MCP servers on a public HTTPS address, so it has its own command. With
[uv](https://docs.astral.sh/uv/) installed (see above), run:

```powershell
uvx research-mcp chatgpt
```

It starts the server on your PC, opens a free [Cloudflare quick tunnel](https://developers.cloudflare.com/cloudflare-one/networks/connectors/cloudflare-tunnel/do-more-with-tunnels/trycloudflare/)
to it (no account; `cloudflared` is downloaded once and checked against its published SHA-256), and prints
a connector URL like `https://<words>.trycloudflare.com/mcp/<secret>`, also copied to the clipboard. Then,
in ChatGPT on the web with developer mode (per [OpenAI](https://help.openai.com/en/articles/12584461-connecting-custom-mcp-servers-in-chatgpt),
full MCP with write tools is for Business and Enterprise/Edu; Pro gets read/fetch tools only):

1. Settings > Apps (or Apps & Connectors) > Advanced settings: turn on **Developer mode**
   (some accounts have it under Settings > Security).
2. Create a connector (**Create app** or **+**): any name, paste the URL, Authentication **No authentication**,
   tick "I trust this application".
3. In a chat, pick it from the **+** menu. Tools that change a document or your library are marked as writes,
   so ChatGPT may ask before running them, depending on its permission settings.

- **The URL is the password.** Anyone who has it can use the tools while the window is open. Don't share it;
  `--new-token` makes a new one (old connectors stop working).
- **Keep the window open** while you use ChatGPT. **Ctrl+C** (or closing the window) stops the server and the tunnel.
- The quick-tunnel address **changes every start**, so you paste the new URL into the connector each time.
  For a fixed address, run your own tunnel or reverse proxy to a fixed port and pass it:
  `uvx research-mcp chatgpt --port 8765 --url https://research.example.com` (https only: the URL holds the
  password).
- `--read-only` exposes only the tools that read (inspect, audit, search Zotero; no page renders, since those
  write image files): a good start. Nothing on disk or in Zotero can be changed, but the tools can still read any
  .docx on this PC (set `RESEARCH_MCP_ROOT_DIR` to limit them to one folder).
- Cloudflare ends a request after about 100 seconds, so a very long tool call (rendering a big document) can
  fail through the tunnel; run it again or render fewer pages.

## What it can do

| Tool | What it does |
|---|---|
| `inspect_document` | paragraphs, tables, comments, revisions and the document's `sha256`; `search_text` finds a passage |
| `audit_document` | read-only package check, including whether Zotero citations can be added |
| `edit_document` | text, paragraph, table-row and section edits; tracked changes (add, accept, reject); Word comments (add, reply, resolve) |
| `list_revisions`, `list_bookmarks` | read tracked changes and bookmarks |
| `undo_document`, `checkpoint_document`, `document_history`, `revert_document` | undo any edit in one call, name a point to return to |
| `search_zotero`, `list_collections`, `set_project_collection` | search your Zotero library, pin a folder to a collection |
| `insert_zotero_citation` | insert a live Zotero citation (Zotero's own field, so the Zotero Word plugin keeps working) |
| `edit_citation_cluster` | change which references one citation cites |
| `refresh_citations` | re-render all citations and the bibliography through Zotero |
| `add_zotero_item_by_doi`, `update_zotero_item` | add a reference by DOI, fix a reference's details (needs a [write key](#zotero-write-key)) |
| `format_tables`, `clean_bookmarks` | table layout, remove stray bookmarks |
| `audit_document_style`, `style_text`, `apply_document_style` | check emphasis, bold/italic one phrase, apply a typography preset |
| `render_document_pages` | page images (PNG) to check layout |
| `report_issue` | a ready-to-file GitHub issue link |

Every write needs the `sha256` from the last read, so it never edits a version it hasn't seen. A
document open in Word is refused (save and close it first); the tools never touch the Word window you
are working in. Undo history lives in `%LOCALAPPDATA%\research_mcp\checkpoints`.

Agent tips: [`skills/word-research/SKILL.md`](skills/word-research/SKILL.md) (works as a Claude skill, or
paste it into your assistant's instructions).

## Zotero settings

Zotero must be running. In Zotero 7: **Edit → Settings → Advanced**, tick **"Allow other applications on
this computer to communicate with Zotero"**. That turns on the local API that searching and citing use
(`http://127.0.0.1:23119/api/`).

### Zotero write key

Optional. Only `add_zotero_item_by_doi` and `update_zotero_item` need it (they write through Zotero's web
API, then Zotero syncs the change to your desktop library).

1. Go to <https://www.zotero.org/settings/keys> and click **Create new private key**.
2. Tick **Allow library access** and **Allow write access**, save, and copy the key.
3. Note **Your user ID for use in API calls** on the same page.
4. Run `uvx research-mcp setup` and paste them in (or enter them in Claude Desktop's extension settings).

## Config file

`%APPDATA%\research-mcp\config.env`, one `KEY=value` per line. Environment variables override it.
`research-mcp setup` writes it for you.

| Key | Meaning |
|---|---|
| `ZOTERO_API_KEY` | Zotero write key (optional) |
| `ZOTERO_LIBRARY_ID` | your Zotero user ID (with the key) |
| `ZOTERO_LIBRARY_TYPE` | `user` (default) or `group` |
| `RESEARCH_MCP_AUTHOR` | author name on Word comments (default: the user name set in Word) |
| `RESEARCH_MCP_AUTHOR_INITIALS` | its initials (default: from the name) |
| `ZOTERO_MCP_CONTACT_EMAIL` | your email, sent to CrossRef and Unpaywall when adding by DOI so an open-access PDF can be attached (optional; see [Privacy](#privacy)) |
| `RESEARCH_MCP_ROOT_DIR` | limit the tools to documents inside this folder (default: anywhere) |
| `RESEARCH_MCP_REMOTE_TOKEN` | the secret in the `research-mcp chatgpt` URL (made on first run) |
| `RESEARCH_MCP_LOG_DIR` | where the local tool-call log goes (default `%LOCALAPPDATA%\research_mcp\logs`) |

## Troubleshooting

```powershell
uvx research-mcp doctor
```

checks Windows, Python, Word, Zotero and the write key, shows the config file, and lists which clients
have the server entry (named `research`). It exits with code 1 when something required is missing, and prints a fix
for each problem. Zotero counts as required: the tools read your library through the Zotero app, so it
has to be running with its local API on.

- **Tools don't show up:** quit the client fully (also from the system tray) and open it again.
- **"Zotero is not running"** (`doctor` exits 1; `setup` still writes the client configs but ends with
  `[FAIL] Zotero` and exit 1): start Zotero and leave it open. No need to run setup again.
- **"local API is off":** see [Zotero settings](#zotero-settings).
- **A write is refused because the file is open:** save and close the document in Word, then retry.
- **New version not picked up:** `uvx research-mcp@latest setup`, or `uv cache clean research-mcp`.

## Privacy

Everything runs on your PC: Word, your documents and Zotero's local API. The only requests that leave it:

- **Zotero web API** (`api.zotero.org`), only when you set a write key: the two tools that change your
  library use it, and `setup` and `doctor` check the key with one read-only call (`doctor --offline` skips it).
- **Adding a reference by DOI** sends the DOI to CrossRef (`api.crossref.org`) for its metadata, with your
  `ZOTERO_MCP_CONTACT_EMAIL` if you set one. With that email set, it also sends the DOI and the email to
  Unpaywall (`api.unpaywall.org`) and, when Unpaywall knows an open-access PDF, downloads it from the site
  Unpaywall names and attaches it to the new item.

- **`research-mcp chatgpt`** sends ChatGPT's tool calls and their results (document and Zotero text)
  through Cloudflare's tunnel to OpenAI, while it runs. The other ways to run it don't.

Tool calls are logged on your PC only, in `%LOCALAPPDATA%\research_mcp\logs\mcp_traffic.jsonl`
(next to the undo history; change it with `RESEARCH_MCP_LOG_DIR`). Each line holds the tool's arguments and
result, so it contains text from your documents and Zotero items. At 20 MB the log moves to
`mcp_traffic.1.jsonl` (replacing the older one), so it never takes more than about 40 MB. `report_issue`
notes go to `ISSUES.md` in the same folder. Delete the folder at any time.

## License

MIT. See [LICENSE](LICENSE). Contributions welcome: [CONTRIBUTING.md](CONTRIBUTING.md).
