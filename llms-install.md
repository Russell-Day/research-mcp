# Installing research-mcp (for AI agents)

These steps are for an AI assistant that can run commands on the user's PC (the ChatGPT or Codex desktop
app, Claude Code, Claude Desktop with a shell tool, Cursor, VS Code). research-mcp is an MCP server that
edits Word documents and cites from Zotero. Run the commands yourself; the steps marked **user** are for
the user.

## 1. Check the PC

- It must be **Windows** with the desktop **Microsoft Word**. On macOS or Linux, stop and tell the user
  research-mcp needs Windows and Word.
- Zotero 7 (desktop) should be installed and running. If it isn't, carry on, and tell the user at the end.

## 2. Install uv if it is missing

Run `uv --version`. If the command isn't found, install it (PowerShell, no admin needed):

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then add it to this shell's PATH: `$env:Path = "$env:USERPROFILE\.local\bin;$env:Path"`.

## 3. Add the server

```powershell
uvx research-mcp setup --yes
```

This adds a server named `research` to every MCP client it finds. To set up only the app you are running
in, pass `--client <name>` instead: `chatgpt` (the ChatGPT desktop app, which shares the Codex config),
`codex`, `claude-desktop`, `claude-code`, `cursor` or `vscode`.

Setup keeps the user's other servers and backs up each file it changes. If it prints `NOT changed` for a
client, show the user that line; don't edit that file by hand. `uvx research-mcp setup --print` shows the
entry for each client without changing anything.

## 4. Check

```powershell
uvx research-mcp doctor
```

Exit code 0 means the install is fine. Each `[FAIL]` line has a `fix:` line under it; pass it on to the user.

## 5. Tell the user

- **user:** Restart the app fully (quit it from the system tray, then open it again). The research tools
  appear after the restart, not in this session.
- **user, ChatGPT app:** after the restart, type `/mcp` in a chat to see `research`. If a chat doesn't show
  the research tools, use a Codex chat.
- **user, Zotero:** keep Zotero open while working. In Zotero: Edit > Settings > Advanced, tick "Allow other
  applications on this computer to communicate with Zotero".
- **user, optional:** a Zotero write key is only needed to add or change library items. They run
  `uvx research-mcp setup` in a terminal themselves, where the key is typed hidden. Don't ask them to paste
  the key into the chat.
