# Contributing

Bug reports and pull requests are welcome.

## Report a problem

Run `uvx research-mcp doctor` and include its output (it never prints your Zotero key). The
`report_issue` tool also builds a ready-to-file issue link from inside your assistant. Please leave out
document text you don't want public.

## Develop

You need Windows, Python 3.12, 3.13 or 3.14 and [uv](https://docs.astral.sh/uv/).

```powershell
git clone https://github.com/Russell-Day/research-mcp
cd research-mcp
uv sync
uv run pytest -m "not private_data"
```

The default test suite runs without Word or Zotero (CI runs it on `windows-latest`). Tests that drive
a real Word instance are opt-in: set `RESEARCH_MCP_WORD_LIVE=1` (they open a hidden, separate Word
instance on throwaway copies). Tests marked `private_data` need documents that are not in the repo.

Run your working copy as the server:

```powershell
uv --directory C:\path\to\research-mcp run research-mcp
```

and point an MCP client at that command, or try the setup flow without touching your real configs:

```powershell
uv run research-mcp setup --dry-run
```

## Pull requests

- One change per pull request, with a test when the behaviour changes.
- Keep writes safe: every tool that changes a document reads its `sha256` first, records an undo entry,
  and refuses when the file is open in Word.
- Don't add new runtime dependencies without a reason in the pull request.

By contributing you agree that your contribution is licensed under the MIT License.
