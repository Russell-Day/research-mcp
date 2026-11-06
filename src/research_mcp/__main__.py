"""``research-mcp`` command: run the server, set it up, or check an install.

    research-mcp                     run the MCP server over stdio (what MCP clients launch)
    research-mcp serve [--transport] the same, with the server's own options
    research-mcp setup               check this PC and add the server to your MCP clients
    research-mcp doctor              check an existing install (exit code 1 when a required check fails)
    research-mcp chatgpt             serve the tools to ChatGPT (HTTP + a public tunnel address)
    research-mcp <rmcp command>      the diagnostic commands of ``rmcp`` (audit, bookmarks, ...)

A first argument that starts with ``-`` (other than ``-h``/``--version``) goes to the server, so launch lines
such as ``research-mcp --transport stdio`` keep working. ``python -m research_mcp`` is the same command.
Imports stay inside the branches: ``setup`` and ``doctor`` must work where Word's COM modules are missing.
"""

from __future__ import annotations

import sys

PROG = "research-mcp"

USAGE = f"""usage: {PROG} [serve | setup | doctor | chatgpt] [options]

Word and Zotero tools for MCP clients (Windows, Microsoft Word, Zotero desktop).

commands:
  serve     run the MCP server over stdio (the default; `{PROG} serve --help` for HTTP options)
  setup     check this PC, store an optional Zotero key, add the server to your MCP clients
  doctor    check an existing install; exit code 1 when a required check fails
  chatgpt   serve the tools to ChatGPT through a public HTTPS address (keep the window open)

diagnostics (same as the `rmcp` command): audit, bookmarks, integration-status, word-worker-status, ...
`{PROG} <command> --help` shows a command's options.
"""


def _version() -> str:
    from research_mcp import __version__

    return __version__


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    command = args[0] if args else "serve"
    if command in ("-h", "--help", "help"):
        print(USAGE)
        return 0
    if command in ("-V", "--version"):
        print(f"{PROG} {_version()}")
        return 0
    if command == "setup":
        from research_mcp.setup_cmd import setup_main

        return setup_main(args[1:])
    if command == "doctor":
        from research_mcp.setup_cmd import doctor_main

        return doctor_main(args[1:])
    if command == "chatgpt":
        from research_mcp.remote import chatgpt_main

        return chatgpt_main(args[1:])
    if command == "serve" or command.startswith("-"):
        from research_mcp.server import main as serve

        serve(args[1:] if command == "serve" else args, prog=f"{PROG} serve")
        return 0
    from research_mcp.cli import build_parser as rmcp_parser, main as rmcp

    if command in _rmcp_commands(rmcp_parser()):
        return rmcp(args)
    print(USAGE, file=sys.stderr)
    print(f"{PROG}: unknown command {command!r}", file=sys.stderr)
    return 2


def _rmcp_commands(parser) -> set[str]:
    import argparse

    for action in parser._actions:  # argparse keeps subcommands on its private subparsers action
        if isinstance(action, argparse._SubParsersAction):
            return set(action.choices)
    return set()


if __name__ == "__main__":
    raise SystemExit(main())
