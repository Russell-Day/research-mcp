from research_mcp.cli import build_parser


def test_cli_requires_subcommand():
    parser = build_parser()
    assert parser.prog == "rmcp"


def test_cli_parses_audit():
    args = build_parser().parse_args(["audit", "fixture.docx"])
    assert args.command == "audit"
    assert args.path == "fixture.docx"

