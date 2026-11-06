"""CLI for diagnostics and deterministic passes through the safe core."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from .audit import audit_docx
from .bookmarks import inventory, release_to_version
from .fixtures import create_table_fixture
from .integration import clear_integration_lock, integration_status
from .tables import MergeGroup, TableFormatSpec, format_to_version


def _load_names(path: str | Path) -> tuple[str, ...]:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, list) or not all(isinstance(item, str) for item in data):
        raise ValueError("bookmark names file must be a JSON list of exact strings")
    return tuple(data)


def _load_table_specs(path: str | Path) -> tuple[TableFormatSpec, ...]:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    specs: list[TableFormatSpec] = []
    for raw in data:
        merges = tuple(MergeGroup(**merge) for merge in raw.get("merges", []))
        specs.append(
            TableFormatSpec(
                index=int(raw["index"]),
                descriptive_columns=tuple(int(value) for value in raw["descriptive_columns"]),
                merges=merges,
            )
        )
    return tuple(specs)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="rmcp")
    subparsers = parser.add_subparsers(dest="command", required=True)

    audit = subparsers.add_parser("audit", help="read-only DOCX package audit")
    audit.add_argument("path")

    fixture = subparsers.add_parser("fixture-table", help="create a synthetic table fixture")
    fixture.add_argument("output")
    fixture.add_argument("--visible", action="store_true")

    bookmark_inventory = subparsers.add_parser("bookmarks", help="read-only bookmark inventory")
    bookmark_inventory.add_argument("path")

    bookmark_clean = subparsers.add_parser("clean-bookmarks", help="delete exact bookmark names into a new version")
    bookmark_clean.add_argument("source")
    bookmark_clean.add_argument("output")
    bookmark_clean.add_argument("--names-json", required=True)
    bookmark_clean.add_argument("--visible", action="store_true")

    tables = subparsers.add_parser("format-tables", help="apply verified table specs into a new version")
    tables.add_argument("source")
    tables.add_argument("output")
    tables.add_argument("--spec-json", required=True)
    tables.add_argument("--visible", action="store_true")

    restore = subparsers.add_parser(
        "restore-citations",
        help="replace (Author Year [N]) markers with live Zotero fields into a new version",
    )
    restore.add_argument("source")
    restore.add_argument("output")
    restore.add_argument("--item-donor", required=True)
    restore.add_argument("--bibliography-donor", required=True)
    restore.add_argument("--expect-markers", type=int, default=None)
    restore.add_argument("--expect-references", type=int, default=None)
    restore.add_argument("--timeout", type=float, default=None)

    subparsers.add_parser("integration-status", help="inspect the persistent Zotero integration lock")
    subparsers.add_parser("zotero-recover", help="dismiss dialogs, restart Zotero, wait for its local API")
    subparsers.add_parser("word-worker-status", help="inspect the project-owned Word worker")
    subparsers.add_parser("word-worker-shutdown", help="gracefully stop the project-owned Word worker")
    clear_lock = subparsers.add_parser("integration-clear-lock", help="clear a checked stale integration lock")
    clear_lock.add_argument("--confirm", choices=("CLEAR",), required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "audit":
        print(audit_docx(args.path).to_json())
    elif args.command == "fixture-table":
        print(create_table_fixture(args.output, visible=args.visible))
    elif args.command == "bookmarks":
        print(json.dumps([info.__dict__ for info in inventory(args.path)], indent=2))
    elif args.command == "clean-bookmarks":
        result = release_to_version(
            args.source,
            args.output,
            _load_names(args.names_json),
            visible=args.visible,
        )
        print(result.to_json())
    elif args.command == "format-tables":
        result = format_to_version(
            args.source,
            args.output,
            _load_table_specs(args.spec_json),
            visible=args.visible,
        )
        print(result.to_json())
    elif args.command == "restore-citations":
        from .citation_ops import restore_citations_to_version

        print(json.dumps(restore_citations_to_version(
            args.source,
            args.output,
            item_donor=args.item_donor,
            bibliography_donor=args.bibliography_donor,
            expected_markers=args.expect_markers,
            expected_references=args.expect_references,
            timeout=args.timeout,
        ), indent=2))
    elif args.command == "integration-status":
        print(json.dumps(integration_status(), indent=2))
    elif args.command == "zotero-recover":
        from .recovery import recover_zotero

        print(json.dumps(recover_zotero(), indent=2))
    elif args.command == "word-worker-status":
        from .word_service import read_worker_state, request_word_worker, worker_state_is_live

        state = read_worker_state()
        public_state = {key: value for key, value in (state or {}).items() if key != "authkey"}
        running = worker_state_is_live(state)
        live = request_word_worker({"command": "ping"}) if running else None
        print(json.dumps({"running": running, "state": public_state or None, "live": live}, indent=2))
    elif args.command == "word-worker-shutdown":
        from .word_service import shutdown_word_worker

        print(json.dumps(shutdown_word_worker(), indent=2))
    elif args.command == "integration-clear-lock":
        print(json.dumps(clear_integration_lock(confirmed=args.confirm == "CLEAR"), indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
