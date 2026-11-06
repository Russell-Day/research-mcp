import asyncio
import re
import sys
from pathlib import Path

from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client
import pytest
from starlette.testclient import TestClient

from research_mcp.errors import RefusedUnsafe
from research_mcp.server import build_parser, clean_bookmarks, format_tables, server


def test_mcp_exports_only_narrow_tools():
    tools = asyncio.run(server.list_tools())
    names = {tool.name for tool in tools}
    assert names == {
        "audit_document",
        "inspect_document",
        "list_bookmarks",
        "list_revisions",
        "clean_bookmarks",
        "format_tables",
        "search_zotero",
        "list_collections",
        "add_zotero_item_by_doi",
        "update_zotero_item",
        "set_project_collection",
        "insert_zotero_citation",
        "edit_citation_cluster",
        "refresh_citations",
        "edit_document",
        "revert_document",
        "undo_document",
        "checkpoint_document",
        "document_history",
        "report_issue",
        "apply_document_style",
        "audit_document_style",
        "style_text",
        "render_document_pages",
    }
    assert "execute_com" not in names
    assert "delete_range" not in names


def test_structured_content_shape_for_read_tools(minimal_docx):
    async def call(name, arguments):
        result = await server.call_tool(name, arguments)
        return result.structured_content

    audit = asyncio.run(call("audit_document", {"path": str(minimal_docx)}))
    assert audit["tables"], audit
    for table in audit["tables"]:
        assert "cell_hashes" not in table
        assert len(table["cell_text_sha256"]) == 64
    # The rest of the audit is intact, including the bookmark names an agent
    # needs to build a safe clean_bookmarks call.
    assert audit["bookmark_names"] == ["KeepMe"]
    assert audit["table_count"] == 1
    assert audit["paragraph_count_basis"].startswith("raw OOXML")

    inspected = asyncio.run(call("inspect_document", {"path": str(minimal_docx)}))
    assert re.fullmatch(r"[0-9A-F]{64}", inspected["source_sha256"])
    assert inspected["paragraph_count_basis"].startswith("Word-compatible")
    assert inspected["window_truncated"] is None


def test_inspect_window_schema_exposes_window_truncation(minimal_docx):
    result = asyncio.run(
        server.call_tool(
            "inspect_document",
            {"path": str(minimal_docx), "first_paragraph": 1, "last_paragraph": 999},
        )
    )
    payload = result.structured_content
    assert payload["window_truncated"] is False
    assert payload["truncated"] is False
    assert payload["returned_paragraph_count"] == len(payload["paragraphs"])


def test_search_zotero_reports_marker_path_when_marker_supplies_scope(tmp_path, monkeypatch):
    folder = tmp_path / "paper"
    folder.mkdir()
    marker = folder / ".zotero-collection"
    marker.write_text("Thesis\n", encoding="utf-8")
    doc = folder / "draft.docx"
    doc.write_bytes(b"")

    class FakeZotero:
        def resolve_collection(self, value):
            assert value == "Thesis"
            return "THESIS01"

        def search(self, query, *, limit, collection_key):
            assert (query, limit, collection_key) == ("term", 10, "THESIS01")
            return ()

    monkeypatch.setattr("research_mcp.server.ZoteroClient", FakeZotero)
    result = asyncio.run(
        server.call_tool("search_zotero", {"query": "term", "path": str(doc)})
    )

    assert result.structured_content == {
        "candidates": [],
        "collection_marker_path": str(marker.resolve()),
    }


def test_set_project_collection_resolves_before_writing_and_returns_identity(tmp_path, monkeypatch):
    folder = tmp_path / "paper"
    folder.mkdir()
    doc = folder / "draft.docx"
    doc.write_bytes(b"")

    class FakeZotero:
        def resolve_collection(self, value):
            assert value == "Thesis"
            return "THESIS01"

        def collections(self):
            return [{"key": "THESIS01", "name": "Thesis", "parent": None}]

    monkeypatch.setattr("research_mcp.server.ZoteroClient", FakeZotero)
    result = asyncio.run(
        server.call_tool(
            "set_project_collection",
            {"path": str(doc), "collection": "Thesis"},
        )
    )

    assert result.structured_content == {
        "success": True,
        "collection": "Thesis",
        "collection_key": "THESIS01",
        "marker_path": str((folder / ".zotero-collection").resolve()),
        "previous_value": None,
        "replaced": False,
    }
    assert (folder / ".zotero-collection").read_text(encoding="utf-8") == "Thesis\n"


def test_set_project_collection_refuses_before_writing_unknown_collection(tmp_path, monkeypatch):
    folder = tmp_path / "paper"
    folder.mkdir()
    doc = folder / "draft.docx"
    doc.write_bytes(b"")

    class FakeZotero:
        def resolve_collection(self, value):
            raise RefusedUnsafe(f"unknown collection: {value}")

    monkeypatch.setattr("research_mcp.server.ZoteroClient", FakeZotero)
    with pytest.raises(Exception, match="unknown collection"):
        asyncio.run(
            server.call_tool(
                "set_project_collection",
                {"path": str(doc), "collection": "Missing"},
            )
        )

    assert not (folder / ".zotero-collection").exists()


def test_list_collections_returns_collections(monkeypatch):
    class FakeZotero:
        def collections(self):
            return [{"key": "COL12345", "name": "Deep Learning", "parent": None}]

    monkeypatch.setattr("research_mcp.server.ZoteroClient", FakeZotero)
    result = asyncio.run(server.call_tool("list_collections", {}))
    assert result.structured_content == {
        "collections": [{"key": "COL12345", "name": "Deep Learning", "parent": None}],
    }


def test_mcp_transport_defaults_and_http_options():
    parser = build_parser()
    assert parser.parse_args([]).transport == "stdio"
    args = parser.parse_args(["--transport", "streamable-http", "--port", "9999"])
    assert args.host == "127.0.0.1"
    assert args.port == 9999


def test_stdio_protocol_handshake_and_tool_call(minimal_docx):
    async def exercise():
        parameters = StdioServerParameters(
            command=sys.executable,
            args=["-m", "research_mcp.server"],
        )
        async with stdio_client(parameters) as (read, write):
            async with ClientSession(read, write) as session:
                initialized = await session.initialize()
                tools = await session.list_tools()
                result = await session.call_tool("audit_document", {"path": str(minimal_docx)})
                return initialized, tools, result

    initialized, tools, result = asyncio.run(exercise())
    assert initialized.protocol_version == "2025-11-25"
    assert "audit_document" in {tool.name for tool in tools.tools}
    assert "apply_document_style" in {tool.name for tool in tools.tools}
    assert "render_document_pages" in {tool.name for tool in tools.tools}
    assert not result.is_error
    assert result.structured_content["table_count"] >= 0


def test_modern_http_request_is_self_contained_and_sessionless():
    app = server.streamable_http_app(stateless_http=True, json_response=True)
    meta = {
        "io.modelcontextprotocol/protocolVersion": "2026-07-28",
        "io.modelcontextprotocol/clientCapabilities": {},
    }
    body = {"jsonrpc": "2.0", "id": 1, "method": "tools/list", "params": {"_meta": meta}}
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "MCP-Protocol-Version": "2026-07-28",
        "Mcp-Method": "tools/list",
        "Host": "127.0.0.1:8765",
    }
    with TestClient(app) as client:
        first = client.post("/mcp", json=body, headers=headers)
        second = client.post("/mcp", json={**body, "id": 2}, headers=headers)
    assert first.status_code == 200
    assert second.status_code == 200
    assert "Mcp-Session-Id" not in first.headers
    assert {tool["name"] for tool in first.json()["result"]["tools"]} == {
        "audit_document",
        "inspect_document",
        "list_bookmarks",
        "list_revisions",
        "clean_bookmarks",
        "format_tables",
        "search_zotero",
        "list_collections",
        "add_zotero_item_by_doi",
        "update_zotero_item",
        "set_project_collection",
        "insert_zotero_citation",
        "edit_citation_cluster",
        "refresh_citations",
        "edit_document",
        "revert_document",
        "undo_document",
        "checkpoint_document",
        "document_history",
        "report_issue",
        "apply_document_style",
        "audit_document_style",
        "style_text",
        "render_document_pages",
    }


def test_edit_document_tool_returns_typed_stale_sha_refusal(minimal_docx, tmp_path):
    async def call():
        operations = [
            {
                "op_id": "op-1",
                "type": "replace_text",
                "paragraph_index": 1,
                "find": "old",
                "replace": "new",
            }
        ]
        result = await server.call_tool(
            "edit_document",
            {
                "source": str(minimal_docx),
                "source_sha256": "0" * 64,
                "output": str(tmp_path / "out.docx"),
                "operations": operations,
            },
        )
        return result

    result = asyncio.run(call())
    assert result.is_error is False
    assert result.structured_content["status"] == "refused"
    assert "sha256" in result.structured_content["message"].lower()


def test_inspect_document_search_text_returns_matches(minimal_docx):
    # Test that search_text doesn't fail validation
    result = asyncio.run(
        server.call_tool(
            "inspect_document",
            {"path": str(minimal_docx), "search_text": "Sample"},
        )
    )
    assert result.is_error is False
    assert "matches" in result.structured_content


def test_inspect_document_window_paging_returns_paragraphs(minimal_docx):
    # Test that bounded window doesn't fail validation
    result = asyncio.run(
        server.call_tool(
            "inspect_document",
            {"path": str(minimal_docx), "first_paragraph": 1, "last_paragraph": 2},
        )
    )
    assert result.is_error is False
    assert "paragraphs" in result.structured_content


def test_clean_bookmarks_worker_failure_does_not_publish(monkeypatch, minimal_docx, tmp_path: Path):
    original = minimal_docx.read_bytes()

    def fail(*args, **kwargs):
        raise RuntimeError("worker failed")

    monkeypatch.setattr("research_mcp.operation_process.run_operation", fail)
    output = tmp_path / "cleaned.docx"
    with pytest.raises(RuntimeError, match="worker failed"):
        clean_bookmarks(str(minimal_docx), str(output), ["MCPFixtureLegacy"])
    assert not output.exists()
    assert minimal_docx.read_bytes() == original


def test_format_tables_worker_failure_does_not_publish(monkeypatch, minimal_docx, tmp_path: Path):
    original = minimal_docx.read_bytes()

    def fail(*args, **kwargs):
        raise RuntimeError("worker failed")

    monkeypatch.setattr("research_mcp.operation_process.run_operation", fail)
    output = tmp_path / "formatted.docx"
    with pytest.raises(RuntimeError, match="worker failed"):
        format_tables(
            str(minimal_docx),
            str(output),
            [{"index": 1, "descriptive_columns": [1], "merges": []}],
        )
    assert not output.exists()
    assert minimal_docx.read_bytes() == original


def test_worker_payloads_use_staged_path_and_keyword_bookmark_names(monkeypatch, minimal_docx, tmp_path: Path):
    calls = []

    def succeed(kind, payload, **kwargs):
        calls.append((kind, payload))
        assert payload["source"] == payload["output"]
        return {}

    monkeypatch.setattr("research_mcp.operation_process.run_operation", succeed)
    clean_bookmarks(str(minimal_docx), str(tmp_path / "cleaned.docx"), ["MCPFixtureLegacy"])
    format_tables(
        str(minimal_docx),
        str(tmp_path / "formatted.docx"),
        [{"index": 1, "descriptive_columns": [1], "merges": []}],
    )
    assert calls[0][0] == "bookmarks"
    assert calls[0][1]["names"] == ["MCPFixtureLegacy"]
    assert calls[1][0] == "tables"
    assert calls[1][1]["specs"][0]["index"] == 1
