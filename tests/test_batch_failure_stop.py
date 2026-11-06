from __future__ import annotations

import shutil

from research_mcp.audit import file_sha256
from research_mcp.edit_batch import edit_document
from research_mcp.edit_models import EditDocumentRequest
from research_mcp.errors import IntegrationRefused


def test_integration_failure_stops_later_operations_and_publishes_nothing(
    monkeypatch, minimal_docx, tmp_path
):
    seen = []

    def fail_first(source, output, operation):
        seen.append(operation.op_id)
        raise IntegrationRefused("Word integration unavailable")

    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fail_first)
    before = file_sha256(minimal_docx)
    output = tmp_path / "should-not-publish.docx"
    request = EditDocumentRequest(
        source=str(minimal_docx),
        source_sha256=before,
        output=str(output),
        operations=[
            {"op_id": "first", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"},
            {"op_id": "second", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "c"},
        ],
    )

    result = edit_document(request)

    assert seen == ["first"]
    assert [item.status for item in result.operations] == ["failed", "failed"]
    assert result.operations[1].error.context["attempted"] is False
    assert result.status == "no_changes"
    assert result.output is None
    assert not output.exists()
    assert file_sha256(minimal_docx) == before
