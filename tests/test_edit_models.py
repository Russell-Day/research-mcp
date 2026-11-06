from __future__ import annotations

import pytest
from pydantic import ValidationError

from research_mcp.edit_models import EditBatchResult, EditDocumentRequest


SHA = "A" * 64


def request(operations):
    return EditDocumentRequest(
        source="source.docx",
        source_sha256=SHA,
        output="output.docx",
        operations=operations,
    )


def test_edit_operation_schema_is_an_eighteen_way_discriminated_union():
    schema = EditDocumentRequest.model_json_schema()
    item = schema["properties"]["operations"]["items"]
    assert item["discriminator"]["propertyName"] == "type"
    assert set(item["discriminator"]["mapping"]) == {
        "replace_text",
        "insert_text",
        "insert_paragraph",
        "insert_table",
        "reply_comment",
        "add_comment",
        "replace_paragraph",
        "replace_table_row",
        "replace_section",
        "tracked_replace_text",
        "tracked_replace_paragraph",
        "accept_revisions",
        "reject_revisions",
        "list_revisions",
        "sweep_characters",
        "resolve_comment",
        "reopen_comment",
        "delete_comment",
    }
    assert len(item["oneOf"]) == 18
    assert schema["properties"]["operations"]["minItems"] == 1
    assert schema["properties"]["operations"]["maxItems"] == 20


def test_request_rejects_bad_sha_unknown_fields_and_duplicate_ids():
    operation = {
        "op_id": "edit1",
        "type": "replace_text",
        "paragraph_index": 1,
        "find": "old",
        "replace": "new",
    }
    with pytest.raises(ValidationError, match="source_sha256"):
        EditDocumentRequest(source="a.docx", source_sha256="bad", output="b.docx", operations=[operation])
    with pytest.raises(ValidationError, match="extra_forbidden"):
        request([operation | {"invented": True}])
    with pytest.raises(ValidationError, match="unique"):
        request([operation, operation])


def test_addressed_reply_requires_known_success_dependency():
    edit = {
        "op_id": "edit1",
        "type": "replace_text",
        "paragraph_index": 1,
        "find": "old",
        "replace": "new",
    }
    reply = {
        "op_id": "reply1",
        "type": "reply_comment",
        "comment_id": "0A5777F8",
        "disposition": "addressed",
        "reply_text": "Clarified.",
        "depends_on": ["edit1"],
    }
    assert len(request([edit, reply]).operations) == 2
    with pytest.raises(ValidationError, match="requires"):
        request([reply | {"depends_on": []}])
    with pytest.raises(ValidationError, match="unknown dependencies"):
        request([reply | {"depends_on": ["missing"]}])


def test_dependency_cycles_are_rejected():
    replies = [
        {
            "op_id": "reply1",
            "type": "reply_comment",
            "comment_id": "A",
            "disposition": "addressed",
            "reply_text": "One",
            "depends_on": ["reply2"],
        },
        {
            "op_id": "reply2",
            "type": "reply_comment",
            "comment_id": "B",
            "disposition": "addressed",
            "reply_text": "Two",
            "depends_on": ["reply1"],
        },
    ]
    with pytest.raises(ValidationError, match="earlier|cycle"):
        request(replies)


def test_dependencies_must_precede_the_dependent_operation():
    reply = {
        "op_id": "reply",
        "type": "reply_comment",
        "comment_id": "A",
        "disposition": "addressed",
        "reply_text": "Done",
        "depends_on": ["later"],
    }
    later = {
        "op_id": "later",
        "type": "replace_text",
        "paragraph_index": 1,
        "find": "a",
        "replace": "b",
    }
    with pytest.raises(ValidationError, match="must appear earlier"):
        request([reply, later])


def test_table_grid_bounds_labels_and_merges_validate():
    valid = {
        "op_id": "table1",
        "type": "insert_table",
        "after_paragraph_index": 1,
        "headers": ["Model", "Value"],
        "rows": [["LoRA", "1"], ["LoRA", "2"]],
        "descriptive_columns": [1],
        "merges": [{"column": 1, "start_row": 2, "end_row": 3, "label": "LoRA"}],
    }
    assert request([valid]).operations[0].headers == ["Model", "Value"]
    with pytest.raises(ValidationError, match="same width"):
        request([valid | {"rows": [["too short"]]}])
    with pytest.raises(ValidationError, match="outside"):
        request([valid | {"descriptive_columns": [3]}])
    with pytest.raises(ValidationError, match="exactly match"):
        request([valid | {"merges": [{"column": 1, "start_row": 2, "end_row": 3, "label": "Wrong"}]}])


def test_typed_result_schema_exposes_status_retry_and_per_operation_errors():
    schema = EditBatchResult.model_json_schema()
    assert set(schema["properties"]) >= {
        "status",
        "source",
        "source_sha256",
        "output",
        "output_sha256",
        "operations",
        "retry",
        "protected",
    }
    operation_ref = schema["properties"]["operations"]["items"]["$ref"]
    assert operation_ref.endswith("/OperationResult")
