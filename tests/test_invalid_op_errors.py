"""Plan 864 decision 18: an invalid edit_document op gets one readable line, never an IndexError.

Whole-model validator errors carry an empty `loc`; each must come back as its own message, naming the op, its
type and what the type takes. delete_citation also accepts `find` as current_text and, with no selector, lists
the paragraph's citations; an empty edit_citation_cluster points at a ready delete_citation op."""

from __future__ import annotations

import asyncio
import json
import os
import re

import pytest

from research_mcp.audit import file_sha256
from research_mcp.edit_models import DEFAULT_MAX_OPERATIONS, EditDocumentRequest
from research_mcp.errors import RefusedUnsafe
from test_citation_delete import LEGEND, doc_xml, make_docx

windows_only = pytest.mark.skipif(os.name != "nt", reason="undo store writes need Windows Python")


def refusal(path, operations, **extra) -> str:
    """The error text edit_document gives the model for these operations (raised or returned as an error)."""

    from research_mcp import server as srv

    async def run():
        return await srv.server.call_tool("edit_document", {"source": str(path), "source_sha256": "0" * 64,
                                                            "output": str(path), "operations": operations,
                                                            **extra})

    try:
        result = asyncio.run(run())
    except Exception as exc:  # noqa: BLE001 - the tool error is what the model sees
        return str(exc)
    assert getattr(result, "is_error", False), result
    return " ".join(getattr(block, "text", "") for block in result.content)


WHOLE_MODEL_CASES = [
    ({"type": "insert_table", "after_paragraph_index": 1, "headers": ["A", "B"], "rows": [["1"]],
      "descriptive_columns": [1]}, "every row must have the same width as headers (2)"),
    ({"type": "insert_table", "after_paragraph_index": 1, "headers": ["A"], "rows": [["1"]],
      "descriptive_columns": [3]}, "descriptive column 3 outside table width 1"),
    ({"type": "replace_section", "heading": "Methods"},
     "give exactly one of rows (table sections) or paragraphs (body sections)"),
    ({"type": "replace_section", "heading": "Methods", "rows": [[]]},
     "rows must be a non-empty list of non-empty cell lists"),
    ({"type": "accept_revisions"}, "give ids, author, a paragraph range or a heading"),
    ({"type": "reject_revisions"}, "give ids, author, a paragraph range or a heading"),
    ({"type": "list_revisions", "until_heading": "Results"}, "until_heading needs heading"),
    ({"type": "sweep_characters", "paragraph_index": 5, "through_paragraph_index": 2},
     "through_paragraph_index must not be before paragraph_index"),
    ({"type": "delete_citation", "paragraph_index": 2, "citation_id": "c1", "current_text": "(1)"},
     "give citation_id or current_text, not both"),
]


@pytest.mark.parametrize(("op", "message"), WHOLE_MODEL_CASES, ids=[c[0]["type"] for c in WHOLE_MODEL_CASES])
def test_whole_model_validator_errors_read_as_their_own_message(tmp_path, op, message):
    text = refusal(tmp_path / "doc.docx", [{"op_id": "bad", **op}])
    assert "tuple index" not in text and "IndexError" not in text
    assert f"operation 'bad' ({op['type']}) invalid: {message}" in text
    assert f"{op['type']} requires:" in text


def test_operation_error_text_with_an_empty_loc():
    from research_mcp.server import MCPEditOperation, _operation_error_text

    op = MCPEditOperation(op_id="x", type="accept_revisions")
    text = _operation_error_text(op, {"loc": (), "msg": "Value error, give ids, author, a paragraph range or a "
                                                        "heading", "type": "value_error"})
    assert text.startswith("operation 'x' (accept_revisions) invalid: give ids, author, a paragraph range or a "
                           "heading. accept_revisions requires:")
    field = _operation_error_text(op, {"loc": ("accept_revisions", "ids"), "msg": "List should have at least 1 item"})
    assert "invalid: field 'ids' List should have at least 1 item" in field


def test_batch_level_errors_keep_their_message(tmp_path):
    text = refusal(tmp_path / "doc.docx", [
        {"op_id": "same", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"},
        {"op_id": "same", "type": "replace_text", "paragraph_index": 2, "find": "a", "replace": "b"}])
    assert "op_id values must be unique within a batch" in text
    too_many = [{"op_id": f"o{n}", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"}
                for n in range(DEFAULT_MAX_OPERATIONS + 1)]
    assert f"at most {DEFAULT_MAX_OPERATIONS} operations per batch" in refusal(tmp_path / "doc.docx", too_many)


def test_track_changes_sugar_keeps_error_positions_on_the_callers_ops(tmp_path):
    """The prepended switch must not shift a later op's error onto its neighbour."""

    text = refusal(tmp_path / "doc.docx", [
        {"op_id": "first", "type": "replace_text", "paragraph_index": 1, "find": "a", "replace": "b"},
        {"op_id": "dep", "type": "replace_text", "paragraph_index": 1, "find": "c", "replace": "d",
         "depends_on": ["nowhere"]}], track_changes=True)
    assert "unknown dependencies: 'nowhere'" in text


# -- delete_citation ---------------------------------------------------------------------------------------

def test_find_is_another_name_for_current_text():
    from research_mcp.server import MCPEditOperation, _strict_operation_payload

    op = _strict_operation_payload(MCPEditOperation(op_id="d", type="delete_citation", paragraph_index=1,
                                                    find="22"))
    assert (op.current_text, op.citation_id) == ("22", None)
    with pytest.raises(RefusedUnsafe, match="find and current_text both given"):
        _strict_operation_payload(MCPEditOperation(op_id="d", type="delete_citation", paragraph_index=1,
                                                   find="22", current_text="22"))


def test_no_selector_lists_the_paragraphs_citations(tmp_path):
    from research_mcp.citation_delete import delete_citation

    src = make_docx(tmp_path / "in.docx", LEGEND)
    with pytest.raises(RefusedUnsafe) as caught:
        delete_citation(src, tmp_path / "out.docx", paragraph_index=1)
    message = str(caught.value)
    assert message.startswith('delete_citation: say which citation to delete with current_text (its rendered '
                              'text, e.g. "8") or citation_id; Zotero citations in paragraph 1: ')
    assert "cX '22' items ['K1']" in message


@windows_only
def test_no_selector_through_edit_document_reports_the_listing(tmp_path):
    from research_mcp.edit_batch import edit_document

    src = make_docx(tmp_path / "in.docx", LEGEND)
    before = file_sha256(src)
    result = edit_document(EditDocumentRequest(source=str(src), source_sha256=before, output=str(src),
                                               operations=[{"op_id": "d", "type": "delete_citation",
                                                            "paragraph_index": 1}]))
    assert result.status == "no_changes"
    assert "Zotero citations in paragraph 1: cX '22'" in result.operations[0].error.message
    assert file_sha256(src) == before


@windows_only
def test_find_alias_deletes_through_the_tool(tmp_path):
    from research_mcp import server as srv

    src = make_docx(tmp_path / "in.docx", LEGEND)

    async def run():
        return await srv.server.call_tool("edit_document", {
            "source": str(src), "source_sha256": file_sha256(src), "output": str(src),
            "operations": [{"op_id": "d", "type": "delete_citation", "paragraph_index": 1, "find": "22"}]})

    result = asyncio.run(run())
    assert result.structured_content["status"] == "complete", result.structured_content
    xml = doc_xml(src)
    assert xml.count('w:author="Claude"') == 1 and "<w:delInstrText" in xml
    inner = xml[xml.index('<w:del '):xml.index("</w:del>")]
    assert "<w:delText>22</w:delText>" in inner and "ZOTERO_ITEM" in inner  # the whole field, one w:del
    assert not re.search(r"<w:instrText[^>]*>[^<]*ZOTERO_ITEM", xml)  # no live citation left
    from research_mcp.paragraphs import read_paragraphs
    assert read_paragraphs(src)[0].text == "Figure 2 legend."  # accepted view: citation and its space gone


def test_empty_item_keys_hint_is_a_ready_op(tmp_path):
    from research_mcp.cluster_ops import edit_cluster_to_version

    src = make_docx(tmp_path / "in.docx", LEGEND)
    with pytest.raises(RefusedUnsafe) as caught:
        edit_cluster_to_version(src, src, source_sha256="0" * 64, paragraph_index=1, item_keys=[],
                                current_text="22")
    ready = str(caught.value).split("by default): ", 1)[1]
    assert json.loads(ready) == {"type": "delete_citation", "paragraph_index": 1, "current_text": "22"}
    with pytest.raises(RefusedUnsafe) as by_id:
        edit_cluster_to_version(src, src, source_sha256="0" * 64, paragraph_index=1, item_keys=[],
                                citation_id="cX")
    assert json.loads(str(by_id.value).split("by default): ", 1)[1])["citation_id"] == "cX"


def test_server_instructions_name_the_reply_disposition():
    """Plan 864 decision 13: the instructions say how to answer a comment in plain words."""

    from research_mcp.server import SERVER_INSTRUCTIONS

    assert "To answer a comment in plain words use disposition `reply`" in SERVER_INSTRUCTIONS
