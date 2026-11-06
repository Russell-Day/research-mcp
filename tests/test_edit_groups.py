"""Offline batch tests for operation groups (D10), the batch consistency check (D11) and the new op envelope."""

from __future__ import annotations

import asyncio
from pathlib import Path
import shutil
import zipfile

import pytest
from pydantic import ValidationError

from research_mcp.server import server
from research_mcp.audit import audit_docx, file_sha256
from research_mcp.edit_batch import edit_document
from research_mcp.edit_models import EditDocumentRequest, ReplaceSectionOperation
from research_mcp.errors import RefusedUnsafe
from research_mcp.fixtures import create_mixed_run_fixture


@pytest.fixture
def mixed(tmp_path: Path) -> Path:
    return create_mixed_run_fixture(tmp_path / "mixed.docx")


def _request(source: Path, output: Path | None, operations: list[dict], *, allow_partial: bool = True):
    return EditDocumentRequest(
        source=str(source),
        source_sha256=file_sha256(source),
        output=str(output) if output else None,
        operations=operations,
        allow_partial=allow_partial,
    )


def _fake_apply(failing: set[str], *, break_date: set[str] = frozenset(), rewrite: dict | None = None):
    calls: list[str] = []

    def apply(source, output, operation):
        calls.append(operation.op_id)
        if operation.op_id in failing:
            raise RefusedUnsafe(f"deliberate refusal of {operation.op_id}")
        edits = dict((rewrite or {}).get(operation.op_id, {}))
        if operation.op_id in break_date:
            edits["03/2020 – Present "] = "Ongoing work"
        if edits:
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            xml = parts["word/document.xml"].decode("utf-8")
            for old, new in edits.items():
                xml = xml.replace(old, new)
            parts["word/document.xml"] = xml.encode("utf-8")
            with zipfile.ZipFile(output, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
        else:
            shutil.copy2(source, output)
        return {"output_sha256": file_sha256(output), "paragraph_count": audit_docx(output).paragraph_count}

    apply.calls = calls
    return apply


def _replace(op_id: str, index: int, find: str, replace: str, **extra) -> dict:
    return {"op_id": op_id, "type": "replace_text", "paragraph_index": index, "find": find, "replace": replace, **extra}


def test_allow_partial_rolls_back_a_whole_table_row_group(monkeypatch, mixed, tmp_path):
    fake = _fake_apply({"date"})
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake)
    result = edit_document(_request(mixed, tmp_path / "out.docx", [
        _replace("title", 17, "Collections", "Colonial History"),
        _replace("date", 20, "Present", "08/2024"),
        _replace("closing", 29, "Plain", "Simple"),
    ]))
    statuses = {item.op_id: item.status for item in result.operations}
    assert statuses == {"title": "rolled_back", "date": "failed", "closing": "applied"}
    assert result.status == "partial"
    title = next(item for item in result.operations if item.op_id == "title")
    assert title.details["group_rolled_back"] and "table 1 row 5" in title.details["rollback_reason"]
    (group,) = result.groups
    assert group.group == "table 1 row 5" and group.status == "rolled_back"
    assert group.op_ids == ["title", "date"]
    assert result.consistency_check == "passed"
    assert "rolled back with their group" in result.message


def test_explicit_group_is_all_or_nothing(monkeypatch, mixed, tmp_path):
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", _fake_apply({"b"}))
    result = edit_document(_request(mixed, tmp_path / "out.docx", [
        _replace("a", 1, "Jordan", "Jordan A.", group="header"),
        _replace("b", 29, "Plain", "Simple", group="header"),
        _replace("c", 2, "Summary", "Profile"),
    ]))
    assert {item.op_id: item.status for item in result.operations} == {"a": "rolled_back", "b": "failed", "c": "applied"}
    assert [(g.group, g.status) for g in result.groups] == [("group 'header'", "rolled_back")]


def test_complete_group_reports_applied(monkeypatch, mixed, tmp_path):
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", _fake_apply(set()))
    result = edit_document(_request(mixed, tmp_path / "out.docx", [
        _replace("title", 17, "Collections", "Colonial History"),
        _replace("date", 20, "Present", "08/2024"),
    ]))
    assert result.status == "complete"
    assert [(g.group, g.status) for g in result.groups] == [("table 1 row 5", "applied")]


@pytest.mark.parametrize("allow_partial", [True, False])
def test_batch_consistency_check_cannot_be_bypassed(monkeypatch, mixed, tmp_path, allow_partial):
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", _fake_apply(set(), break_date={"closing"}))
    output = tmp_path / "out.docx"
    result = edit_document(_request(mixed, output, [_replace("closing", 29, "Plain", "Simple")],
                                    allow_partial=allow_partial))
    assert result.status == "refused"
    assert result.message.startswith("Batch consistency check failed; nothing was published")
    assert "Ongoing work" in result.message
    assert not output.exists()
    assert result.operations[0].status == "rolled_back"


def test_in_place_batch_reports_the_one_undo_call(monkeypatch, mixed):
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation",
                        _fake_apply(set(), rewrite={"closing": {"Plain closing": "Simple closing"}}))
    before = file_sha256(mixed)
    result = edit_document(_request(mixed, None, [_replace("closing", 29, "Plain", "Simple")]))
    assert result.status == "complete" and result.undo_token
    assert result.undo_call == f'undo_document(path="{str(mixed.resolve())}", undo_token="{result.undo_token}")'.replace(
        "\\", "\\\\")
    assert result.sha256_before == before and result.sha256_after == file_sha256(mixed) != before


def test_replace_section_needs_exactly_one_content_kind():
    base = {"op_id": "s", "type": "replace_section", "heading": "WORK EXPERIENCE"}
    with pytest.raises(ValidationError, match="exactly one"):
        ReplaceSectionOperation(**base)
    with pytest.raises(ValidationError, match="exactly one"):
        ReplaceSectionOperation(**base, rows=[["a", "2024"]], paragraphs=["a"])
    with pytest.raises(ValidationError, match="non-empty"):
        ReplaceSectionOperation(**base, rows=[[]])
    assert ReplaceSectionOperation(**base, paragraphs=["One", "Two"]).paragraphs == ["One", "Two"]


def test_server_envelope_accepts_new_operations(mixed, tmp_path):
    operations = [
        {"op_id": "p", "type": "replace_paragraph", "paragraph_index": 29, "text": "Closing."},
        {"op_id": "r", "type": "replace_table_row", "match": "07/2024", "cells": ["Norhaven", "07/2024"],
         "group": "g"},
        {"op_id": "s", "type": "replace_section", "heading": "WORK EXPERIENCE",
         "rows": [["Rivera Lab", "01/2024 – Present"]]},
    ]

    async def call():
        return await server.call_tool("edit_document", {
            "source": str(mixed), "source_sha256": "0" * 64,
            "output": str(tmp_path / "out.docx"), "operations": operations,
        })

    result = asyncio.run(call())
    assert result.is_error is False
    assert result.structured_content["status"] == "refused"
    assert "sha256" in result.structured_content["message"].lower()


def test_row_matched_by_a_staged_title_groups_with_the_title_edit(monkeypatch, mixed, tmp_path):
    """Review r1 finding 2: op 2 names the row by the title op 1 just wrote."""

    title = "Delta Park Museum, Department of Collections, Outreach Volunteer"
    renamed = title.replace("Delta Park Museum", "Delta Park Center")
    fake = _fake_apply({"date"}, rewrite={"title": {"Delta Park Museum": "Delta Park Center"},
                                           "closing": {"Plain closing": "Simple closing"}})
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake)
    output = tmp_path / "out.docx"
    result = edit_document(_request(mixed, output, [
        _replace("title", 17, "Delta Park Museum", "Delta Park Center"),
        {"op_id": "date", "type": "replace_table_row", "match": renamed, "cells": [renamed, "Ongoing work"]},
        _replace("closing", 29, "Plain", "Simple"),
    ]))
    statuses = {item.op_id: item.status for item in result.operations}
    assert statuses == {"title": "rolled_back", "date": "failed", "closing": "applied"}
    (group,) = result.groups
    assert group.group == "table 1 row 5" and group.op_ids == ["title", "date"] and group.status == "rolled_back"
    from research_mcp import sections as S
    published = S.read_layout(output)
    assert published.para(17).text == title and published.para(29).text.startswith("Simple closing")


def test_unresolvable_row_op_refuses_partial_publication(monkeypatch, mixed, tmp_path):
    """A row op whose row is found nowhere cannot be grouped, so nothing is published."""

    fake = _fake_apply({"date"}, rewrite={"closing": {"Plain closing": "Simple closing"}})
    monkeypatch.setattr("research_mcp.edit_batch._apply_operation", fake)
    output = tmp_path / "out.docx"
    result = edit_document(_request(mixed, output, [
        {"op_id": "date", "type": "replace_table_row", "match": "No such row", "cells": ["a", "b"]},
        _replace("closing", 29, "Plain", "Simple"),
    ]))
    assert result.status == "refused" and not output.exists()
    assert "replace_table_row 'date'" in result.message and "nothing was published" in result.message
    assert {item.op_id: item.status for item in result.operations} == {"date": "failed", "closing": "rolled_back"}


def test_staged_row_identity_survives_a_section_that_adds_rows(mixed):
    """Rows below a section that grew keep their source identity; new rows take the section's."""

    import copy
    import xml.etree.ElementTree as ET

    from research_mcp import sections as S
    from research_mcp.edit_batch import _advance_row_ids, source_row_ids
    from research_mcp.ooxml import NS

    before = S.read_layout(mixed)
    ids = source_row_ids(before)
    grown = mixed.with_name("grown.docx")
    with zipfile.ZipFile(mixed) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = ET.fromstring(parts["word/document.xml"])
    table = document.find("w:body/w:tbl", NS)
    rows = table.findall("w:tr", NS)
    table.insert(list(table).index(rows[3]) + 1, copy.deepcopy(rows[3]))  # a section op copying a row
    parts["word/document.xml"] = ET.tostring(document, xml_declaration=True, encoding="UTF-8")
    with zipfile.ZipFile(grown, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    after = S.read_layout(grown)
    advanced = _advance_row_ids(before, ids, after, frozenset({"section op"}))
    assert len(advanced[0]) == len(ids[0]) + 1
    assert advanced[0][:4] == ids[0][:4]
    assert advanced[0][5:] == ids[0][4:]  # rows below keep their source selves
    assert advanced[0][4] == frozenset({"section op"})
