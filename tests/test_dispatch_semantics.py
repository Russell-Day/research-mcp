from pathlib import Path
import pytest
from research_mcp.edit_batch import _apply_operation_local
from research_mcp.edit_models import InsertTableOperation, ReplaceTextOperation
from research_mcp.errors import RefusedUnsafe


def test_insert_table_merge_dictionary_reaches_core(monkeypatch):
    seen = {}
    def insert(source, output, **kwargs):
        seen.update(kwargs)
        return {}
    monkeypatch.setattr("research_mcp.edit_batch.insert_table_to_version", insert)
    operation = InsertTableOperation(op_id="table", type="insert_table", after_paragraph_index=1, headers=["Group"], rows=[["A"], ["A"]], descriptive_columns=[1], merges=[dict(column=1, start_row=2, end_row=3, label="A")])
    _apply_operation_local(Path("source.docx"), Path("output.docx"), operation)
    assert seen["merges"][0].label == "A"
    assert seen["merges"][0].end_row == 3


@pytest.mark.parametrize("options", [dict(match_case=False), dict(occurrence=2), dict(expected_count=2)])
def test_replace_options_are_never_silently_ignored(options):
    operation = ReplaceTextOperation(op_id="r", type="replace_text", paragraph_index=1, find="A", replace="B", **options)
    with pytest.raises(RefusedUnsafe, match="unique, case-sensitive"):
        _apply_operation_local(Path("source.docx"), Path("output.docx"), operation)
