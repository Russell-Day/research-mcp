from __future__ import annotations

from pathlib import Path

import pytest

from research_mcp.audit import audit_docx
from research_mcp.bookmarks import validate_release
from research_mcp.constants import word_bgr
from research_mcp.errors import RefusedUnsafe
from research_mcp.word import create_versioned_copy


def test_word_bgr():
    assert word_bgr(0x11, 0x22, 0x33) == 0x332211
    with pytest.raises(ValueError):
        word_bgr(256, 0, 0)


def test_versioned_copy_refuses_overwrite(minimal_docx, tmp_path: Path):
    output = tmp_path / "output.docx"
    create_versioned_copy(minimal_docx, output)
    assert output.read_bytes() == minimal_docx.read_bytes()
    with pytest.raises(RefusedUnsafe):
        create_versioned_copy(minimal_docx, output)


def test_referenced_bookmark_cannot_be_released(minimal_docx):
    audit = audit_docx(minimal_docx)
    with pytest.raises(RefusedUnsafe, match="referenced"):
        validate_release(audit, ["KeepMe"])

