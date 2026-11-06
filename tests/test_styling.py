"""Tests for document styling and rendering tools."""

import asyncio
from pathlib import Path
from mcp.server.mcpserver.exceptions import ToolError
import pytest

from research_mcp.errors import RefusedUnsafe
from research_mcp.errors import VerificationFailed
from research_mcp.rendering import find_pdftoppm
from research_mcp.server import server
from research_mcp.styling import (
    PROFILES,
    _apply_line_spacing,
    _strip_headers_footers,
    apply_journal_style,
)


def test_style_profiles_exist_and_configured():
    assert "radiology_ai" in PROFILES
    assert "editing_compact" in PROFILES
    assert "compact_single" in PROFILES
    assert "standard_double" in PROFILES

    rad_profile = PROFILES["radiology_ai"]
    assert rad_profile.name == "radiology_ai"
    assert rad_profile.font_name == "Times New Roman"
    assert rad_profile.font_size_pt == 11.0
    assert rad_profile.line_spacing_rule == 2
    assert rad_profile.margin_top_inches == 1.0
    assert rad_profile.strip_header_footers is True

    compact = PROFILES["editing_compact"]
    assert compact.name == "editing_compact"
    assert compact.font_name == "Times New Roman"
    assert compact.font_size_pt == 11.0
    assert compact.line_spacing_rule == 0  # Single spacing
    assert compact.space_before_pt == 0.0
    assert compact.space_after_pt == 0.0
    assert compact.margin_top_inches == 1.0
    assert compact.strip_header_footers is True


def test_find_pdftoppm_returns_a_path_or_refuses_clearly():
    try:
        result = find_pdftoppm()
    except RefusedUnsafe as exc:  # machines without Poppler/MiKTeX (CI)
        assert "pdftoppm executable not found" in str(exc)
    else:
        assert isinstance(result, str) and result


def test_apply_journal_style_rejects_unknown_profile(tmp_path):
    src = tmp_path / "dummy.docx"
    out = tmp_path / "out.docx"
    src.write_bytes(b"")
    with pytest.raises(RefusedUnsafe, match="unknown style profile"):
        apply_journal_style(str(src), str(out), profile_name="nonexistent_profile")


def test_apply_document_style_tool_validates_stale_sha(minimal_docx, tmp_path):
    out = tmp_path / "styled.docx"
    with pytest.raises(ToolError, match="mismatch"):
        asyncio.run(
            server.call_tool(
                "apply_document_style",
                {
                    "source": str(minimal_docx),
                    "output": str(out),
                    "profile": "editing_compact",
                    "source_sha256": "0" * 64,
                },
            )
        )


def test_render_document_pages_validates_path(tmp_path):
    missing = tmp_path / "does_not_exist.docx"
    with pytest.raises(ToolError, match="document not found"):
        asyncio.run(
            server.call_tool(
                "render_document_pages",
                {"path": str(missing)},
            )
        )


class _FakeFormat:
    def __init__(self):
        self.assignments = []

    def __setattr__(self, name, value):
        if name != "assignments":
            self.assignments.append((name, value))
        object.__setattr__(self, name, value)


def test_named_line_spacing_rules_do_not_overwrite_the_rule_with_numeric_spacing():
    fmt = _FakeFormat()

    _apply_line_spacing(fmt, PROFILES["radiology_ai"])
    assert fmt.assignments == [("LineSpacingRule", 2)]

    fmt = _FakeFormat()
    _apply_line_spacing(fmt, PROFILES["editing_compact"])
    assert fmt.assignments == [("LineSpacingRule", 0)]


class _FakeHeaderFooter:
    def __init__(self, exists=True, error=None):
        self.Exists = exists
        self.error = None
        self.Range = self
        self.Text = "identifying content"
        self.error = error

    def __setattr__(self, name, value):
        if name == "Text" and getattr(self, "error", None):
            raise self.error
        object.__setattr__(self, name, value)


class _FakeSection:
    def __init__(self, error=None):
        self._headers = {i: _FakeHeaderFooter(error=error) for i in (1, 2, 3)}
        self._footers = {i: _FakeHeaderFooter(error=error) for i in (1, 2, 3)}

    def Headers(self, index):
        return self._headers[index]

    def Footers(self, index):
        return self._footers[index]


class _FakeDocument:
    def __init__(self, error=None):
        self.Sections = [_FakeSection(error=error)]


def test_strip_headers_and_footers_clears_all_existing_parts():
    doc = _FakeDocument()
    _strip_headers_footers(doc)
    for section in doc.Sections:
        for part in (*section._headers.values(), *section._footers.values()):
            assert part.Text == ""


def test_strip_headers_and_footers_surfaces_com_failures():
    with pytest.raises(VerificationFailed, match="could not clear all headers/footers"):
        _strip_headers_footers(_FakeDocument(error=RuntimeError("protected")))
