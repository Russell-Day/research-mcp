"""How long a style pass may run in Word (plan 973 S1 r1). With Track Changes on, Word records a revision for every
paragraph and run it formats; measured on two real documents, 79 paragraphs took 31 s and 117 paragraphs 96 s,
past the 30 s every pass used to get. The tracked pass's time scales with the paragraph count. Fixtures are
synthetic; Word is faked."""

from __future__ import annotations

import pytest

from research_mcp import operation_process, styling
from research_mcp.audit import audit_docx, file_sha256
from research_mcp.styling import word_timeout_seconds, worker_timeout_seconds

from test_tables_grid import build_cv


class _Stop(Exception):
    pass


def test_the_untracked_pass_keeps_30_seconds():
    assert word_timeout_seconds(10, False) == word_timeout_seconds(5000, False) == 30.0
    assert worker_timeout_seconds(5000, False) == 120.0  # as before


def test_the_tracked_pass_scales_with_the_paragraphs():
    assert word_timeout_seconds(79, True) > 31.5  # measured 31.0-31.5 s
    assert word_timeout_seconds(117, True) > 96.4  # measured 96.4 s
    assert word_timeout_seconds(5, True) == 60.0  # floor
    assert word_timeout_seconds(100_000, True) == 600.0  # cap
    assert worker_timeout_seconds(117, True) == word_timeout_seconds(117, True) + 60.0
    assert worker_timeout_seconds(5, True) == 120.0


@pytest.mark.parametrize("tracked", [False, True])
def test_the_word_session_gets_the_pass_timeout(tmp_path, monkeypatch, tracked):
    source = build_cv(tmp_path / "cv.docx")
    seen = {}

    def isolated(path, **kwargs):
        seen.update(kwargs)
        raise _Stop

    monkeypatch.setattr(styling, "IsolatedWordDocument", isolated)
    with pytest.raises(_Stop):
        styling._style_document_local(source, "compact_single", tracked=tracked)
    assert seen["timeout_seconds"] == word_timeout_seconds(audit_docx(source).paragraph_count, tracked)
    assert seen["timeout_seconds"] == (60.0 if tracked else 30.0)


def test_the_tracked_worker_gets_the_scaled_timeout(tmp_path, monkeypatch):
    source = build_cv(tmp_path / "cv.docx")
    from docx import Document

    document = Document(source)
    for number in range(60):
        document.add_paragraph(f"Line {number}.")
    document.save(source)
    paragraphs = audit_docx(source).paragraph_count
    seen = {}

    def run_operation(name, payload, *, timeout_seconds):
        seen.update(name=name, timeout=timeout_seconds)
        raise _Stop

    monkeypatch.setattr(operation_process, "run_operation", run_operation)
    before = file_sha256(source)
    with pytest.raises(_Stop):
        styling.apply_journal_style(source, tmp_path / "out.docx", "compact_single", tracked=True)
    assert seen == {"name": "style", "timeout": worker_timeout_seconds(paragraphs, True)}
    assert seen["timeout"] == 30.0 + 2.0 * paragraphs + 60.0 > 120.0
    assert not (tmp_path / "out.docx").exists() and file_sha256(source) == before
