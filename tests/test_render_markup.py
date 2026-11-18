"""render_document_pages(markup=...) (plan 864 decision 17): Item=7 with markup, Item=0 by default."""
from __future__ import annotations

from pathlib import Path
import subprocess

import pytest

from research_mcp import rendering
from research_mcp.schema_models import RenderDocumentResult


class FakeView:
    def __init__(self):
        self.ShowRevisionsAndComments = False
        self.RevisionsView = 1
        self.MarkupMode = 1

        class Filter:
            Markup = 0
            View = 1

        self.RevisionsFilter = Filter()


class FakeWindow:
    def __init__(self):
        self.View = FakeView()


def _fake_word(monkeypatch, calls: dict, *, with_window: bool = True):
    window = FakeWindow()

    class FakeDocument:
        def Windows(self, index):
            if not with_window:
                raise RuntimeError("no window")
            assert index == 1
            return window

        def ExportAsFixedFormat(self, **kwargs):
            calls["export"] = kwargs
            calls["view_at_export"] = {
                "ShowRevisionsAndComments": window.View.ShowRevisionsAndComments,
                "RevisionsView": window.View.RevisionsView,
                "MarkupMode": window.View.MarkupMode,
                "Filter.Markup": window.View.RevisionsFilter.Markup,
                "Filter.View": window.View.RevisionsFilter.View,
            }
            Path(kwargs["OutputFileName"]).write_bytes(b"pdf")

    class FakeSession:
        document = FakeDocument()

        def __init__(self, *args, **kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

    def fake_run(cmd, **kwargs):
        Path(cmd[-1] + "-1.png").write_bytes(b"png")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(rendering, "IsolatedWordDocument", FakeSession)
    monkeypatch.setattr(rendering, "find_pdftoppm", lambda: "pdftoppm.exe")
    monkeypatch.setattr(rendering.subprocess, "run", fake_run)
    return window


def test_markup_true_exports_with_item_7_and_all_markup_view(monkeypatch, tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    calls: dict = {}
    _fake_word(monkeypatch, calls)

    result = rendering.render_docx_pages(source, output_dir=tmp_path / "out", markup=True)

    assert calls["export"]["Item"] == 7
    assert calls["view_at_export"] == {
        "ShowRevisionsAndComments": True, "RevisionsView": 0, "MarkupMode": 0,
        "Filter.Markup": 2, "Filter.View": 0,
    }
    assert result["markup"] is True
    assert "warnings" not in result
    RenderDocumentResult.model_validate(result)


def test_markup_false_is_the_default_and_exports_item_0_without_touching_the_view(monkeypatch, tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    calls: dict = {}
    window = _fake_word(monkeypatch, calls)

    result = rendering.render_docx_pages(source, output_dir=tmp_path / "out")

    assert calls["export"]["Item"] == 0
    assert window.View.ShowRevisionsAndComments is False and window.View.MarkupMode == 1
    assert result["markup"] is False
    RenderDocumentResult.model_validate(result)


def test_markup_export_still_runs_when_the_window_view_cannot_be_set(monkeypatch, tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    calls: dict = {}
    _fake_word(monkeypatch, calls, with_window=False)

    result = rendering.render_docx_pages(source, output_dir=tmp_path / "out", markup=True)

    assert calls["export"]["Item"] == 7
    # terra61 S5 review P2: the pages may lack markup, so the result says so instead of reporting markup silently.
    assert result["warnings"] == [rendering.MARKUP_VIEW_WARNING]
    RenderDocumentResult.model_validate(result)


def test_markup_warns_when_the_markup_filter_cannot_be_set(monkeypatch, tmp_path: Path):
    """terra61 S5 r2: a filter left at Simple or No Markup hides changes, so failing to set it also warns."""
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    calls: dict = {}
    window = _fake_word(monkeypatch, calls)

    class StuckFilter:
        Markup = 0  # Word's "Simple Markup" left in place
        View = 1

        def __setattr__(self, name, value):
            raise RuntimeError("filter is read-only")

    window.View.RevisionsFilter = StuckFilter()
    result = rendering.render_docx_pages(source, output_dir=tmp_path / "out", markup=True)

    assert calls["export"]["Item"] == 7
    assert result["warnings"] == [rendering.MARKUP_VIEW_WARNING]


@pytest.mark.parametrize("markup, folder", [(False, "rendered"), (True, "rendered_markup")])
def test_default_output_folder_is_keyed_by_markup(monkeypatch, tmp_path: Path, markup, folder):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    _fake_word(monkeypatch, {})

    class Audit:
        sha256 = "abcdef0123456789"

    monkeypatch.setattr(rendering, "audit_docx", lambda path: Audit())
    result = rendering.render_docx_pages(source, markup=markup)

    assert Path(result["output_directory"]) == tmp_path / f"paper_{folder}_abcdef01"


@pytest.mark.parametrize("kwargs, expected", [({}, False), ({"markup": True}, True)])
def test_server_passes_markup_to_the_render_worker(monkeypatch, tmp_path: Path, kwargs, expected):
    from research_mcp import operation_process, server

    seen: dict = {}

    def fake_run_operation(kind, payload, timeout_seconds):
        seen["kind"], seen["payload"] = kind, payload
        return {"source": payload["docx_path"], "output_directory": str(tmp_path), "total_rendered": 0,
                "dpi": payload["dpi"], "markup": payload["markup"], "pages": []}

    monkeypatch.setattr(operation_process, "run_operation", fake_run_operation)
    server.render_document_pages(str(tmp_path / "paper.docx"), **kwargs)

    assert seen["kind"] == "render"
    assert seen["payload"]["markup"] is expected


def test_export_pdf_stays_clean(monkeypatch, tmp_path: Path):
    """export_pdf is unchanged by decision 17: always Item=0."""
    calls: dict = {}

    class FakeDocument:
        def ExportAsFixedFormat(self, **kwargs):
            calls["export"] = kwargs

    rendering._export_pdf(FakeDocument(), tmp_path / "x.pdf")
    assert calls["export"]["Item"] == 0
