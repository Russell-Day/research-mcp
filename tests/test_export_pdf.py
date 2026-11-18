"""export_pdf: a new PDF beside the .docx, never over an existing file, source untouched (Word faked)."""

from __future__ import annotations

import asyncio
import dataclasses
from pathlib import Path

from mcp.server.mcpserver.exceptions import ToolError
import pytest

from research_mcp import editing, operation_process, rendering
from research_mcp.audit import file_sha256
from research_mcp.errors import RefusedUnsafe
from research_mcp.server import server


# Three page objects (one with its keys in another order) and the page-tree root, which is not a page.
FAKE_PDF = b"%PDF-1.7 << /Type /Pages /Count 3 >> << /Type/Page >> << /Type /Page/Parent 2 0 R >> << /Parent 2 0 R /Type/Page >>"


class FakeSession:
    """Stands in for IsolatedWordDocument; ``export`` decides what ExportAsFixedFormat does."""

    calls: list[dict] = []
    export = staticmethod(lambda path: Path(path).write_bytes(FAKE_PDF))

    def __init__(self, path, *, visible=False, read_only=False, timeout_seconds=30.0):
        FakeSession.calls.append(
            {"path": path, "visible": visible, "read_only": read_only, "timeout_seconds": timeout_seconds}
        )
        outer = self

        class Document:
            def ExportAsFixedFormat(self, **kwargs):
                outer.export_kwargs = kwargs
                FakeSession.last_export = kwargs
                FakeSession.export(kwargs["OutputFileName"])

        self.document = Document()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


@pytest.fixture
def fake_word(monkeypatch):
    FakeSession.calls = []
    FakeSession.export = staticmethod(lambda path: Path(path).write_bytes(FAKE_PDF))
    monkeypatch.setattr(rendering, "IsolatedWordDocument", FakeSession)
    return FakeSession


def _leftovers(folder: Path) -> list[str]:
    return sorted(p.name for p in folder.iterdir() if p.name.startswith(".~research-mcp-export-"))


def test_export_writes_pdf_beside_the_docx_read_only(fake_word, minimal_docx):
    before = file_sha256(minimal_docx)

    result = rendering.export_docx_pdf(minimal_docx)

    target = minimal_docx.with_suffix(".pdf")
    assert result["output"] == str(target)
    assert target.read_bytes() == FAKE_PDF
    assert result["size_bytes"] == len(FAKE_PDF)
    assert result["page_count"] == 3
    assert result["source_sha256"] == before == file_sha256(minimal_docx)
    assert result["warnings"] == []
    assert result["elapsed_seconds"] >= 0
    call = fake_word.calls[0]
    assert call["read_only"] is True and call["visible"] is False
    assert call["timeout_seconds"] == rendering.WORD_EXPORT_TIMEOUT_SECONDS
    exported = fake_word.last_export
    assert exported["Item"] == 0 and exported["ExportFormat"] == 17
    # Word writes a hidden temporary in the target folder, published afterwards; none is left.
    assert Path(exported["OutputFileName"]).parent == target.parent
    assert Path(exported["OutputFileName"]).name.startswith(".~research-mcp-export-")
    assert _leftovers(target.parent) == []


def test_export_to_a_named_output(fake_word, minimal_docx, tmp_path):
    folder = tmp_path / "out"
    folder.mkdir()
    result = rendering.export_docx_pdf(minimal_docx, folder / "PS2_hand_in.PDF")
    assert result["output"] == str((folder / "PS2_hand_in.PDF").resolve())
    assert (folder / "PS2_hand_in.PDF").is_file()
    assert not minimal_docx.with_suffix(".pdf").exists()


def test_existing_target_is_refused_before_word_starts(fake_word, minimal_docx):
    target = minimal_docx.with_suffix(".pdf")
    target.write_bytes(b"the user's checked copy")

    with pytest.raises(RefusedUnsafe, match="already exists and is never overwritten"):
        rendering.export_docx_pdf(minimal_docx)

    assert target.read_bytes() == b"the user's checked copy"
    assert fake_word.calls == []
    assert _leftovers(target.parent) == []


@pytest.mark.parametrize(
    ("output", "message"),
    [("notes.docx", "must end in .pdf"), ("missing/dir/out.pdf", "output folder does not exist")],
)
def test_bad_output_is_refused(fake_word, minimal_docx, tmp_path, output, message):
    with pytest.raises(RefusedUnsafe, match=message):
        rendering.export_docx_pdf(minimal_docx, tmp_path / output)
    assert fake_word.calls == []
    assert not (tmp_path / output).exists()


@pytest.mark.skipif(not hasattr(__import__("os"), "symlink"), reason="no symlinks")
def test_dangling_symlink_output_is_refused(fake_word, minimal_docx, tmp_path):
    link = tmp_path / "link.pdf"
    try:
        link.symlink_to(tmp_path / "elsewhere.pdf")
    except OSError:
        pytest.skip("symlink creation needs a privilege here")
    with pytest.raises(RefusedUnsafe, match="already exists"):
        rendering.export_docx_pdf(minimal_docx, link)
    assert not (tmp_path / "elsewhere.pdf").exists()
    assert fake_word.calls == []


def test_missing_document_is_refused(fake_word, tmp_path):
    with pytest.raises(FileNotFoundError, match="document not found"):
        rendering.export_docx_pdf(tmp_path / "nope.docx")


def test_target_that_appears_during_export_is_left_alone(fake_word, minimal_docx):
    target = minimal_docx.with_suffix(".pdf")

    def export_while_someone_saves(path):
        Path(path).write_bytes(b"%PDF-1.7 ours")
        target.write_bytes(b"theirs")

    fake_word.export = staticmethod(export_while_someone_saves)
    with pytest.raises(RefusedUnsafe, match="appeared while Word was exporting"):
        rendering.export_docx_pdf(minimal_docx)

    assert target.read_bytes() == b"theirs"
    assert _leftovers(target.parent) == []


def test_failed_export_leaves_no_files(fake_word, minimal_docx):
    def export_then_fail(path):
        Path(path).write_bytes(b"partial")
        raise RuntimeError("Word crashed")

    fake_word.export = staticmethod(export_then_fail)
    with pytest.raises(RuntimeError, match="Word crashed"):
        rendering.export_docx_pdf(minimal_docx)

    assert not minimal_docx.with_suffix(".pdf").exists()
    assert _leftovers(minimal_docx.parent) == []


def test_empty_export_is_refused(fake_word, minimal_docx):
    fake_word.export = staticmethod(lambda path: Path(path).write_bytes(b""))
    with pytest.raises(RuntimeError, match="did not export a PDF"):
        rendering.export_docx_pdf(minimal_docx)
    assert not minimal_docx.with_suffix(".pdf").exists()
    assert _leftovers(minimal_docx.parent) == []


def test_source_changed_during_export_publishes_nothing(fake_word, minimal_docx):
    def export_while_word_saves(path):
        Path(path).write_bytes(b"%PDF-1.7 fake")
        with minimal_docx.open("ab") as handle:
            handle.write(b"saved elsewhere")

    fake_word.export = staticmethod(export_while_word_saves)
    with pytest.raises(RefusedUnsafe, match="changed on disk during the export"):
        rendering.export_docx_pdf(minimal_docx)
    assert not minimal_docx.with_suffix(".pdf").exists()
    assert _leftovers(minimal_docx.parent) == []


def test_warnings_name_revisions_comments_and_an_open_document(fake_word, minimal_docx, monkeypatch):
    real_audit = rendering.audit_docx
    monkeypatch.setattr(
        rendering,
        "audit_docx",
        lambda path: dataclasses.replace(real_audit(path), comment_count=2, is_locked=True),
    )
    monkeypatch.setattr(editing, "revision_total", lambda path: 5)

    result = rendering.export_docx_pdf(minimal_docx)

    assert result["revision_count"] == 5
    assert result["comment_count"] == 2
    joined = " | ".join(result["warnings"])
    assert "5 tracked revision" in joined and "every change accepted and no markup" in joined
    assert "2 comments" in joined
    assert "open in Word" in joined


def test_page_count_is_none_without_page_objects(tmp_path):
    pdf = tmp_path / "x.pdf"
    pdf.write_bytes(b"%PDF-1.7 /Type/ObjStm")
    assert rendering.pdf_page_count(pdf) is None


def test_render_shares_the_export_helper(monkeypatch, tmp_path):
    exported: list[Path] = []
    monkeypatch.setattr(rendering, "_export_pdf", lambda document, path: exported.append(path))
    monkeypatch.setattr(rendering, "IsolatedWordDocument", FakeSession)
    monkeypatch.setattr(rendering, "find_pdftoppm", lambda: "pdftoppm")
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    with pytest.raises(RuntimeError, match="failed to export PDF"):
        rendering.render_docx_pages(source, output_dir=tmp_path / "png")
    assert len(exported) == 1


def test_worker_dispatches_export_pdf(monkeypatch):
    seen = {}
    monkeypatch.setattr(rendering, "export_docx_pdf", lambda **payload: seen.update(payload) or {"ok": 1})
    assert operation_process._dispatch("export_pdf", {"docx_path": "a.docx", "output": None}) == {"ok": 1}
    assert seen == {"docx_path": "a.docx", "output": None}


def test_tool_is_a_write_left_out_of_read_only_mode():
    tools = {tool.name: tool for tool in asyncio.run(server.list_tools())}
    annotations = tools["export_pdf"].annotations
    assert annotations.read_only_hint is False
    assert annotations.destructive_hint is False


def test_tool_refuses_an_existing_target_without_starting_a_worker(minimal_docx, monkeypatch):
    target = minimal_docx.with_suffix(".pdf")
    target.write_bytes(b"keep")
    monkeypatch.setattr(
        operation_process, "run_operation", lambda *a, **k: pytest.fail("worker started for a refused export")
    )
    with pytest.raises(ToolError, match="already exists"):
        asyncio.run(server.call_tool("export_pdf", {"path": str(minimal_docx)}))
    assert target.read_bytes() == b"keep"


def test_tool_refuses_output_outside_root(minimal_docx, tmp_path, monkeypatch):
    from research_mcp import server as server_module

    root = minimal_docx.parent / "root_only"
    root.mkdir()
    inside = root / "paper.docx"
    inside.write_bytes(minimal_docx.read_bytes())
    monkeypatch.setattr(server_module, "ROOT_DIR", root.resolve())
    with pytest.raises(ToolError, match="outside the allowed root"):
        asyncio.run(server.call_tool("export_pdf", {"path": str(inside), "output": str(tmp_path / "x.pdf")}))
    assert not (tmp_path / "x.pdf").exists()
