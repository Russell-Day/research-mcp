"""Plan 916 S1: Word 97-2003 .doc files.

Reading tools read a .doc through a temporary .docx; write tools refuse it with one line naming
convert_document; convert_document saves a new .docx beside it, never over an existing file. Detection is by
the OLE2 header, not the extension. Word is faked here (the conversion copies a prepared .docx); the live
test at the end (RESEARCH_MCP_WORD_LIVE=1) makes a real .doc with Word and runs the same calls.
"""

from __future__ import annotations

import asyncio
import os
from pathlib import Path
import shutil

from mcp.server.mcpserver.exceptions import ToolError
import pytest

from research_mcp import legacy_doc, operation_process
from research_mcp.audit import audit_docx, document_structure_errors, file_sha256
from research_mcp.errors import RefusedUnsafe
from research_mcp.legacy_doc import (
    OLE2_HEADER,
    TEMP_PREFIX,
    is_legacy_doc,
    macro_refusal,
    not_a_package,
    password_refusal,
)
from research_mcp.paragraphs import inspect_document as inspect_docx, read_snapshot
from research_mcp.server import server

LIVE = os.environ.get("RESEARCH_MCP_WORD_LIVE") == "1"

# Tools that take a document path but must not refuse a .doc: they read it through a temporary .docx, or
# (convert_document) exist to convert it; set_project_collection's path is a folder.
READS_DOC = {"inspect_document", "audit_document", "render_document_pages", "export_pdf", "convert_document"}
NOT_A_DOCUMENT = {"set_project_collection"}
# Read-only tools that parse the .docx package themselves refuse a .doc with read_refusal.
PARSES_PACKAGE = {"list_bookmarks", "list_revisions", "audit_document_style"}


def fake_ole(path: Path) -> Path:
    path.write_bytes(OLE2_HEADER + bytes(range(256)) * 8)
    return path


def call(name: str, arguments: dict):
    return asyncio.run(server.call_tool(name, arguments)).structured_content


def refused(name: str, arguments: dict) -> str:
    with pytest.raises(ToolError) as excinfo:
        asyncio.run(server.call_tool(name, arguments))
    return str(excinfo.value)


def listing(folder: Path) -> dict[str, str]:
    return {p.name: (file_sha256(p) if p.is_file() else "dir") for p in folder.iterdir()}


class FakeWord:
    """Stands in for the Word session in the conversion: SaveAs2 copies ``docx`` to the requested file."""

    def __init__(self, docx: Path):
        self.docx = docx
        self.sessions: list[dict] = []
        self.saves: list[dict] = []
        self.has_vb_project = False
        self.open_error: Exception | None = None
        self.on_save = None  # called with the source path while Word "saves" (a save racing the conversion)

    def session_class(self):
        fake = self

        class Session:
            def __init__(self, path, *, visible=False, read_only=False, timeout_seconds=30.0):
                fake.sessions.append({"path": Path(path), "read_only": read_only, "visible": visible})

                class Document:
                    HasVBProject = fake.has_vb_project

                    def SaveAs2(self, **kwargs):
                        fake.saves.append(kwargs)
                        shutil.copyfile(fake.docx, kwargs["FileName"])
                        if fake.on_save is not None:
                            fake.on_save(Path(path))

                self.document = Document()

            def __enter__(self):
                if fake.open_error is not None:
                    raise fake.open_error
                return self

            def __exit__(self, *exc):
                return False

        return Session


@pytest.fixture
def fake_word(monkeypatch, minimal_docx):
    """Fake Word, and run_operation in this process: convert_doc really runs convert_to_docx; render and
    export (Word) write a fake page or PDF where asked and answer with a plausible result. ``before_convert``,
    ``after_convert`` and ``during_word`` (called with the operation's kind) let a test change the source at
    each boundary."""
    word = FakeWord(minimal_docx)
    monkeypatch.setattr(legacy_doc, "_legacy_session_class", word.session_class)
    word.operations = []
    word.before_convert = word.after_convert = word.during_word = None

    def run_operation(kind, payload, *, timeout_seconds=120):
        word.operations.append((kind, dict(payload)))
        if kind == "convert_doc":
            if word.before_convert:
                word.before_convert()
            result = operation_process._dispatch(kind, payload)
            if word.after_convert:
                word.after_convert()
            return result
        docx = Path(payload["docx_path"])
        assert docx.is_file(), "the temporary .docx must exist while Word reads it"
        word.read_docx = docx
        if kind == "render":
            folder = Path(payload["output_dir"])
            folder.mkdir(parents=True, exist_ok=True)
            (folder / "page-1.png").write_bytes(b"png")
            result = {"source": str(docx), "output_directory": str(folder), "total_rendered": 1,
                      "dpi": payload["dpi"], "markup": payload["markup"], "warnings": [],
                      "pages": [{"page_number": 1, "image_path": str(folder / "page-1.png"), "size_bytes": 3}]}
        elif kind == "export_pdf":
            Path(payload["output"]).write_bytes(b"%PDF-1.7 fake")
            result = {"source": str(docx), "source_sha256": file_sha256(docx), "output": payload["output"],
                      "size_bytes": 13, "page_count": 1, "revision_count": 0, "comment_count": 0, "warnings": [],
                      "elapsed_seconds": 0.0}
        else:
            raise AssertionError(kind)
        if word.during_word:
            word.during_word(kind)
        return result

    monkeypatch.setattr(operation_process, "run_operation", run_operation)
    return word


@pytest.fixture
def doc_file(tmp_path) -> Path:
    folder = tmp_path / "docs"
    folder.mkdir()
    return fake_ole(folder / "cv.doc")


# --- B1: detection by header --------------------------------------------------------------------------------


def test_detection_is_by_header_not_extension(tmp_path, minimal_docx):
    assert is_legacy_doc(fake_ole(tmp_path / "a.doc"))
    assert is_legacy_doc(fake_ole(tmp_path / "b.docx"))
    zip_named_doc = tmp_path / "c.doc"
    shutil.copyfile(minimal_docx, zip_named_doc)
    assert not is_legacy_doc(zip_named_doc)
    assert not is_legacy_doc(minimal_docx)
    assert not is_legacy_doc(tmp_path / "missing.doc")


def test_low_level_readers_name_convert_document_instead_of_a_zip_error(doc_file):
    # The real .doc that prompted plan 916 failed with "There is no item named 'word/document.xml'".
    before = file_sha256(doc_file)
    with pytest.raises(RefusedUnsafe) as excinfo:
        audit_docx(doc_file)
    assert str(excinfo.value) == not_a_package(doc_file.resolve())
    assert document_structure_errors(doc_file) == (f"package: {not_a_package(doc_file.resolve())}",)
    with pytest.raises(RefusedUnsafe, match="convert_document"):
        read_snapshot(doc_file)
    assert file_sha256(doc_file) == before


# --- B2/B3/B4: reading tools read a .doc through a temporary .docx -------------------------------------------


def test_inspect_reads_a_doc_through_a_temporary_docx(fake_word, doc_file, minimal_docx):
    before, around = file_sha256(doc_file), listing(doc_file.parent)

    result = call("inspect_document", {"path": str(doc_file), "first_paragraph": 1, "last_paragraph": 50})

    expected = inspect_docx(minimal_docx, first_paragraph=1, last_paragraph=50)
    assert result["converted_from_doc"] is True
    assert result["path"] == str(doc_file.resolve())
    assert result["source_sha256"] == before
    assert result["paragraph_count"] == expected["paragraph_count"] > 0
    assert [p["text"] for p in result["paragraphs"]] == [p["text"] for p in expected["paragraphs"]]
    # Word opened the .doc read-only and saved Word's default format; the temporary is gone.
    assert fake_word.sessions == [{"path": doc_file.resolve(), "read_only": True, "visible": False}]
    assert fake_word.saves[0]["FileFormat"] == 16
    temporary = Path(fake_word.operations[0][1]["output"])
    assert not temporary.exists() and not temporary.parent.exists()
    assert not Path(fake_word.operations[0][1]["temporary"]).exists()
    assert file_sha256(doc_file) == before and listing(doc_file.parent) == around


def test_audit_reads_a_doc(fake_word, doc_file, minimal_docx):
    before = file_sha256(doc_file)

    result = call("audit_document", {"path": str(doc_file)})

    expected = audit_docx(minimal_docx)
    assert result["converted_from_doc"] is True
    assert result["path"] == str(doc_file.resolve())
    assert result["sha256"] == before
    assert result["paragraph_count"] == expected.paragraph_count
    assert result["table_count"] == expected.table_count
    assert file_sha256(doc_file) == before


def test_render_and_export_a_doc_write_beside_the_doc(fake_word, doc_file):
    before = file_sha256(doc_file)

    rendered = call("render_document_pages", {"path": str(doc_file)})
    assert rendered["converted_from_doc"] is True
    assert rendered["source"] == str(doc_file.resolve())
    folder = doc_file.parent.resolve() / f"cv_rendered_{before[:8]}"
    assert rendered["output_directory"] == str(folder)
    assert [p["image_path"] for p in rendered["pages"]] == [str(folder / "page-1.png")]
    assert (folder / "page-1.png").read_bytes() == b"png"
    marked = call("render_document_pages", {"path": str(doc_file), "markup": True})
    assert marked["output_directory"] == str(doc_file.parent.resolve() / f"cv_rendered_markup_{before[:8]}")

    exported = call("export_pdf", {"path": str(doc_file)})
    assert exported["converted_from_doc"] is True
    assert exported["source"] == str(doc_file.resolve())
    assert exported["source_sha256"] == before
    pdf = doc_file.with_suffix(".pdf").resolve()
    assert exported["output"] == str(pdf) and pdf.read_bytes() == b"%PDF-1.7 fake"
    assert exported["size_bytes"] == pdf.stat().st_size
    assert fake_word.read_docx.suffix == ".docx" and not fake_word.read_docx.exists()
    assert not [p for p in doc_file.parent.iterdir() if p.name.startswith(".~research-mcp-")]
    assert file_sha256(doc_file) == before


def test_export_of_a_doc_refuses_an_existing_pdf_before_converting(fake_word, doc_file):
    before = file_sha256(doc_file)
    doc_file.with_suffix(".pdf").write_bytes(b"keep")
    message = refused("export_pdf", {"path": str(doc_file)})
    assert "never overwritten" in message
    assert fake_word.operations == [] and doc_file.with_suffix(".pdf").read_bytes() == b"keep"
    assert file_sha256(doc_file) == before


def test_a_docx_is_read_as_before(fake_word, minimal_docx):
    before = file_sha256(minimal_docx)
    result = call("inspect_document", {"path": str(minimal_docx)})
    assert result["converted_from_doc"] is None  # null for a .docx, never false or absent
    assert "converted_from_doc" in result
    assert call("audit_document", {"path": str(minimal_docx)})["converted_from_doc"] is None
    assert fake_word.operations == []
    assert file_sha256(minimal_docx) == before


# --- B11: the result's sha is the bytes Word converted; a later change refuses -------------------------------


def _resave(doc: Path) -> str:
    """A Word or OneDrive save of the .doc: new bytes, still a .doc. Returns the new sha256."""
    doc.write_bytes(doc.read_bytes() + b"saved again")
    return file_sha256(doc)


def test_a_save_before_the_conversion_is_reported_with_the_sha_word_converted(fake_word, doc_file):
    first = file_sha256(doc_file)
    saved = {}
    fake_word.before_convert = lambda: saved.setdefault("sha", _resave(doc_file))

    inspected = call("inspect_document", {"path": str(doc_file)})

    assert saved["sha"] != first
    assert inspected["source_sha256"] == saved["sha"] == file_sha256(doc_file)
    fake_word.before_convert = lambda: saved.__setitem__("sha", _resave(doc_file))
    audited = call("audit_document", {"path": str(doc_file)})
    assert audited["sha256"] == saved["sha"] == file_sha256(doc_file)


@pytest.mark.parametrize("tool", ["inspect_document", "audit_document"])
def test_a_save_after_the_conversion_refuses_the_read(fake_word, doc_file, tool):
    converted = file_sha256(doc_file)
    changed = {}
    fake_word.after_convert = lambda: changed.setdefault("sha", _resave(doc_file))

    message = refused(tool, {"path": str(doc_file)})

    assert message.endswith(
        f"{tool}: cv.doc changed on disk while it was being read (sha256 {converted} -> {changed['sha']}); "
        "nothing was returned or written. Try again.")


def test_a_save_after_the_conversion_publishes_no_pdf(fake_word, doc_file):
    converted = file_sha256(doc_file)
    changed = {}
    fake_word.during_word = lambda kind: changed.setdefault("sha", _resave(doc_file))

    message = refused("export_pdf", {"path": str(doc_file)})

    assert message.endswith(
        f"export_pdf: cv.doc changed on disk while it was being read (sha256 {converted} -> {changed['sha']}); "
        "nothing was returned or written. Try again.")
    assert sorted(p.name for p in doc_file.parent.iterdir()) == ["cv.doc"]


def test_a_save_after_the_conversion_publishes_no_pages(fake_word, doc_file):
    converted = file_sha256(doc_file)
    changed = {}
    fake_word.during_word = lambda kind: changed.setdefault("sha", _resave(doc_file))

    message = refused("render_document_pages", {"path": str(doc_file)})

    assert message.endswith(
        f"render_document_pages: cv.doc changed on disk while it was being read (sha256 {converted} -> "
        f"{changed['sha']}); nothing was returned or written. Try again.")
    assert sorted(p.name for p in doc_file.parent.iterdir()) == ["cv.doc"]


def test_a_save_during_the_conversion_writes_no_docx(fake_word, doc_file):
    first = file_sha256(doc_file)
    changed = {}
    fake_word.on_save = lambda path: changed.setdefault("sha", _resave(path))

    message = refused("convert_document", {"path": str(doc_file)})

    assert message.endswith(
        f"convert_document: cv.doc changed on disk during the conversion (sha256 {first} -> {changed['sha']}); "
        "nothing was written. Try again.")
    assert sorted(p.name for p in doc_file.parent.iterdir()) == ["cv.doc"]


def test_a_worker_stopped_at_its_deadline_leaves_no_temporary(fake_word, doc_file, monkeypatch):
    from research_mcp.errors import IntegrationRefused

    before = file_sha256(doc_file)

    def stopped(kind, payload, *, timeout_seconds=120):
        Path(payload["temporary"]).write_bytes(b"half a .docx")  # Word was saving when the worker was killed
        raise IntegrationRefused("worker exceeded timeout_seconds=180")

    monkeypatch.setattr(operation_process, "run_operation", stopped)
    assert "timeout_seconds" in refused("convert_document", {"path": str(doc_file)})
    assert sorted(p.name for p in doc_file.parent.iterdir()) == ["cv.doc"]
    assert file_sha256(doc_file) == before


# --- B5: every write tool refuses a .doc ----------------------------------------------------------------------


def _arguments(tool, path: Path, tmp_path: Path) -> dict:
    """The required arguments of ``tool``, with ``path`` as its document."""
    args = {}
    for name in tool.input_schema.get("required", []):
        schema = tool.input_schema["properties"][name]
        kinds = {schema.get("type")} | {s.get("type") for s in schema.get("anyOf", [])}
        if name in {"source", "path"}:
            args[name] = str(path)
        elif name == "source_sha256":
            args[name] = "0" * 64
        elif name == "output":
            args[name] = str(tmp_path / "out.docx")
        elif "array" in kinds:
            args[name] = []
        elif "integer" in kinds:
            args[name] = 1
        elif "boolean" in kinds:
            args[name] = False
        else:
            args[name] = "x"
    return args


def _document_tools():
    for tool in asyncio.run(server.list_tools()):
        # A document tool requires its document; an optional ``path`` (search_zotero, add_zotero_item_by_doi)
        # only locates a collection marker.
        param = next((p for p in ("source", "path") if p in tool.input_schema.get("required", [])), None)
        if param and tool.name not in READS_DOC | NOT_A_DOCUMENT:
            yield tool, param


def test_every_write_tool_refuses_a_doc_with_one_line(fake_word, doc_file, tmp_path):
    before, around = file_sha256(doc_file), listing(doc_file.parent)
    checked = set()

    for tool, param in _document_tools():
        if tool.annotations.read_only_hint:
            continue
        message = refused(tool.name, _arguments(tool, doc_file, tmp_path))
        # The required text, written out here (not built with the code's own write_refusal).
        expected = f"{tool.name} can't edit a .doc; call convert_document(path) to make cv.docx beside it and edit that."
        assert message == f"Error executing tool {tool.name}: {expected}", (tool.name, message)
        checked.add(tool.name)

    # The enumeration found the write tools (a tool added later is checked without editing this test).
    assert {"edit_document", "style_text", "format_tables", "clean_bookmarks", "insert_zotero_citation",
            "edit_citation_cluster", "refresh_citations", "set_citation_style", "apply_document_style",
            "undo_document", "revert_document", "checkpoint_document"} <= checked
    assert fake_word.operations == [] and fake_word.sessions == []
    assert file_sha256(doc_file) == before and listing(doc_file.parent) == around
    assert not (tmp_path / "out.docx").exists()


def test_read_tools_that_parse_the_package_refuse_a_doc(fake_word, doc_file, tmp_path):
    before = file_sha256(doc_file)
    read_only = {tool.name for tool, _ in _document_tools() if tool.annotations.read_only_hint}
    assert PARSES_PACKAGE <= read_only
    for name in sorted(PARSES_PACKAGE):
        tool = next(t for t, _ in _document_tools() if t.name == name)
        message = refused(name, _arguments(tool, doc_file, tmp_path))
        assert message == (
            f"Error executing tool {name}: {name} reads .docx files and cv.doc is a Word 97-2003 .doc; call "
            "convert_document(path) to make a .docx and read that (inspect_document, audit_document, "
            "render_document_pages and export_pdf read a .doc as is)."), (name, message)
    assert file_sha256(doc_file) == before


# --- B6/B7/B8: convert_document ------------------------------------------------------------------------------


def test_convert_writes_a_docx_beside_the_doc(fake_word, doc_file, minimal_docx, isolated_undo_store):
    before = file_sha256(doc_file)

    result = call("convert_document", {"path": str(doc_file)})

    # No undo entry: nothing that existed changed (the store of this test is its own, see conftest).
    store = isolated_undo_store / "checkpoints"
    assert not store.exists() or not [p for p in store.rglob("*") if p.is_file()]
    assert call("document_history", {"path": result["output"]})["entries"] == []

    target = doc_file.with_suffix(".docx").resolve()
    assert result["output"] == str(target)
    assert result["output_sha256"] == file_sha256(target) == file_sha256(minimal_docx)
    assert result["source"] == str(doc_file.resolve()) and result["source_sha256"] == before
    assert result["size_bytes"] == target.stat().st_size
    assert fake_word.sessions[0]["read_only"] is True
    saved = fake_word.saves[0]
    assert saved["FileFormat"] == 16 and saved["AddToRecentFiles"] is False
    # Word saved a hidden temporary in the target folder, published with a no-replace rename.
    assert Path(saved["FileName"]).parent == target.parent and Path(saved["FileName"]).name.startswith(TEMP_PREFIX)
    assert sorted(p.name for p in doc_file.parent.iterdir()) == ["cv.doc", "cv.docx"]
    assert file_sha256(doc_file) == before
    # The new file is a .docx every tool reads.
    assert call("inspect_document", {"path": result["output"]})["converted_from_doc"] is None


def test_second_convert_refuses_and_writes_nothing(fake_word, doc_file):
    first = call("convert_document", {"path": str(doc_file)})
    around = listing(doc_file.parent)

    message = refused("convert_document", {"path": str(doc_file)})

    assert "already exists and is never overwritten" in message
    assert len(fake_word.sessions) == 1  # refused before Word started
    assert listing(doc_file.parent) == around
    assert file_sha256(first["output"]) == first["output_sha256"]
    assert file_sha256(doc_file) == first["source_sha256"]


def test_convert_output_elsewhere(fake_word, doc_file, tmp_path):
    other = tmp_path / "elsewhere"
    other.mkdir()
    before = file_sha256(doc_file)

    result = call("convert_document", {"path": str(doc_file), "output": str(other / "new name.docx")})
    assert result["output"] == str((other / "new name.docx").resolve())
    assert (other / "new name.docx").is_file() and not doc_file.with_suffix(".docx").exists()

    assert "must end in .docx" in refused("convert_document", {"path": str(doc_file), "output": str(other / "x.pdf")})
    assert "folder does not exist" in refused(
        "convert_document", {"path": str(doc_file), "output": str(tmp_path / "missing" / "x.docx")})
    assert sorted(p.name for p in other.iterdir()) == ["new name.docx"]
    assert len(fake_word.sessions) == 1
    assert file_sha256(doc_file) == before


def test_failed_conversion_leaves_no_file(fake_word, doc_file):
    fake_word.docx = doc_file  # Word "saves" something that is no .docx package
    before, around = file_sha256(doc_file), listing(doc_file.parent)
    message = refused("convert_document", {"path": str(doc_file)})
    assert "package check" in message
    assert listing(doc_file.parent) == around and file_sha256(doc_file) == before


# --- B9: header wins over the name ---------------------------------------------------------------------------


def test_a_zip_named_doc_is_handled_as_a_docx(fake_word, minimal_docx, tmp_path):
    named_doc = tmp_path / "package.doc"
    shutil.copyfile(minimal_docx, named_doc)
    before = file_sha256(named_doc)

    assert call("inspect_document", {"path": str(named_doc)})["converted_from_doc"] is None
    assert call("audit_document", {"path": str(named_doc)})["converted_from_doc"] is None
    edit = call("edit_document", {"source": str(named_doc), "output": str(named_doc), "source_sha256": "0" * 64, "operations": [
        {"op_id": "a", "type": "replace_text", "paragraph_index": 1, "find": "x", "replace": "y"}]})
    assert edit["status"] == "refused" and "convert_document" not in edit["message"]
    assert "already a .docx package" in refused("convert_document", {"path": str(named_doc)})
    assert fake_word.operations == []
    assert file_sha256(named_doc) == before


def test_a_doc_named_docx_is_detected(fake_word, tmp_path):
    disguised = fake_ole(tmp_path / "old.docx")
    before = file_sha256(disguised)

    assert call("inspect_document", {"path": str(disguised)})["converted_from_doc"] is True
    message = refused("edit_document", {"source": str(disguised), "output": str(disguised),
                                        "source_sha256": "0" * 64, "operations": []})
    assert message == (
        "Error executing tool edit_document: edit_document can't edit old.docx: it is a Word 97-2003 .doc file "
        "named .docx; call convert_document(path, output=<new .docx path>) and edit that.")
    assert "pass output=" in refused("convert_document", {"path": str(disguised)})
    result = call("convert_document", {"path": str(disguised), "output": str(tmp_path / "old-converted.docx")})
    assert Path(result["output"]).is_file()
    assert file_sha256(disguised) == before


# --- B10: protected or macro documents ------------------------------------------------------------------------


def test_macro_doc_is_refused(fake_word, doc_file):
    fake_word.has_vb_project = True
    before, around = file_sha256(doc_file), listing(doc_file.parent)
    assert refused("convert_document", {"path": str(doc_file)}).endswith(macro_refusal(doc_file.resolve()))
    assert refused("inspect_document", {"path": str(doc_file)}).endswith(macro_refusal(doc_file.resolve()))
    assert fake_word.saves == []
    assert file_sha256(doc_file) == before and listing(doc_file.parent) == around


def test_password_error_from_word_becomes_a_clear_refusal(monkeypatch, doc_file):
    """The real LegacyWordDocument session, with Word's COM layer faked: a password error is a refusal."""
    from research_mcp import word, word_com

    opened = {}

    class FakeDocuments:
        def Open(self, **kwargs):
            opened.update(kwargs)
            raise Exception("The password is incorrect. Word cannot open the document.")

    class App:
        Documents = FakeDocuments()

    monkeypatch.setattr(word_com, "co_initialize", lambda: None)
    monkeypatch.setattr(word_com, "create_word_application", lambda visible=False: App())
    monkeypatch.setattr(word_com, "word_pid", lambda app: None)
    monkeypatch.setattr(word.IsolatedWordDocument, "_start_watchdog", lambda self: None)
    monkeypatch.setattr(word.IsolatedWordDocument, "_force_cleanup", lambda self: None)

    before = file_sha256(doc_file)
    session = legacy_doc._legacy_session_class()(doc_file, read_only=True)
    with pytest.raises(RefusedUnsafe) as excinfo:
        session.__enter__()
    assert file_sha256(doc_file) == before
    assert str(excinfo.value) == password_refusal(doc_file.resolve())
    assert opened["ReadOnly"] is True and opened["AddToRecentFiles"] is False
    assert opened["PasswordDocument"] == legacy_doc.NO_PASSWORD  # never a prompt


def test_password_refusal_reaches_the_reading_tools(fake_word, doc_file):
    fake_word.open_error = RefusedUnsafe(password_refusal(doc_file.resolve()))
    before = file_sha256(doc_file)
    assert refused("inspect_document", {"path": str(doc_file)}).endswith(password_refusal(doc_file.resolve()))
    assert refused("convert_document", {"path": str(doc_file)}).endswith(password_refusal(doc_file.resolve()))
    assert not doc_file.with_suffix(".docx").exists()
    assert file_sha256(doc_file) == before


# --- live: a real .doc made by Word ---------------------------------------------------------------------------

CV_LINES = [
    "Alex Example",
    "Education",
    "Example State University, B.S. in Biology, 2024",
    "Experience",
    "Research Assistant, Example Lab, 2023-2024",
    "Built a data pipeline for imaging studies.",
]


def _make_doc(tmp_path: Path, name: str, *, password: str | None = None) -> Path:
    import docx  # python-docx, dev dependency
    from research_mcp.word import IsolatedWordDocument

    source = tmp_path / f"{name}-source.docx"
    document = docx.Document()
    for line in CV_LINES:
        document.add_paragraph(line)
    document.core_properties.author = "Reviewer B"
    document.save(source)
    target = tmp_path / f"{name}.doc"
    with IsolatedWordDocument(source, read_only=True, timeout_seconds=90) as session:
        if password:
            session.document.Password = password  # SaveAs2(Password=...) on a read-only document is ignored
        session.document.SaveAs2(FileName=str(target), FileFormat=0, AddToRecentFiles=False)
    assert is_legacy_doc(target)
    return target


@pytest.mark.skipif(not LIVE, reason="requires Word (RESEARCH_MCP_WORD_LIVE=1)")
def test_live_doc_read_refused_converted(tmp_path):
    doc = _make_doc(tmp_path, "cv")
    before = file_sha256(doc)

    inspected = call("inspect_document", {"path": str(doc), "first_paragraph": 1, "last_paragraph": 50})
    assert inspected["converted_from_doc"] is True and inspected["source_sha256"] == before
    assert [p["text"] for p in inspected["paragraphs"] if p["text"]] == CV_LINES
    audited = call("audit_document", {"path": str(doc)})
    assert audited["converted_from_doc"] is True and audited["sha256"] == before

    rendered = call("render_document_pages", {"path": str(doc), "pages": [1], "dpi": 50})
    assert rendered["total_rendered"] == 1 and Path(rendered["output_directory"]).parent == doc.parent.resolve()
    exported = call("export_pdf", {"path": str(doc)})
    assert Path(exported["output"]) == doc.with_suffix(".pdf").resolve() and exported["page_count"] == 1

    message = refused("edit_document", {"source": str(doc), "output": str(doc), "source_sha256": before, "operations": [
        {"op_id": "a", "type": "replace_text", "paragraph_index": 1, "find": "Alex", "replace": "Sam"}]})
    assert message == ("Error executing tool edit_document: edit_document can't edit a .doc; call "
                       "convert_document(path) to make cv.docx beside it and edit that.")
    assert file_sha256(doc) == before

    converted = call("convert_document", {"path": str(doc)})
    target = doc.with_suffix(".docx").resolve()
    assert converted["output"] == str(target) and converted["output_sha256"] == file_sha256(target)
    assert not document_structure_errors(target)
    assert [p["text"] for p in inspect_docx(target, first_paragraph=1, last_paragraph=50)["paragraphs"]
            if p["text"]] == CV_LINES
    assert "already exists" in refused("convert_document", {"path": str(doc)})
    assert file_sha256(target) == converted["output_sha256"]
    assert not [p for p in doc.parent.iterdir() if p.name.startswith(".~research-mcp-")]
    assert file_sha256(doc) == before


@pytest.mark.skipif(not LIVE, reason="requires Word (RESEARCH_MCP_WORD_LIVE=1)")
def test_live_password_doc_is_refused(tmp_path):
    doc = _make_doc(tmp_path, "locked", password="example-pw")
    before = file_sha256(doc)
    assert refused("inspect_document", {"path": str(doc)}).endswith(password_refusal(doc.resolve()))
    assert refused("convert_document", {"path": str(doc)}).endswith(password_refusal(doc.resolve()))
    assert not doc.with_suffix(".docx").exists()
    assert file_sha256(doc) == before
