from __future__ import annotations

from pathlib import Path
import subprocess

from research_mcp import rendering


def test_rendering_opens_word_read_only_and_hides_pdftoppm(monkeypatch, tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    calls: dict[str, object] = {}

    class FakeDocument:
        def ExportAsFixedFormat(self, **kwargs):
            calls["export"] = kwargs
            Path(kwargs["OutputFileName"]).write_bytes(b"pdf")

    class FakeSession:
        document = FakeDocument()

        def __init__(self, path, *, visible=False, read_only=False):
            calls["path"] = path
            calls["visible"] = visible
            calls["read_only"] = read_only

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

    def fake_run(cmd, **kwargs):
        calls["cmd"] = cmd
        calls["run_kwargs"] = kwargs
        Path(cmd[-1] + "-1.png").write_bytes(b"png")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(rendering, "IsolatedWordDocument", FakeSession)
    monkeypatch.setattr(rendering, "find_pdftoppm", lambda: "pdftoppm.exe")
    monkeypatch.setattr(rendering.subprocess, "run", fake_run)

    result = rendering.render_docx_pages(source, output_dir=tmp_path / "rendered", pages=[1], dpi=300)

    assert calls["path"] == source.resolve()
    assert calls["visible"] is False
    assert calls["read_only"] is True
    assert calls["export"]["Item"] == 0
    assert calls["run_kwargs"]["creationflags"] == getattr(subprocess, "CREATE_NO_WINDOW", 0)
    assert calls["run_kwargs"]["timeout"] == 120
    assert result["total_rendered"] == 1
    assert result["pages"][0]["page_number"] == 1


def test_rendering_passes_requested_page_range_to_pdftoppm(monkeypatch, tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    commands: list[list[str]] = []

    class FakeDocument:
        def ExportAsFixedFormat(self, **kwargs):
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
        commands.append(cmd)
        Path(cmd[-1] + "-2.png").write_bytes(b"png")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(rendering, "IsolatedWordDocument", FakeSession)
    monkeypatch.setattr(rendering, "find_pdftoppm", lambda: "pdftoppm")
    monkeypatch.setattr(rendering.subprocess, "run", fake_run)

    rendering.render_docx_pages(source, output_dir=tmp_path / "rendered", pages=[2, 4])

    assert commands[0][1:7] == ["-png", "-r", "200", "-f", "2", "-l"]
    assert commands[0][7] == "4"


def test_rendering_reused_output_does_not_return_stale_pages(monkeypatch, tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"fixture")
    output_dir = tmp_path / "rendered"
    output_dir.mkdir()
    stale = output_dir / "page-99.png"
    stale.write_bytes(b"stale")

    class FakeDocument:
        def ExportAsFixedFormat(self, **kwargs):
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
        # The page is written beside the isolated temporary prefix supplied
        # to pdftoppm, never beside the reused public output directory.
        Path(cmd[-1] + "-1.png").write_bytes(b"current")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(rendering, "IsolatedWordDocument", FakeSession)
    monkeypatch.setattr(rendering, "find_pdftoppm", lambda: "pdftoppm")
    monkeypatch.setattr(rendering.subprocess, "run", fake_run)

    result = rendering.render_docx_pages(source, output_dir=output_dir)

    assert result["total_rendered"] == 1
    assert [page["page_number"] for page in result["pages"]] == [1]
    assert (output_dir / "page-1.png").read_bytes() == b"current"
    assert stale.read_bytes() == b"stale"
