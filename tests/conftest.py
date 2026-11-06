from __future__ import annotations

import os
from pathlib import Path
import zipfile

import pytest


DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Fixture paragraph.</w:t></w:r></w:p>
    <w:tbl>
      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>LoRA</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Early</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>Middle</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>Late</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="KeepMe"/>
      <w:r><w:t>Bookmarked.</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    <w:p><w:hyperlink w:anchor="KeepMe"><w:r><w:t>jump</w:t></w:r></w:hyperlink></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"""


@pytest.fixture(autouse=True)
def isolated_undo_store(tmp_path_factory):
    """Keep the checkpoint store of every test out of the real %LOCALAPPDATA%\\research_mcp.

    Its own MonkeyPatch, so a test's ``monkeypatch.undo()`` cannot point the store back at the real one."""
    home = tmp_path_factory.mktemp("research_mcp_home")
    with pytest.MonkeyPatch.context() as mp:
        mp.setenv("RESEARCH_MCP_HOME", str(home))
        yield home


TEST_AUTHOR = "Test Author"  # every test's default comment author (initials "TA")


@pytest.fixture(autouse=True)
def isolated_config(tmp_path_factory):
    """Point the config file at a scratch file and pin the comment author, so the user's own
    ``%APPDATA%\\research-mcp\\config.env`` and Word user name never change a test's result.

    The env vars reach operation subprocesses too. Live Zotero runs (RESEARCH_MCP_ZOTERO_LIVE=1) copy
    only the ``ZOTERO_*`` lines of the real file into the scratch one (values are never printed)."""
    from research_mcp import config

    scratch = tmp_path_factory.mktemp("research_mcp_config") / "config.env"
    with pytest.MonkeyPatch.context() as mp:
        real = config.config_path()
        if os.environ.get("RESEARCH_MCP_ZOTERO_LIVE") == "1" and real.exists():
            lines = [line for line in real.read_text(encoding="utf-8-sig").splitlines()
                     if line.strip().startswith("ZOTERO_")]
            scratch.write_text("\n".join(lines) + "\n", encoding="utf-8")
        mp.setenv("RESEARCH_MCP_CONFIG", str(scratch))
        mp.setenv("RESEARCH_MCP_AUTHOR", TEST_AUTHOR)
        for name in ("RESEARCH_MCP_AUTHOR_INITIALS", "RESEARCH_MCP_ISSUES"):
            mp.delenv(name, raising=False)
        config.reload()
        yield scratch
    config.reload()


@pytest.fixture
def minimal_docx(tmp_path: Path) -> Path:
    path = tmp_path / "fixture.docx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>")
        archive.writestr("word/document.xml", DOCUMENT_XML)
        archive.writestr("word/media/image1.png", b"fixture")
    return path

