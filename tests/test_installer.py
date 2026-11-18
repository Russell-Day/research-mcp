"""installer/research-mcp-install.cmd, the double-click installer attached to each GitHub release."""

from __future__ import annotations

import re
from pathlib import Path

CMD = Path(__file__).resolve().parents[1] / "installer" / "research-mcp-install.cmd"


def test_installer_is_plain_ascii_with_crlf_line_endings():
    data = CMD.read_bytes()
    data.decode("ascii")
    assert b"\r\n" in data and b"\n" not in data.replace(b"\r\n", b"")


def test_installer_installs_uv_then_runs_setup_and_pauses():
    text = CMD.read_text(encoding="ascii")
    assert "https://astral.sh/uv/install.ps1" in text
    assert "uvx research-mcp@latest setup" in text
    assert text.rstrip().splitlines()[-2:] == ["pause", "exit /b %rc%"]  # a script calling it sees setup's exit code
    # Only astral.sh is fetched directly; everything else comes through uv.
    assert set(re.findall(r"https?://[^/\s\"]+", text)) <= {"https://astral.sh", "https://docs.astral.sh"}


def test_installer_blocks_have_no_stray_parentheses():
    """A ")" inside an echo in a ( ... ) block ends the block early in cmd.exe."""
    depth = 0
    for line in CMD.read_text(encoding="ascii").splitlines():
        stripped = line.strip()
        if stripped.lower().startswith("rem "):
            continue
        if depth and stripped.lower().startswith("echo"):
            assert ")" not in stripped and "(" not in stripped, line
        depth += stripped.count("(") - stripped.count(")")
        assert depth >= 0, line
    assert depth == 0
