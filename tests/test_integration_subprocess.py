from __future__ import annotations

import subprocess

from research_mcp.integration import _hidden_subprocess_kwargs


def test_integration_helper_uses_hidden_windows_subprocess_options():
    options = _hidden_subprocess_kwargs()
    assert options["creationflags"] & subprocess.CREATE_NO_WINDOW
    assert options["startupinfo"].lpDesktop == r"WinSta0\Default"
