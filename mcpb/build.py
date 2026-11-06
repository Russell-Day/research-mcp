"""Build the Claude Desktop bundle: ``research-mcp-<version>.mcpb``.

    python mcpb/build.py [--out DIR] [--keep-stage]

The bundle uses the MCPB ``uv`` server type: it holds the project source (pyproject.toml, uv.lock,
src/research_mcp) and Claude Desktop runs it with uv, installing Python and the dependencies itself.
Nothing is downloaded from PyPI for the package, so a bundle works before (or without) a PyPI release.

Needs Node.js (``npx``). Uses the official CLI, ``npx @anthropic-ai/mcpb``, to validate and pack.
Standard library only, so it runs under any Python 3.12+.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

MCPB_CLI = "@anthropic-ai/mcpb@2.1.2"
HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
# Files staged next to the manifest. README.md and LICENSE come from public/ when that overlay exists
# (the private tree), otherwise from the repo root (the public repo).
TOP_FILES = ("pyproject.toml", "uv.lock", "README.md", "LICENSE")
SKIP = shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo")


def package_version(root: Path = ROOT) -> str:
    text = (root / "src" / "research_mcp" / "__init__.py").read_text(encoding="utf-8")
    match = re.search(r'^__version__\s*=\s*"([^"]+)"', text, re.M)
    if not match:
        raise SystemExit("could not read __version__ from src/research_mcp/__init__.py")
    return match.group(1)


def _source(name: str, root: Path) -> Path:
    overlay = root / "public" / name
    return overlay if overlay.exists() else root / name


def stage(dest: Path, root: Path = ROOT) -> Path:
    """Copy the bundle's files into ``dest`` and write its manifest with the package version."""
    manifest = json.loads((HERE / "manifest.json").read_text(encoding="utf-8"))
    manifest["version"] = package_version(root)
    dest.mkdir(parents=True, exist_ok=True)
    (dest / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    for name in TOP_FILES:
        src = _source(name, root)
        if not src.exists():
            raise SystemExit(f"missing {name} (looked in {root / 'public'} and {root})")
        shutil.copy2(src, dest / name)
    shutil.copytree(root / "src" / "research_mcp", dest / "src" / "research_mcp", ignore=SKIP)
    return dest


def _npx() -> str:
    npx = shutil.which("npx") or shutil.which("npx.cmd")
    if not npx:
        raise SystemExit("npx not found: install Node.js (https://nodejs.org) to build the bundle")
    return npx


def mcpb(*args: str) -> None:
    cmd = [_npx(), "--yes", MCPB_CLI, *args]
    print("+ " + " ".join(cmd[1:]), flush=True)
    subprocess.run(cmd, check=True)


def build(out: Path, keep_stage: bool = False) -> Path:
    out = out.resolve()
    out.mkdir(parents=True, exist_ok=True)
    bundle = out / f"research-mcp-{package_version()}.mcpb"
    work = Path(tempfile.mkdtemp(prefix="research-mcp-mcpb-"))
    try:
        staged = stage(work / "research-mcp")
        mcpb("validate", str(staged / "manifest.json"))
        mcpb("pack", str(staged), str(bundle))
        mcpb("info", str(bundle))
    finally:
        if keep_stage:
            print(f"staged files kept in {work}")
        else:
            shutil.rmtree(work, ignore_errors=True)
    return bundle


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="mcpb/build.py", description=__doc__.splitlines()[0])
    parser.add_argument("--out", type=Path, default=ROOT / "dist", help="output folder (default: dist)")
    parser.add_argument("--keep-stage", action="store_true", help="keep the staged folder for inspection")
    args = parser.parse_args(argv)
    bundle = build(args.out, args.keep_stage)
    print(f"built {bundle}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
