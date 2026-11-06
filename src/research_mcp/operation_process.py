"""Hard process boundary for Word work; callers retain publication ownership."""
from __future__ import annotations
import json
import math
import os
from pathlib import Path
import subprocess
import sys
import tempfile
from typing import Any

from .errors import IntegrationRefused, RefusedUnsafe, VerificationFailed
from .processes import process_identity_matches


def _cleanup_owned_word(registry: Path) -> None:
    # Never enumerate or kill arbitrary WINWORD processes.
    try:
        identity = json.loads(registry.read_text(encoding="utf-8"))
        pid, created = int(identity["pid"]), identity["created"]
        if process_identity_matches(pid, created):
            from .word_com import kill_process
            kill_process(pid)
    except (OSError, ValueError, KeyError, TypeError):
        pass


def run_operation(kind: str, payload: dict[str, Any], *, timeout_seconds: float = 120) -> dict:
    if kind not in {"edit", "style", "render", "tables", "bookmarks", "inline_style"}:
        raise ValueError(f"Unknown isolated operation: {kind}")
    if not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
        raise ValueError("timeout_seconds must be positive")
    with tempfile.TemporaryDirectory(prefix="research-mcp-worker-") as temporary:
        root = Path(temporary)
        request, response, registry = root / "request.json", root / "response.json", root / "word.json"
        request.write_text(json.dumps({"kind": kind, "payload": payload}), encoding="utf-8")
        environment = os.environ.copy()
        environment["RESEARCH_MCP_OWNED_WORD_FILE"] = str(registry)
        # Windows venv python.exe is a launcher that spawns the actual runtime.
        # Killing only that launcher leaves its worker alive. Launch the real
        # interpreter directly and carry over this environment's import paths.
        executable = getattr(sys, "_base_executable", sys.executable) if os.name == "nt" else sys.executable
        environment["PYTHONPATH"] = os.pathsep.join(str(Path(entry).resolve()) for entry in sys.path if entry)

        try:
            completed = subprocess.run(
                [executable, "-m", "research_mcp.operation_process", str(request), str(response)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, text=True, env=environment,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
                timeout=timeout_seconds, check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise IntegrationRefused(
                f"Word {kind} worker exceeded {timeout_seconds:g} seconds and was stopped. "
                "No staged result was accepted; inspect before retrying."
            ) from exc
        finally:
            _cleanup_owned_word(registry)
        if completed.returncode != 0 or not response.is_file():
            raise IntegrationRefused(f"Word {kind} worker exited without a verified result (exit {completed.returncode}).")
        try:
            result = json.loads(response.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise IntegrationRefused("Word worker returned an unreadable result") from exc
        if not result.get("ok"):
            error = result.get("error", {})
            error_type = {"RefusedUnsafe": RefusedUnsafe, "VerificationFailed": VerificationFailed}.get(error.get("type"), IntegrationRefused)
            raise error_type(error.get("message", "Word worker failed"))
        return result["result"]


def _dispatch(kind: str, payload: dict) -> dict:
    if kind == "edit":
        from pydantic import TypeAdapter
        from .edit_models import EditOperation
        from .edit_batch import _apply_operation_local
        return _apply_operation_local(Path(payload["source"]), Path(payload["output"]), TypeAdapter(EditOperation).validate_python(payload["operation"]))
    if kind == "inline_style":
        from .inline_style import style_exact_text
        return style_exact_text(**payload)
    if kind == "style":
        from .styling import _style_document_local
        return _style_document_local(payload["source"], payload["profile_name"], payload.get("visible", False), payload.get("emphasis_policy", "preserve"))
    if kind == "tables":
        from .tables import TableFormatSpec, MergeGroup, format_to_version
        specs = tuple(TableFormatSpec(index=raw["index"], descriptive_columns=tuple(raw["descriptive_columns"]), merges=tuple(MergeGroup(**merge) for merge in raw.get("merges", []))) for raw in payload["specs"])
        return format_to_version(payload["source"], payload["source"], specs).to_dict()
    if kind == "bookmarks":
        from .bookmarks import release_to_version
        return release_to_version(payload["source"], payload["source"], names=tuple(payload["names"])).to_dict()
    if kind == "render":
        from .rendering import render_docx_pages
        return render_docx_pages(**payload)
    raise RefusedUnsafe("Unknown worker operation")


def main() -> None:
    request, response = map(Path, sys.argv[1:3])
    try:
        data = json.loads(request.read_text(encoding="utf-8"))
        result = {"ok": True, "result": _dispatch(data["kind"], data["payload"])}
    except Exception as exc:
        result = {"ok": False, "error": {"type": type(exc).__name__, "message": str(exc)}}
    response.write_text(json.dumps(result), encoding="utf-8")


if __name__ == "__main__":
    main()
