"""Comprehensive logging for MCP traffic and agent issue reporting.

Records all tool inputs, outputs, execution latencies, and failures to an
append-only JSONL log (`logs/mcp_traffic.jsonl`), and provides a mechanism for
agents to log friction, unexpected defects, or workflow issues to `logs/ISSUES.md`.
"""

from __future__ import annotations

from dataclasses import asdict, is_dataclass
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import threading
import traceback
from typing import Any
import uuid

_LOCK = threading.Lock()


ROLLOVER_BYTES = 20 * 1024 * 1024  # the per-user default log rolls over once at this size (one .1 file kept)


def _checkout_root() -> Path | None:
    """The repository root when running from a git source checkout (``<repo>/src/research_mcp``), else None."""
    root = Path(__file__).resolve().parent.parent.parent
    return root if (root / ".git").exists() and (root / "pyproject.toml").exists() else None


def _default_log_dir() -> Path:
    """``<repo>/logs`` in a source checkout; otherwise ``%LOCALAPPDATA%\\research_mcp\\logs`` (next to the
    undo store), never the install folder, which uv may delete or a user may not be able to write."""
    root = _checkout_root()
    if root is not None:
        return root / "logs"
    from .checkpoints import home

    return home() / "logs"


def get_log_dir() -> Path:
    """Resolve the directory used for MCP traffic logs and agent issue logs (``RESEARCH_MCP_LOG_DIR``
    from the environment or config.env, else the default above)."""
    from . import config

    configured = config.get("RESEARCH_MCP_LOG_DIR")
    p = Path(configured).resolve() if configured else _default_log_dir()
    p.mkdir(parents=True, exist_ok=True)
    return p


def _rolls_over() -> bool:
    """Only the per-user default rolls over; a checkout's or a configured log is left as the user keeps it."""
    from . import config

    return not config.get("RESEARCH_MCP_LOG_DIR") and _checkout_root() is None


def _roll_over(path: Path) -> None:
    try:
        if path.stat().st_size >= ROLLOVER_BYTES:
            os.replace(path, path.with_name(path.stem + ".1" + path.suffix))
    except OSError:
        pass


def get_traffic_log_path() -> Path:
    return get_log_dir() / "mcp_traffic.jsonl"


def get_issues_md_path() -> Path:
    return get_log_dir() / "ISSUES.md"


def _to_serializable(val: Any) -> Any:
    """Convert arbitrary objects (Pydantic models, CallToolResult, Paths) to JSON-safe structures."""
    if val is None or isinstance(val, (int, float, str, bool)):
        return val
    if isinstance(val, Path):
        return str(val)
    if isinstance(val, bytes):
        return f"<bytes len={len(val)}>"
    if isinstance(val, (list, tuple, set)):
        return [_to_serializable(x) for x in val]
    if isinstance(val, dict):
        return {str(k): _to_serializable(v) for k, v in val.items()}
    if hasattr(val, "model_dump") and callable(val.model_dump):
        try:
            return val.model_dump(mode="json")
        except Exception:
            pass
    if is_dataclass(val):
        try:
            return _to_serializable(asdict(val))
        except Exception:
            pass
    # Handle MCP CallToolResult
    if hasattr(val, "structured_content") and val.structured_content is not None:
        return _to_serializable(val.structured_content)
    if hasattr(val, "content"):
        items = []
        for c in val.content:
            if hasattr(c, "text"):
                try:
                    items.append(json.loads(c.text))
                except Exception:
                    items.append(c.text)
            else:
                items.append(repr(c))
        return items if len(items) != 1 else items[0]
    if isinstance(val, Exception):
        return {
            "type": val.__class__.__name__,
            "message": str(val),
        }
    return repr(val)


def log_tool_call(
    tool_name: str,
    arguments: dict[str, Any],
    result: Any = None,
    duration_ms: float = 0.0,
    error: Exception | str | None = None,
) -> None:
    """Append a tool invocation record to logs/mcp_traffic.jsonl.

    Never raises: failures in logging are caught silently to prevent disrupting
    active tool calls.
    """
    try:
        entry = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "tool": tool_name,
            "duration_ms": round(duration_ms, 2),
            "status": "error" if error is not None else "success",
            "arguments": _to_serializable(arguments),
        }
        if error is not None:
            if isinstance(error, Exception):
                entry["error"] = {
                    "type": error.__class__.__name__,
                    "message": str(error),
                    "traceback": traceback.format_exc(),
                }
            else:
                entry["error"] = {"message": str(error)}
        else:
            entry["result"] = _to_serializable(result)

        line = json.dumps(entry, ensure_ascii=False)
        target = get_traffic_log_path()
        with _LOCK:
            if _rolls_over():
                _roll_over(target)
            with open(target, "a", encoding="utf-8") as f:
                f.write(line + "\n")
    except Exception:
        # Logging must never crash the primary execution path
        pass


def record_agent_issue(
    title: str,
    details: str,
    category: str = "friction",
    document_path: str | None = None,
    context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Append a structured issue or friction report to logs/ISSUES.md."""

    now = datetime.now(timezone.utc)
    date_str = now.strftime("%Y%m%d")
    short_uuid = uuid.uuid4().hex[:6].upper()
    issue_id = f"ISSUE-{date_str}-{short_uuid}"
    md_path: Path | str = "(log folder unavailable)"

    try:
        md_path = get_issues_md_path()
        with _LOCK:
            needs_header = not md_path.exists() or md_path.stat().st_size == 0
            with open(md_path, "a", encoding="utf-8") as f:
                if needs_header:
                    f.write("# Research MCP Agent Issues & Feedback Log\n\n")
                    f.write(
                        "Tracked friction points, formatting defects, unexpected tool outcomes, "
                        "and workflow improvements reported during agent execution.\n\n---\n\n"
                    )

                f.write(f"## [{issue_id}] {title}\n")
                f.write(f"- **Reported At:** `{now.isoformat()}`\n")
                f.write(f"- **Category:** `{category}`\n")
                if document_path:
                    f.write(f"- **Document Path:** `{document_path}`\n")
                f.write("- **Status:** `open`\n\n")

                f.write("### Details\n")
                f.write(f"{details.strip()}\n\n")

                if context:
                    f.write("### Additional Context\n```json\n")
                    f.write(json.dumps(_to_serializable(context), indent=2, ensure_ascii=False))
                    f.write("\n```\n\n")

                f.write("---\n\n")

        # Also log to mcp_traffic.jsonl
        log_tool_call(
            tool_name="report_issue",
            arguments={
                "title": title,
                "category": category,
                "document_path": document_path,
                "details": details,
                "context": context,
            },
            result={"issue_id": issue_id, "logged_to": str(md_path)},
            duration_ms=0.0,
        )

        return {
            "ok": True,
            "issue_id": issue_id,
            "logged_to": str(md_path),
            "message": f"Issue '{issue_id}' successfully recorded in {Path(md_path).name}.",
        }
    except Exception as exc:
        return {
            "ok": False,
            "issue_id": issue_id,
            "logged_to": str(md_path),
            "message": f"Failed to write issue to log: {exc}",
        }
