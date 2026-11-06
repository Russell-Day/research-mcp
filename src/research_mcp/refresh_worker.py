"""One-shot supervised relay to the project-owned persistent Word worker."""

from __future__ import annotations

import json
from pathlib import Path
import sys

from .word_service import request_word_worker


def main(argv: list[str] | None = None) -> int:
    args = argv if argv is not None else sys.argv[1:]
    if len(args) not in (3, 4):
        raise SystemExit(
            "usage: python -m research_mcp.refresh_worker "
            "DOCUMENT.docx IDS_JSON WORKER_TIMEOUT [PIPE_TIMEOUT]"
        )
    worker_timeout = float(args[2])
    # The relay must give up before its supervisor does. Otherwise the relay
    # never exits, the supervisor's "helper exited non-zero" branch never runs,
    # and saved-package reconciliation stays unreachable on the one failure
    # path that actually occurs.
    pipe_timeout = float(args[3]) if len(args) == 4 else worker_timeout + 15.0
    response = request_word_worker(
        {
            "command": "refresh",
            "path": str(Path(args[0]).resolve(strict=True)),
            "citation_ids": list(json.loads(args[1])),
            "timeout": worker_timeout,
        },
        timeout=pipe_timeout,
    )
    print(json.dumps({key: value for key, value in response.items() if key != "ok"}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
