"""Cross-process Zotero integration lease and supervised refresh dispatch."""

from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import secrets
import subprocess
import threading
import sys
import time

from .errors import IntegrationBusy, ZoteroModal
from .audit import file_sha256
from .ooxml import ZoteroPackageFields, zotero_package_fields
from .processes import process_creation_time, process_identity_matches
from .windows import is_suspicious_zotero_window, visible_dialogs_for_pids
from .word_service import refresh_state_path


def default_lock_path() -> Path:
    root = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "research_mcp"
    return root / "zotero-integration.lock"


@dataclass
class IntegrationLease:
    path: Path
    lease_id: str
    acquired: bool = False

    @classmethod
    def acquire(cls, path: str | Path | None = None) -> "IntegrationLease":
        lock_path = Path(path) if path else default_lock_path()
        lock_path.parent.mkdir(parents=True, exist_ok=True)
        lease_id = secrets.token_hex(16)
        payload = json.dumps({
            "pid": os.getpid(),
            "process_created": process_creation_time(os.getpid()),
            "lease_id": lease_id,
            "created": time.time(),
        })
        try:
            descriptor = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError as exc:
            detail = lock_path.read_text(encoding="utf-8", errors="replace")
            raise IntegrationBusy(
                f"Zotero integration is locked or requires recovery: {lock_path} {detail}"
            ) from exc
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(payload)
        return cls(path=lock_path, lease_id=lease_id, acquired=True)

    def _assert_ownership(self) -> None:
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (FileNotFoundError, ValueError, OSError) as exc:
            raise IntegrationBusy("integration lease disappeared or became invalid") from exc
        if payload.get("lease_id") != self.lease_id:
            raise IntegrationBusy("integration lease ownership changed; refusing deletion")

    def release_success(self) -> None:
        if self.acquired:
            self._assert_ownership()
            self.path.unlink(missing_ok=False)
            self.acquired = False

    def preserve_failure(self, reason: str) -> None:
        if self.acquired:
            payload = {
                "pid": os.getpid(),
                "process_created": process_creation_time(os.getpid()),
                "lease_id": self.lease_id,
                "failed": time.time(),
                "reason": reason,
            }
            self.path.write_text(json.dumps(payload), encoding="utf-8")


def zotero_pids() -> set[int]:
    import subprocess as sp

    result = sp.run(
        ["powershell", "-NoProfile", "-Command", "Get-Process zotero -ErrorAction SilentlyContinue | ForEach-Object Id"],
        capture_output=True,
        text=True,
        timeout=10,
        check=False,
    )
    return {int(line.strip()) for line in result.stdout.splitlines() if line.strip().isdigit()}


def process_exists(pid: int) -> bool:
    return process_creation_time(pid) is not None


def integration_status(path: str | Path | None = None) -> dict:
    lock_path = Path(path) if path else default_lock_path()
    if not lock_path.exists():
        return {"locked": False, "path": str(lock_path)}
    try:
        payload = json.loads(lock_path.read_text(encoding="utf-8"))
    except Exception:
        payload = {"invalid": True}
    pid = int(payload.get("pid", 0)) if isinstance(payload, dict) else 0
    expected_created = payload.get("process_created") if isinstance(payload, dict) else None
    identity_matches = process_identity_matches(pid, expected_created)
    return {
        "locked": True,
        "path": str(lock_path),
        "owner_pid": pid or None,
        "owner_running": process_exists(pid),
        "owner_identity_matches": identity_matches,
        "payload": payload,
    }


def clear_integration_lock(path: str | Path | None = None, *, confirmed: bool = False) -> dict:
    if not confirmed:
        raise IntegrationBusy("explicit recovery confirmation is required")
    status = integration_status(path)
    if not status["locked"]:
        return status
    if status["payload"].get("invalid") or not status["owner_pid"]:
        raise IntegrationBusy("refusing to clear an invalid or partially written integration lock")
    if status["owner_identity_matches"]:
        raise IntegrationBusy(f"refusing to clear lock owned by the live process {status['owner_pid']}")
    if status["owner_running"] and status["payload"].get("process_created") is None:
        raise IntegrationBusy(
            f"legacy lock PID {status['owner_pid']} is running but has no creation-time identity; inspect manually"
        )
    suspicious = [
        window.title
        for window in visible_dialogs_for_pids(zotero_pids())
        if is_suspicious_zotero_window(window)
    ]
    if suspicious:
        raise ZoteroModal(f"refusing to clear lock while Zotero modal is visible: {suspicious}")
    # Narrow the stale-check/delete race: recovery must still be looking at the
    # exact payload it inspected. Lease owners also verify their token before
    # deleting, so neither side silently removes a changed ownership record.
    current = integration_status(path)
    if not current["locked"] or current["payload"] != status["payload"]:
        raise IntegrationBusy("integration lock changed during recovery; inspect again")
    Path(status["path"]).unlink(missing_ok=False)
    return {**status, "locked": False, "cleared": True}


def read_refresh_progress() -> dict | None:
    """Last progress record the Word worker published, if any."""

    try:
        data = json.loads(refresh_state_path().read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError, OSError):
        return None
    return data if isinstance(data, dict) else None


def _hidden_subprocess_kwargs() -> dict:
    """Return Windows startup options that prevent helper console windows."""

    if sys.platform != "win32":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.lpDesktop = r"WinSta0\Default"
    return {
        "startupinfo": startupinfo,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
    }


def _reconcile_saved_refresh(
    document_path: Path,
    citation_ids: tuple[str, ...],
    before_sha256: str,
    before_fields: ZoteroPackageFields,
) -> dict:
    """Accept a timed-out helper only when the closed package proves progress.

    This path is intentionally unavailable for an already-rendered idempotent
    refresh: at least one requested result must have been pending before
    dispatch. That prevents an unchanged healthy-looking file from masking a
    refresh that never ran.
    """

    before_pairs = list(before_fields.item_results)
    before_rendered = dict(before_pairs)
    if len(before_rendered) != len(before_pairs):
        raise RuntimeError("pre-dispatch package contains duplicate Zotero citation IDs")
    requested = set(citation_ids)
    if not requested.issubset(before_rendered):
        raise RuntimeError("requested citation IDs were absent before dispatch")

    after_sha256 = file_sha256(document_path)
    if after_sha256 == before_sha256:
        raise RuntimeError("saved package did not change during refresh")
    after_fields = zotero_package_fields(document_path)
    pairs = list(after_fields.item_results)
    rendered = dict(pairs)
    if len(rendered) != len(pairs):
        raise RuntimeError("saved package contains duplicate Zotero citation IDs")
    if not requested.issubset(rendered):
        raise RuntimeError(
            "saved package is missing requested citation IDs: "
            f"expected_subset={sorted(citation_ids)} actual={sorted(rendered)}"
        )
    unrendered = [
        citation_id for citation_id in citation_ids
        for result in (rendered[citation_id],)
        if not result or result == "{Citation}"
    ]
    if unrendered:
        raise RuntimeError(f"saved package still has unrendered citations: {unrendered[:5]}")
    if len(after_fields.bibliography_results) != len(before_fields.bibliography_results):
        raise RuntimeError("saved package bibliography field count changed during refresh")
    if any(result == "{Bibliography}" for result in after_fields.bibliography_results):
        raise RuntimeError("saved package still has an unrendered bibliography")

    targeted_transition = any(
        (not before_rendered[citation_id] or before_rendered[citation_id] == "{Citation}")
        and rendered[citation_id]
        and rendered[citation_id] != "{Citation}"
        for citation_id in citation_ids
    )
    bibliography_transition = (
        not citation_ids
        and any(result == "{Bibliography}" for result in before_fields.bibliography_results)
        and all(result != "{Bibliography}" for result in after_fields.bibliography_results)
    )
    if not targeted_transition and not bibliography_transition:
        raise RuntimeError("saved-package reconciliation requires a targeted pending-to-rendered transition")
    return {
        "rendered": {citation_id: rendered[citation_id] for citation_id in citation_ids},
        "bibliography_results": list(after_fields.bibliography_results),
        "reconciled_from_saved_package": True,
        "output_sha256": after_sha256,
    }


def _validate_helper_result(data: object, citation_ids: tuple[str, ...]) -> dict:
    if not isinstance(data, dict):
        raise RuntimeError("refresh helper JSON must be an object")
    rendered = data.get("rendered")
    bibliographies = data.get("bibliography_results")
    if not isinstance(rendered, dict) or set(rendered) != set(citation_ids):
        raise RuntimeError("refresh helper returned the wrong citation IDs")
    if any(not isinstance(value, str) or not value.strip() or value.strip() == "{Citation}" for value in rendered.values()):
        raise RuntimeError("refresh helper returned unrendered citation results")
    if not isinstance(bibliographies, list) or any(
        not isinstance(value, str) or value.strip() == "{Bibliography}"
        for value in bibliographies
    ):
        raise RuntimeError("refresh helper returned invalid bibliography results")
    return data


def dispatch_refresh(
    document_path: str | Path,
    *,
    citation_ids: tuple[str, ...] = (),
    dispatch_timeout: float = 75.0,
    poll_interval: float = 0.2,
) -> dict:
    """Have one supervised helper own, refresh, verify, and save the document."""

    from .recovery import zotero_is_suspect

    suspect = zotero_is_suspect()
    if suspect:
        raise IntegrationBusy(
            f"Zotero integration is marked suspect ({suspect}); run `rmcp zotero-recover` before refresh"
        )

    # Reserve a real cleanup/IPC margin. Five seconds was insufficient for the
    # 36-citation manuscript: the worker reached a correct saved state near its
    # own deadline, then the outer caller timed out while the helper unwound.
    cleanup_margin = min(30.0, max(5.0, dispatch_timeout * 0.1))
    worker_timeout = max(1.0, dispatch_timeout - cleanup_margin)
    # Strictly ordered deadlines: worker loop < relay pipe < this supervisor.
    # Any other ordering leaves the relay blocked when the supervisor gives up,
    # which is what made saved-package reconciliation unreachable.
    relay_timeout = max(worker_timeout + 1.0, dispatch_timeout - cleanup_margin * 0.5)
    command = [
        sys.executable,
        "-m",
        "research_mcp.refresh_worker",
        str(Path(document_path).resolve()),
        json.dumps(citation_ids),
        str(worker_timeout),
        str(relay_timeout),
    ]
    preexisting = [
        window.title
        for window in visible_dialogs_for_pids(zotero_pids())
        if is_suspicious_zotero_window(window)
    ]
    if preexisting:
        raise ZoteroModal(
            f"refusing refresh while Zotero modal window(s) are already visible: {preexisting}"
        )
    target = Path(document_path).resolve(strict=True)
    before_sha256 = file_sha256(target)
    before_fields = zotero_package_fields(target)
    helper = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        **_hidden_subprocess_kwargs(),
    )
    # MEASURED: this OS pipe buffers between 4096 and 7382 bytes, and a
    # 36-citation refresh response is 7382 bytes of JSON. Polling poll() while
    # nobody drains stdout deadlocks the helper mid-write, so the refresh
    # appears to hang forever no matter how large the timeout is -- 144s, 300s
    # and 900s all failed identically on a document that was already complete
    # and saved. One citation fits in the buffer; thirty-six do not. Drain both
    # streams concurrently, which also makes stderr readable on timeout.
    stdout_chunks: list[str] = []
    stderr_chunks: list[str] = []

    def _drain(stream, sink: list[str]) -> None:
        try:
            for line in stream:
                sink.append(line)
        except Exception:
            pass
        finally:
            try:
                stream.close()
            except Exception:
                pass

    readers = [
        threading.Thread(target=_drain, args=(helper.stdout, stdout_chunks), daemon=True),
        threading.Thread(target=_drain, args=(helper.stderr, stderr_chunks), daemon=True),
    ]
    for reader in readers:
        reader.start()
    deadline = time.monotonic() + dispatch_timeout
    while helper.poll() is None and time.monotonic() < deadline:
        dialogs = [
            window
            for window in visible_dialogs_for_pids(zotero_pids())
            if is_suspicious_zotero_window(window)
        ]
        if dialogs:
            titles = [window.title for window in dialogs]
            raise ZoteroModal(
                f"Zotero opened modal window(s) {titles}; helper PID {helper.pid} retained; "
                f"integration lock must be recovered explicitly"
            )
        time.sleep(poll_interval)
    if helper.poll() is None:
        # The worker's internal deadline is deliberately shorter than this
        # outer deadline, but COM cleanup and pipe delivery can still straddle
        # it. Give only that tail a bounded grace period; never terminate Word
        # or the helper automatically.
        grace_deadline = time.monotonic() + min(30.0, max(2.0, dispatch_timeout * 0.1))
        while helper.poll() is None and time.monotonic() < grace_deadline:
            dialogs = [
                window
                for window in visible_dialogs_for_pids(zotero_pids())
                if is_suspicious_zotero_window(window)
            ]
            if dialogs:
                titles = [window.title for window in dialogs]
                raise ZoteroModal(
                    f"Zotero opened modal window(s) {titles}; helper PID {helper.pid} retained; "
                    "integration lock must be recovered explicitly"
                )
            time.sleep(poll_interval)
    if helper.poll() is None:
        # The helper is still running, so its stdout/stderr cannot be read
        # without blocking. Two things are still readable: the worker's
        # progress sidecar, and the saved package itself. Zotero saves the
        # document as part of its own refresh, so a supervisor timeout does
        # NOT imply the work failed -- three recorded runs discarded a
        # complete, correctly rendered manuscript here.
        progress = read_refresh_progress()
        partial_stderr = "".join(stderr_chunks).strip()[-2000:]
        try:
            reconciled = _reconcile_saved_refresh(
                target, citation_ids, before_sha256, before_fields
            )
        except Exception as reconciliation_error:
            raise ZoteroModal(
                f"Zotero refresh dispatch exceeded {dispatch_timeout}s; helper PID {helper.pid} retained; "
                f"integration lock must be recovered explicitly; "
                f"worker progress: {progress}; helper stderr: {partial_stderr or '(none)'}; "
                f"saved-package reconciliation failed: {reconciliation_error}"
            ) from reconciliation_error
        # The document is provably correct, but the environment is not clean:
        # the helper still runs and still owns the lease. Say so in the result
        # so no caller mistakes this for a tidy success.
        return {
            **reconciled,
            "helper_pid_retained": helper.pid,
            "helper_still_running": True,
            "integration_lock_requires_recovery": True,
            "dispatch_timed_out_seconds": dispatch_timeout,
            "worker_progress": progress,
            "helper_stderr": partial_stderr,
        }
    for reader in readers:
        reader.join(timeout=10.0)
    stdout = "".join(stdout_chunks)
    stderr = "".join(stderr_chunks)
    if helper.returncode != 0:
        try:
            reconciled = _reconcile_saved_refresh(
                target, citation_ids, before_sha256, before_fields
            )
        except Exception as reconciliation_error:
            raise RuntimeError(
                f"refresh helper failed ({helper.returncode}): {stderr.strip()}; "
                f"saved-package reconciliation failed: {reconciliation_error}"
            ) from reconciliation_error
        return {
            **reconciled,
            "helper_returncode": helper.returncode,
            "helper_error": stderr.strip(),
        }
    try:
        return _validate_helper_result(json.loads(stdout), citation_ids)
    except (json.JSONDecodeError, TypeError, RuntimeError) as output_error:
        try:
            reconciled = _reconcile_saved_refresh(
                target, citation_ids, before_sha256, before_fields
            )
        except Exception as reconciliation_error:
            raise RuntimeError(
                f"refresh helper output was invalid ({output_error}); "
                f"saved-package reconciliation failed: {reconciliation_error}"
            ) from output_error
        return {**reconciled, "helper_output_error": str(output_error)}
