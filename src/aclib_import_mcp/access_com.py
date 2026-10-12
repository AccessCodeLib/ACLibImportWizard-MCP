"""COM access to Microsoft Access.

All COM calls run on one dedicated STA thread, one call after another. Each call
is additionally guarded by a named mutex per database, so several server
processes never work on the same database at the same time.

Instance handling per call:

1. A running Access instance that has exactly this database open (looked up in
   the Running Object Table) is reused and never closed.
2. Otherwise, if the lock file of the database is in use by another process,
   the call fails - a second Access instance is never opened on the same file.
3. Otherwise an own hidden instance is started, the database is opened, and
   both are closed again after the call.
"""

from __future__ import annotations

import hashlib
import logging
import os
import queue
import threading
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any, TypeVar

from . import dialogs
from .dialogs import DialogInfo

log = logging.getLogger(__name__)

T = TypeVar("T")

DEFAULT_CALL_TIMEOUT = 120.0
CALL_TIMEOUT_ENV = "ACLIB_CALL_TIMEOUT"
DEFAULT_DIALOG_TIMEOUT = 20.0
DIALOG_TIMEOUT_ENV = "ACLIB_DIALOG_TIMEOUT"
DIALOG_POLL_SECONDS = 1.0

MUTEX_PREFIX = "Local\\aclib-import-mcp-db-"
MUTEX_POLL_SECONDS = 0.2
# Extra time the caller waits beyond the job deadline, so the worker can report a
# busy database itself instead of the caller raising a generic timeout.
RESULT_GRACE_SECONDS = 1.0

AC_QUIT_SAVE_ALL = 1

LOCK_RELEASE_TIMEOUT_SECONDS = 10.0
LOCK_RELEASE_POLL_SECONDS = 0.1

_LOCK_FILE_SUFFIXES = {
    ".accdb": ".laccdb",
    ".accda": ".laccdb",
    ".accde": ".laccdb",
    ".accdr": ".laccdb",
    ".mdb": ".ldb",
    ".mda": ".ldb",
    ".mde": ".ldb",
}


class AccessComError(Exception):
    """Base class of all errors raised by this module."""


class DatabaseNotFoundError(AccessComError):
    """The database file does not exist."""


class DatabaseInUseError(AccessComError):
    """The database is opened by a process that is not reachable via the ROT."""


class DatabaseBusyError(AccessComError):
    """Another server process is working on the database."""


class CallTimeoutError(AccessComError):
    """The call did not finish within the timeout."""


class DialogBlockingError(AccessComError):
    """A modal dialog of the Access instance blocks the call."""


class WorkerBlockedError(AccessComError):
    """A previous call is still blocked (e.g. by a dialog); no further calls are possible."""


def _get_seconds(env: str, default: float) -> float:
    raw = os.environ.get(env, "").strip()
    if not raw:
        return default
    try:
        value = float(raw)
    except ValueError:
        raise ValueError(f"{env} must be a number of seconds, got {raw!r}") from None
    if value <= 0:
        raise ValueError(f"{env} must be greater than 0, got {raw!r}")
    return value


def get_call_timeout() -> float:
    """Return the call timeout in seconds from ``ACLIB_CALL_TIMEOUT`` or the default."""
    return _get_seconds(CALL_TIMEOUT_ENV, DEFAULT_CALL_TIMEOUT)


def get_dialog_timeout() -> float:
    """Return how long a dialog may stay open before the call fails (``ACLIB_DIALOG_TIMEOUT``)."""
    return _get_seconds(DIALOG_TIMEOUT_ENV, DEFAULT_DIALOG_TIMEOUT)


def normalize_path(path: str | os.PathLike[str]) -> str:
    """Return the absolute, case-normalized form of ``path`` for comparisons."""
    return os.path.normcase(os.path.abspath(os.fspath(path)))


def lock_file_path(database_path: str | os.PathLike[str]) -> Path:
    """Return the path of the Access lock file (``.laccdb``/``.ldb``) of a database."""
    path = Path(database_path)
    suffix = _LOCK_FILE_SUFFIXES.get(path.suffix.lower(), ".laccdb")
    return path.with_suffix(suffix)


def mutex_name(database_path: str | os.PathLike[str]) -> str:
    """Return the name of the cross-process mutex guarding a database."""
    digest = hashlib.sha256(normalize_path(database_path).encode("utf-8")).hexdigest()
    return MUTEX_PREFIX + digest


class ComBackend:
    """Thin wrapper around pywin32; replaced by a fake in unit tests."""

    def initialize(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()

    def uninitialize(self) -> None:
        import pythoncom

        pythoncom.CoUninitialize()

    def find_running_app(self, database_path: str) -> Any | None:
        """Return the Access application that has ``database_path`` open, or None.

        Only looks up the Running Object Table; never starts Access.
        """
        import pythoncom
        import win32com.client.dynamic

        target = normalize_path(database_path)
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        for moniker in rot.EnumRunning():
            try:
                name = moniker.GetDisplayName(ctx, None)
            except pythoncom.com_error:
                continue
            if normalize_path(name) != target:
                continue
            try:
                unknown = rot.GetObject(moniker)
                disp = win32com.client.dynamic.Dispatch(
                    unknown.QueryInterface(pythoncom.IID_IDispatch)
                )
            except pythoncom.com_error:
                log.debug("ROT entry %s could not be bound", name, exc_info=True)
                continue
            app = _as_access_application(disp)
            if app is not None and _current_database(app) == target:
                return app
        return None

    def create_app(self) -> Any:
        import win32com.client

        return win32com.client.DispatchEx("Access.Application")

    def process_id(self, app: Any) -> int | None:
        """Return the process id of an Access application, or None if unknown."""
        import win32process

        try:
            return win32process.GetWindowThreadProcessId(access_window_handle(app))[1]
        except Exception:
            log.warning("Could not determine the process id of Access", exc_info=True)
            return None

    def is_lock_file_in_use(self, lock_path: str) -> bool:
        """Return True if another process holds the lock file open."""
        import pywintypes
        import win32con
        import win32file
        import winerror

        try:
            handle = win32file.CreateFile(
                lock_path,
                win32con.GENERIC_READ,
                0,  # no sharing: fails while Access holds the file open
                None,
                win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL,
                None,
            )
        except pywintypes.error as exc:
            # Sharing violation means in use; anything else is treated as in use too.
            return exc.winerror not in (
                winerror.ERROR_FILE_NOT_FOUND,
                winerror.ERROR_PATH_NOT_FOUND,
            )
        handle.Close()
        return False

    def create_mutex(self, name: str) -> Any:
        import win32event

        return win32event.CreateMutex(None, False, name)

    def wait_mutex(self, handle: Any, timeout_seconds: float) -> bool:
        import win32event

        result = win32event.WaitForSingleObject(handle, int(timeout_seconds * 1000))
        return result in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED)

    def release_mutex(self, handle: Any) -> None:
        import win32event

        win32event.ReleaseMutex(handle)

    def close_handle(self, handle: Any) -> None:
        handle.Close()


def access_window_handle(app: Any) -> int:
    """Return the main window handle of Access.

    ``hWndAccessApp`` is a method; with dynamic dispatch the attribute is a bound method.
    """
    handle = app.hWndAccessApp
    if callable(handle):
        handle = handle()
    return int(handle)


def _as_access_application(disp: Any) -> Any | None:
    """Return the Access application object for a ROT object, or None."""
    for candidate in (disp, getattr(disp, "Application", None)):
        if candidate is None:
            continue
        try:
            _ = candidate.CurrentProject.FullName
        except Exception:
            continue
        return candidate
    return None


def _current_database(app: Any) -> str | None:
    try:
        return normalize_path(app.CurrentProject.FullName)
    except Exception:
        return None


class _Job:
    def __init__(self, database_path: str, action: Callable[[Any], Any], timeout: float) -> None:
        self.database_path = database_path
        self.action = action
        self.deadline = time.monotonic() + timeout
        self.done = threading.Event()
        self.abandoned = threading.Event()
        self.process_id: int | None = None
        self.result: Any = None
        self.error: BaseException | None = None


DialogFinder = Callable[[int], list[DialogInfo]]


class AccessWorker:
    """Runs actions against an Access database on one dedicated STA thread."""

    def __init__(
        self,
        backend: ComBackend | None = None,
        timeout: float | None = None,
        dialog_timeout: float | None = None,
        dialog_finder: DialogFinder | None = None,
    ) -> None:
        self._backend = backend or ComBackend()
        self._timeout = timeout
        self._dialog_timeout = dialog_timeout
        self._find_dialogs = dialog_finder or dialogs.find_dialogs
        self._queue: queue.Queue[_Job | None] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._start_lock = threading.Lock()
        self._current: _Job | None = None

    @property
    def timeout(self) -> float:
        return self._timeout if self._timeout is not None else get_call_timeout()

    @property
    def dialog_timeout(self) -> float:
        if self._dialog_timeout is not None:
            return self._dialog_timeout
        return get_dialog_timeout()

    def execute(self, database_path: str | os.PathLike[str], action: Callable[[Any], T]) -> T:
        """Run ``action(app)`` on the worker thread with ``database_path`` opened in ``app``.

        Blocks until the action finished, the timeout elapsed, or a dialog of the
        Access instance stayed open longer than the dialog timeout.
        """
        path = os.path.abspath(os.fspath(database_path))
        if not os.path.isfile(path):
            raise DatabaseNotFoundError(f"Database not found: {path}")
        self._check_not_blocked()

        timeout = self.timeout
        dialog_timeout = self.dialog_timeout
        job = _Job(path, action, timeout)
        self._ensure_thread()
        self._queue.put(job)

        wait_until = time.monotonic() + timeout + RESULT_GRACE_SECONDS
        dialog_since: float | None = None
        while not job.done.wait(DIALOG_POLL_SECONDS):
            now = time.monotonic()
            found = self._dialogs_of(job)
            if found:
                dialog_since = dialog_since if dialog_since is not None else now
                if now - dialog_since >= dialog_timeout:
                    job.abandoned.set()
                    raise DialogBlockingError(
                        f"Call on {path} is blocked by a dialog of Access (open for more than "
                        f"{dialog_timeout:g} seconds, configurable via {DIALOG_TIMEOUT_ENV}). "
                        "The dialog is not answered automatically; close it in Access. "
                        "Until then no further calls are possible.\n"
                        + dialogs.describe_dialogs(found)
                    )
            else:
                dialog_since = None
            if now >= wait_until:
                job.abandoned.set()
                message = (
                    f"Call on {path} did not finish within {timeout:g} seconds "
                    f"(configurable via {CALL_TIMEOUT_ENV})."
                )
                if found:
                    message += "\n" + dialogs.describe_dialogs(found)
                raise CallTimeoutError(message)
        if job.error is not None:
            raise job.error
        return job.result

    def _dialogs_of(self, job: _Job) -> list[DialogInfo]:
        if job.process_id is None:
            return []
        try:
            return self._find_dialogs(job.process_id)
        except Exception:
            log.warning("Dialog detection failed", exc_info=True)
            return []

    def _check_not_blocked(self) -> None:
        current = self._current
        if current is None or not current.abandoned.is_set() or current.done.is_set():
            return
        message = (
            f"A previous call on {current.database_path} is still running or blocked in Access; "
            "close an open dialog in Access, or wait, and try again."
        )
        found = self._dialogs_of(current)
        if found:
            message += "\n" + dialogs.describe_dialogs(found)
        raise WorkerBlockedError(message)

    def close(self) -> None:
        """Stop the worker thread after the queued jobs are processed."""
        thread = self._thread
        if thread is None:
            return
        self._queue.put(None)
        thread.join()
        self._thread = None

    def _ensure_thread(self) -> None:
        with self._start_lock:
            if self._thread is None:
                self._thread = threading.Thread(
                    target=self._run, name="aclib-com-worker", daemon=True
                )
                self._thread.start()

    def _run(self) -> None:
        self._backend.initialize()
        try:
            while True:
                job = self._queue.get()
                if job is None:
                    break
                if job.abandoned.is_set():
                    continue
                self._current = job
                try:
                    job.result = self._run_job(job)
                except BaseException as exc:  # handed over to the calling thread
                    job.error = exc
                finally:
                    job.done.set()
                    self._current = None
        finally:
            self._backend.uninitialize()

    def _run_job(self, job: _Job) -> Any:
        backend = self._backend
        handle = backend.create_mutex(mutex_name(job.database_path))
        try:
            if not self._acquire_mutex(handle, job):
                raise DatabaseBusyError(
                    f"Database {job.database_path} is in use by another aclib-import-mcp process."
                )
            try:
                return self._run_action(job)
            finally:
                backend.release_mutex(handle)
        finally:
            backend.close_handle(handle)

    def _acquire_mutex(self, handle: Any, job: _Job) -> bool:
        while not job.abandoned.is_set():
            remaining = job.deadline - time.monotonic()
            if remaining <= 0:
                return False
            if self._backend.wait_mutex(handle, min(remaining, MUTEX_POLL_SECONDS)):
                return True
        return False

    def _run_action(self, job: _Job) -> Any:
        backend = self._backend
        app = backend.find_running_app(job.database_path)
        if app is not None:
            log.info("Using running Access instance with %s", job.database_path)
            job.process_id = backend.process_id(app)
            return job.action(app)

        lock_path = lock_file_path(job.database_path)
        if lock_path.exists():
            if backend.is_lock_file_in_use(str(lock_path)):
                raise DatabaseInUseError(
                    f"Database {job.database_path} is opened by another process "
                    "that cannot be attached to. Please close it and try again."
                )
            log.info("Ignoring stale lock file %s", lock_path)

        log.info("Starting own Access instance for %s", job.database_path)
        app = backend.create_app()
        try:
            # Before opening: a security prompt may already appear in OpenCurrentDatabase.
            job.process_id = backend.process_id(app)
            app.OpenCurrentDatabase(job.database_path)
            return job.action(app)
        finally:
            _close_own_instance(app)
            _wait_for_lock_release(lock_path)


def _close_own_instance(app: Any) -> None:
    try:
        app.CloseCurrentDatabase()
    except Exception:
        log.warning("CloseCurrentDatabase failed", exc_info=True)
    try:
        app.Quit(AC_QUIT_SAVE_ALL)
    except Exception:
        log.warning("Quit of own Access instance failed", exc_info=True)


def _wait_for_lock_release(lock_path: Path) -> None:
    """Wait until Access removed the lock file; Quit returns before Access has exited."""
    deadline = time.monotonic() + LOCK_RELEASE_TIMEOUT_SECONDS
    while lock_path.exists():
        if time.monotonic() >= deadline:
            log.warning("Lock file %s still exists after closing Access", lock_path)
            return
        time.sleep(LOCK_RELEASE_POLL_SECONDS)
