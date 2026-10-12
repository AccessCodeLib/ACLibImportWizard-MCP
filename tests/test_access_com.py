import threading
import time

import pytest

from aclib_import_mcp import access_com
from aclib_import_mcp.access_com import (
    AccessWorker,
    CallTimeoutError,
    DatabaseBusyError,
    DatabaseInUseError,
    DatabaseNotFoundError,
    DialogBlockingError,
    WorkerBlockedError,
)
from aclib_import_mcp.dialogs import DialogInfo


class FakeApp:
    def __init__(self):
        self.calls = []

    def OpenCurrentDatabase(self, path):
        self.calls.append(("OpenCurrentDatabase", path))

    def CloseCurrentDatabase(self):
        self.calls.append(("CloseCurrentDatabase",))

    def Quit(self, option):
        self.calls.append(("Quit", option))


class FakeBackend:
    def __init__(self, running_app=None, lock_in_use=False, mutex_free=True):
        self.running_app = running_app
        self.lock_in_use = lock_in_use
        self.mutex_free = mutex_free
        self.created_apps = []
        self.mutex_events = []
        self.threads = set()
        self.initialized_thread = None

    def initialize(self):
        self.initialized_thread = threading.get_ident()

    def uninitialize(self):
        pass

    def find_running_app(self, database_path):
        self.threads.add(threading.get_ident())
        return self.running_app

    def create_app(self):
        app = FakeApp()
        self.created_apps.append(app)
        return app

    def process_id(self, app):
        app.calls.append(("process_id",))
        return 4711

    def is_lock_file_in_use(self, lock_path):
        return self.lock_in_use

    def create_mutex(self, name):
        self.mutex_events.append(("create", name))
        return name

    def wait_mutex(self, handle, timeout_seconds):
        if self.mutex_free:
            self.mutex_events.append(("acquire", handle))
            return True
        time.sleep(timeout_seconds)
        return False

    def release_mutex(self, handle):
        self.mutex_events.append(("release", handle))

    def close_handle(self, handle):
        self.mutex_events.append(("close", handle))


@pytest.fixture
def database(tmp_path):
    path = tmp_path / "test.accdb"
    path.write_bytes(b"")
    return path


@pytest.fixture(autouse=True)
def fast_polling(monkeypatch):
    monkeypatch.setattr(access_com, "DIALOG_POLL_SECONDS", 0.02)


@pytest.fixture
def make_worker():
    workers = []

    def factory(backend, timeout=5.0, dialog_timeout=5.0, dialog_finder=lambda pid: []):
        worker = AccessWorker(
            backend=backend,
            timeout=timeout,
            dialog_timeout=dialog_timeout,
            dialog_finder=dialog_finder,
        )
        workers.append(worker)
        return worker

    yield factory
    for worker in workers:
        worker.close()


SAVE_PROMPT = DialogInfo(
    hwnd=1,
    class_name="#32770",
    title="Microsoft Access",
    text="Möchten Sie die Änderungen am Entwurf von Modul 'String2' speichern?",
)


def test_get_call_timeout_default(monkeypatch):
    monkeypatch.delenv(access_com.CALL_TIMEOUT_ENV, raising=False)
    assert access_com.get_call_timeout() == access_com.DEFAULT_CALL_TIMEOUT


def test_get_call_timeout_from_env(monkeypatch):
    monkeypatch.setenv(access_com.CALL_TIMEOUT_ENV, "300")
    assert access_com.get_call_timeout() == 300.0


@pytest.mark.parametrize("raw", ["abc", "0", "-5"])
def test_get_call_timeout_invalid(monkeypatch, raw):
    monkeypatch.setenv(access_com.CALL_TIMEOUT_ENV, raw)
    with pytest.raises(ValueError):
        access_com.get_call_timeout()


@pytest.mark.parametrize(
    ("name", "lock"),
    [("a.accdb", "a.laccdb"), ("a.ACCDB", "a.laccdb"), ("a.mdb", "a.ldb"), ("a.accda", "a.laccdb")],
)
def test_lock_file_path(tmp_path, name, lock):
    assert access_com.lock_file_path(tmp_path / name) == tmp_path / lock


def test_mutex_name_is_case_insensitive(tmp_path):
    lower = access_com.mutex_name(tmp_path / "test.accdb")
    upper = access_com.mutex_name(tmp_path / "TEST.ACCDB")
    assert lower == upper
    assert lower.startswith("Local\\aclib-import-mcp-db-")
    assert "\\" not in lower.removeprefix("Local\\")


def test_missing_database(make_worker, tmp_path):
    worker = make_worker(FakeBackend())
    with pytest.raises(DatabaseNotFoundError):
        worker.execute(tmp_path / "missing.accdb", lambda app: None)


def test_uses_running_instance_without_closing(make_worker, database):
    running = FakeApp()
    backend = FakeBackend(running_app=running)
    worker = make_worker(backend)

    result = worker.execute(database, lambda app: app)

    assert result is running
    assert running.calls == [("process_id",)]  # no Open/Close/Quit
    assert backend.created_apps == []


def test_own_instance_is_opened_and_closed(make_worker, database):
    backend = FakeBackend()
    worker = make_worker(backend)

    worker.execute(database, lambda app: "done")

    (app,) = backend.created_apps
    assert app.calls == [
        ("process_id",),
        ("OpenCurrentDatabase", str(database)),
        ("CloseCurrentDatabase",),
        ("Quit", access_com.AC_QUIT_SAVE_ALL),
    ]


def test_waits_for_lock_release_after_quit(make_worker, database):
    lock = access_com.lock_file_path(database)
    backend = FakeBackend()
    original_create_app = backend.create_app

    def create_app():
        app = original_create_app()
        lock.write_bytes(b"")  # Access creates the lock file when opening the database
        quit_app = app.Quit

        def delayed_quit(option):
            quit_app(option)
            threading.Timer(0.3, lock.unlink).start()  # removed after Quit returned

        app.Quit = delayed_quit
        return app

    backend.create_app = create_app
    worker = make_worker(backend)

    worker.execute(database, lambda app: None)

    assert not lock.exists()


def test_own_instance_is_closed_on_error(make_worker, database):
    backend = FakeBackend()
    worker = make_worker(backend)

    def action(app):
        raise RuntimeError("boom")

    with pytest.raises(RuntimeError, match="boom"):
        worker.execute(database, action)

    (app,) = backend.created_apps
    assert ("Quit", access_com.AC_QUIT_SAVE_ALL) in app.calls


def test_lock_file_in_use_prevents_second_instance(make_worker, database):
    access_com.lock_file_path(database).write_bytes(b"")
    backend = FakeBackend(lock_in_use=True)
    worker = make_worker(backend)

    with pytest.raises(DatabaseInUseError):
        worker.execute(database, lambda app: None)

    assert backend.created_apps == []


def test_stale_lock_file_is_ignored(make_worker, database, monkeypatch):
    monkeypatch.setattr(access_com, "LOCK_RELEASE_TIMEOUT_SECONDS", 0.1)
    access_com.lock_file_path(database).write_bytes(b"")
    backend = FakeBackend(lock_in_use=False)
    worker = make_worker(backend)

    assert worker.execute(database, lambda app: "done") == "done"
    assert len(backend.created_apps) == 1


def test_mutex_is_released_after_error(make_worker, database):
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(backend)

    def action(app):
        raise RuntimeError("boom")

    with pytest.raises(RuntimeError):
        worker.execute(database, action)

    kinds = [kind for kind, _ in backend.mutex_events]
    assert kinds == ["create", "acquire", "release", "close"]
    assert backend.mutex_events[0][1] == access_com.mutex_name(database)


def test_busy_mutex_raises_and_closes_handle(make_worker, database, monkeypatch):
    monkeypatch.setattr(access_com, "MUTEX_POLL_SECONDS", 0.05)
    backend = FakeBackend(mutex_free=False)
    worker = make_worker(backend, timeout=0.3)

    with pytest.raises(DatabaseBusyError):
        worker.execute(database, lambda app: None)

    kinds = [kind for kind, _ in backend.mutex_events]
    assert "acquire" not in kinds
    assert "release" not in kinds
    assert kinds[-1] == "close"
    assert backend.created_apps == []


def test_timeout(make_worker, database, monkeypatch):
    monkeypatch.setattr(access_com, "RESULT_GRACE_SECONDS", 0.0)
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(backend, timeout=0.1)

    with pytest.raises(CallTimeoutError, match="ACLIB_CALL_TIMEOUT"):
        worker.execute(database, lambda app: time.sleep(0.5))


def test_abandoned_queued_job_is_skipped(make_worker, database, monkeypatch):
    monkeypatch.setattr(access_com, "RESULT_GRACE_SECONDS", 0.0)
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(backend, timeout=0.3)
    release = threading.Event()
    ran = []
    errors = []

    def first():
        try:
            worker.execute(database, lambda app: release.wait(5))
        except CallTimeoutError as exc:
            errors.append(exc)

    caller = threading.Thread(target=first)
    caller.start()
    time.sleep(0.05)  # first job is running, not yet abandoned -> second job is queued
    with pytest.raises(CallTimeoutError):
        worker.execute(database, lambda app: ran.append(True))
    caller.join()
    release.set()

    worker.close()
    assert len(errors) == 1
    assert ran == []


def test_call_after_abandoned_running_job_fails_fast(make_worker, database, monkeypatch):
    monkeypatch.setattr(access_com, "RESULT_GRACE_SECONDS", 0.0)
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(backend, timeout=0.1)
    release = threading.Event()

    with pytest.raises(CallTimeoutError):
        worker.execute(database, lambda app: release.wait(5))
    with pytest.raises(WorkerBlockedError, match="still running or blocked"):
        worker.execute(database, lambda app: None)
    release.set()


def test_calls_run_sequentially_on_one_thread(make_worker, database):
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(backend)
    active = 0
    max_active = 0
    action_threads = set()
    counter_lock = threading.Lock()

    def action(app):
        nonlocal active, max_active
        with counter_lock:
            active += 1
            max_active = max(max_active, active)
        action_threads.add(threading.get_ident())
        time.sleep(0.05)
        with counter_lock:
            active -= 1

    callers = [threading.Thread(target=worker.execute, args=(database, action)) for _ in range(5)]
    for caller in callers:
        caller.start()
    for caller in callers:
        caller.join()

    assert max_active == 1
    assert len(action_threads) == 1
    assert action_threads == {backend.initialized_thread}
    assert threading.get_ident() not in action_threads


def test_get_dialog_timeout(monkeypatch):
    env = access_com.DIALOG_TIMEOUT_ENV
    monkeypatch.delenv(env, raising=False)
    assert access_com.get_dialog_timeout() == access_com.DEFAULT_DIALOG_TIMEOUT == 20.0
    monkeypatch.setenv(env, "45")
    assert access_com.get_dialog_timeout() == 45.0


def test_dialog_reports_early_and_blocks_further_calls(make_worker, database):
    release = threading.Event()
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(
        backend,
        timeout=30.0,
        dialog_timeout=0.2,
        dialog_finder=lambda pid: [] if release.is_set() else [SAVE_PROMPT],
    )

    start = time.monotonic()
    with pytest.raises(DialogBlockingError) as info:
        worker.execute(database, lambda app: release.wait(5))
    assert time.monotonic() - start < 5
    message = str(info.value)
    assert "String2" in message
    assert "does not compile" in message
    assert "ACLIB_DIALOG_TIMEOUT" in message

    with pytest.raises(WorkerBlockedError, match="String2"):
        worker.execute(database, lambda app: "next")

    release.set()  # user closed the dialog
    deadline = time.monotonic() + 5
    while worker._current is not None and time.monotonic() < deadline:
        time.sleep(0.02)
    assert worker.execute(database, lambda app: "next") == "next"


def test_short_dialog_does_not_fail(make_worker, database):
    shown_until = time.monotonic() + 0.1
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(
        backend,
        dialog_timeout=0.3,
        dialog_finder=lambda pid: [SAVE_PROMPT] if time.monotonic() < shown_until else [],
    )

    assert worker.execute(database, lambda app: time.sleep(0.5) or "done") == "done"


def test_dialog_finder_gets_process_id(make_worker, database):
    seen = []
    backend = FakeBackend()
    worker = make_worker(backend, dialog_finder=lambda pid: seen.append(pid) or [])

    worker.execute(database, lambda app: time.sleep(0.1))

    assert set(seen) == {4711}


def test_timeout_message_lists_dialogs(make_worker, database, monkeypatch):
    monkeypatch.setattr(access_com, "RESULT_GRACE_SECONDS", 0.0)
    release = threading.Event()
    backend = FakeBackend(running_app=FakeApp())
    worker = make_worker(
        backend, timeout=0.2, dialog_timeout=10.0, dialog_finder=lambda pid: [SAVE_PROMPT]
    )

    with pytest.raises(CallTimeoutError, match="String2"):
        worker.execute(database, lambda app: release.wait(5))
    release.set()


class MethodHandleApp:
    def hWndAccessApp(self):
        return 2428106


class PropertyHandleApp:
    hWndAccessApp = 4711


def test_access_window_handle_calls_method():
    assert access_com.access_window_handle(MethodHandleApp()) == 2428106


def test_access_window_handle_accepts_property():
    assert access_com.access_window_handle(PropertyHandleApp()) == 4711
