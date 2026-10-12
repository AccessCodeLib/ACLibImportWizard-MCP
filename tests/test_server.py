import json

import anyio
import pytest
from mcp.server.fastmcp.exceptions import ToolError

from aclib_import_mcp import addin, server
from aclib_import_mcp.access_com import DatabaseInUseError

DB = r"C:\Data\test.accdb"
VERSION_OK = json.dumps({"ok": True, "version": "1.5.0", "api": 1})


class FakeApp:
    def __init__(self, results):
        self.results = results
        self.calls = []

    def Run(self, procedure, command, *args):
        self.calls.append((command, *args))
        return self.results[command]


class FakeWorker:
    def __init__(self, app=None, error=None):
        self.app = app
        self.error = error
        self.database_paths = []

    def execute(self, database_path, action):
        self.database_paths.append(database_path)
        if self.error is not None:
            raise self.error
        return action(self.app)


@pytest.fixture
def use_worker(monkeypatch, tmp_path):
    addin_file = tmp_path / "ACLibImportWizard.accda"
    addin_file.write_bytes(b"")
    monkeypatch.setenv(addin.ADDIN_PATH_ENV, str(addin_file))

    def factory(worker):
        monkeypatch.setattr(server, "get_worker", lambda: worker)
        return worker

    return factory


def call(tool, **kwargs):
    async def run():
        return await tool(**kwargs)

    return anyio.run(run)


def test_build_options_defaults():
    assert server.build_options() == "mode=missing;tests=0;examples=0;source=config"


def test_build_options_all_values():
    options = server.build_options("all", True, True, "github", "owner/repo@dev")
    assert options == "mode=all;tests=1;examples=1;source=github;repo=owner/repo@dev"


@pytest.mark.parametrize("repo", ["owner", "owner/repo;mode=all", "owner/repo@x=y", "/repo"])
def test_invalid_repo(repo):
    with pytest.raises(ToolError):
        server.build_source_options("github", repo)


def test_repo_with_local_source():
    with pytest.raises(ToolError, match="source='local'"):
        server.build_source_options("local", "owner/repo")


def test_validate_files_normalizes_backslashes():
    assert server.validate_files([" data\\SqlTools.cls ", "text/String2.cls"]) == [
        "data/SqlTools.cls",
        "text/String2.cls",
    ]


@pytest.mark.parametrize(
    "file",
    ["", "a.cls;b.cls", "../secret.bas", "data/../../x.bas", "/abs.bas", r"C:\x.bas", "a=b.cls"],
)
def test_validate_files_rejects(file):
    with pytest.raises(ToolError):
        server.validate_files([file])


@pytest.mark.parametrize("path", ["test.accdb", r"C:\Data\test.xlsx", r"C:\Data\test.accde"])
def test_validate_database_path_rejects(path):
    with pytest.raises(ToolError):
        server.validate_database_path(path)


def test_get_version(use_worker):
    app = FakeApp({"GetVersion": VERSION_OK})
    use_worker(FakeWorker(app))

    result = call(server.aclib_get_version, database_path=DB)

    assert result["api"] == 1
    assert app.calls == [("GetVersion",)]


def test_list_installed_checks_version_first(use_worker):
    app = FakeApp(
        {"GetVersion": VERSION_OK, "ListInstalled": json.dumps({"ok": True, "modules": []})}
    )
    worker = use_worker(FakeWorker(app))

    result = call(server.aclib_list_installed, database_path=DB)

    assert result == {"ok": True, "modules": []}
    assert app.calls == [("GetVersion",), ("ListInstalled",)]
    assert worker.database_paths == [DB]


def test_resolve_passes_files_and_options(use_worker):
    app = FakeApp({"GetVersion": VERSION_OK, "Resolve": json.dumps({"ok": True, "files": []})})
    use_worker(FakeWorker(app))

    call(
        server.aclib_resolve,
        database_path=DB,
        files=["data/SqlTools.cls", "text/String2.cls"],
        tests=True,
    )

    assert app.calls[-1] == (
        "Resolve",
        "data/SqlTools.cls;text/String2.cls",
        "mode=missing;tests=1;examples=0;source=config",
    )


def test_import_passes_files_and_options(use_worker):
    app = FakeApp({"GetVersion": VERSION_OK, "Import": json.dumps({"ok": True, "files": []})})
    use_worker(FakeWorker(app))

    call(
        server.aclib_import,
        database_path=DB,
        files=["data/SqlTools.cls"],
        mode="all",
        repo="AccessCodeLib/AccessCodeLib@dev",
    )

    assert app.calls[-1] == (
        "Import",
        "data/SqlTools.cls",
        "mode=all;tests=0;examples=0;source=config;repo=AccessCodeLib/AccessCodeLib@dev",
    )


def test_get_config_passes_source_options(use_worker):
    app = FakeApp({"GetVersion": VERSION_OK, "GetConfig": json.dumps({"ok": True})})
    use_worker(FakeWorker(app))

    call(server.aclib_get_config, database_path=DB, source="github")

    assert app.calls[-1] == ("GetConfig", "source=github")


def test_old_addin_without_api_version(use_worker):
    app = FakeApp({"GetVersion": json.dumps({"ok": True, "version": "1.4.0"})})
    use_worker(FakeWorker(app))

    with pytest.raises(ToolError, match="API version"):
        call(server.aclib_list_installed, database_path=DB)
    assert app.calls == [("GetVersion",)]


def test_api_error_becomes_tool_error(use_worker):
    error = json.dumps(
        {"ok": False, "error": {"number": 5, "source": "Configure", "description": "Bad option"}}
    )
    app = FakeApp({"GetVersion": VERSION_OK, "Resolve": error})
    use_worker(FakeWorker(app))

    with pytest.raises(ToolError, match="Bad option.*number=5"):
        call(server.aclib_resolve, database_path=DB, files=["data/SqlTools.cls"])


def test_access_error_becomes_tool_error(use_worker):
    use_worker(FakeWorker(error=DatabaseInUseError("opened by another process")))

    with pytest.raises(ToolError, match="opened by another process"):
        call(server.aclib_list_installed, database_path=DB)


def test_tools_are_registered_with_annotations():
    tools = {tool.name: tool for tool in anyio.run(server.mcp.list_tools)}

    assert set(tools) == {
        "aclib_get_version",
        "aclib_get_config",
        "aclib_list_installed",
        "aclib_resolve",
        "aclib_import",
        "aclib_compile",
    }
    for name in ("aclib_get_version", "aclib_get_config", "aclib_list_installed", "aclib_resolve"):
        assert tools[name].annotations.readOnlyHint is True
    assert tools["aclib_import"].annotations.readOnlyHint is False
    assert tools["aclib_import"].annotations.destructiveHint is True
    assert tools["aclib_compile"].annotations.readOnlyHint is False


class CompileApp:
    def __init__(self, compiled, error=None):
        self.IsCompiled = compiled
        self.error = error
        self.commands = []

    def RunCommand(self, command):
        self.commands.append(command)
        if self.error:
            raise self.error


def test_compile_project_ok():
    app = CompileApp(compiled=True)
    assert server.compile_project(app) == {"ok": True, "compiled": True}
    assert app.commands == [server.AC_CMD_COMPILE_AND_SAVE_ALL_MODULES]


def test_compile_project_not_compiled_adds_hint():
    result = server.compile_project(CompileApp(compiled=False))
    assert result == {"ok": True, "compiled": False, "message": server.NOT_COMPILED_MESSAGE}


def test_compile_project_not_compiled_with_message():
    app = CompileApp(compiled=False, error=RuntimeError("command not available"))
    result = server.compile_project(app)
    assert result == {"ok": True, "compiled": False, "message": "command not available"}


def test_compile_tool_uses_worker(use_worker):
    app = CompileApp(compiled=True)
    worker = use_worker(FakeWorker(app))

    result = call(server.aclib_compile, database_path=DB)

    assert result["compiled"] is True
    assert worker.database_paths == [DB]
