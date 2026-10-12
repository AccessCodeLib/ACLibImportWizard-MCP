import json
from pathlib import Path

import pytest

from aclib_import_mcp import addin
from aclib_import_mcp.addin import AclibApiError, AddinNotFoundError, AddinVersionError


class FakeApp:
    def __init__(self, result):
        self.result = result
        self.calls = []

    def Run(self, *args):
        self.calls.append(args)
        return self.result


class FakeWorker:
    def __init__(self, app):
        self.app = app
        self.database_paths = []

    def execute(self, database_path, action):
        self.database_paths.append(database_path)
        return action(self.app)


@pytest.fixture
def addin_file(tmp_path):
    path = tmp_path / "ACLibImportWizard.accda"
    path.write_bytes(b"")
    return path


def test_default_addin_path_uses_appdata(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert (
        addin.default_addin_path() == tmp_path / "Microsoft" / "AddIns" / "ACLibImportWizard.accda"
    )


def test_get_addin_path_default_location(monkeypatch, tmp_path):
    monkeypatch.delenv(addin.ADDIN_PATH_ENV, raising=False)
    monkeypatch.setenv("APPDATA", str(tmp_path))
    expected = tmp_path / "Microsoft" / "AddIns" / "ACLibImportWizard.accda"
    expected.parent.mkdir(parents=True)
    expected.write_bytes(b"")
    assert addin.get_addin_path() == expected


def test_get_addin_path_from_env(monkeypatch, addin_file):
    monkeypatch.setenv(addin.ADDIN_PATH_ENV, str(addin_file))
    assert addin.get_addin_path() == addin_file


def test_get_addin_path_missing_file(monkeypatch, tmp_path):
    monkeypatch.setenv(addin.ADDIN_PATH_ENV, str(tmp_path / "missing.accda"))
    with pytest.raises(AddinNotFoundError, match="ACLIB_ADDIN_PATH"):
        addin.get_addin_path()


def test_api_procedure_strips_extension():
    path = Path(r"C:\AddIns\ACLibImportWizard.accda")
    assert addin.api_procedure(path) == r"C:\AddIns\ACLibImportWizard.API"


def test_parse_result_ok():
    assert addin.parse_result('{"ok":true,"api":1}') == {"ok": True, "api": 1}


def test_parse_result_error():
    raw = json.dumps(
        {"ok": False, "error": {"number": 5, "source": "modApi", "description": "Bad command"}}
    )
    with pytest.raises(AclibApiError) as info:
        addin.parse_result(raw)
    assert info.value.number == 5
    assert info.value.source == "modApi"
    assert info.value.description == "Bad command"
    assert "Bad command" in str(info.value)


@pytest.mark.parametrize("raw", ["not json", "[1, 2]", None, '{"ok": false}'])
def test_parse_result_invalid(raw):
    with pytest.raises(AclibApiError):
        addin.parse_result(raw)


def test_check_version_ok():
    assert addin.check_version({"ok": True, "api": 1}) == 1


@pytest.mark.parametrize(
    "result", [{"ok": True}, {"ok": True, "api": 0}, {"ok": True, "api": True}]
)
def test_check_version_invalid(result):
    with pytest.raises(AddinVersionError):
        addin.check_version(result)


def test_call_api_runs_api_procedure(addin_file, tmp_path):
    app = FakeApp('{"ok":true,"files":[]}')
    worker = FakeWorker(app)
    db = tmp_path / "test.accdb"

    result = addin.call_api(
        worker, db, "Resolve", "data/SqlTools.cls", "mode=missing", addin_path=addin_file
    )

    assert result == {"ok": True, "files": []}
    assert worker.database_paths == [db]
    assert app.calls == [
        (str(addin_file.with_suffix("")) + ".API", "Resolve", "data/SqlTools.cls", "mode=missing")
    ]


def test_call_api_unwraps_byref_tuple(addin_file, tmp_path):
    # pywin32 returns (result, *byref_args) for Application.Run
    app = FakeApp(('{"ok":true,"api":1}', "GetVersion", None, None))
    result = addin.call_api(
        FakeWorker(app), tmp_path / "test.accdb", "GetVersion", addin_path=addin_file
    )
    assert result == {"ok": True, "api": 1}


def test_get_version_checks_api(addin_file, tmp_path):
    worker = FakeWorker(FakeApp('{"ok":true,"version":"1.4.0"}'))
    with pytest.raises(AddinVersionError):
        addin.get_version(worker, tmp_path / "test.accdb", addin_path=addin_file)
