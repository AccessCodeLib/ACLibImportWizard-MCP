import aclib_import_mcp
from aclib_import_mcp import server


def test_version_is_set():
    assert aclib_import_mcp.__version__


def test_server_has_entry_point():
    assert callable(server.main)
    assert server.mcp.name == "aclib-import"
