import os

from aclib_import_mcp import dialogs
from aclib_import_mcp.dialogs import DialogInfo


def dialog(title="Microsoft Access", text=""):
    return DialogInfo(hwnd=1, class_name="#32770", title=title, text=text)


def test_describe_lists_title_and_text():
    message = dialogs.describe_dialogs([dialog(text="Something happened")])
    assert "'Microsoft Access': Something happened" in message


def test_describe_without_title_uses_class_name():
    message = dialogs.describe_dialogs([DialogInfo(1, "NUIDialog", "", "")])
    assert "[NUIDialog]" in message


def test_hint_save_prompt_german():
    text = "Möchten Sie die Änderungen am Entwurf von Modul 'String2' speichern?"
    message = dialogs.describe_dialogs([dialog(text=text)])
    assert "does not compile" in message


def test_hint_security_notice():
    message = dialogs.describe_dialogs([dialog(title="Microsoft Access-Sicherheitshinweis")])
    assert "trusted locations" in message


def test_hint_compile_error_english():
    message = dialogs.describe_dialogs(
        [
            dialog(
                title="Microsoft Visual Basic for Applications", text="Compile error: Syntax error"
            )
        ]
    )
    assert "compile error" in message.lower()
    assert "VBA editor" in message


def test_no_hint_for_unknown_dialog():
    assert "Hint" not in dialogs.describe_dialogs([dialog(text="Hello")])


def test_find_dialogs_of_own_process_has_none():
    # The test process has no visible dialogs; only checks that enumeration works.
    assert dialogs.find_dialogs(os.getpid()) == []


def window(class_name, visible=True, enabled=True, hwnd=1):
    return dialogs._Window(hwnd=hwnd, class_name=class_name, visible=visible, enabled=enabled)


def test_select_visible_dialog():
    windows = [window("OMain", visible=False, enabled=False), window("#32770", hwnd=2)]
    assert [w.hwnd for w in dialogs.select_dialogs(windows)] == [2]


def test_select_hidden_dialog_when_main_window_disabled():
    windows = [window("OMain", visible=False, enabled=False), window("#32770", visible=False)]
    assert len(dialogs.select_dialogs(windows)) == 1


def test_ignore_hidden_dialog_when_main_window_enabled():
    windows = [window("OMain", visible=False, enabled=True), window("#32770", visible=False)]
    assert dialogs.select_dialogs(windows) == []


def test_ignore_disabled_and_other_windows():
    windows = [
        window("OMain", visible=False, enabled=False),
        window("#32770", enabled=False),
        window("wndclass_desked_gsk"),
    ]
    assert dialogs.select_dialogs(windows) == []


def test_hidden_dialog_gets_hint():
    message = dialogs.describe_dialogs([DialogInfo(1, "#32770", "Speichern", "", visible=False)])
    assert "(not visible)" in message
    assert "Task Manager" in message
