"""Detection of modal dialogs of the Access instance (read only, never clicks).

Uses plain Win32 window enumeration (no COM), so it can run on any thread while
the COM worker thread is blocked by the dialog.
"""

from __future__ import annotations

import contextlib
from dataclasses import dataclass

# Standard dialog/message box class and the Office "NUI" dialogs (e.g. security notice).
DIALOG_CLASSES = frozenset({"#32770", "NUIDialog"})
ACCESS_MAIN_WINDOW_CLASS = "OMain"

_HINTS = (
    (
        ("speichern", "save"),
        "Access asks to save changes. This usually means the VBA project does not compile "
        "(saving modules via DoCmd fails); fix the compile error, then answer the prompts.",
    ),
    (
        ("sicherheit", "security", "vertrauens", "trust"),
        "Security prompt: put the add-in and the database into trusted locations.",
    ),
    (
        ("kompilier", "compile", "syntax"),
        "VBA compile error: fix the reported code line in the VBA editor.",
    ),
)

_HIDDEN_HINT = (
    "The dialog is not visible (hidden Access instance). Bring the Access window to the "
    "foreground or end the Access process in the Task Manager."
)


@dataclass(frozen=True)
class DialogInfo:
    hwnd: int
    class_name: str
    title: str
    text: str
    visible: bool = True

    def describe(self) -> str:
        parts = [f"'{self.title}'" if self.title else f"[{self.class_name}]"]
        if self.text:
            parts.append(self.text)
        description = ": ".join(parts)
        return description if self.visible else description + " (not visible)"


@dataclass(frozen=True)
class _Window:
    hwnd: int
    class_name: str
    visible: bool
    enabled: bool


def find_dialogs(process_id: int) -> list[DialogInfo]:
    """Return the dialogs blocking the Access instance of the given process.

    Visible dialogs always count. Invisible dialogs only count if the Access main
    window is disabled (modal state) and no visible dialog exists.
    """
    import win32gui

    chosen = select_dialogs(_top_level_windows(process_id))
    return [
        DialogInfo(
            hwnd=w.hwnd,
            class_name=w.class_name,
            title=win32gui.GetWindowText(w.hwnd),
            text=_dialog_text(w.hwnd),
            visible=w.visible,
        )
        for w in chosen
    ]


def _top_level_windows(process_id: int) -> list[_Window]:
    import win32gui
    import win32process

    windows: list[_Window] = []

    def on_window(hwnd: int, _: object) -> bool:
        if win32process.GetWindowThreadProcessId(hwnd)[1] == process_id:
            windows.append(
                _Window(
                    hwnd=hwnd,
                    class_name=win32gui.GetClassName(hwnd),
                    visible=bool(win32gui.IsWindowVisible(hwnd)),
                    enabled=bool(win32gui.IsWindowEnabled(hwnd)),
                )
            )
        return True

    win32gui.EnumWindows(on_window, None)
    return windows


def select_dialogs(windows: list[_Window]) -> list[_Window]:
    """Pick the blocking dialogs from the top-level windows of an Access process."""
    candidates = [w for w in windows if w.class_name in DIALOG_CLASSES and w.enabled]
    visible = [w for w in candidates if w.visible]
    if visible:
        return visible
    main_disabled = any(w.class_name == ACCESS_MAIN_WINDOW_CLASS and not w.enabled for w in windows)
    return candidates if main_disabled else []


def _dialog_text(hwnd: int) -> str:
    """Return the visible static texts of a dialog (the message of a message box)."""
    import win32gui

    texts: list[str] = []

    def on_child(child: int, _: object) -> bool:
        if win32gui.GetClassName(child) == "Static" and win32gui.IsWindowVisible(child):
            text = win32gui.GetWindowText(child).strip()
            if text:
                texts.append(" ".join(text.replace("&", "").split()))
        return True

    with contextlib.suppress(win32gui.error):
        win32gui.EnumChildWindows(hwnd, on_child, None)
    return " ".join(texts)


def describe_dialogs(dialogs: list[DialogInfo]) -> str:
    """Return a message listing the dialogs and hints on how to resolve them."""
    lines = [f"- {dialog.describe()}" for dialog in dialogs]
    combined = " ".join(f"{dialog.title} {dialog.text}" for dialog in dialogs).lower()
    hints = [hint for keywords, hint in _HINTS if any(word in combined for word in keywords)]
    if any(not dialog.visible for dialog in dialogs):
        hints.append(_HIDDEN_HINT)
    lines.extend(f"Hint: {hint}" for hint in hints)
    return "\n".join(lines)
