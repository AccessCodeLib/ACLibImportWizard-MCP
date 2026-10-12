"""ACLib Import Wizard add-in: path from configuration and API version check.

The add-in path is only taken from configuration (``ACLIB_ADDIN_PATH`` or the
default location), never from tool parameters, and only the add-in's ``API``
function is called.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Protocol

ADDIN_PATH_ENV = "ACLIB_ADDIN_PATH"
ADDIN_FILE_NAME = "ACLibImportWizard.accda"
API_FUNCTION = "API"
MIN_API_VERSION = 1


class AddinError(Exception):
    """Base class of all add-in related errors."""


class AddinNotFoundError(AddinError):
    """The add-in file does not exist."""


class AddinVersionError(AddinError):
    """The add-in does not provide the required API version."""


class AclibApiError(AddinError):
    """The add-in API returned ``ok: false`` or an invalid result."""

    def __init__(self, description: str, number: int | None = None, source: str | None = None):
        super().__init__(description)
        self.description = description
        self.number = number
        self.source = source

    def __str__(self) -> str:
        parts = [self.description]
        if self.number is not None:
            parts.append(f"number={self.number}")
        if self.source:
            parts.append(f"source={self.source}")
        return " | ".join(parts)


class Worker(Protocol):
    def execute(self, database_path: str | os.PathLike[str], action: Any) -> Any: ...


def default_addin_path() -> Path:
    """Return ``%APPDATA%\\Microsoft\\AddIns\\ACLibImportWizard.accda``."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise AddinNotFoundError("APPDATA is not set; set ACLIB_ADDIN_PATH instead.")
    return Path(appdata) / "Microsoft" / "AddIns" / ADDIN_FILE_NAME


def get_addin_path() -> Path:
    """Return the configured add-in path; raise if the file does not exist."""
    raw = os.environ.get(ADDIN_PATH_ENV, "").strip()
    path = Path(os.path.expandvars(raw)) if raw else default_addin_path()
    if not path.is_file():
        source = ADDIN_PATH_ENV if raw else "default location"
        raise AddinNotFoundError(f"ACLib Import Wizard add-in not found ({source}): {path}")
    return path


def api_procedure(addin_path: str | os.PathLike[str]) -> str:
    """Return the procedure name for ``Application.Run``: add-in path without extension + ``.API``."""
    return f"{Path(addin_path).with_suffix('')}.{API_FUNCTION}"


def parse_result(raw: Any) -> dict[str, Any]:
    """Parse the JSON result of the API; raise :class:`AclibApiError` on ``ok: false``."""
    if not isinstance(raw, str):
        raise AclibApiError(f"Unexpected API result type: {type(raw).__name__}")
    try:
        result = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise AclibApiError(f"API result is not valid JSON: {exc}") from None
    if not isinstance(result, dict):
        raise AclibApiError("API result is not a JSON object")
    if result.get("ok") is True:
        return result
    error = result.get("error")
    if not isinstance(error, dict):
        raise AclibApiError("API returned ok=false without error details")
    number = error.get("number")
    raise AclibApiError(
        str(error.get("description") or "Unknown API error"),
        number=number if isinstance(number, int) else None,
        source=error.get("source"),
    )


def check_version(result: dict[str, Any]) -> int:
    """Return the API version of a ``GetVersion`` result; raise if it is too old."""
    api = result.get("api")
    if not isinstance(api, int) or isinstance(api, bool):
        raise AddinVersionError(
            "The ACLib Import Wizard add-in does not report an API version; "
            "an add-in version with automation API is required."
        )
    if api < MIN_API_VERSION:
        raise AddinVersionError(
            f"The ACLib Import Wizard add-in provides API {api}; "
            f"at least API {MIN_API_VERSION} is required."
        )
    return api


def call_api(
    worker: Worker,
    database_path: str | os.PathLike[str],
    command: str,
    *args: str,
    addin_path: str | os.PathLike[str] | None = None,
    check_api_version: bool = False,
) -> dict[str, Any]:
    """Call ``<add-in>.API(command, *args)`` in the given database and return the parsed result.

    With ``check_api_version``, ``GetVersion`` is called first in the same Access session.
    """
    procedure = api_procedure(addin_path if addin_path is not None else get_addin_path())

    def action(app: Any) -> Any:
        if check_api_version:
            check_version(parse_result(_run(app, procedure, "GetVersion")))
        return _run(app, procedure, command, *args)

    return parse_result(worker.execute(database_path, action))


def _run(app: Any, procedure: str, *args: str) -> Any:
    """Call ``Application.Run`` and return the function result.

    pywin32 returns ``(result, *byref_args)`` because the arguments of ``Run`` are ByRef.
    """
    result = app.Run(procedure, *args)
    if isinstance(result, tuple):
        return result[0] if result else None
    return result


def get_version(
    worker: Worker,
    database_path: str | os.PathLike[str],
    addin_path: str | os.PathLike[str] | None = None,
) -> dict[str, Any]:
    """Call ``GetVersion`` and check that the add-in provides the required API version."""
    result = call_api(worker, database_path, "GetVersion", addin_path=addin_path)
    check_version(result)
    return result
