"""MCP tools, parameter validation and building the API options string."""

from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path, PurePosixPath
from typing import Annotated, Any, Literal

import anyio.to_thread
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from . import addin, catalog
from .access_com import AccessComError, AccessWorker
from .addin import AclibApiError, AddinError

INSTRUCTIONS = """\
ACLib Import Wizard MCP server: import or update AccessCodeLib (ACLib) modules in a Microsoft \
Access database.

Use the skill aclib-import before working with the aclib_* tools. If it is not available in the \
project, tell the user and offer to continue without it only after they agree.

Essential rules if the skill is not available:
- To find a module for a task use aclib_find (search in the module catalog). Never guess file \
paths: take them from aclib_find, the catalog (_codelib/catalog/catalog.md) or from the user.
- Order: aclib_get_version, aclib_list_installed, aclib_resolve (show the planned actions to the \
user), then aclib_import and aclib_compile only after the user approved them.
- mode=all/selected and the action overwrite only after explicit confirmation: local edits of a \
module in the project would be lost.
"""

mcp = FastMCP("aclib-import", instructions=INSTRUCTIONS)

Mode = Literal["missing", "all", "selected"]
Source = Literal["config", "local", "github"]

DATABASE_EXTENSIONS = {".accdb", ".mdb"}
AC_CMD_COMPILE_AND_SAVE_ALL_MODULES = 126
NOT_COMPILED_MESSAGE = (
    "Project does not compile; open the VBA editor (Debug > Compile) to see the error."
)
REPO_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9._-]+(@[A-Za-z0-9._/-]+)?$")
_FORBIDDEN_FILE_CHARS = re.compile(r"[;=:*?\"<>|\x00-\x1f]")

DatabasePath = Annotated[
    str,
    Field(description="Full path of the target Access database (.accdb/.mdb)."),
]
Files = Annotated[
    list[str],
    Field(
        description="ACLib repository paths, e.g. ['data/SqlTools.cls', 'text/String2.cls'].",
        min_length=1,
    ),
]
ModeParam = Annotated[
    Mode,
    Field(
        description="missing: import only modules not yet in the database (default); "
        "all: also overwrite existing dependencies; selected: only the given files, no dependencies."
    ),
]
TestsParam = Annotated[bool, Field(description="Also import the test modules.")]
ExamplesParam = Annotated[bool, Field(description="Also import the example modules.")]
SourceParam = Annotated[
    Source,
    Field(
        description="config: local repository if configured, otherwise GitHub (default); "
        "local: configured local repository; github: GitHub repository."
    ),
]
RepoParam = Annotated[
    str | None,
    Field(
        description="GitHub repository 'owner/name[@branch]' (implies source=github). "
        "Default: AccessCodeLib/AccessCodeLib@master."
    ),
]

_worker: AccessWorker | None = None


def get_worker() -> AccessWorker:
    global _worker
    if _worker is None:
        _worker = AccessWorker()
    return _worker


def validate_database_path(database_path: str) -> str:
    path = Path(database_path.strip())
    if not path.is_absolute():
        raise ToolError(f"database_path must be an absolute path: {database_path!r}")
    if path.suffix.lower() not in DATABASE_EXTENSIONS:
        raise ToolError(f"database_path must be an .accdb or .mdb file: {database_path!r}")
    return str(path)


def validate_files(files: list[str]) -> list[str]:
    result = []
    for raw in files:
        file = raw.strip().replace("\\", "/")
        if not file:
            raise ToolError("files must not contain empty entries")
        if _FORBIDDEN_FILE_CHARS.search(file):
            raise ToolError(f"Invalid character in file {raw!r}")
        parts = PurePosixPath(file).parts
        if file.startswith("/") or ".." in parts:
            raise ToolError(f"files must be relative repository paths without '..': {raw!r}")
        result.append(file)
    if not result:
        raise ToolError("files must contain at least one entry")
    return result


def validate_repo(repo: str | None, source: Source) -> str | None:
    if repo is None or not repo.strip():
        return None
    repo = repo.strip()
    if not REPO_PATTERN.match(repo):
        raise ToolError(f"repo must have the form 'owner/name[@branch]': {repo!r}")
    if source == "local":
        raise ToolError("repo can not be combined with source='local'")
    return repo


def build_source_options(source: Source = "config", repo: str | None = None) -> str:
    """Return the ``source``/``repo`` part of the API options string."""
    repo = validate_repo(repo, source)
    items = [f"source={source}"]
    if repo:
        items.append(f"repo={repo}")
    return ";".join(items)


def build_options(
    mode: Mode = "missing",
    tests: bool = False,
    examples: bool = False,
    source: Source = "config",
    repo: str | None = None,
) -> str:
    """Return the API options string, e.g. ``mode=missing;tests=0;examples=0;source=config``."""
    return ";".join(
        [
            f"mode={mode}",
            f"tests={int(tests)}",
            f"examples={int(examples)}",
            build_source_options(source, repo),
        ]
    )


async def _run_in_worker(run: Callable[[], dict[str, Any]]) -> dict[str, Any]:
    try:
        return await anyio.to_thread.run_sync(run)
    except AclibApiError as exc:
        raise ToolError(f"ACLib Import Wizard API error: {exc}") from exc
    except (AddinError, AccessComError, ValueError) as exc:
        raise ToolError(str(exc)) from exc


async def _call_api(
    database_path: str, command: str, *args: str, check_api_version: bool = True
) -> dict[str, Any]:
    database_path = validate_database_path(database_path)
    worker = get_worker()

    def run() -> dict[str, Any]:
        if command == "GetVersion":
            return addin.get_version(worker, database_path)
        return addin.call_api(
            worker, database_path, command, *args, check_api_version=check_api_version
        )

    return await _run_in_worker(run)


def compile_project(app: Any) -> dict[str, Any]:
    """Compile and save all modules, then report ``IsCompiled``."""
    result: dict[str, Any] = {"ok": True}
    try:
        app.RunCommand(AC_CMD_COMPILE_AND_SAVE_ALL_MODULES)
    except Exception as exc:
        result["message"] = str(exc)
    result["compiled"] = bool(app.IsCompiled)
    if not result["compiled"] and "message" not in result:
        result["message"] = NOT_COMPILED_MESSAGE
    return result


READ_ONLY = ToolAnnotations(readOnlyHint=True, destructiveHint=False, idempotentHint=True)


@mcp.tool(annotations=READ_ONLY.model_copy(update={"openWorldHint": False}))
async def aclib_get_version(database_path: DatabasePath) -> dict[str, Any]:
    """Return the version of the ACLib Import Wizard add-in and its automation API."""
    return await _call_api(database_path, "GetVersion")


@mcp.tool(annotations=READ_ONLY.model_copy(update={"openWorldHint": False}))
async def aclib_get_config(
    database_path: DatabasePath,
    source: SourceParam = "config",
    repo: RepoParam = None,
) -> dict[str, Any]:
    """Return the wizard configuration: repositories and the resulting source.

    The module catalog is at <localRepositoryPath>/_codelib/catalog/catalog.md (or on GitHub
    raw); to search it use aclib_find.
    """
    return await _call_api(database_path, "GetConfig", build_source_options(source, repo))


@mcp.tool(annotations=READ_ONLY.model_copy(update={"openWorldHint": True}))
async def aclib_find(
    database_path: DatabasePath,
    query: Annotated[
        str,
        Field(
            description="English search words and synonyms (any may match, more matching words "
            "rank higher), e.g. 'search filter criteria form'. The catalog is English: "
            "translate the task into several synonyms.",
            min_length=1,
        ),
    ],
    source: SourceParam = "config",
    repo: RepoParam = None,
    limit: Annotated[
        int, Field(description="Maximum number of modules returned.", ge=1, le=25)
    ] = 8,
    overview: Annotated[
        bool,
        Field(description="Also return a compact list of all modules (path and description)."),
    ] = False,
) -> dict[str, Any]:
    """Find ACLib modules in the module catalog (instead of reading the whole catalog).

    The match is literal text, not meaning: pass several English synonyms of the task. If
    nothing matches (or overview=true), a compact list of all modules is returned; choose
    from it by meaning and do not conclude that no suitable module exists.

    Returns per match: path, description, uses (dependencies), replaces, sameObjectName
    (variants that cannot be imported together), tests, examples and public members.
    Show the candidates to the user and let them choose; pass the chosen paths to
    aclib_resolve. Reads only; does not change the database.
    """
    config = await _call_api(database_path, "GetConfig", build_source_options(source, repo))

    def run() -> dict[str, Any]:
        origin = catalog.catalog_origin(config, source, validate_repo(repo, source))
        result = catalog.search_modules(catalog.load_catalog(origin), query, limit, overview)
        return {"ok": True, "catalog": origin, **result}

    try:
        return await anyio.to_thread.run_sync(run)
    except (catalog.CatalogError, ValueError) as exc:
        raise ToolError(str(exc)) from exc


@mcp.tool(annotations=READ_ONLY.model_copy(update={"openWorldHint": False}))
async def aclib_list_installed(database_path: DatabasePath) -> dict[str, Any]:
    """List the ACLib modules installed in the database (from their codelib blocks)."""
    return await _call_api(database_path, "ListInstalled")


@mcp.tool(annotations=READ_ONLY.model_copy(update={"openWorldHint": True}))
async def aclib_resolve(
    database_path: DatabasePath,
    files: Files,
    mode: ModeParam = "missing",
    tests: TestsParam = False,
    examples: ExamplesParam = False,
    source: SourceParam = "config",
    repo: RepoParam = None,
) -> dict[str, Any]:
    """Plan an import without changing the database. Use the skill aclib-import first.

    Take file paths only from aclib_find, the catalog or the user, never guess them.
    Returns the files incl. dependencies, references, tests/examples, the planned action per
    file (import/skip/overwrite/replaced) and missing files. Call before aclib_import and
    show the result to the user.
    """
    file_list = ";".join(validate_files(files))
    options = build_options(mode, tests, examples, source, repo)
    return await _call_api(database_path, "Resolve", file_list, options)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=True
    )
)
async def aclib_import(
    database_path: DatabasePath,
    files: Files,
    mode: ModeParam = "missing",
    tests: TestsParam = False,
    examples: ExamplesParam = False,
    source: SourceParam = "config",
    repo: RepoParam = None,
) -> dict[str, Any]:
    """Import ACLib modules incl. dependencies and references into the database.

    Use the skill aclib-import first. Only after the user approved the plan of aclib_resolve;
    take file paths only from aclib_find, the catalog or the user, never guess them.
    Existing modules may be overwritten (mode=all/selected). Use the same arguments as a
    previous aclib_resolve. Compile the project afterwards.
    """
    file_list = ";".join(validate_files(files))
    options = build_options(mode, tests, examples, source, repo)
    return await _call_api(database_path, "Import", file_list, options)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=False, destructiveHint=False, idempotentHint=True, openWorldHint=False
    )
)
async def aclib_compile(database_path: DatabasePath) -> dict[str, Any]:
    """Compile and save all VBA modules of the database and report whether it compiles.

    Returns compiled=true/false; the compile error itself is not reported (compile in the VBA
    editor to see it). Should Access show a dialog, it is reported as an error (not answered
    automatically) and has to be closed in Access.
    """
    database_path = validate_database_path(database_path)
    worker = get_worker()
    return await _run_in_worker(lambda: worker.execute(database_path, compile_project))


def main() -> None:
    """Entry point of the ``aclib-import-mcp`` command (stdio transport)."""
    mcp.run()


if __name__ == "__main__":
    main()
