# ACLibImportWizard-MCP

MCP server for the automation API of the [ACLib Import Wizard](https://github.com/AccessCodeLib/ACLibImportWizard).
It lets AI agents import and update [AccessCodeLib](https://github.com/AccessCodeLib/AccessCodeLib) modules
in Microsoft Access databases – including `<use>` dependencies, `<ref>` references and the license module –
without the wizard UI.

The server only needs the Import Wizard add-in. No other Access MCP server is required.

> Status: early development (0.1.0, first release).

## Prerequisites

- Windows with Microsoft Access
- ACLib Import Wizard with automation API (version 1.5.0 or later), installed as add-in
  (default: `%APPDATA%\Microsoft\AddIns\ACLibImportWizard.accda`)
- Add-in **and** target database in trusted locations (Access: File > Options > Trust Center)
  – otherwise Access shows a security prompt or disables code
- [uv](https://docs.astral.sh/uv/)

## Installation

### From GitHub (uvx)

```json
"aclib-import": {
  "command": "uvx",
  "args": ["--from", "git+https://github.com/AccessCodeLib/ACLibImportWizard-MCP", "aclib-import-mcp"]
}
```

### Local clone (Claude Code)

```
claude mcp add aclib-import --scope user -- uv run --directory C:/path/to/ACLibImportWizard-MCP aclib-import-mcp
```

Use forward slashes in the path when running the command from a bash shell (backslashes get removed).

## Configuration

| Environment variable | Default | Description |
|---|---|---|
| `ACLIB_ADDIN_PATH` | `%APPDATA%\Microsoft\AddIns\ACLibImportWizard.accda` | Path of the wizard add-in |
| `ACLIB_CALL_TIMEOUT` | `120` | Timeout of an API call in seconds (GitHub downloads can take a while) |
| `ACLIB_DIALOG_TIMEOUT` | `20` | Seconds a dialog of Access may stay open before the call fails (dialogs are never answered automatically) |

Set them in the MCP server configuration, e.g.:

```json
"aclib-import": {
  "command": "uvx",
  "args": ["--from", "git+https://github.com/AccessCodeLib/ACLibImportWizard-MCP", "aclib-import-mcp"],
  "env": { "ACLIB_CALL_TIMEOUT": "300" }
}
```

## Tools

| Tool | Parameters | Writes | Result |
|---|---|---|---|
| `aclib_get_version` | `database_path` | no | add-in and API version |
| `aclib_get_config` | `database_path`, `source?`, `repo?` | no | configured repositories, resulting source |
| `aclib_list_installed` | `database_path` | no | ACLib modules in the database (from their `<codelib>` blocks) |
| `aclib_resolve` | `database_path`, `files[]`, `mode?`, `tests?`, `examples?`, `source?`, `repo?` | no | files incl. dependencies, references, tests/examples, planned action, missing files |
| `aclib_import` | same as `aclib_resolve` | yes (may overwrite) | imported/skipped files, missing dependencies, added/removed references, messages |
| `aclib_compile` | `database_path` | yes | `compiled: true/false` |

- `files`: repository paths, e.g. `["data/SqlTools.cls", "text/StringCollection.cls"]`
- `mode`: `missing` (default), `all` (also overwrite existing dependencies), `selected` (only the given files)
- `source`: `config` (default: local repository if configured, otherwise GitHub), `local`, `github`
- `repo`: GitHub repository `owner/name[@branch]` (default `AccessCodeLib/AccessCodeLib@master`)

The read-only tools can be allowed permanently in the MCP client; `aclib_import` and `aclib_compile`
should require approval.

## Skill

[`skills/aclib-import/SKILL.md`](skills/aclib-import/SKILL.md) describes the workflow for agents
(check version → list installed → resolve → import after approval → compile). For Claude Code, copy the
folder `skills/aclib-import` to `~/.claude/skills/`.

## How it works

- The server calls `Application.Run "<add-in path without extension>.API", <Command>, ...` in the target
  database.
- If the database is open in a running Access instance (e.g. the user's window), that instance is used and
  left open. Otherwise the server starts its own hidden instance and closes it after the call. A second
  Access instance is never opened on a database that is in use by another process.
- All calls run one after another on one COM thread; a named mutex per database prevents parallel calls
  from several server processes.
- Modal dialogs of Access (e.g. save prompts, security notices) are detected and reported after
  `ACLIB_DIALOG_TIMEOUT`; the user has to close them in Access.

## Security

- Only the `API` function of the configured add-in is called – there is no tool for arbitrary VBA code.
- The add-in path comes only from the configuration (`ACLIB_ADDIN_PATH` or the default location), never
  from tool parameters.
- `AutomationSecurity` is not lowered; untrusted locations result in an error instead of a silent bypass.
- Dialogs are only detected, never clicked away.
- `files` must be relative repository paths (no `..`, no absolute paths).

## Troubleshooting

- **"blocked by a dialog" with a save prompt ("Speichern unter" / "Save As"):** the VBA project does not
  compile, so saving the imported modules fails. Close the dialogs in Access, then fix the compile error.
- **"blocked by a dialog" with a security notice:** put add-in and database into trusted locations.
- **"opened by another process":** close the database in the other process.
- **Timeout:** raise `ACLIB_CALL_TIMEOUT` (e.g. for imports from GitHub).

## Development

```
uv sync --extra dev
uv run pytest -m "not integration"
uv run ruff check .
uv run ruff format --check .
```

Tests marked `integration` need Microsoft Access. Test databases go into `testdata\` (git-ignored, must be
a trusted location).

## License

[MIT](LICENSE)
