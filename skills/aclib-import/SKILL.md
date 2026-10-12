---
name: aclib-import
description: Import or update AccessCodeLib (ACLib) modules in a Microsoft Access database via the aclib-import MCP server (ACLib Import Wizard automation API). Use when the user wants to include or update an ACLib module (e.g. SqlTools, FilterStringBuilder, filter form, StringCollection), needs a module with a <codelib> block, wants dynamic SQL secured with ACLib data modules, or asks which ACLib modules are installed.
---

# ACLib import via the aclib-import MCP server

Imports modules from the AccessCodeLib (local repository or GitHub) into an Access database,
including `<use>` dependencies, `<ref>` references and the license module, without the wizard UI.
Only the tools of the `aclib-import` MCP server are used.

## Prerequisites
- MCP server `aclib-import` is connected (tools `aclib_*`).
- ACLib Import Wizard with automation API is installed (`aclib_get_version` returns `api >= 1`).
- Add-in **and** target database are in trusted locations. Otherwise a security prompt blocks the
  call or code is disabled.
- `database_path` is the full path of the target `.accdb`/`.mdb`.

## Tools

| Tool | Parameters | Writes |
|---|---|---|
| `aclib_get_version` | `database_path` | no |
| `aclib_get_config` | `database_path`, `source?`, `repo?` | no |
| `aclib_list_installed` | `database_path` | no |
| `aclib_resolve` | `database_path`, `files[]`, `mode?`, `tests?`, `examples?`, `source?`, `repo?` | no |
| `aclib_import` | same as `aclib_resolve` | yes (may overwrite) |
| `aclib_compile` | `database_path` | yes (compiles and saves all modules) |

- `files`: repository paths, e.g. `["data/SqlTools.cls", "text/StringCollection.cls"]`.
- `mode`: `missing` (default: only modules not yet in the database), `all` (also overwrite existing
  dependencies), `selected` (only the given files, no dependencies).
- `tests`, `examples`: also import test/example modules (default `false`).
- `source`: `config` (default: local repository if configured, otherwise GitHub), `local`, `github`.
- `repo`: GitHub repository `owner/name[@branch]` (implies `source=github`, default
  `AccessCodeLib/AccessCodeLib@master`); not combinable with `source=local`.

## Workflow
1. **Check API:** `aclib_get_version`. On error: stop, see Troubleshooting.
2. **Installed modules:** `aclib_list_installed`.
3. **Plan (read only):** `aclib_resolve` with the requested files and options.
   Show the user a table: name, file, action, dependencies, references (status), tests/examples,
   plus `missing` and `executes`. Point out `overwrite`, `replaced`, `missing` and reference changes.
4. **After approval:** `aclib_import` with the same arguments as `aclib_resolve`.
   Report `files` (action per file), `missing`, `references.added/removed`, `executes`, `messages`.
5. **Compile (after approval):** `aclib_compile`. On `compiled: false`: stop and ask the user to compile
   in the VBA editor (Debug > Compile) and send the error line.
6. **Summary:** which modules were imported/overwritten, reference changes, compile result.

## Rules
- Each writing call (`aclib_import`, `aclib_compile`) only after user approval.
- `mode=all` / `mode=selected` and actions `overwrite` only after explicit confirmation.
- `tests=true` imports AccUnit test classes; they need the AccUnit test environment in the database,
  otherwise the project will not compile (the dependency is not listed in `references`). Ask the user
  whether AccUnit is set up before importing tests.
- Never edit ACLib files in the project (keeps them updatable via the wizard). Bugs in ACLib code
  are reported for the AccessCodeLib repository, not fixed in the project.
- `Option Compare Database` in ACLib modules stays as it is.
- Tool error -> stop and report the message; no workarounds, do not retry blindly.
- Missing dependencies (`missing` with `requiredBy`) are skipped by the import -> report them,
  the project will probably not compile.
- Never run `aclib_import` / `aclib_compile` while other tools or the user change the same database.

## Behaviour of the server
- If the database is open in a running Access instance (e.g. the user's window), the server uses that
  instance and leaves it open. Otherwise it starts its own hidden instance and closes it after the call.
- All calls run one after another; calls on the same database never run in parallel.
- Dialogs of Access are never answered automatically. If a dialog stays open longer than
  `ACLIB_DIALOG_TIMEOUT` (default 20 s), the call fails with the dialog title/text and a hint.

## Troubleshooting
- **"blocked by a dialog" with "Speichern unter" / "Save As" (module name):** the VBA project does not
  compile, so saving the imported modules fails and Access asks per module. Ask the user to close the
  dialogs in Access (they appear where Access was last shown, possibly on another monitor), then run
  `aclib_compile` and fix the compile error. Do not call the import again before that.
- **"blocked by a dialog" with a security notice:** add-in or database not in a trusted location.
- **"still running or blocked":** a previous call has not finished (usually an open dialog). Ask the user
  to close it, then retry.
- **"opened by another process":** the database is open in a process the server cannot attach to.
  Ask the user to close it.
- **"in use by another aclib-import-mcp process":** another session works on the database; wait.
- **Timeout (`ACLIB_CALL_TIMEOUT`, default 120 s):** e.g. slow GitHub downloads; ask the user to raise
  the value in the MCP server configuration.
- **API error (`ok:false`):** report number, source and description.

## Result fields (API)
- `aclib_get_version`: `version`, `api`.
- `aclib_get_config`: `source`, `repository`, `localRepositoryPath`, `privateRepositoryPath`,
  `github{owner,repo,branch}`, `hasGitHubToken`, `importTestsDefault`.
- `aclib_list_installed`: `database`, `modules[{name,type,file}]`.
- `aclib_resolve` file entry: `name`, `file`, `requestedPath`, `type` (module/class/form/report/package),
  `exists`, `action`, `description`, `license`, `replaces`, `dependencies`,
  `references[{name,guid,major,minor,status}]` (status `present`/`add`/`update`), `tests`, `examples`,
  `executes`; plus `missing[]`, `executes[]`.
- `aclib_import`: `files[{name,file,type,action}]`, `missing[]`, `references{added,removed}`,
  `executes[]`, `messages[]`.
- `aclib_compile`: `compiled`, optional `message`.
- Actions: `import` (new), `skip` (exists, mode=missing), `overwrite`, `replaced` (superseded via
  `<replace>`), `none` (package).
- `missing` entry: `file`, `requiredBy` (null = selected file), `message`. `aclib_import` stops before
  importing if a selected file is missing.
