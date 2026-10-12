# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added
- Skill `aclib-import`: find suitable modules for a task in the ACLib catalog
  (`_codelib/catalog/catalog.md` / `catalog.json`, local repository or GitHub) before resolving the import.

## [0.1.0] - 2026-10-12

### Added
- MCP tools `aclib_get_version`, `aclib_get_config`, `aclib_list_installed`, `aclib_resolve`,
  `aclib_import` and `aclib_compile` for the ACLib Import Wizard automation API.
- Reuse of a running Access instance with the target database; otherwise an own hidden instance that is
  closed after each call.
- Serialized calls on one COM thread and a named mutex per database (no parallel changes).
- Detection of modal Access dialogs, reported after `ACLIB_DIALOG_TIMEOUT` (never answered automatically).
- Configuration via `ACLIB_ADDIN_PATH`, `ACLIB_CALL_TIMEOUT` and `ACLIB_DIALOG_TIMEOUT`.
- Skill `skills/aclib-import` for agents.
- Requires ACLib Import Wizard 1.5.0 or later (automation API 1).
