# Contributing

## Branches

| Branch | Purpose |
|---|---|
| `main` | Releases only – one merge commit per release, tagged `vX.Y.Z` |
| `dev` | Integration branch for the next release |
| `feature/<name>`, `fix/<name>` | Work branches, created from `dev` |

1. Create a branch from `dev`: `git switch dev && git pull && git switch -c feature/<name>`.
2. Open a pull request into `dev`. CI (ruff, pytest) must be green.
3. Merge it with **Squash and merge** – one commit per feature or fix on `dev`.
4. Do not commit directly to `main` or `dev`.

## Release

1. Update `CHANGELOG.md` (move `[Unreleased]` to `[X.Y.Z] - YYYY-MM-DD`) and `__version__` in
   `src/aclib_import_mcp/__init__.py` on `dev` (via a `release/X.Y.Z` branch and pull request).
2. Pull request `dev` -> `main`, merged with **Create a merge commit** (title `Release X.Y.Z`).
   No squash, so `main` and `dev` keep a shared history and no merge back is needed.
3. Tag the merge commit on `main` and push only this tag:
   ```
   git switch main && git pull
   git tag -a vX.Y.Z -m "Release X.Y.Z"
   git push origin vX.Y.Z
   ```

`git log --first-parent main` lists the releases only.

## Development

See [README – Development](README.md#development) for setup, tests and linting.

- Code, comments and commit messages in English; commit subject in imperative mood, max. ~72 characters.
- Text files use CRLF line endings and UTF-8 without BOM.
- Tests that need Microsoft Access are marked `integration`; use copies of databases in `testdata\`
  (git-ignored, trusted location), never your real databases.
