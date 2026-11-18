"""Search in the generated ACLib module catalog (``_codelib/catalog/catalog.json``).

The catalog is read from the local repository (path from the wizard configuration) or from
GitHub raw. Only a fixed host and validated repository names are used.
"""

from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

CATALOG_RELATIVE_PATH = "_codelib/catalog/catalog.json"
GITHUB_RAW_URL = (
    "https://raw.githubusercontent.com/{owner}/{repo}/{branch}/" + CATALOG_RELATIVE_PATH
)
DEFAULT_GITHUB = {"owner": "AccessCodeLib", "repo": "AccessCodeLib", "branch": "master"}
HTTP_TIMEOUT_SECONDS = 30
MAX_MEMBERS_SHOWN = 8
_NAME_WEIGHT = 5
_DESCRIPTION_WEIGHT = 3
_MEMBER_WEIGHT = 1


class CatalogError(Exception):
    """The catalog could not be loaded or parsed."""


def github_location(config: dict[str, Any], repo: str | None) -> dict[str, str]:
    """Return ``{owner, repo, branch}``: from ``repo`` ('owner/name[@branch]') or the config."""
    if repo:
        name, _, branch = repo.partition("@")
        owner, _, repo_name = name.partition("/")
        return {"owner": owner, "repo": repo_name, "branch": branch or "master"}
    github = config.get("github")
    if isinstance(github, dict) and all(github.get(k) for k in ("owner", "repo", "branch")):
        return {k: str(github[k]) for k in ("owner", "repo", "branch")}
    return dict(DEFAULT_GITHUB)


def catalog_origin(config: dict[str, Any], source: str, repo: str | None) -> str:
    """Return the local catalog file path or the GitHub raw URL for the effective source."""
    local = str(config.get("localRepositoryPath") or "").strip()
    use_local = not repo and (source == "local" or (source == "config" and local))
    if use_local:
        if not local:
            raise CatalogError("No local repository configured in the ACLib Import Wizard.")
        return str(Path(local) / CATALOG_RELATIVE_PATH)
    location = github_location(config, repo)
    return GITHUB_RAW_URL.format(**location)


def load_catalog(origin: str) -> dict[str, Any]:
    """Load and parse the catalog from a file path or an https URL."""
    try:
        if origin.startswith("https://raw.githubusercontent.com/"):
            with urllib.request.urlopen(origin, timeout=HTTP_TIMEOUT_SECONDS) as response:  # noqa: S310
                raw = response.read()
        else:
            raw = Path(origin).read_bytes()
        data = json.loads(raw.decode("utf-8-sig"))
    except urllib.error.HTTPError as exc:
        raise CatalogError(f"Catalog not found ({exc.code}): {origin}") from None
    except (OSError, urllib.error.URLError, ValueError) as exc:
        raise CatalogError(f"Catalog could not be read ({exc}): {origin}") from None
    if not isinstance(data, dict) or not isinstance(data.get("modules"), list):
        raise CatalogError(f"Catalog has an unexpected format: {origin}")
    return data


def _terms(query: str) -> list[str]:
    seen: dict[str, None] = {}
    for term in re.split(r"[\s,;]+", query.lower().strip()):
        if term:
            seen[term] = None
    return list(seen)


def _hit(term: str, text: str) -> bool:
    """True if ``term`` starts a word in ``text`` (so 'filter' finds 'filters', not 'prefilter')."""
    return re.search(r"(?<![a-z0-9])" + re.escape(term), text) is not None


def _score(module: dict[str, Any], terms: list[str]) -> tuple[int, int] | None:
    """Return ``(matched terms, weighted score)`` if at least one term matches, else ``None``."""
    name = f"{module.get('path') or ''} {module.get('name') or ''}".lower()
    description = str(module.get("description") or "").lower()
    members = " ".join(str(m) for m in module.get("publicMembers") or []).lower()
    matched = score = 0
    for term in terms:
        points = (
            _NAME_WEIGHT * _hit(term, name)
            + _DESCRIPTION_WEIGHT * _hit(term, description)
            + _MEMBER_WEIGHT * _hit(term, members)
        )
        if points:
            matched += 1
            score += points
    return (matched, score) if matched else None


def _entry(module: dict[str, Any], terms: list[str]) -> dict[str, Any]:
    members = [str(m) for m in module.get("publicMembers") or []]
    matching = [m for m in members if any(_hit(t, m.lower()) for t in terms)]
    shown = (matching or members)[:MAX_MEMBERS_SHOWN]
    entry: dict[str, Any] = {
        "path": module.get("path"),
        "kind": module.get("kind"),
        "description": module.get("description"),
        "descriptionSource": module.get("descriptionSource"),
    }
    for key in ("predeclaredId",):
        if module.get(key):
            entry[key] = True
    for key in ("uses", "replaces", "sameObjectName", "tests", "examples"):
        if module.get(key):
            entry[key] = module[key]
    entry["publicMembers"] = shown
    if len(matching or members) > len(shown):
        entry["publicMembersTruncated"] = True
    return entry


def overview(catalog: dict[str, Any]) -> list[str]:
    """Return one compact ``path - description`` line per module."""
    return [
        f"{m.get('path')} - {m.get('description') or ''}".rstrip(" -")
        for m in catalog["modules"]
        if isinstance(m, dict)
    ]


def search_modules(
    catalog: dict[str, Any], query: str, limit: int, want_overview: bool = False
) -> dict[str, Any]:
    """Return the best matches for ``query`` (OR over the words, ranked by matched words).

    Without any match, or with ``want_overview``, a compact list of all modules is returned
    instead, so that the caller can choose by meaning.
    """
    terms = _terms(query)
    if not terms:
        raise ValueError("query must contain at least one word")
    scored = []
    for module in catalog["modules"]:
        if isinstance(module, dict):
            score = _score(module, terms)
            if score is not None:
                scored.append((score, str(module.get("path")), module))
    scored.sort(key=lambda item: (-item[0][0], -item[0][1], item[1]))
    result: dict[str, Any] = {
        "total": len(scored),
        "shown": min(limit, len(scored)),
        "modules": [_entry(m, terms) for _, _, m in scored[:limit]],
    }
    if want_overview or not scored:
        result["overview"] = overview(catalog)
        if not scored:
            result["note"] = "No module matched; choose from the overview by meaning."
    return result
