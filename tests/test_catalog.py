import pytest

from aclib_import_mcp import catalog

CATALOG = {
    "modules": [
        {
            "path": "data/SqlTools.cls",
            "name": "SqlTools",
            "kind": "class",
            "description": "Builds SQL strings",
            "publicMembers": ["Function BuildCriteria() As String"],
        },
        {
            "path": "form/filter/FilterControl.cls",
            "name": "FilterControl",
            "kind": "class",
            "description": "Filter form control with SQL output",
            "uses": ["data/SqlTools.cls"],
            "sameObjectName": ["form/filter/FilterControl_DotNetLib.cls"],
            "publicMembers": [f"Sub Member{i}()" for i in range(12)],
        },
    ]
}


def test_any_term_matches_and_more_matches_rank_higher():
    result = catalog.search_modules(CATALOG, "filter sql", 10)
    assert [m["path"] for m in result["modules"]] == [
        "form/filter/FilterControl.cls",
        "data/SqlTools.cls",
    ]


def test_synonyms_find_module_without_literal_query_word():
    result = catalog.search_modules(CATALOG, "search criteria filter", 10)
    assert result["modules"][0]["path"] == "form/filter/FilterControl.cls"
    assert "overview" not in result


def test_terms_match_word_starts_only():
    assert catalog.search_modules(CATALOG, "prefilter", 10)["total"] == 0
    assert catalog.search_modules(CATALOG, "filt", 10)["total"] == 1


def test_no_match_returns_overview():
    result = catalog.search_modules(CATALOG, "pdf", 10)
    assert result["total"] == 0
    assert result["overview"] == [
        "data/SqlTools.cls - Builds SQL strings",
        "form/filter/FilterControl.cls - Filter form control with SQL output",
    ]
    assert "overview" in result["note"]


def test_overview_on_request():
    result = catalog.search_modules(CATALOG, "sql", 1, want_overview=True)
    assert result["shown"] == 1
    assert len(result["overview"]) == 2
    assert "note" not in result


def test_ranking_prefers_name_and_description_over_members():
    result = catalog.search_modules(CATALOG, "sql", 10)
    assert [m["path"] for m in result["modules"]] == [
        "data/SqlTools.cls",
        "form/filter/FilterControl.cls",
    ]
    assert result["total"] == 2


def test_limit_and_truncation():
    result = catalog.search_modules(CATALOG, "filter", 1)
    assert result["shown"] == 1
    entry = result["modules"][0]
    assert len(entry["publicMembers"]) == catalog.MAX_MEMBERS_SHOWN
    assert entry["publicMembersTruncated"] is True
    assert entry["sameObjectName"] == ["form/filter/FilterControl_DotNetLib.cls"]


def test_empty_query_is_rejected():
    with pytest.raises(ValueError):
        catalog.search_modules(CATALOG, "   ", 5)


def test_origin_prefers_local_repository_for_config_source(tmp_path):
    origin = catalog.catalog_origin({"localRepositoryPath": str(tmp_path)}, "config", None)
    assert origin.endswith("catalog.json")
    assert str(tmp_path) in origin


def test_origin_github_uses_repo_parameter():
    origin = catalog.catalog_origin({"localRepositoryPath": "C:/x"}, "config", "me/lib@dev")
    assert origin == "https://raw.githubusercontent.com/me/lib/dev/_codelib/catalog/catalog.json"


def test_origin_github_default_without_config():
    origin = catalog.catalog_origin({}, "github", None)
    assert "AccessCodeLib/AccessCodeLib/master" in origin


def test_origin_local_without_path_fails():
    with pytest.raises(catalog.CatalogError):
        catalog.catalog_origin({}, "local", None)
