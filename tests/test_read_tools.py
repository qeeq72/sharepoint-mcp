"""Tests for the read-only MCP tools (list_sites, search_sharepoint).

Tools are exercised end-to-end: extracted from a FastMCP registry and run
against a GraphClient whose HTTP layer is backed by httpx.MockTransport.
"""

import json
from datetime import datetime, timedelta
from unittest.mock import MagicMock

import httpx
import pytest
from mcp.server.fastmcp import FastMCP

from auth.sharepoint_auth import SharePointContext
from config import settings
from config.settings import SHAREPOINT_CONFIG
from tools.read_tools import register_read_tools

GRAPH = "https://graph.microsoft.com/v1.0"

HR_SITE = {
    "id": "contoso.sharepoint.com,hr1,hr2",
    "name": "hr",
    "displayName": "HR",
    "webUrl": "https://contoso.sharepoint.com/sites/hr",
}
LEGAL_SITE = {
    "id": "contoso.sharepoint.com,lg1,lg2",
    "name": "legal",
    "displayName": "Legal",
    "webUrl": "https://contoso.sharepoint.com/sites/legal",
}
DOCS_SITE = {
    "id": "contoso.sharepoint.com,dc1,dc2",
    "name": "docs",
    "displayName": "Docs",
    "webUrl": "https://contoso.sharepoint.com/sites/docs",
}


def search_hit(name, more_results=False):
    """Build a Graph search response with a single driveItem hit."""
    return {
        "value": [
            {
                "hitsContainers": [
                    {
                        "moreResultsAvailable": more_results,
                        "hits": [
                            {
                                "summary": f"summary of {name}",
                                "resource": {
                                    "name": name,
                                    "webUrl": f"https://contoso.sharepoint.com/{name}",
                                    "@odata.type": "#microsoft.graph.driveItem",
                                },
                            }
                        ],
                    }
                ]
            }
        ]
    }


@pytest.fixture(autouse=True)
def _pin_env_dependent_settings(monkeypatch):
    """Isolate tests from the developer's local .env values."""
    monkeypatch.setattr(settings, "ALLOWED_SITES", [])
    monkeypatch.setitem(SHAREPOINT_CONFIG, "search_sites", [])
    monkeypatch.setitem(SHAREPOINT_CONFIG, "search_region", "")
    monkeypatch.setitem(
        SHAREPOINT_CONFIG, "site_url", "https://contoso.sharepoint.com/sites/hr"
    )


@pytest.fixture
def tool_fns():
    """Register read tools on a fresh FastMCP and return them by name."""
    mcp = FastMCP("test")
    register_read_tools(mcp)
    return {
        name: mcp._tool_manager.get_tool(name).fn
        for name in ("list_sites", "search_sharepoint", "list_document_libraries")
    }


@pytest.fixture
def fake_ctx():
    """Build a fake MCP Context carrying a valid SharePointContext."""
    sp_ctx = SharePointContext(
        access_token="test_token", token_expiry=datetime.now() + timedelta(hours=1)
    )
    ctx = MagicMock()
    ctx.request_context.lifespan_context = sp_ctx
    return ctx


@pytest.fixture
def graph(monkeypatch):
    """Route Graph API calls to canned responses via httpx.MockTransport.

    Returns a routing dict the test can mutate:
      sites: list returned by GET /sites?search=...
      search: {site_web_url: httpx.Response | dict} for POST /search/query
        (the site is recovered from the KQL path filter in the request body)
      requests: every httpx.Request made
    """
    routing = {"sites": [], "search": {}, "requests": []}

    def handler(request):
        routing["requests"].append(request)
        path = request.url.path
        if request.method == "GET" and path == "/v1.0/sites":
            return httpx.Response(200, json={"value": routing["sites"]})
        if request.method == "GET" and path.endswith("/lists"):
            return httpx.Response(200, json={"value": []})
        if request.method == "GET" and path.endswith("/drives"):
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "drive-1",
                            "name": "Documents",
                            "driveType": "documentLibrary",
                            "webUrl": "https://contoso.sharepoint.com/sites/x/Docs",
                        }
                    ]
                },
            )
        if request.method == "GET" and path.startswith("/v1.0/sites/"):
            # Site resolution by URL (.../sites/{domain}:/sites/{name})
            # or by ID (.../sites/{site_id})
            tail = path.rsplit("/", 1)[-1]
            for site in (HR_SITE, LEGAL_SITE, DOCS_SITE):
                if site["name"] == tail or path.endswith(site["id"]):
                    return httpx.Response(200, json=site)
            return httpx.Response(404, text="site not found")
        if request.method == "POST" and path == "/v1.0/search/query":
            query_string = json.loads(request.content)["requests"][0]["query"][
                "queryString"
            ]
            for web_url, outcome in routing["search"].items():
                if f'path:"{web_url}"' in query_string:
                    if isinstance(outcome, httpx.Response):
                        return outcome
                    return httpx.Response(200, json=outcome)
            return httpx.Response(404, text=f"no search stub for: {query_string}")
        return httpx.Response(404, text=f"unstubbed: {request.method} {path}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    monkeypatch.setattr("utils._graph_http.get_http_client", lambda: client)
    return routing


async def test_list_sites_tool(tool_fns, fake_ctx, graph):
    """Test the list_sites tool returns formatted site entries."""
    graph["sites"] = [HR_SITE, LEGAL_SITE]

    result = json.loads(await tool_fns["list_sites"](fake_ctx))
    assert [s["name"] for s in result] == ["hr", "legal"]
    assert result[0] == {
        "id": HR_SITE["id"],
        "name": "hr",
        "display_name": "HR",
        "web_url": HR_SITE["webUrl"],
    }
    assert str(graph["requests"][0].url) == f"{GRAPH}/sites?search=%2A"


async def test_list_sites_tool_passes_query(tool_fns, fake_ctx, graph):
    """Test the list_sites tool forwards the name filter."""
    graph["sites"] = [HR_SITE]
    await tool_fns["list_sites"](fake_ctx, query="hr")
    assert str(graph["requests"][0].url) == f"{GRAPH}/sites?search=hr"


async def test_search_with_explicit_site_urls(tool_fns, fake_ctx, graph):
    """Test that site URLs are resolved and results tagged per site."""
    graph["search"] = {
        HR_SITE["webUrl"]: search_hit("vacation.docx"),
        LEGAL_SITE["webUrl"]: search_hit("contract.docx"),
    }

    result = json.loads(
        await tool_fns["search_sharepoint"](
            fake_ctx,
            query="test",
            sites=[HR_SITE["webUrl"], LEGAL_SITE["webUrl"]],
        )
    )
    assert result["sites_searched"] == 2
    assert result["errors"] == []
    assert result["truncated"] is False
    by_site = {hit["site"]: hit["title"] for hit in result["results"]}
    assert by_site == {
        HR_SITE["webUrl"]: "vacation.docx",
        LEGAL_SITE["webUrl"]: "contract.docx",
    }


async def test_search_with_site_ids(tool_fns, fake_ctx, graph):
    """Test that site IDs are resolved to their web URL for the path filter."""
    graph["search"] = {HR_SITE["webUrl"]: search_hit("doc.docx")}

    result = json.loads(
        await tool_fns["search_sharepoint"](fake_ctx, query="q", sites=[HR_SITE["id"]])
    )
    assert result["sites_searched"] == 1
    assert len(result["results"]) == 1
    assert result["results"][0]["site"] == HR_SITE["webUrl"]
    # The ID was resolved via GET /sites/{id} before searching
    assert [r.method for r in graph["requests"]] == ["GET", "POST"]


async def test_search_scope_falls_back_to_env(tool_fns, fake_ctx, graph, monkeypatch):
    """Test that SEARCH_SITES config defines the scope when sites is empty."""
    monkeypatch.setitem(SHAREPOINT_CONFIG, "search_sites", [HR_SITE["id"]])
    graph["search"] = {HR_SITE["webUrl"]: search_hit("doc.docx")}

    result = json.loads(await tool_fns["search_sharepoint"](fake_ctx, query="q"))
    assert result["sites_searched"] == 1
    # Scope came from config: no tenant-wide site listing was requested
    assert all(r.url.path != "/v1.0/sites" for r in graph["requests"])


async def test_search_scope_defaults_to_all_sites(
    tool_fns, fake_ctx, graph, monkeypatch
):
    """Test that with no sites and no env scope, all tenant sites are searched."""
    monkeypatch.setitem(SHAREPOINT_CONFIG, "search_sites", [])
    graph["sites"] = [HR_SITE, LEGAL_SITE]
    graph["search"] = {
        HR_SITE["webUrl"]: search_hit("a.docx"),
        LEGAL_SITE["webUrl"]: search_hit("b.docx"),
    }

    result = json.loads(await tool_fns["search_sharepoint"](fake_ctx, query="q"))
    assert result["sites_searched"] == 2
    assert result["truncated"] is False
    assert len(result["results"]) == 2


async def test_search_truncates_at_max_sites(tool_fns, fake_ctx, graph, monkeypatch):
    """Test that tenant-wide scope is capped at max_sites and flagged."""
    monkeypatch.setitem(SHAREPOINT_CONFIG, "search_sites", [])
    graph["sites"] = [HR_SITE, LEGAL_SITE, DOCS_SITE]
    graph["search"] = {
        HR_SITE["webUrl"]: search_hit("a.docx"),
        LEGAL_SITE["webUrl"]: search_hit("b.docx"),
    }

    result = json.loads(
        await tool_fns["search_sharepoint"](fake_ctx, query="q", max_sites=2)
    )
    assert result["sites_searched"] == 2
    assert result["truncated"] is True
    # The third site was never searched
    search_bodies = [
        r.content.decode() for r in graph["requests"] if r.method == "POST"
    ]
    assert all(DOCS_SITE["webUrl"] not in body for body in search_bodies)


async def test_search_tolerates_per_site_errors(tool_fns, fake_ctx, graph):
    """Test that one failing site lands in errors without failing the search."""
    graph["search"] = {
        HR_SITE["webUrl"]: search_hit("doc.docx"),
        LEGAL_SITE["webUrl"]: httpx.Response(403, text="Access denied"),
    }

    result = json.loads(
        await tool_fns["search_sharepoint"](
            fake_ctx, query="q", sites=[HR_SITE["id"], LEGAL_SITE["id"]]
        )
    )
    assert result["sites_searched"] == 2
    assert len(result["results"]) == 1
    assert result["results"][0]["site"] == HR_SITE["webUrl"]
    assert len(result["errors"]) == 1
    assert result["errors"][0]["site"] == LEGAL_SITE["webUrl"]
    assert "403" in result["errors"][0]["error"]


async def test_search_reports_more_results(tool_fns, fake_ctx, graph):
    """Test that sites with more matches than size are flagged in the response."""
    graph["search"] = {
        HR_SITE["webUrl"]: search_hit("a.docx", more_results=True),
        LEGAL_SITE["webUrl"]: search_hit("b.docx"),
    }

    result = json.loads(
        await tool_fns["search_sharepoint"](
            fake_ctx, query="q", sites=[HR_SITE["id"], LEGAL_SITE["id"]]
        )
    )
    assert result["more_results_available_on"] == [HR_SITE["webUrl"]]


async def test_search_passes_size(tool_fns, fake_ctx, graph):
    """Test that the size argument reaches the Graph search request body."""
    graph["search"] = {HR_SITE["webUrl"]: search_hit("a.docx")}

    await tool_fns["search_sharepoint"](
        fake_ctx, query="q", sites=[HR_SITE["webUrl"]], size=5
    )
    sent_body = json.loads(graph["requests"][-1].content)
    assert sent_body["requests"][0]["size"] == 5


async def test_list_sites_filtered_by_allowlist(tool_fns, fake_ctx, graph, monkeypatch):
    """Test that list_sites hides sites outside MCP_ALLOWED_SITES."""
    monkeypatch.setattr(settings, "ALLOWED_SITES", [HR_SITE["webUrl"]])
    graph["sites"] = [HR_SITE, LEGAL_SITE]

    result = json.loads(await tool_fns["list_sites"](fake_ctx))
    assert [s["name"] for s in result] == ["hr"]


async def test_search_rejects_disallowed_site(tool_fns, fake_ctx, graph, monkeypatch):
    """Test that explicit sites outside the allowlist land in errors."""
    monkeypatch.setattr(settings, "ALLOWED_SITES", [HR_SITE["webUrl"]])
    graph["search"] = {HR_SITE["webUrl"]: search_hit("doc.docx")}

    result = json.loads(
        await tool_fns["search_sharepoint"](
            fake_ctx, query="q", sites=[HR_SITE["webUrl"], LEGAL_SITE["webUrl"]]
        )
    )
    assert result["sites_searched"] == 1
    assert len(result["results"]) == 1
    assert result["errors"] == [
        {"site": LEGAL_SITE["webUrl"], "error": "Not allowed by MCP_ALLOWED_SITES"}
    ]


async def test_search_default_scope_is_allowlist(
    tool_fns, fake_ctx, graph, monkeypatch
):
    """Test that with no sites and no SEARCH_SITES the allowlist is the scope."""
    monkeypatch.setitem(SHAREPOINT_CONFIG, "search_sites", [])
    monkeypatch.setattr(settings, "ALLOWED_SITES", [HR_SITE["webUrl"]])
    graph["search"] = {HR_SITE["webUrl"]: search_hit("doc.docx")}

    result = json.loads(await tool_fns["search_sharepoint"](fake_ctx, query="q"))
    assert result["sites_searched"] == 1
    # Tenant-wide listing was never requested
    assert all(r.url.path != "/v1.0/sites" for r in graph["requests"])


async def test_per_site_tool_rejects_disallowed_site(fake_ctx, graph, monkeypatch):
    """Test that a per-site tool refuses a site outside the allowlist."""
    import pytest as _pytest

    from mcp.server.fastmcp import FastMCP
    from tools.read_tools import register_read_tools as _register

    monkeypatch.setattr(settings, "ALLOWED_SITES", [HR_SITE["webUrl"]])
    mcp = FastMCP("test-allow")
    _register(mcp)
    get_lists = mcp._tool_manager.get_tool("get_lists").fn

    # Disallowed: resolves LEGAL and refuses
    with _pytest.raises(Exception, match="not allowed by MCP_ALLOWED_SITES"):
        await get_lists(fake_ctx, site_id=LEGAL_SITE["id"])

    # Allowed: resolves HR and proceeds to the lists request
    result = json.loads(await get_lists(fake_ctx, site_id=HR_SITE["id"]))
    assert result == []


async def test_allowlist_accepts_bare_site_names(
    tool_fns, fake_ctx, graph, monkeypatch
):
    """Test that allowlist entries can be bare site names and sites resolve."""
    monkeypatch.setattr(settings, "ALLOWED_SITES", ["hr"])
    graph["sites"] = [HR_SITE, LEGAL_SITE]
    graph["search"] = {HR_SITE["webUrl"]: search_hit("doc.docx")}

    # list_sites keeps only the named site
    listed = json.loads(await tool_fns["list_sites"](fake_ctx))
    assert [s["name"] for s in listed] == ["hr"]

    # default scope resolves the bare name via the tenant domain
    result = json.loads(await tool_fns["search_sharepoint"](fake_ctx, query="q"))
    assert result["sites_searched"] == 1
    assert result["errors"] == []
    assert result["results"][0]["site"] == HR_SITE["webUrl"]


async def test_search_tolerates_unresolvable_site(tool_fns, fake_ctx, graph):
    """Test that a site that fails to resolve lands in errors, not a crash."""
    graph["search"] = {HR_SITE["webUrl"]: search_hit("doc.docx")}

    result = json.loads(
        await tool_fns["search_sharepoint"](
            fake_ctx,
            query="q",
            sites=[
                HR_SITE["webUrl"],
                "https://wrong-tenant.sharepoint.com/sites/nope",
            ],
        )
    )
    assert result["sites_searched"] == 1
    assert len(result["results"]) == 1
    assert len(result["errors"]) == 1
    assert "wrong-tenant" in result["errors"][0]["site"]


async def test_list_document_libraries_accepts_site(tool_fns, fake_ctx, graph):
    """Test that list_document_libraries takes any site, not only SITE_URL."""
    result = json.loads(
        await tool_fns["list_document_libraries"](fake_ctx, site=LEGAL_SITE["webUrl"])
    )
    assert result[0]["id"] == "drive-1"
    assert result[0]["name"] == "Documents"
    # The drives request was made against the resolved LEGAL site
    drives_calls = [r for r in graph["requests"] if str(r.url).endswith("/drives")]
    assert LEGAL_SITE["id"] in str(drives_calls[0].url)


async def test_list_document_libraries_defaults_to_site_url(tool_fns, fake_ctx, graph):
    """Test that without a site argument the configured SITE_URL is used."""
    result = json.loads(await tool_fns["list_document_libraries"](fake_ctx))
    assert result[0]["id"] == "drive-1"
    # SITE_URL is pinned to the HR site in the autouse fixture
    drives_calls = [r for r in graph["requests"] if str(r.url).endswith("/drives")]
    assert HR_SITE["id"] in str(drives_calls[0].url)


async def test_list_document_libraries_respects_allowlist(
    tool_fns, fake_ctx, graph, monkeypatch
):
    """Test that a site outside the allowlist is refused."""
    monkeypatch.setattr(settings, "ALLOWED_SITES", [HR_SITE["webUrl"]])
    with pytest.raises(Exception, match="not allowed by MCP_ALLOWED_SITES"):
        await tool_fns["list_document_libraries"](fake_ctx, site=LEGAL_SITE["webUrl"])


async def test_search_handles_empty_results(tool_fns, fake_ctx, graph):
    """Test that an empty search response yields empty results, not a crash."""
    graph["search"] = {HR_SITE["webUrl"]: {"value": []}}

    result = json.loads(
        await tool_fns["search_sharepoint"](fake_ctx, query="q", sites=[HR_SITE["id"]])
    )
    assert result["results"] == []
    assert result["errors"] == []
