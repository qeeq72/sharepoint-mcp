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


def search_hit(name):
    """Build a Graph search response with a single driveItem hit."""
    return {
        "value": [
            {
                "hitsContainers": [
                    {
                        "hits": [
                            {
                                "summary": f"summary of {name}",
                                "resource": {
                                    "name": name,
                                    "webUrl": f"https://contoso.sharepoint.com/{name}",
                                    "@odata.type": "#microsoft.graph.driveItem",
                                },
                            }
                        ]
                    }
                ]
            }
        ]
    }


@pytest.fixture
def tool_fns():
    """Register read tools on a fresh FastMCP and return them by name."""
    mcp = FastMCP("test")
    register_read_tools(mcp)
    return {
        name: mcp._tool_manager.get_tool(name).fn
        for name in ("list_sites", "search_sharepoint")
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
      search: {site_id: httpx.Response | dict} for POST /sites/{id}/search
      requests: every httpx.Request made
    """
    routing = {"sites": [], "search": {}, "requests": []}

    def handler(request):
        routing["requests"].append(request)
        path = request.url.path
        if request.method == "GET" and path == "/v1.0/sites":
            return httpx.Response(200, json={"value": routing["sites"]})
        if request.method == "GET" and path.startswith("/v1.0/sites/"):
            # Site resolution by URL: /v1.0/sites/{domain}:/sites/{name}
            site_name = path.rsplit("/", 1)[-1]
            for site in (HR_SITE, LEGAL_SITE, DOCS_SITE):
                if site["name"] == site_name:
                    return httpx.Response(200, json=site)
            return httpx.Response(404, text="site not found")
        if request.method == "POST" and path.endswith("/search"):
            site_id = path.removeprefix("/v1.0/sites/").removesuffix("/search")
            outcome = routing["search"].get(site_id)
            if isinstance(outcome, httpx.Response):
                return outcome
            if outcome is not None:
                return httpx.Response(200, json=outcome)
            return httpx.Response(404, text=f"no search stub for {site_id}")
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
    """Test that site URLs are resolved to IDs and results tagged per site."""
    graph["search"] = {
        HR_SITE["id"]: search_hit("vacation.docx"),
        LEGAL_SITE["id"]: search_hit("contract.docx"),
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


async def test_search_with_site_ids_skips_resolution(tool_fns, fake_ctx, graph):
    """Test that site IDs are used as-is, without a get_site_info round-trip."""
    graph["search"] = {HR_SITE["id"]: search_hit("doc.docx")}

    result = json.loads(
        await tool_fns["search_sharepoint"](fake_ctx, query="q", sites=[HR_SITE["id"]])
    )
    assert result["sites_searched"] == 1
    assert len(result["results"]) == 1
    # Only the search POST was made — no site-info GET
    assert [r.method for r in graph["requests"]] == ["POST"]


async def test_search_scope_falls_back_to_env(tool_fns, fake_ctx, graph, monkeypatch):
    """Test that SEARCH_SITES config defines the scope when sites is empty."""
    monkeypatch.setitem(SHAREPOINT_CONFIG, "search_sites", [HR_SITE["id"]])
    graph["search"] = {HR_SITE["id"]: search_hit("doc.docx")}

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
        HR_SITE["id"]: search_hit("a.docx"),
        LEGAL_SITE["id"]: search_hit("b.docx"),
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
        HR_SITE["id"]: search_hit("a.docx"),
        LEGAL_SITE["id"]: search_hit("b.docx"),
    }

    result = json.loads(
        await tool_fns["search_sharepoint"](fake_ctx, query="q", max_sites=2)
    )
    assert result["sites_searched"] == 2
    assert result["truncated"] is True
    # The third site was never searched
    assert all(DOCS_SITE["id"] not in r.url.path for r in graph["requests"])


async def test_search_tolerates_per_site_errors(tool_fns, fake_ctx, graph):
    """Test that one failing site lands in errors without failing the search."""
    graph["search"] = {
        HR_SITE["id"]: search_hit("doc.docx"),
        LEGAL_SITE["id"]: httpx.Response(403, text="Access denied"),
    }

    result = json.loads(
        await tool_fns["search_sharepoint"](
            fake_ctx, query="q", sites=[HR_SITE["id"], LEGAL_SITE["id"]]
        )
    )
    assert result["sites_searched"] == 2
    assert len(result["results"]) == 1
    assert result["results"][0]["site"] == HR_SITE["id"]
    assert len(result["errors"]) == 1
    assert result["errors"][0]["site"] == LEGAL_SITE["id"]
    assert "403" in result["errors"][0]["error"]


async def test_search_handles_empty_results(tool_fns, fake_ctx, graph):
    """Test that an empty search response yields empty results, not a crash."""
    graph["search"] = {HR_SITE["id"]: {"value": []}}

    result = json.loads(
        await tool_fns["search_sharepoint"](fake_ctx, query="q", sites=[HR_SITE["id"]])
    )
    assert result["results"] == []
    assert result["errors"] == []
