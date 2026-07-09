# Requirements — Multi-Site Search

## Feature Description

Extend the MCP server so an LLM can search content across multiple SharePoint sites in one tool call, and discover which sites exist in the tenant.

Current limitation: `search_sharepoint` searches only the single site configured via the `SITE_URL` environment variable.

## User Stories

1. As an LLM client, I want to list all sites in the tenant (optionally filtered by name) so I can decide where to search.
2. As an LLM client, I want to pass a list of sites (URLs or site IDs) to `search_sharepoint` and get merged results labeled with their source site.
3. As an operator, I want an optional `SEARCH_SITES` environment variable that restricts the default search scope when the tool is called without an explicit site list.

## Acceptance Criteria

- New `list_sites` tool returns `{id, name, displayName, webUrl}` for tenant sites via `GET /sites?search=...`.
- `search_sharepoint` accepts `sites: list[str]` (URLs or site IDs) and `max_sites: int`.
- Scope resolution priority: tool param `sites` → env `SEARCH_SITES` → all tenant sites (capped at `max_sites`, response flags truncation).
- A failure on one site (e.g. 403) does not fail the whole search; errors are reported per-site in the response.
- Each result hit carries a `site` field identifying its source.
- Empty search response (`value: []`) does not raise (fixes existing `IndexError`).
- All quality checks pass: `black`, `ruff`, `pytest`.

## Constraints

- Auth stays app-only (client credentials). The delegated-only `POST /search/query` endpoint is out of scope; multi-site search is implemented as fan-out over `POST /sites/{id}/search`.
- Application permission `Sites.Read.All` (already granted) is sufficient; no permission changes.
