# Design — Multi-Site Search

## Approach

Keep app-only auth. Fan out `POST /sites/{site_id}/search` across a resolved list of sites and merge results. Site discovery via `GET /sites?search=*` (works with application permission `Sites.Read.All`; the tenant-wide `POST /search/query` endpoint requires delegated auth and is not used).

## Components to Change

| File | Change |
|------|--------|
| `config/settings.py` | New `search_sites` key in `SHAREPOINT_CONFIG` parsed from optional `SEARCH_SITES` env var (comma-separated URLs/IDs). |
| `.env.example` | Document `SEARCH_SITES`. |
| `utils/_graph_site_ops.py` | New `GraphClient` methods: `list_sites(query)` → `GET sites?search={query}`; `search_site(site_id, query, entity_types)` → `POST sites/{id}/search`. |
| `tools/read_tools.py` | New `list_sites` tool; refactor `search_sharepoint` to multi-site. |
| `tests/test_graph_client.py` | Unit tests for the new GraphClient methods. |
| `README.md` | Document new tool and behavior. |

## search_sharepoint Flow

1. Resolve scope: `sites` param → `SEARCH_SITES` env → `list_sites()` capped at `max_sites`.
2. Resolve each entry to a `site_id`: entries that look like URLs are parsed (domain/site_name) and resolved via `get_site_info`; anything else is treated as an already-valid site ID.
3. `asyncio.gather` over per-site `search_site()` calls, `return_exceptions=True`; per-site errors collected, not raised.
4. Merge hits, tagging each with its source `site`. Iterate over ALL `value[]` entries (fixes `IndexError` on empty responses).
5. Return `{"results", "errors", "sites_searched", "truncated"}`.

## Impact Analysis

- Behavior change: `search_sharepoint` without `sites` previously searched the single `SITE_URL` site; now it searches `SEARCH_SITES` or the whole tenant. Documented in the tool docstring and README.
- No changes to auth, permissions, or other tools.
