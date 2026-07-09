# Design — HTTP Layer Hardening

## Approach

Migrate only the JSON verb methods (`get`/`post`/`patch`/`delete`) of `_GraphHttpMixin` to a shared `httpx.AsyncClient`, funneled through a single `_request()` coroutine that owns URL building, timeouts, and the retry loop. Upload/download and auth paths deliberately stay on `requests` (httpx does not follow redirects by default; Graph `/content` responds 302 — unverifiable without a live tenant).

## Components

| File | Change |
|------|--------|
| `utils/_graph_http.py` | Lazy module-level `httpx.AsyncClient` singleton (`get_http_client`/`close_http_client`, `follow_redirects=True`); `_parse_retry_after`; central `_request()`; `get_paged()`; verbs become thin wrappers. `upload_file`/`_upload_in_chunks` untouched. |
| `utils/graph_client.py` | `__init__(context, http_client=None)` — test injection seam. |
| `utils/_graph_site_ops.py` | `list_sites` uses `get_paged`. |
| `server.py` | `await close_http_client()` in lifespan `finally`. |
| `auth/sharepoint_auth.py` | `timeout=15` on the 5 existing `requests` calls; no migration. |
| `pytest.ini` | `pythonpath = .` |
| `requirements.txt` | `httpx>=0.27.0` added explicitly; `requests` stays. |
| `tests/test_graph_client.py` | 14 tests converted to `httpx.MockTransport` via `make_client()` helper; upload/download tests unchanged; 9 new retry/pagination tests. |

## Retry Policy

- Retryable: 429 for all methods; 502/503/504 for GET only.
- Delay: numeric `Retry-After` (capped 60s), else `min(2**attempt, 60)`.
- `MAX_RETRIES = 3` (4 attempts total). Fresh `context.headers` per attempt.

## Impact Analysis

- Public GraphClient method signatures and response shapes unchanged — tools and resources untouched.
- Error strings preserved verbatim (tests assert them).
- Behavior additions only: requests now time out (30s / connect 10s), retry on throttling, `list_sites` returns all pages.
