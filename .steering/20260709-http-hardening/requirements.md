# Requirements — HTTP Layer Hardening

## Feature Description

Make the HTTP layer support the multi-site search feature reliably: true async parallelism, timeouts, throttling-aware retries, and complete tenant site listings.

## Problems Addressed

1. Async GraphClient methods call blocking `requests` — multi-site fan-out is effectively sequential and blocks the event loop.
2. No timeouts anywhere — a hung request hangs the whole server.
3. No retry on 429 — Graph throttles fan-out; sites drop into the `errors` field spuriously.
4. `list_sites` returns only the first page of `GET /sites?search=` (no `@odata.nextLink` handling).
5. Tests require manual `PYTHONPATH=.`.

## Acceptance Criteria

- `get`/`post`/`patch`/`delete` use a shared `httpx.AsyncClient` with timeouts.
- 429 responses are retried (all methods) honoring numeric `Retry-After` capped at 60s, with exponential backoff fallback, max 3 retries. GET is additionally retried on 502/503/504; non-GET is NOT retried on 5xx (duplicate-create protection).
- `list_sites` follows `@odata.nextLink` (up to 10 pages) and returns merged results.
- Error message format unchanged: `Graph API error: {status} - {text}`.
- Upload/download paths (`upload_file`, `_upload_in_chunks`, drive_ops downloads) and auth remain on `requests` — deliberately out of scope until live-tenant verification is possible (httpx redirect behavior differs). Auth calls gain `timeout=15`.
- `pytest` runs without manual PYTHONPATH.
- All quality checks pass: `black`, `ruff`, `pytest`.

## Explicitly Out of Scope

- Pagination for `get_list_items`/`get_lists`/`list_folder_contents`/`list_document_libraries` (changes `$top` semantics, bloats LLM output).
- Migration of upload/download and auth paths to httpx.
- Search result pagination (`from`/`size`).
