# Tasklist — HTTP Layer Hardening

- [x] `pytest.ini`: `pythonpath = .`
- [x] `requirements.txt`: add `httpx>=0.27.0`
- [x] `utils/_graph_http.py`: client singleton, `_parse_retry_after`, `_request`, `get_paged`, verb wrappers
- [x] `utils/graph_client.py`: `http_client` injection seam
- [x] `utils/_graph_site_ops.py`: `list_sites` → `get_paged`
- [x] `server.py`: close client in lifespan
- [x] `auth/sharepoint_auth.py`: `timeout=15` on requests calls
- [x] Tests: convert 14 to MockTransport, add retry/pagination cases
- [x] Quality checks: `black .`, `ruff check .`, `pytest` (31 passed)

## Completion Criteria

All acceptance criteria in `requirements.md` met; all checks green.
