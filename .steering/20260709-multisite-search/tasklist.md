# Tasklist — Multi-Site Search

- [ ] `config/settings.py`: add `search_sites` from `SEARCH_SITES` env var
- [ ] `.env.example`: document `SEARCH_SITES`
- [ ] `utils/_graph_site_ops.py`: add `list_sites`, `search_site`
- [ ] `tools/read_tools.py`: add `list_sites` tool
- [ ] `tools/read_tools.py`: refactor `search_sharepoint` (multi-site, error tolerance, empty-value fix)
- [ ] Tests: `test_list_sites`, `test_search_site`, `test_search_site_empty_value`
- [ ] `README.md`: document new tool and behavior
- [ ] Quality checks: `black .`, `ruff check .`, `pytest`

## Completion Criteria

All acceptance criteria in `requirements.md` met; all checks green.
