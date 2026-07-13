[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/demodorigatsuo-sharepoint-mcp-badge.png)](https://mseep.ai/app/demodorigatsuo-sharepoint-mcp)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://docs.anthropic.com/claude/docs/model-context-protocol)
[![Built with Claude Code](https://img.shields.io/badge/Built%20with-Claude%20Code-blueviolet.svg)](https://claude.ai/claude-code)

# SharePoint MCP Server

> **DISCLAIMER**: This project is not affiliated with, endorsed by, or related to Microsoft Corporation. SharePoint and Microsoft Graph API are trademarks of Microsoft Corporation. This is an independent, community-driven project.

SharePoint MCP Server is a [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that connects LLM applications such as Claude to your SharePoint site via the Microsoft Graph API. Use natural language to query documents, manage lists, upload files, and more — directly from your AI assistant.

---

## Features

| Category | Capability |
|----------|------------|
| **Site** | Get site information |
| **Libraries** | Browse document libraries, list folder contents |
| **Documents** | Read DOCX, PDF, XLSX, CSV, TXT; browse by path; get item metadata; upload files |
| **Search** | Full-text search across all site content |
| **Lists** | Create lists with AI-optimized schemas; create, update list items |
| **Pages** | Create modern pages and news posts |
| **Provisioning** | Create new SharePoint sites and advanced document libraries |
| **Transport** | stdio (local), SSE, streamable-http (web / Docker) |

---

## Prerequisites

- Python 3.10 or higher
- A SharePoint site with Microsoft 365
- An Azure AD application registration with the required Graph API permissions (see [docs/auth_guide.md](docs/auth_guide.md))

---

## Quickstart

### 1. Clone and install

```bash
git clone https://github.com/DEmodoriGatsuO/sharepoint-mcp.git
cd sharepoint-mcp

python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

pip install -r requirements.txt
```

### 2. Configure

```bash
cp .env.example .env
# Edit .env with your Azure AD credentials and SharePoint site URL
```

Required variables in `.env`:

| Variable | Description |
|----------|-------------|
| `TENANT_ID` | Azure AD tenant ID |
| `CLIENT_ID` | Azure AD application (client) ID |
| `CLIENT_SECRET` | Azure AD client secret |
| `SITE_URL` | SharePoint site URL (`https://{tenant}.sharepoint.com/sites/{name}`) |

### 3. Verify your setup (optional)

```bash
python config_checker.py   # Validate configuration
python auth-diagnostic.py  # Test authentication
```

### 4. Start the server

```bash
# stdio — default, for Claude Desktop / MCP Inspector
python server.py

# HTTP streamable-http — for web services and Copilot agents
python server.py --transport streamable-http --port 8000

# Docker
docker-compose up
```

---

## Usage

### Claude Desktop

Install the server into Claude Desktop:

```bash
mcp install server.py --name "SharePoint Assistant"
```

Or add it manually to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "python",
      "args": ["/absolute/path/to/sharepoint-mcp/server.py"],
      "env": {
        "TENANT_ID": "...",
        "CLIENT_ID": "...",
        "CLIENT_SECRET": "...",
        "SITE_URL": "..."
      }
    }
  }
}
```

### MCP Inspector (development)

```bash
mcp dev server.py
```

### HTTP Server

```bash
# streamable-http (recommended for Copilot agents and web clients)
python server.py --transport streamable-http --host 0.0.0.0 --port 8000

# SSE
python server.py --transport sse --host 0.0.0.0 --port 8000

# Via environment variables
MCP_TRANSPORT=streamable-http MCP_PORT=8000 python server.py
```

### Docker

```bash
# Build and start (defaults to streamable-http on port 8000)
docker-compose up

# Or run manually
docker build -t sharepoint-mcp .
docker run --env-file .env -p 8000:8000 sharepoint-mcp
```

---

## Available Tools

The following MCP tools are exposed to the LLM:

| Tool | Description |
|------|-------------|
| `get_site_info` | Get name, description, URL, and metadata of the SharePoint site |
| `list_document_libraries` | List all document libraries (drives) in the site |
| `list_folder_contents` | Browse files and folders within a document library by path |
| `get_document_content`† | Read and parse DOCX, PDF, XLSX, CSV, or TXT files |
| `get_document_by_path`† | Retrieve document content by file path |
| `get_item_metadata` | Get metadata for a file or folder |
| `list_sites` | List sites in the tenant, optionally filtered by name |
| `search_sharepoint` | Full-text search across one or more sites (see below) |
| `upload_document` | Upload a file to a document library |
| `create_list_item` | Create a new item in a SharePoint list |
| `update_list_item` | Update an existing item in a SharePoint list |
| `create_intelligent_list`* | Provision a list with an AI-optimized schema |
| `create_advanced_document_library`* | Create a document library with rich metadata |
| `create_modern_page`* | Publish a modern SharePoint page |
| `create_news_post`* | Publish a news article to the site |
| `create_sharepoint_site`* | Provision a new SharePoint team site |

\* Provisioning tools are **disabled by default**; set
`MCP_ENABLE_PROVISIONING_TOOLS=True` to register them.

† Document parsing tools are **disabled by default** (use `download_file`
for raw content); set `MCP_ENABLE_DOCUMENT_PARSING_TOOLS=True` to register
them.

Individual tools can also be hidden by name via
`MCP_DISABLED_TOOLS=tool_one,tool_two` (applied after the group flags).

`MCP_ALLOWED_SITES` (comma-separated site URLs or IDs) restricts every tool
to the listed sites: `list_sites` filters its output, search defaults to the
allowlist and rejects sites outside it, and per-site tools refuse other
sites. This is an application-level guard — for a hard security boundary use
the `Sites.Selected` application permission in Entra ID instead of
`Sites.Read.All`/`Sites.ReadWrite.All`.

### Multi-site search

`search_sharepoint` accepts an optional `sites` argument (a list of site URLs or
site IDs) and searches each site in a single call, merging results tagged with
their source site. The search scope is resolved in priority order:

1. The `sites` argument, when provided.
2. The `SEARCH_SITES` environment variable (comma-separated site URLs or IDs), when set.
3. All sites in the tenant, capped at `max_sites` (default 20; the response
   sets `truncated: true` when the cap applies).

A failure on one site (e.g. no access) does not fail the whole search — it is
reported in the `errors` field of the response. Use `list_sites` to discover
available sites and their IDs.

Each site returns at most `size` results (default 25). Sites that matched more
than that are listed in the `more_results_available_on` response field —
narrow the query or raise `size` to see more.

For detailed usage examples and example prompts, see [docs/usage.md](docs/usage.md).

---

## Monitoring and Troubleshooting

### Logs

The server writes logs to stdout. Set `DEBUG=True` in `.env` to enable verbose logging.

### Common Issues

| Symptom | Resolution |
|---------|------------|
| Authentication failure | Run `python auth-diagnostic.py` to diagnose |
| Permission errors | Verify your Azure AD app has the required Graph API permissions |
| Token issues | Run `python token-decoder.py` to inspect token claims |

---

## Contributing

Contributions are welcome. Please open an issue first to discuss significant changes. See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

All contributions must pass the quality checks before merge:

```bash
black .       # Formatting
ruff check .  # Linting
pytest        # Tests
```

---

## License

Released under the MIT License. See [LICENSE](LICENSE) for details.
