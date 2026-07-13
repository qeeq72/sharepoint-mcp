"""Configuration settings for the SharePoint MCP Server."""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Basic settings
APP_NAME = "SharePoint MCP"
DEBUG = os.getenv("DEBUG", "False").lower() in ("true", "1", "t")

# Provisioning tools (create sites/lists/libraries/pages) are the most
# destructive tool group and are disabled unless explicitly enabled.
ENABLE_PROVISIONING_TOOLS = os.getenv(
    "MCP_ENABLE_PROVISIONING_TOOLS", "False"
).lower() in ("true", "1", "t")

# Individual tools to hide regardless of group flags: comma-separated
# tool names, e.g. "upload_document,update_list_item".
DISABLED_TOOLS = [
    s.strip() for s in os.getenv("MCP_DISABLED_TOOLS", "").split(",") if s.strip()
]

# Site allowlist: comma-separated site URLs or site IDs. When set, every
# tool refuses to touch sites outside this list, list_sites filters its
# output, and the default search scope becomes this list. Empty = no
# restriction. NOTE: this is enforced in application code only — for a
# hard security boundary use the Sites.Selected application permission.
ALLOWED_SITES = [
    s.strip() for s in os.getenv("MCP_ALLOWED_SITES", "").split(",") if s.strip()
]

# SharePoint connection settings
SHAREPOINT_CONFIG = {
    "tenant_id": os.getenv("TENANT_ID", ""),
    "client_id": os.getenv("CLIENT_ID", ""),
    "client_secret": os.getenv("CLIENT_SECRET", ""),
    "site_url": os.getenv("SITE_URL", ""),
    # Optional default scope for multi-site search: comma-separated site URLs
    # or site IDs. Empty means "all sites in the tenant" (capped by max_sites).
    "search_sites": [
        s.strip() for s in os.getenv("SEARCH_SITES", "").split(",") if s.strip()
    ],
    # Tenant search region for /search/query with app-only tokens (e.g. NAM,
    # EMEA). Optional: auto-detected from the Graph error response when empty.
    "search_region": os.getenv("SEARCH_REGION", ""),
    "scope": [
        "https://graph.microsoft.com/.default",
        # The application must have these permissions:
        # - Sites.Read.All (for reading site content)
        # - Sites.ReadWrite.All (for modifying site content)
        # - Sites.Manage.All (for creating sites)
        # - Files.ReadWrite.All (for document operations)
    ],
}

# Microsoft Graph API settings
GRAPH_API_VERSION = "v1.0"
GRAPH_BASE_URL = f"https://graph.microsoft.com/{GRAPH_API_VERSION}"

# Document processing settings
DOCUMENT_PROCESSING = {
    "max_text_preview_length": 5000,  # Maximum characters for text preview
    "max_rows_preview": 50,  # Maximum rows for CSV/Excel preview
    "supported_extensions": [
        "csv",
        "xlsx",
        "xls",
        "docx",
        "pdf",
        "txt",
        "md",
        "html",
        "htm",
    ],
}

# Content generation settings
CONTENT_GENERATION = {
    "default_audience": "general",
    "default_purpose": "general",
    "enable_rich_layout": True,
}
