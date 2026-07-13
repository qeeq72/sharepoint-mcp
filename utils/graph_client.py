"""Microsoft Graph API client for SharePoint MCP server."""

import logging

from auth.sharepoint_auth import SharePointContext
from utils._graph_constants import LARGE_FILE_THRESHOLD, UPLOAD_CHUNK_SIZE  # noqa: F401
from utils._graph_drive_ops import _GraphDriveOpsMixin
from utils._graph_http import _GraphHttpMixin
from utils._graph_list_ops import _GraphListOpsMixin
from utils._graph_page_ops import _GraphPageOpsMixin
from utils._graph_site_ops import _GraphSiteOpsMixin

# Logging is configured by the entry point (server.py); library modules
# must not call basicConfig, or they override the DEBUG-aware setup.
logger = logging.getLogger("graph_client")


class GraphClient(
    _GraphSiteOpsMixin,
    _GraphListOpsMixin,
    _GraphPageOpsMixin,
    _GraphDriveOpsMixin,
    _GraphHttpMixin,
):
    """Client for interacting with Microsoft Graph API."""

    def __init__(self, context: SharePointContext, http_client=None):
        """Initialize Graph client with SharePoint context.

        Args:
            context: Authenticated SharePoint context
            http_client: Optional httpx.AsyncClient override (used in tests);
                defaults to the shared module-level client
        """
        self.context = context
        self.base_url = context.graph_url
        self._http = http_client
        logger.debug(f"GraphClient initialized with base URL: {self.base_url}")
