"""SharePoint site tools — thin delegator to sub-modules."""

import logging

from mcp.server.fastmcp import FastMCP

from config import settings
from tools.provisioning_tools import register_provisioning_tools
from tools.read_tools import register_read_tools
from tools.write_tools import register_write_tools

logger = logging.getLogger("sharepoint_tools")


def register_site_tools(mcp: FastMCP):
    """Register SharePoint tools with the MCP server.

    Provisioning tools (site/list/library/page creation) are registered
    only when MCP_ENABLE_PROVISIONING_TOOLS is set — they carry the
    highest blast radius (Sites.Manage.All) and most deployments never
    need them.
    """
    register_read_tools(mcp)
    register_write_tools(mcp)
    if settings.ENABLE_PROVISIONING_TOOLS:
        register_provisioning_tools(mcp)
        logger.info("Provisioning tools enabled")
    else:
        logger.info(
            "Provisioning tools disabled "
            "(set MCP_ENABLE_PROVISIONING_TOOLS=True to enable)"
        )

    if not settings.ENABLE_DOCUMENT_PARSING_TOOLS:
        for name in ("get_document_content", "get_document_by_path"):
            mcp.remove_tool(name)
        logger.info(
            "Document parsing tools disabled "
            "(set MCP_ENABLE_DOCUMENT_PARSING_TOOLS=True to enable)"
        )

    for name in settings.DISABLED_TOOLS:
        try:
            mcp.remove_tool(name)
            logger.info(f"Tool disabled via MCP_DISABLED_TOOLS: {name}")
        except Exception:
            logger.warning(
                f"MCP_DISABLED_TOOLS names unknown or already absent tool: {name}"
            )
