"""Main implementation of the SharePoint MCP Server."""

import argparse
import os
import sys
import logging
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
from datetime import datetime, timedelta

import uvicorn
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings

from auth.sharepoint_auth import SharePointContext, get_auth_context
from config.settings import APP_NAME, DEBUG
from utils._graph_http import close_http_client
from utils.asgi_auth import BearerAuthMiddleware
from tools.site_tools import register_site_tools

# Set logging level
logging_level = logging.DEBUG if DEBUG else logging.INFO
logging.basicConfig(
    level=logging_level, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("sharepoint_mcp")


@asynccontextmanager
async def sharepoint_lifespan(server: FastMCP) -> AsyncIterator[SharePointContext]:
    """Manage SharePoint connection lifecycle."""
    logger.info("Initializing SharePoint connection...")

    try:
        # Get SharePoint authentication context
        logger.debug("Attempting to get authentication context...")
        context = await get_auth_context()
        logger.info(f"Authentication successful. Token expiry: {context.token_expiry}")

        # Yield context for use in the application
        yield context

    except Exception as e:
        logger.error(f"Error during SharePoint authentication: {e}")

        # Create error context
        error_context = SharePointContext(
            access_token="error",
            token_expiry=datetime.now() + timedelta(seconds=10),  # Short expiry
            graph_url="https://graph.microsoft.com/v1.0",
        )

        logger.warning("Using error context due to authentication failure")
        yield error_context

    finally:
        await close_http_client()
        logger.info("Ending SharePoint connection...")


def build_transport_security() -> TransportSecuritySettings | None:
    """Build DNS-rebinding protection settings from environment variables.

    MCP_ALLOWED_HOSTS / MCP_ALLOWED_ORIGINS: comma-separated values that
    clients are allowed to send in the Host / Origin headers (include the
    port, e.g. "mcp.example.com:8001"). "*" disables the protection
    entirely. Unset keeps the SDK default (localhost-only), which returns
    421 Invalid Host header to any other hostname.
    """
    hosts = [
        h.strip() for h in os.getenv("MCP_ALLOWED_HOSTS", "").split(",") if h.strip()
    ]
    origins = [
        o.strip() for o in os.getenv("MCP_ALLOWED_ORIGINS", "").split(",") if o.strip()
    ]
    if not hosts and not origins:
        return None
    if "*" in hosts:
        return TransportSecuritySettings(enable_dns_rebinding_protection=False)
    return TransportSecuritySettings(allowed_hosts=hosts, allowed_origins=origins)


# Create MCP server at module level so CLI can find it
mcp = FastMCP(
    APP_NAME,
    lifespan=sharepoint_lifespan,
    transport_security=build_transport_security(),
)

# Register tools
register_site_tools(mcp)


def main():
    """Main entry point for the SharePoint MCP server."""
    parser = argparse.ArgumentParser(description="SharePoint MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse", "streamable-http"],
        default=os.getenv("MCP_TRANSPORT", "stdio"),
        help="Transport protocol (default: stdio)",
    )
    parser.add_argument(
        "--host",
        default=os.getenv("MCP_HOST", "0.0.0.0"),
        help="Bind host for HTTP transports (default: 0.0.0.0)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("MCP_PORT", "8000")),
        help="Bind port for HTTP transports (default: 8000)",
    )
    args = parser.parse_args()

    try:
        logger.info(f"Starting {APP_NAME} server (transport={args.transport})...")
        if args.transport == "stdio":
            mcp.run(transport="stdio")
            return

        mcp.settings.host = args.host
        mcp.settings.port = args.port
        logger.info(f"HTTP server binding to {args.host}:{args.port}")
        app = (
            mcp.streamable_http_app()
            if args.transport == "streamable-http"
            else mcp.sse_app()
        )
        auth_token = os.getenv("MCP_AUTH_TOKEN", "")
        if auth_token:
            app = BearerAuthMiddleware(app, auth_token)
            logger.info("Bearer authentication enabled for the HTTP endpoint")
        else:
            logger.warning(
                "MCP_AUTH_TOKEN is not set — the HTTP endpoint is unauthenticated"
            )
        uvicorn.run(app, host=args.host, port=args.port)
    except Exception as e:
        logger.error(f"Error occurred during MCP server startup: {e}")
        raise


# Main execution
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Fatal error in SharePoint MCP server: {e}")
        sys.exit(1)
