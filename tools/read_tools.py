"""SharePoint read-only tools."""

import asyncio
import base64
import json
import logging
from typing import List, Optional

from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from config.settings import SHAREPOINT_CONFIG
from tools._tool_helpers import _check_auth
from utils.document_processor import DocumentProcessor
from utils.graph_client import GraphClient

logger = logging.getLogger("sharepoint_tools")


def register_read_tools(mcp: FastMCP):
    """Register read-only SharePoint tools with the MCP server."""

    @mcp.tool()
    async def get_site_info(ctx: Context) -> str:
        """Get basic information about the SharePoint site."""
        logger.info("Tool called: get_site_info")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            site_parts = (
                SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            )
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            logger.info(f"Getting info for site: {site_name} in domain: {domain}")

            site_info = await graph_client.get_site_info(domain, site_name)
            result = {
                "name": site_info.get("displayName", "Unknown"),
                "description": site_info.get("description", "No description"),
                "created": site_info.get("createdDateTime", "Unknown"),
                "last_modified": site_info.get("lastModifiedDateTime", "Unknown"),
                "web_url": site_info.get("webUrl", SHAREPOINT_CONFIG["site_url"]),
                "id": site_info.get("id", "Unknown"),
            }
            logger.info(f"Successfully retrieved site info for: {result['name']}")
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in get_site_info: {str(e)}")
            raise

    @mcp.tool()
    async def list_document_libraries(ctx: Context) -> str:
        """List all document libraries in the SharePoint site."""
        logger.info("Tool called: list_document_libraries")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            site_parts = (
                SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            )
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            logger.info(
                f"Listing document libraries for site: {site_name} in domain: {domain}"
            )

            result = await graph_client.list_document_libraries(domain, site_name)
            drives = result.get("value", [])
            formatted_drives = [
                {
                    "name": drive.get("name", "Unknown"),
                    "description": drive.get("description", "No description"),
                    "web_url": drive.get("webUrl", "Unknown"),
                    "drive_type": drive.get("driveType", "Unknown"),
                    "id": drive.get("id", "Unknown"),
                }
                for drive in drives
            ]
            logger.info(
                f"Successfully retrieved {len(formatted_drives)} document libraries"
            )
            return json.dumps(formatted_drives, indent=2)
        except Exception as e:
            logger.error(f"Error in list_document_libraries: {str(e)}")
            raise

    @mcp.tool()
    async def list_sites(ctx: Context, query: str = "") -> str:
        """List SharePoint sites in the tenant.

        Args:
            query: Optional name filter; empty returns all sites
        """
        logger.info(f"Tool called: list_sites with query: {query}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            result = await graph_client.list_sites(query or "*")
            formatted_sites = [
                {
                    "id": site.get("id", "Unknown"),
                    "name": site.get("name", "Unknown"),
                    "display_name": site.get("displayName", "Unknown"),
                    "web_url": site.get("webUrl", "Unknown"),
                }
                for site in result.get("value", [])
            ]
            logger.info(f"Successfully retrieved {len(formatted_sites)} sites")
            return json.dumps(formatted_sites, indent=2)
        except Exception as e:
            logger.error(f"Error in list_sites: {str(e)}")
            raise

    async def _resolve_site(graph_client: GraphClient, site: str) -> dict:
        """Resolve a site URL or site ID to {"id", "web_url"}.

        The web URL is required in both cases: the search request scopes to
        a site with a KQL path filter, which takes a URL, not an ID.
        """
        if "sharepoint.com" in site and ("://" in site or "/" in site):
            site_parts = site.replace("https://", "").replace("http://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            site_info = await graph_client.get_site_info(domain, site_name)
        else:
            site_info = await graph_client.get(f"sites/{site}")
        site_id = site_info.get("id")
        web_url = site_info.get("webUrl")
        if not site_id or not web_url:
            raise Exception(f"Could not resolve site: {site}")
        return {"id": site_id, "web_url": web_url}

    @mcp.tool()
    async def search_sharepoint(
        ctx: Context,
        query: str,
        sites: Optional[List[str]] = None,
        max_sites: int = 20,
        size: int = 25,
    ) -> str:
        """Search content across one or more SharePoint sites.

        Search scope is resolved in priority order:
        1. The `sites` argument (site URLs or site IDs).
        2. The SEARCH_SITES environment variable (comma-separated), if set.
        3. All sites in the tenant, capped at `max_sites`.

        Sites that have more matches than `size` are listed in the
        `more_results_available_on` field of the response — narrow the
        query or raise `size` to see more.

        Args:
            query: Search query string
            sites: Optional list of site URLs or site IDs to search
            max_sites: Maximum number of sites searched when scope is the
                whole tenant (default 20)
            size: Maximum number of results per site (default 25)
        """
        logger.info(f"Tool called: search_sharepoint with query: {query}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            truncated = False
            if sites:
                targets = [await _resolve_site(graph_client, s) for s in sites]
            elif SHAREPOINT_CONFIG["search_sites"]:
                targets = [
                    await _resolve_site(graph_client, s)
                    for s in SHAREPOINT_CONFIG["search_sites"]
                ]
            else:
                all_sites = (await graph_client.list_sites()).get("value", [])
                if len(all_sites) > max_sites:
                    truncated = True
                    all_sites = all_sites[:max_sites]
                targets = [
                    {"id": s["id"], "web_url": s["webUrl"]}
                    for s in all_sites
                    if s.get("id") and s.get("webUrl")
                ]
            logger.info(f"Searching for '{query}' across {len(targets)} site(s)")

            region = SHAREPOINT_CONFIG["search_region"] or None
            search_outcomes = await asyncio.gather(
                *[
                    graph_client.search_site(
                        t["web_url"], query, size=size, region=region
                    )
                    for t in targets
                ],
                return_exceptions=True,
            )

            formatted_results = []
            errors = []
            more_results_available_on = []
            for target, outcome in zip(targets, search_outcomes):
                if isinstance(outcome, BaseException):
                    logger.error(f"Search failed for {target['web_url']}: {outcome}")
                    errors.append({"site": target["web_url"], "error": str(outcome)})
                    continue
                for response in outcome.get("value", []):
                    for container in response.get("hitsContainers", []):
                        if container.get("moreResultsAvailable"):
                            more_results_available_on.append(target["web_url"])
                        for hit in container.get("hits", []):
                            resource = hit.get("resource", {})
                            formatted_results.append(
                                {
                                    "title": resource.get("name", "Unknown"),
                                    "url": resource.get("webUrl", "Unknown"),
                                    "type": resource.get("@odata.type", "Unknown"),
                                    "summary": hit.get(
                                        "summary", "No summary available"
                                    ),
                                    "site": target["web_url"],
                                }
                            )
            logger.info(
                f"Search returned {len(formatted_results)} results "
                f"from {len(targets)} site(s), {len(errors)} error(s)"
            )
            return json.dumps(
                {
                    "results": formatted_results,
                    "errors": errors,
                    "sites_searched": len(targets),
                    "truncated": truncated,
                    "more_results_available_on": sorted(set(more_results_available_on)),
                },
                indent=2,
            )
        except Exception as e:
            logger.error(f"Error in search_sharepoint: {str(e)}")
            raise

    @mcp.tool()
    async def get_document_content(
        ctx: Context, site_id: str, drive_id: str, item_id: str, filename: str
    ) -> str:
        """Get and process content from a SharePoint document.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the document
            filename: Name of the file (for content type detection)
        """
        logger.info(f"Tool called: get_document_content for file: {filename}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            content = await graph_client.get_document_content(
                site_id, drive_id, item_id
            )
            processed_content = DocumentProcessor.process_document(content, filename)
            logger.info(f"Successfully processed document content for: {filename}")
            return json.dumps(processed_content, indent=2)
        except Exception as e:
            logger.error(f"Error in get_document_content: {str(e)}")
            raise

    @mcp.tool()
    async def list_folder_contents(
        ctx: Context, site_id: str, drive_id: str, folder_path: str = ""
    ) -> str:
        """List files and folders at a given path in a SharePoint document library.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library (drive)
            folder_path: Folder path relative to drive root (e.g. "General" or
                "Docs/2026"). Leave empty to list the root of the drive.
        """
        logger.info(f"Tool called: list_folder_contents path='{folder_path or '/'}'")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            result = await graph_client.list_folder_contents(
                site_id, drive_id, folder_path
            )
            items = result.get("value", [])
            formatted = [
                {
                    "name": item.get("name", "Unknown"),
                    "type": "folder" if "folder" in item else "file",
                    "size": item.get("size", 0),
                    "id": item.get("id", "Unknown"),
                    "web_url": item.get("webUrl", "Unknown"),
                    "last_modified": item.get("lastModifiedDateTime", "Unknown"),
                }
                for item in items
            ]
            logger.info(
                f"Successfully listed {len(formatted)} items at path '{folder_path or '/'}'"
            )
            return json.dumps(formatted, indent=2)
        except Exception as e:
            logger.error(f"Error in list_folder_contents: {str(e)}")
            raise

    @mcp.tool()
    async def get_document_by_path(
        ctx: Context, site_id: str, drive_id: str, file_path: str, filename: str
    ) -> str:
        """Get and process the content of a SharePoint document by its path.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library (drive)
            file_path: File path relative to drive root (e.g. "General/report.docx")
            filename: File name used to detect the document type (e.g. "report.docx")
        """
        logger.info(f"Tool called: get_document_by_path path='{file_path}'")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            content = await graph_client.get_document_content_by_path(
                site_id, drive_id, file_path
            )
            processed_content = DocumentProcessor.process_document(content, filename)
            logger.info(
                f"Successfully processed document content for path: '{file_path}'"
            )
            return json.dumps(processed_content, indent=2)
        except Exception as e:
            logger.error(f"Error in get_document_by_path: {str(e)}")
            raise

    @mcp.tool()
    async def get_item_metadata(
        ctx: Context, site_id: str, drive_id: str, item_path: str
    ) -> str:
        """Get metadata of a file or folder by its path in a SharePoint document library.

        Returns the item's id, name, size, web URL, and timestamps.
        Use the returned id with get_document_content to retrieve file content.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library (drive)
            item_path: Item path relative to drive root (e.g. "General/report.docx"
                or "General")
        """
        logger.info(f"Tool called: get_item_metadata path='{item_path}'")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            item = await graph_client.get_item_metadata_by_path(
                site_id, drive_id, item_path
            )
            result = {
                "id": item.get("id", "Unknown"),
                "name": item.get("name", "Unknown"),
                "size": item.get("size", 0),
                "web_url": item.get("webUrl", "Unknown"),
                "created_by": item.get("createdBy", {})
                .get("user", {})
                .get("displayName", "Unknown"),
                "created_datetime": item.get("createdDateTime", "Unknown"),
                "last_modified_datetime": item.get("lastModifiedDateTime", "Unknown"),
            }
            if "folder" in item:
                result["type"] = "folder"
                result["child_count"] = item["folder"].get("childCount", 0)
            elif "file" in item:
                result["type"] = "file"
                result["mime_type"] = item["file"].get("mimeType", "Unknown")

            logger.info(f"Successfully retrieved metadata for path: '{item_path}'")
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in get_item_metadata: {str(e)}")
            raise

    @mcp.tool()
    async def download_file(
        ctx: Context, site_id: str, drive_id: str, item_id: str, filename: str
    ) -> str:
        """Download a file from SharePoint and return its content as base64.

        Use this to retrieve binary files (docx, xlsx, pdf, etc.) so they can
        be edited and re-uploaded. Pair with upload_document to complete edits.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the file
            filename: Name of the file (used for logging)
        """
        logger.info(f"Tool called: download_file for file: {filename}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            content = await graph_client.get_document_content(
                site_id, drive_id, item_id
            )
            encoded = base64.b64encode(content).decode("utf-8")
            logger.info(
                f"Successfully downloaded file: {filename} ({len(content)} bytes)"
            )
            return json.dumps(
                {
                    "filename": filename,
                    "size_bytes": len(content),
                    "content_base64": encoded,
                },
                indent=2,
            )
        except Exception as e:
            logger.error(f"Error in download_file: {str(e)}")
            raise

    @mcp.tool()
    async def get_lists(ctx: Context, site_id: str) -> str:
        """List all SharePoint lists (and document libraries) in a site.

        Returns list names, IDs, and templates so you can query items.
        For subsites, use the compound site ID format: siteCollectionId,webId

        Args:
            site_id: The site ID. For subsites use compound format
                (e.g. "f474e3b9-...,0b580fc8-...")
        """
        logger.info(f"Tool called: get_lists for site: {site_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            result = await graph_client.get_lists(site_id)
            lists = result.get("value", [])
            formatted = [
                {
                    "id": lst.get("id", "Unknown"),
                    "name": lst.get("displayName", "Unknown"),
                    "description": lst.get("description", ""),
                    "web_url": lst.get("webUrl", "Unknown"),
                    "template": lst.get("list", {}).get("template", "Unknown"),
                    "created": lst.get("createdDateTime", "Unknown"),
                    "last_modified": lst.get("lastModifiedDateTime", "Unknown"),
                }
                for lst in lists
            ]
            logger.info(f"Successfully retrieved {len(formatted)} lists")
            return json.dumps(formatted, indent=2)
        except Exception as e:
            logger.error(f"Error in get_lists: {str(e)}")
            raise

    @mcp.tool()
    async def get_list_items(
        ctx: Context,
        site_id: str,
        list_id: str,
        top: int = 100,
        filter_query: str = "",
    ) -> str:
        """Get items from a SharePoint list with all their column/field values.

        This queries the Graph API /sites/{siteId}/lists/{listId}/items endpoint
        with $expand=fields to return all column values including Status, Title, etc.

        Args:
            site_id: The site ID. For subsites use compound format
                (e.g. "f474e3b9-...,0b580fc8-...")
            list_id: The list ID (GUID) or display name
            top: Maximum items to return (default 100, max 5000)
            filter_query: Optional OData filter (e.g. "fields/Status eq 'Active'")
        """
        logger.info(f"Tool called: get_list_items for list: {list_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            result = await graph_client.get_list_items(
                site_id, list_id, top=top, filter_query=filter_query
            )
            items = result.get("value", [])
            formatted = [
                {
                    "id": item.get("id", "Unknown"),
                    "web_url": item.get("webUrl", ""),
                    "created": item.get("createdDateTime", ""),
                    "last_modified": item.get("lastModifiedDateTime", ""),
                    "fields": item.get("fields", {}),
                }
                for item in items
            ]
            logger.info(f"Successfully retrieved {len(formatted)} list items")
            return json.dumps(formatted, indent=2)
        except Exception as e:
            logger.error(f"Error in get_list_items: {str(e)}")
            raise
