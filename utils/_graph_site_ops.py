"""Site and document library operations mixin for GraphClient."""

import logging
from typing import Dict, Any, List, Optional
from urllib.parse import quote

logger = logging.getLogger("graph_client")


class _GraphSiteOpsMixin:
    """Site and top-level library operations for the Microsoft Graph API."""

    async def get_site_info(self, domain: str, site_name: str) -> Dict[str, Any]:
        """Get SharePoint site information."""
        if site_name == "root" or not site_name:
            endpoint = f"sites/{domain}:"
        else:
            endpoint = f"sites/{domain}:/sites/{site_name}"
        logger.info(f"Getting site info for domain: {domain}, site: {site_name}")
        return await self.get(endpoint)

    async def list_document_libraries(
        self, domain: str, site_name: str
    ) -> Dict[str, Any]:
        """List all document libraries in the site."""
        site_info = await self.get_site_info(domain, site_name)
        site_id = site_info.get("id")

        if not site_id:
            raise Exception(
                f"Failed to get site ID for domain: {domain}, site: {site_name}"
            )

        endpoint = f"sites/{site_id}/drives"
        logger.info(f"Listing document libraries for site ID: {site_id}")
        return await self.get(endpoint)

    async def list_sites(self, query: str = "*") -> Dict[str, Any]:
        """List sites in the tenant, optionally filtered by name.

        Uses GET /sites?search={query}; "*" returns all sites the
        application has access to.
        """
        endpoint = f"sites?search={quote(query or '*')}"
        logger.info(f"Listing tenant sites with query: {query or '*'}")
        return await self.get(endpoint)

    async def search_site(
        self,
        site_id: str,
        query: str,
        entity_types: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Search content within a single site via POST /sites/{id}/search."""
        endpoint = f"sites/{site_id}/search"
        data = {
            "requests": [
                {
                    "entityTypes": entity_types or ["driveItem", "listItem", "list"],
                    "query": {"queryString": query},
                }
            ]
        }
        logger.info(f"Searching site {site_id} for: {query}")
        return await self.post(endpoint, data)

    async def create_site(
        self, display_name: str, alias: str, description: str = ""
    ) -> Dict[str, Any]:
        """Create a new SharePoint site."""
        endpoint = "sites/root/sites"
        data = {"displayName": display_name, "alias": alias, "description": description}
        logger.info(f"Creating new site with name: {display_name}, alias: {alias}")
        return await self.post(endpoint, data)
