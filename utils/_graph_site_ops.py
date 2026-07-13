"""Site and document library operations mixin for GraphClient."""

import logging
import re
from typing import Dict, Any, List, Optional
from urllib.parse import quote

logger = logging.getLogger("graph_client")

# Tenant search region detected from a Graph error response ("Only valid
# regions are XXX"). Cached for the process lifetime; app-only calls to
# /search/query require it.
_detected_region: Optional[str] = None


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
        return await self.get_paged(endpoint)

    async def search_site(
        self,
        site_url: str,
        query: str,
        entity_types: Optional[List[str]] = None,
        size: int = 25,
        region: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Search content within a single site via POST /search/query.

        The search is scoped to the site with a KQL path filter. App-only
        tokens require the `region` request property; if it is not supplied
        (or wrong), the region is auto-detected from the Graph error message
        and the request is retried once.

        Args:
            site_url: Web URL of the site to search (used as the path filter)
            query: Search query string
            entity_types: Entity types to search (default: driveItem/listItem/list)
            size: Maximum number of hits to return (Graph default is 25)
            region: Tenant search region (e.g. "NAM", "EMEA"); auto-detected
                when omitted
        """
        global _detected_region

        def build_request(request_region: Optional[str]) -> Dict[str, Any]:
            request: Dict[str, Any] = {
                "entityTypes": entity_types or ["driveItem", "listItem", "list"],
                "query": {"queryString": f'({query}) AND path:"{site_url}"'},
                "size": size,
            }
            if request_region:
                request["region"] = request_region
            return {"requests": [request]}

        logger.info(f"Searching site {site_url} for: {query}")
        try:
            return await self.post(
                "search/query", build_request(region or _detected_region)
            )
        except Exception as e:
            match = re.search(r"Only valid regions are (\w+)", str(e))
            if match:
                _detected_region = match.group(1)
                logger.info(f"Detected tenant search region: {_detected_region}")
                return await self.post("search/query", build_request(_detected_region))
            if "Region is required" not in str(e):
                raise
            # Graph names the valid region only in response to a WRONG one,
            # so probe with a deliberate guess to elicit it.
            try:
                result = await self.post("search/query", build_request("NAM"))
                _detected_region = "NAM"
                return result
            except Exception as probe_error:
                match = re.search(r"Only valid regions are (\w+)", str(probe_error))
                if not match:
                    raise
                _detected_region = match.group(1)
                logger.info(f"Detected tenant search region: {_detected_region}")
                return await self.post("search/query", build_request(_detected_region))

    async def create_site(
        self, display_name: str, alias: str, description: str = ""
    ) -> Dict[str, Any]:
        """Create a new SharePoint site."""
        endpoint = "sites/root/sites"
        data = {"displayName": display_name, "alias": alias, "description": description}
        logger.info(f"Creating new site with name: {display_name}, alias: {alias}")
        return await self.post(endpoint, data)
