"""Shared helpers for MCP tool modules."""

from config import settings


def _check_auth(sp_ctx) -> None:
    """Check if authentication context is valid, raise exception if not."""
    if not sp_ctx or sp_ctx.access_token == "error" or not sp_ctx.is_token_valid():
        raise Exception(
            "SharePoint authentication failed. Please check your configuration "
            "(CLIENT_ID, CLIENT_SECRET, TENANT_ID, SITE_URL)."
        )


def _normalize_site_ref(value: str) -> str:
    """Normalize a site URL or ID for allowlist comparison."""
    return value.strip().rstrip("/").lower()


def is_site_allowed(site_id: str = "", web_url: str = "") -> bool:
    """Check a site (by ID and/or URL) against MCP_ALLOWED_SITES.

    An empty allowlist means no restriction.
    """
    allowed = settings.ALLOWED_SITES
    if not allowed:
        return True
    normalized = {_normalize_site_ref(entry) for entry in allowed}
    candidates = {_normalize_site_ref(v) for v in (site_id, web_url) if v}
    return bool(candidates & normalized)


async def resolve_site(graph_client, site: str) -> dict:
    """Resolve a site URL or site ID to {"id", "web_url"}."""
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


async def ensure_site_allowed(graph_client, site: str) -> None:
    """Raise unless `site` (URL or ID) is permitted by MCP_ALLOWED_SITES.

    Fast path: the reference itself matches an allowlist entry. Otherwise
    the site is resolved once so both its ID and URL can be compared.
    """
    if not settings.ALLOWED_SITES:
        return
    if is_site_allowed(site_id=site, web_url=site):
        return
    resolved = await resolve_site(graph_client, site)
    if is_site_allowed(site_id=resolved["id"], web_url=resolved["web_url"]):
        return
    raise Exception(f"Access to site '{site}' is not allowed by MCP_ALLOWED_SITES")
