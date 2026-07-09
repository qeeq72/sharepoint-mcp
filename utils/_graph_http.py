"""HTTP transport mixin for GraphClient."""

import asyncio
import logging
from typing import Any, Dict, Optional, Sequence, Union, BinaryIO

import httpx
import requests

from utils._graph_constants import UPLOAD_CHUNK_SIZE

logger = logging.getLogger("graph_client")

DEFAULT_TIMEOUT = httpx.Timeout(30.0, connect=10.0)
MAX_RETRIES = 3
MAX_RETRY_AFTER = 60.0
# Non-GET requests are retried only on 429 (throttled requests are never
# processed, so a retry cannot create duplicates); GET is also safe on 5xx.
RETRY_STATUSES_GET = frozenset({429, 502, 503, 504})
RETRY_STATUSES_OTHER = frozenset({429})

_http_client: Optional[httpx.AsyncClient] = None


def get_http_client() -> httpx.AsyncClient:
    """Return the shared AsyncClient, creating it lazily if needed."""
    global _http_client
    if _http_client is None or _http_client.is_closed:
        _http_client = httpx.AsyncClient(follow_redirects=True)
    return _http_client


async def close_http_client() -> None:
    """Close the shared AsyncClient if it exists."""
    global _http_client
    if _http_client is not None and not _http_client.is_closed:
        await _http_client.aclose()
    _http_client = None


def _parse_retry_after(value: Optional[str]) -> Optional[float]:
    """Parse a numeric Retry-After header value, capped at MAX_RETRY_AFTER."""
    if not value:
        return None
    try:
        seconds = float(value)
    except ValueError:
        return None
    if seconds < 0:
        return None
    return min(seconds, MAX_RETRY_AFTER)


class _GraphHttpMixin:
    """Base HTTP methods for the Microsoft Graph API."""

    async def _request(
        self,
        method: str,
        endpoint_or_url: str,
        *,
        ok_statuses: Sequence[int] = (200,),
        json: Optional[Dict[str, Any]] = None,
        timeout: httpx.Timeout = DEFAULT_TIMEOUT,
    ) -> httpx.Response:
        """Send a request with timeout and throttling-aware retries."""
        if endpoint_or_url.startswith("http"):
            url = endpoint_or_url
        else:
            url = f"{self.base_url}/{endpoint_or_url.lstrip('/')}"
        logger.debug(f"Making {method} request to: {url}")
        if json is not None:
            logger.debug(f"With data: {json}")

        client = self._http or get_http_client()
        retry_statuses = RETRY_STATUSES_GET if method == "GET" else RETRY_STATUSES_OTHER
        kwargs: Dict[str, Any] = {"timeout": timeout}
        if json is not None:
            kwargs["json"] = json

        for attempt in range(MAX_RETRIES + 1):
            response = await client.request(
                method, url, headers=self.context.headers, **kwargs
            )
            logger.debug(f"Response status code: {response.status_code}")
            if response.status_code in retry_statuses and attempt < MAX_RETRIES:
                delay = _parse_retry_after(response.headers.get("Retry-After"))
                if delay is None:
                    delay = min(2.0**attempt, MAX_RETRY_AFTER)
                logger.warning(
                    f"Got {response.status_code} from {url}, "
                    f"retrying in {delay}s (attempt {attempt + 1}/{MAX_RETRIES})"
                )
                await asyncio.sleep(delay)
                continue
            break

        if response.status_code not in ok_statuses:
            error_text = response.text
            logger.error(f"Graph API error: {response.status_code} - {error_text}")
            if response.status_code in (401, 403):
                logger.error("Authentication or authorization error detected")
                if "scp or roles claim" in error_text:
                    logger.error("Token does not have required claims (scp or roles)")
                    logger.error("Please check application permissions in Azure AD")
            raise Exception(f"Graph API error: {response.status_code} - {error_text}")

        return response

    async def get(self, endpoint: str) -> Dict[str, Any]:
        """Send GET request to Graph API. Accepts absolute URLs (nextLink)."""
        response = await self._request("GET", endpoint)
        return response.json()

    async def get_paged(self, endpoint: str, max_pages: int = 10) -> Dict[str, Any]:
        """GET a collection, following @odata.nextLink up to max_pages pages.

        Returns the same {"value": [...]} shape as get(), with the values of
        all fetched pages concatenated.
        """
        url = endpoint
        items = []
        first_page: Optional[Dict[str, Any]] = None
        for _ in range(max_pages):
            data = await self.get(url)
            if first_page is None:
                first_page = data
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            if not url:
                break
        else:
            logger.warning(
                f"Pagination cap of {max_pages} pages reached for {endpoint}"
            )
        result = dict(first_page or {})
        result["value"] = items
        result.pop("@odata.nextLink", None)
        return result

    async def post(self, endpoint: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """Send POST request to Graph API."""
        response = await self._request(
            "POST", endpoint, ok_statuses=(200, 201), json=data
        )
        return response.json()

    async def patch(self, endpoint: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """Send PATCH request to Graph API."""
        response = await self._request(
            "PATCH", endpoint, ok_statuses=(200, 201, 204), json=data
        )
        if response.status_code == 204:
            return {"status": "success"}
        return response.json()

    async def delete(self, endpoint: str) -> Dict[str, Any]:
        """Send DELETE request to Graph API."""
        await self._request("DELETE", endpoint, ok_statuses=(200, 201, 204))
        return {"status": "success"}

    async def upload_file(
        self,
        endpoint: str,
        file_content: Union[bytes, BinaryIO],
        content_type: str = None,
    ) -> Dict[str, Any]:
        """Upload file content to Graph API via simple PUT."""
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        logger.debug(f"Uploading file to: {url}")

        headers = self.context.headers.copy()
        if content_type:
            headers["Content-Type"] = content_type

        response = requests.put(url, headers=headers, data=file_content)
        logger.debug(f"Response status code: {response.status_code}")

        if response.status_code not in (200, 201, 204):
            error_text = response.text
            logger.error(f"Graph API error: {response.status_code} - {error_text}")
            raise Exception(f"Graph API error: {response.status_code} - {error_text}")

        if response.status_code == 204:
            return {"status": "success"}
        return response.json()

    async def _upload_in_chunks(
        self,
        upload_url: str,
        file_content: bytes,
        content_type: str = None,
    ) -> Dict[str, Any]:
        """Upload file content to an upload session URL in chunks."""
        total_size = len(file_content)
        start = 0
        result: Dict[str, Any] = {}

        while start < total_size:
            end = min(start + UPLOAD_CHUNK_SIZE - 1, total_size - 1)
            chunk = file_content[start : end + 1]

            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end}/{total_size}",
            }
            if content_type:
                headers["Content-Type"] = content_type

            logger.debug(f"Uploading chunk: bytes {start}-{end}/{total_size}")
            response = requests.put(upload_url, headers=headers, data=chunk)

            if response.status_code not in (200, 201, 202):
                raise Exception(
                    f"Chunk upload failed: {response.status_code} - {response.text}"
                )

            if response.status_code in (200, 201):
                result = response.json()

            start = end + 1

        return result
