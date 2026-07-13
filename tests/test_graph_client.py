import json

import httpx
import pytest
from unittest.mock import patch, MagicMock
from datetime import datetime, timedelta

from auth.sharepoint_auth import SharePointContext
from utils.graph_client import GraphClient, LARGE_FILE_THRESHOLD, UPLOAD_CHUNK_SIZE


@pytest.fixture
def mock_context():
    """Create a mock SharePoint context for testing."""
    return SharePointContext(
        access_token="test_token", token_expiry=datetime.now() + timedelta(hours=1)
    )


@pytest.fixture
def graph_client(mock_context):
    """Create a GraphClient instance with mock context."""
    return GraphClient(mock_context)


def make_client(mock_context, responses):
    """Create a GraphClient backed by a MockTransport serving `responses` in order.

    Returns (client, calls) where calls collects each httpx.Request made.
    """
    calls = []

    def handler(request):
        calls.append(request)
        return responses.pop(0)

    http = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    return GraphClient(mock_context, http_client=http), calls


@pytest.fixture
def no_sleep(monkeypatch):
    """Replace the retry sleep with a recording no-op; returns recorded delays."""
    delays = []

    async def fake_sleep(seconds):
        delays.append(seconds)

    monkeypatch.setattr("utils._graph_http.asyncio.sleep", fake_sleep)
    return delays


async def test_get(mock_context):
    """Test the GET method of GraphClient."""
    client, calls = make_client(
        mock_context, [httpx.Response(200, json={"value": "test_data"})]
    )
    result = await client.get("endpoint/test")
    assert result == {"value": "test_data"}
    assert str(calls[0].url) == "https://graph.microsoft.com/v1.0/endpoint/test"
    assert calls[0].headers["Authorization"] == "Bearer test_token"

    # Test error response
    client, calls = make_client(mock_context, [httpx.Response(404, text="Not Found")])
    with pytest.raises(Exception) as excinfo:
        await client.get("endpoint/error")
    assert "Graph API error: 404" in str(excinfo.value)


async def test_get_absolute_url(mock_context):
    """Test that get() accepts absolute URLs (used for @odata.nextLink)."""
    client, calls = make_client(mock_context, [httpx.Response(200, json={"value": []})])
    await client.get("https://graph.microsoft.com/v1.0/sites?$skiptoken=abc")
    assert str(calls[0].url) == "https://graph.microsoft.com/v1.0/sites?$skiptoken=abc"


async def test_post(mock_context):
    """Test the POST method of GraphClient."""
    client, calls = make_client(
        mock_context, [httpx.Response(201, json={"id": "new_item_id"})]
    )
    test_data = {"name": "test_item"}
    result = await client.post("endpoint/create", test_data)
    assert result == {"id": "new_item_id"}
    assert str(calls[0].url) == "https://graph.microsoft.com/v1.0/endpoint/create"
    assert json.loads(calls[0].content) == test_data

    # Test error response
    client, calls = make_client(mock_context, [httpx.Response(400, text="Bad Request")])
    with pytest.raises(Exception) as excinfo:
        await client.post("endpoint/error", test_data)
    assert "Graph API error: 400" in str(excinfo.value)


# ---------------------------------------------------------------------------
# Retry behavior
# ---------------------------------------------------------------------------


async def test_retry_on_429_honors_retry_after(mock_context, no_sleep):
    """Test that a 429 is retried after the Retry-After delay."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(429, headers={"Retry-After": "2"}, text="throttled"),
            httpx.Response(200, json={"value": "ok"}),
        ],
    )
    result = await client.get("endpoint/throttled")
    assert result == {"value": "ok"}
    assert len(calls) == 2
    assert no_sleep == [2.0]


async def test_retry_exhaustion_raises(mock_context, no_sleep):
    """Test that retries are capped and the final 429 raises."""
    client, calls = make_client(
        mock_context, [httpx.Response(429, text="throttled")] * 4
    )
    with pytest.raises(Exception) as excinfo:
        await client.get("endpoint/throttled")
    assert "Graph API error: 429" in str(excinfo.value)
    assert len(calls) == 4  # initial attempt + 3 retries


async def test_retry_after_is_capped(mock_context, no_sleep):
    """Test that an excessive Retry-After value is capped at 60s."""
    client, _ = make_client(
        mock_context,
        [
            httpx.Response(429, headers={"Retry-After": "300"}, text="throttled"),
            httpx.Response(200, json={}),
        ],
    )
    await client.get("endpoint/throttled")
    assert no_sleep == [60.0]


async def test_invalid_retry_after_falls_back_to_backoff(mock_context, no_sleep):
    """Test exponential backoff when Retry-After is unparseable."""
    client, _ = make_client(
        mock_context,
        [
            httpx.Response(429, headers={"Retry-After": "soon"}, text="x"),
            httpx.Response(429, text="x"),
            httpx.Response(200, json={}),
        ],
    )
    await client.get("endpoint/throttled")
    assert no_sleep == [1.0, 2.0]


async def test_get_retries_on_503(mock_context, no_sleep):
    """Test that GET is retried on 503."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(503, text="unavailable"),
            httpx.Response(200, json={"value": "ok"}),
        ],
    )
    result = await client.get("endpoint/flaky")
    assert result == {"value": "ok"}
    assert len(calls) == 2


async def test_post_not_retried_on_503(mock_context, no_sleep):
    """Test that non-idempotent POST is NOT retried on 5xx."""
    client, calls = make_client(mock_context, [httpx.Response(503, text="unavailable")])
    with pytest.raises(Exception) as excinfo:
        await client.post("endpoint/create", {"name": "x"})
    assert "Graph API error: 503" in str(excinfo.value)
    assert len(calls) == 1
    assert no_sleep == []


# ---------------------------------------------------------------------------
# Pagination
# ---------------------------------------------------------------------------


async def test_get_paged_merges_pages(mock_context):
    """Test that get_paged follows @odata.nextLink and merges values."""
    next_url = "https://graph.microsoft.com/v1.0/sites?$skiptoken=page2"
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={"value": [{"id": "a"}], "@odata.nextLink": next_url},
            ),
            httpx.Response(200, json={"value": [{"id": "b"}]}),
        ],
    )
    result = await client.get_paged("sites?search=%2A")
    assert [item["id"] for item in result["value"]] == ["a", "b"]
    assert "@odata.nextLink" not in result
    assert len(calls) == 2
    assert str(calls[1].url) == next_url


async def test_get_paged_respects_page_cap(mock_context):
    """Test that get_paged stops at max_pages even if nextLink continues."""
    endless = httpx.Response(
        200,
        json={
            "value": [{"id": "x"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites?$skiptoken=n",
        },
    )
    client, calls = make_client(mock_context, [endless] * 10)
    result = await client.get_paged("sites?search=%2A", max_pages=3)
    assert len(result["value"]) == 3
    assert len(calls) == 3


async def test_list_folder_contents_root(mock_context):
    """Test list_folder_contents with an empty path (root listing)."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200, json={"value": [{"name": "General", "folder": {}, "id": "abc123"}]}
            )
        ],
    )
    result = await client.list_folder_contents("site1", "drive1", "")
    assert result["value"][0]["name"] == "General"
    call_url = str(calls[0].url)
    assert "root/children" in call_url
    assert "root:/" not in call_url


async def test_list_folder_contents_subfolder(mock_context):
    """Test list_folder_contents with a subfolder path."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={"value": [{"name": "report.docx", "file": {}, "id": "def456"}]},
            )
        ],
    )
    result = await client.list_folder_contents("site1", "drive1", "General")
    assert result["value"][0]["name"] == "report.docx"
    assert "root:/General:/children" in str(calls[0].url)


@patch("requests.get")
async def test_get_document_content_by_path(mock_get, graph_client):
    """Test get_document_content_by_path returns bytes content."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.content = b"file content bytes"
    mock_get.return_value = mock_response

    result = await graph_client.get_document_content_by_path(
        "site1", "drive1", "General/report.docx"
    )
    assert result == b"file content bytes"
    call_url = mock_get.call_args[0][0]
    assert "root:/General/report.docx:/content" in call_url


async def test_get_item_metadata_by_path(mock_context):
    """Test get_item_metadata_by_path returns item metadata dict."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={
                    "id": "abc123",
                    "name": "report.docx",
                    "size": 4096,
                    "webUrl": "https://contoso.sharepoint.com/sites/test/report.docx",
                },
            )
        ],
    )
    result = await client.get_item_metadata_by_path(
        "site1", "drive1", "General/report.docx"
    )
    assert result["id"] == "abc123"
    assert result["name"] == "report.docx"
    assert "root:/General/report.docx" in str(calls[0].url)


# ---------------------------------------------------------------------------
# Write tool tests
# ---------------------------------------------------------------------------


@patch("requests.put")
async def test_upload_file(mock_put, graph_client):
    """Test upload_file sends a PUT request with correct URL and content."""
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {"id": "file_id", "name": "test.txt"}
    mock_put.return_value = mock_response

    result = await graph_client.upload_file(
        "sites/site1/drives/drive1/root:/test.txt:/content",
        b"hello",
        "text/plain",
    )
    assert result["id"] == "file_id"
    call_url = mock_put.call_args[0][0]
    assert "root:/test.txt:/content" in call_url

    # Error case
    mock_put.reset_mock()
    mock_response.status_code = 409
    mock_response.text = "Conflict"
    with pytest.raises(Exception) as excinfo:
        await graph_client.upload_file("endpoint/bad", b"data")
    assert "Graph API error: 409" in str(excinfo.value)


@patch("requests.put")
async def test_upload_document_small(mock_put, graph_client):
    """Test upload_document uses simple PUT for files below the threshold."""
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {"id": "doc_id", "name": "small.txt"}
    mock_put.return_value = mock_response

    small_content = b"x" * (LARGE_FILE_THRESHOLD - 1)
    result = await graph_client.upload_document(
        "site1", "drive1", "General", "small.txt", small_content
    )
    assert result["id"] == "doc_id"
    call_url = mock_put.call_args[0][0]
    assert "root:/General/small.txt:/content" in call_url


@patch("requests.put")
@patch("requests.post")
async def test_upload_document_large(mock_post, mock_put, graph_client):
    """Test upload_document uses an upload session for files at or above the threshold."""
    # Mock the upload session creation
    session_response = MagicMock()
    session_response.status_code = 200
    session_response.json.return_value = {
        "uploadUrl": "https://upload.example.com/session123"
    }
    mock_post.return_value = session_response

    # Mock the chunk PUT (two chunks: first returns 202, second returns 201)
    chunk_response_202 = MagicMock()
    chunk_response_202.status_code = 202

    chunk_response_201 = MagicMock()
    chunk_response_201.status_code = 201
    chunk_response_201.json.return_value = {"id": "large_doc_id", "name": "large.bin"}

    mock_put.side_effect = [chunk_response_202, chunk_response_201]

    large_content = b"y" * (UPLOAD_CHUNK_SIZE + 1)
    result = await graph_client.upload_document(
        "site1", "drive1", "Docs", "large.bin", large_content
    )

    assert result["id"] == "large_doc_id"

    # Upload session was created via POST
    session_url = mock_post.call_args[0][0]
    assert "createUploadSession" in session_url
    assert "large.bin" in session_url

    # Two PUT calls for two chunks
    assert mock_put.call_count == 2
    first_headers = mock_put.call_args_list[0][1]["headers"]
    assert first_headers["Content-Range"].startswith("bytes 0-")
    second_headers = mock_put.call_args_list[1][1]["headers"]
    assert f"/{len(large_content)}" in second_headers["Content-Range"]


async def test_create_list_item(mock_context):
    """Test create_list_item sends POST with correct endpoint and body."""
    client, calls = make_client(
        mock_context,
        [httpx.Response(201, json={"id": "item1", "fields": {"Title": "Test"}})],
    )
    result = await client.create_list_item("site1", "list1", {"Title": "Test"})
    assert result["id"] == "item1"
    assert "sites/site1/lists/list1/items" in str(calls[0].url)
    assert json.loads(calls[0].content) == {"fields": {"Title": "Test"}}


async def test_update_list_item(mock_context):
    """Test update_list_item sends PATCH to the item fields endpoint."""
    client, calls = make_client(
        mock_context, [httpx.Response(200, json={"Title": "Updated"})]
    )
    result = await client.update_list_item(
        "site1", "list1", "item1", {"Title": "Updated"}
    )
    assert result["Title"] == "Updated"
    assert "sites/site1/lists/list1/items/item1/fields" in str(calls[0].url)
    assert json.loads(calls[0].content) == {"Title": "Updated"}

    # Error case
    client, _ = make_client(mock_context, [httpx.Response(403, text="Forbidden")])
    with pytest.raises(Exception) as excinfo:
        await client.update_list_item("site1", "list1", "item1", {})
    assert "Graph API error: 403" in str(excinfo.value)


async def test_get_lists(mock_context):
    """Test get_lists returns all lists for a site."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "list1",
                            "displayName": "Tasks",
                            "list": {"template": "genericList"},
                        },
                        {
                            "id": "list2",
                            "displayName": "Documents",
                            "list": {"template": "documentLibrary"},
                        },
                    ]
                },
            )
        ],
    )
    result = await client.get_lists("site1")
    assert len(result["value"]) == 2
    assert result["value"][0]["displayName"] == "Tasks"
    assert "sites/site1/lists" in str(calls[0].url)


async def test_get_list_items(mock_context):
    """Test get_list_items fetches items with fields expanded."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={
                    "value": [
                        {"id": "1", "fields": {"Title": "Item A", "Status": "Active"}},
                        {"id": "2", "fields": {"Title": "Item B", "Status": "Done"}},
                    ]
                },
            )
        ],
    )
    result = await client.get_list_items("site1", "list1")
    assert len(result["value"]) == 2
    assert result["value"][0]["fields"]["Title"] == "Item A"
    call_url = str(calls[0].url)
    assert "sites/site1/lists/list1/items" in call_url
    assert "$expand=fields" in call_url
    assert "$top=100" in call_url


async def test_get_list_items_with_filter(mock_context):
    """Test get_list_items applies OData filter when provided."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={
                    "value": [
                        {"id": "1", "fields": {"Title": "Item A", "Status": "Active"}}
                    ]
                },
            )
        ],
    )
    result = await client.get_list_items(
        "site1", "list1", filter_query="fields/Status eq 'Active'"
    )
    assert len(result["value"]) == 1
    assert "$filter=" in str(calls[0].url)


async def test_list_sites(mock_context):
    """Test list_sites queries the tenant-wide sites endpoint."""
    site = {
        "id": "contoso.sharepoint.com,guid1,guid2",
        "name": "hr",
        "displayName": "HR",
        "webUrl": "https://contoso.sharepoint.com/sites/hr",
    }
    client, calls = make_client(
        mock_context, [httpx.Response(200, json={"value": [site]})]
    )
    result = await client.list_sites()
    assert result["value"][0]["name"] == "hr"
    assert "sites?search=%2A" in str(calls[0].url)

    # Name filter is passed through
    client, calls = make_client(
        mock_context, [httpx.Response(200, json={"value": [site]})]
    )
    await client.list_sites("hr team")
    assert "sites?search=hr%20team" in str(calls[0].url)


async def test_list_sites_paginates(mock_context):
    """Test list_sites follows @odata.nextLink."""
    next_url = "https://graph.microsoft.com/v1.0/sites?$skiptoken=p2"
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={"value": [{"id": "s1"}], "@odata.nextLink": next_url},
            ),
            httpx.Response(200, json={"value": [{"id": "s2"}]}),
        ],
    )
    result = await client.list_sites()
    assert [s["id"] for s in result["value"]] == ["s1", "s2"]
    assert len(calls) == 2


async def test_search_site(mock_context):
    """Test search_site sends the site-scoped search request."""
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "hitsContainers": [
                                {
                                    "hits": [
                                        {
                                            "summary": "match",
                                            "resource": {
                                                "name": "doc.docx",
                                                "webUrl": "url",
                                            },
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                },
            )
        ],
    )
    result = await client.search_site(
        "https://contoso.sharepoint.com/sites/hr", "report"
    )
    hits = result["value"][0]["hitsContainers"][0]["hits"]
    assert hits[0]["resource"]["name"] == "doc.docx"
    assert str(calls[0].url).endswith("/search/query")
    sent_body = json.loads(calls[0].content)
    assert sent_body["requests"][0]["query"]["queryString"] == (
        '(report) AND path:"https://contoso.sharepoint.com/sites/hr"'
    )
    assert sent_body["requests"][0]["size"] == 25
    assert "region" not in sent_body["requests"][0]
    assert sent_body["requests"][0]["entityTypes"] == [
        "driveItem",
        "listItem",
        "list",
    ]


async def test_search_site_custom_size_and_region(mock_context):
    """Test search_site passes size and an explicit region through."""
    client, calls = make_client(mock_context, [httpx.Response(200, json={"value": []})])
    await client.search_site(
        "https://contoso.sharepoint.com/sites/hr", "report", size=5, region="EMEA"
    )
    sent_body = json.loads(calls[0].content)
    assert sent_body["requests"][0]["size"] == 5
    assert sent_body["requests"][0]["region"] == "EMEA"


async def test_search_site_detects_region_and_retries(mock_context, monkeypatch):
    """Test that the region is parsed from the Graph error and retried once."""
    monkeypatch.setattr("utils._graph_site_ops._detected_region", None)
    client, calls = make_client(
        mock_context,
        [
            httpx.Response(
                400,
                json={
                    "error": {
                        "code": "BadRequest",
                        "message": "Requested region  not found. "
                        "Only valid regions are QAT.",
                    }
                },
            ),
            httpx.Response(200, json={"value": []}),
        ],
    )
    result = await client.search_site(
        "https://contoso.sharepoint.com/sites/hr", "report"
    )
    assert result["value"] == []
    assert len(calls) == 2
    first_body = json.loads(calls[0].content)
    assert "region" not in first_body["requests"][0]
    retry_body = json.loads(calls[1].content)
    assert retry_body["requests"][0]["region"] == "QAT"

    # The detected region is cached and used on subsequent calls
    client, calls = make_client(mock_context, [httpx.Response(200, json={"value": []})])
    await client.search_site("https://contoso.sharepoint.com/sites/hr", "next")
    assert json.loads(calls[0].content)["requests"][0]["region"] == "QAT"


async def test_search_site_region_required_probe(mock_context, monkeypatch):
    """Test region discovery when Graph only says 'Region is required'."""
    monkeypatch.setattr("utils._graph_site_ops._detected_region", None)
    required = httpx.Response(
        400,
        json={
            "error": {
                "code": "BadRequest",
                "message": "SearchRequest Invalid (Region is required when "
                "request with application permission.)",
            }
        },
    )
    wrong_region = httpx.Response(
        400,
        json={
            "error": {
                "code": "BadRequest",
                "message": "Requested region  not found. "
                "Only valid regions are QAT.",
            }
        },
    )
    client, calls = make_client(
        mock_context, [required, wrong_region, httpx.Response(200, json={"value": []})]
    )
    result = await client.search_site(
        "https://contoso.sharepoint.com/sites/hr", "report"
    )
    assert result["value"] == []
    assert len(calls) == 3
    assert "region" not in json.loads(calls[0].content)["requests"][0]
    assert json.loads(calls[1].content)["requests"][0]["region"] == "NAM"
    assert json.loads(calls[2].content)["requests"][0]["region"] == "QAT"


async def test_search_site_empty_value(mock_context):
    """Test search_site tolerates an empty value array (no results)."""
    client, _ = make_client(mock_context, [httpx.Response(200, json={"value": []})])
    result = await client.search_site(
        "https://contoso.sharepoint.com/sites/hr", "nothing"
    )
    assert result["value"] == []


async def test_create_site(mock_context):
    """Test create_site sends POST with correct display name and alias."""
    client, calls = make_client(
        mock_context,
        [httpx.Response(201, json={"id": "new_site_id", "displayName": "My Site"})],
    )
    result = await client.create_site("My Site", "mysite", "A test site")
    assert result["id"] == "new_site_id"
    sent_body = json.loads(calls[0].content)
    assert sent_body["displayName"] == "My Site"
    assert sent_body["alias"] == "mysite"
    assert sent_body["description"] == "A test site"
