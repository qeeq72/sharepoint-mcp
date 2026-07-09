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


@patch("requests.get")
async def test_get(mock_get, graph_client):
    """Test the GET method of GraphClient."""
    # Setup mock response
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": "test_data"}
    mock_get.return_value = mock_response

    # Test successful request
    result = await graph_client.get("endpoint/test")
    assert result == {"value": "test_data"}
    mock_get.assert_called_once_with(
        "https://graph.microsoft.com/v1.0/endpoint/test",
        headers=graph_client.context.headers,
    )

    # Test error response
    mock_get.reset_mock()
    mock_response.status_code = 404
    mock_response.text = "Not Found"

    with pytest.raises(Exception) as excinfo:
        await graph_client.get("endpoint/error")
    assert "Graph API error: 404" in str(excinfo.value)


@patch("requests.post")
async def test_post(mock_post, graph_client):
    """Test the POST method of GraphClient."""
    # Setup mock response
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {"id": "new_item_id"}
    mock_post.return_value = mock_response

    # Test data
    test_data = {"name": "test_item"}

    # Test successful request
    result = await graph_client.post("endpoint/create", test_data)
    assert result == {"id": "new_item_id"}
    mock_post.assert_called_once_with(
        "https://graph.microsoft.com/v1.0/endpoint/create",
        headers=graph_client.context.headers,
        json=test_data,
    )

    # Test error response
    mock_post.reset_mock()
    mock_response.status_code = 400
    mock_response.text = "Bad Request"

    with pytest.raises(Exception) as excinfo:
        await graph_client.post("endpoint/error", test_data)
    assert "Graph API error: 400" in str(excinfo.value)


@patch("requests.get")
async def test_list_folder_contents_root(mock_get, graph_client):
    """Test list_folder_contents with an empty path (root listing)."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "value": [{"name": "General", "folder": {}, "id": "abc123"}]
    }
    mock_get.return_value = mock_response

    result = await graph_client.list_folder_contents("site1", "drive1", "")
    assert result["value"][0]["name"] == "General"
    call_url = mock_get.call_args[0][0]
    assert "root/children" in call_url
    assert "root:/" not in call_url


@patch("requests.get")
async def test_list_folder_contents_subfolder(mock_get, graph_client):
    """Test list_folder_contents with a subfolder path."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "value": [{"name": "report.docx", "file": {}, "id": "def456"}]
    }
    mock_get.return_value = mock_response

    result = await graph_client.list_folder_contents("site1", "drive1", "General")
    assert result["value"][0]["name"] == "report.docx"
    call_url = mock_get.call_args[0][0]
    assert "root:/General:/children" in call_url


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


@patch("requests.get")
async def test_get_item_metadata_by_path(mock_get, graph_client):
    """Test get_item_metadata_by_path returns item metadata dict."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "id": "abc123",
        "name": "report.docx",
        "size": 4096,
        "webUrl": "https://contoso.sharepoint.com/sites/test/report.docx",
    }
    mock_get.return_value = mock_response

    result = await graph_client.get_item_metadata_by_path(
        "site1", "drive1", "General/report.docx"
    )
    assert result["id"] == "abc123"
    assert result["name"] == "report.docx"
    call_url = mock_get.call_args[0][0]
    assert "root:/General/report.docx" in call_url


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


@patch("requests.post")
async def test_create_list_item(mock_post, graph_client):
    """Test create_list_item sends POST with correct endpoint and body."""
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {"id": "item1", "fields": {"Title": "Test"}}
    mock_post.return_value = mock_response

    result = await graph_client.create_list_item("site1", "list1", {"Title": "Test"})
    assert result["id"] == "item1"
    call_url = mock_post.call_args[0][0]
    assert "sites/site1/lists/list1/items" in call_url
    sent_body = mock_post.call_args[1]["json"]
    assert sent_body == {"fields": {"Title": "Test"}}


@patch("requests.patch")
async def test_update_list_item(mock_patch, graph_client):
    """Test update_list_item sends PATCH to the item fields endpoint."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"Title": "Updated"}
    mock_patch.return_value = mock_response

    result = await graph_client.update_list_item(
        "site1", "list1", "item1", {"Title": "Updated"}
    )
    assert result["Title"] == "Updated"
    call_url = mock_patch.call_args[0][0]
    assert "sites/site1/lists/list1/items/item1/fields" in call_url
    sent_body = mock_patch.call_args[1]["json"]
    assert sent_body == {"Title": "Updated"}

    # Error case
    mock_patch.reset_mock()
    mock_response.status_code = 403
    mock_response.text = "Forbidden"
    with pytest.raises(Exception) as excinfo:
        await graph_client.update_list_item("site1", "list1", "item1", {})
    assert "Graph API error: 403" in str(excinfo.value)


@patch("requests.get")
async def test_get_lists(mock_get, graph_client):
    """Test get_lists returns all lists for a site."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
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
    }
    mock_get.return_value = mock_response

    result = await graph_client.get_lists("site1")
    assert len(result["value"]) == 2
    assert result["value"][0]["displayName"] == "Tasks"
    call_url = mock_get.call_args[0][0]
    assert "sites/site1/lists" in call_url


@patch("requests.get")
async def test_get_list_items(mock_get, graph_client):
    """Test get_list_items fetches items with fields expanded."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "value": [
            {"id": "1", "fields": {"Title": "Item A", "Status": "Active"}},
            {"id": "2", "fields": {"Title": "Item B", "Status": "Done"}},
        ]
    }
    mock_get.return_value = mock_response

    result = await graph_client.get_list_items("site1", "list1")
    assert len(result["value"]) == 2
    assert result["value"][0]["fields"]["Title"] == "Item A"
    call_url = mock_get.call_args[0][0]
    assert "sites/site1/lists/list1/items" in call_url
    assert "$expand=fields" in call_url
    assert "$top=100" in call_url


@patch("requests.get")
async def test_get_list_items_with_filter(mock_get, graph_client):
    """Test get_list_items applies OData filter when provided."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "value": [{"id": "1", "fields": {"Title": "Item A", "Status": "Active"}}]
    }
    mock_get.return_value = mock_response

    result = await graph_client.get_list_items(
        "site1", "list1", filter_query="fields/Status eq 'Active'"
    )
    assert len(result["value"]) == 1
    call_url = mock_get.call_args[0][0]
    assert "$filter=" in call_url


@patch("requests.get")
async def test_list_sites(mock_get, graph_client):
    """Test list_sites queries the tenant-wide sites endpoint."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "value": [
            {
                "id": "contoso.sharepoint.com,guid1,guid2",
                "name": "hr",
                "displayName": "HR",
                "webUrl": "https://contoso.sharepoint.com/sites/hr",
            }
        ]
    }
    mock_get.return_value = mock_response

    result = await graph_client.list_sites()
    assert result["value"][0]["name"] == "hr"
    call_url = mock_get.call_args[0][0]
    assert "sites?search=%2A" in call_url

    # Name filter is passed through
    mock_get.reset_mock()
    mock_get.return_value = mock_response
    await graph_client.list_sites("hr team")
    call_url = mock_get.call_args[0][0]
    assert "sites?search=hr%20team" in call_url


@patch("requests.post")
async def test_search_site(mock_post, graph_client):
    """Test search_site sends the site-scoped search request."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "value": [
            {
                "hitsContainers": [
                    {
                        "hits": [
                            {
                                "summary": "match",
                                "resource": {"name": "doc.docx", "webUrl": "url"},
                            }
                        ]
                    }
                ]
            }
        ]
    }
    mock_post.return_value = mock_response

    result = await graph_client.search_site("site1", "report")
    hits = result["value"][0]["hitsContainers"][0]["hits"]
    assert hits[0]["resource"]["name"] == "doc.docx"
    call_url = mock_post.call_args[0][0]
    assert "sites/site1/search" in call_url
    sent_body = mock_post.call_args[1]["json"]
    assert sent_body["requests"][0]["query"]["queryString"] == "report"
    assert sent_body["requests"][0]["entityTypes"] == [
        "driveItem",
        "listItem",
        "list",
    ]


@patch("requests.post")
async def test_search_site_empty_value(mock_post, graph_client):
    """Test search_site tolerates an empty value array (no results)."""
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": []}
    mock_post.return_value = mock_response

    result = await graph_client.search_site("site1", "nothing")
    assert result["value"] == []


@patch("requests.post")
async def test_create_site(mock_post, graph_client):
    """Test create_site sends POST with correct display name and alias."""
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {"id": "new_site_id", "displayName": "My Site"}
    mock_post.return_value = mock_response

    result = await graph_client.create_site("My Site", "mysite", "A test site")
    assert result["id"] == "new_site_id"
    sent_body = mock_post.call_args[1]["json"]
    assert sent_body["displayName"] == "My Site"
    assert sent_body["alias"] == "mysite"
    assert sent_body["description"] == "A test site"
