"""Tests for the bearer-token ASGI middleware."""

import httpx
import pytest

from utils.asgi_auth import BearerAuthMiddleware


async def ok_app(scope, receive, send):
    """Minimal ASGI app answering 200 'ok' to any http request."""
    await send(
        {
            "type": "http.response.start",
            "status": 200,
            "headers": [(b"content-type", b"text/plain")],
        }
    )
    await send({"type": "http.response.body", "body": b"ok"})


def make_http_client(token="secret"):
    app = BearerAuthMiddleware(ok_app, token)
    return httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app), base_url="http://test"
    )


async def test_missing_header_rejected():
    """Test that a request without Authorization gets 401."""
    async with make_http_client() as client:
        response = await client.get("/mcp")
    assert response.status_code == 401
    assert response.headers["www-authenticate"] == "Bearer"
    assert response.json() == {"error": "unauthorized"}


async def test_wrong_token_rejected():
    """Test that a wrong bearer token gets 401."""
    async with make_http_client() as client:
        response = await client.get("/mcp", headers={"Authorization": "Bearer nope"})
    assert response.status_code == 401


async def test_wrong_scheme_rejected():
    """Test that a non-bearer scheme gets 401 even with the right secret."""
    async with make_http_client() as client:
        response = await client.get("/mcp", headers={"Authorization": "Basic secret"})
    assert response.status_code == 401


async def test_correct_token_passes():
    """Test that the correct bearer token reaches the wrapped app."""
    async with make_http_client() as client:
        response = await client.get("/mcp", headers={"Authorization": "Bearer secret"})
    assert response.status_code == 200
    assert response.text == "ok"


async def test_non_http_scope_passes_through():
    """Test that non-http scopes (e.g. lifespan) bypass the auth check."""
    seen = []

    async def recorder(scope, receive, send):
        seen.append(scope["type"])

    middleware = BearerAuthMiddleware(recorder, "secret")
    await middleware({"type": "lifespan"}, None, None)
    assert seen == ["lifespan"]


def test_empty_token_forbidden():
    """Test that constructing the middleware without a token fails loudly."""
    with pytest.raises(ValueError):
        BearerAuthMiddleware(ok_app, "")
