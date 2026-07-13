"""Bearer-token authentication middleware for HTTP transports."""

import hmac
import json
import logging

logger = logging.getLogger("asgi_auth")


class BearerAuthMiddleware:
    """Reject HTTP requests that lack the expected bearer token.

    Pure ASGI middleware: wraps any ASGI app, checks the Authorization
    header on every http-scope request, and answers 401 itself when the
    token is missing or wrong. Non-http scopes (lifespan, websocket) pass
    through untouched.
    """

    def __init__(self, app, token: str):
        if not token:
            raise ValueError("BearerAuthMiddleware requires a non-empty token")
        self.app = app
        self.token = token

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        authorization = ""
        for name, value in scope.get("headers", []):
            if name == b"authorization":
                authorization = value.decode("latin-1")
                break

        scheme, _, credentials = authorization.partition(" ")
        if scheme.lower() == "bearer" and hmac.compare_digest(
            credentials.strip(), self.token
        ):
            await self.app(scope, receive, send)
            return

        client = scope.get("client")
        client_host = client[0] if client else "unknown"
        logger.warning(f"Rejected unauthenticated request from {client_host}")
        body = json.dumps({"error": "unauthorized"}).encode()
        await send(
            {
                "type": "http.response.start",
                "status": 401,
                "headers": [
                    (b"content-type", b"application/json"),
                    (b"content-length", str(len(body)).encode()),
                    (b"www-authenticate", b"Bearer"),
                ],
            }
        )
        await send({"type": "http.response.body", "body": body})
