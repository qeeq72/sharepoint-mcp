"""Tests for DNS-rebinding protection configuration."""

from server import build_transport_security


def test_unset_keeps_sdk_default(monkeypatch):
    """Test that without env vars the SDK default (localhost-only) is kept."""
    monkeypatch.delenv("MCP_ALLOWED_HOSTS", raising=False)
    monkeypatch.delenv("MCP_ALLOWED_ORIGINS", raising=False)
    assert build_transport_security() is None


def test_hosts_list_is_passed_through(monkeypatch):
    """Test that a host list lands in allowed_hosts with protection on."""
    monkeypatch.setenv("MCP_ALLOWED_HOSTS", "mcp.example.com:8001, other.host:9000")
    monkeypatch.delenv("MCP_ALLOWED_ORIGINS", raising=False)
    settings = build_transport_security()
    assert settings.enable_dns_rebinding_protection is True
    assert settings.allowed_hosts == ["mcp.example.com:8001", "other.host:9000"]


def test_wildcard_disables_protection(monkeypatch):
    """Test that '*' turns DNS-rebinding protection off entirely."""
    monkeypatch.setenv("MCP_ALLOWED_HOSTS", "*")
    settings = build_transport_security()
    assert settings.enable_dns_rebinding_protection is False


def test_origins_only(monkeypatch):
    """Test that origins can be set without hosts."""
    monkeypatch.delenv("MCP_ALLOWED_HOSTS", raising=False)
    monkeypatch.setenv("MCP_ALLOWED_ORIGINS", "https://app.example.com")
    settings = build_transport_security()
    assert settings.allowed_origins == ["https://app.example.com"]
