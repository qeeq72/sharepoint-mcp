"""Tests for conditional tool registration."""

from mcp.server.fastmcp import FastMCP

from config import settings
from tools.site_tools import register_site_tools

PROVISIONING_TOOLS = {
    "create_sharepoint_site",
    "create_intelligent_list",
    "create_advanced_document_library",
    "create_modern_page",
    "create_news_post",
}


def registered_tool_names(mcp):
    return {tool.name for tool in mcp._tool_manager.list_tools()}


def test_provisioning_tools_disabled_by_default(monkeypatch):
    """Test that provisioning tools are not registered by default."""
    monkeypatch.setattr(settings, "ENABLE_PROVISIONING_TOOLS", False)
    monkeypatch.setattr(settings, "DISABLED_TOOLS", [])
    mcp = FastMCP("test")
    register_site_tools(mcp)
    names = registered_tool_names(mcp)
    assert names.isdisjoint(PROVISIONING_TOOLS)
    # Read and write tools are still there
    assert "search_sharepoint" in names
    assert "upload_document" in names
    assert len(names) == 14


def test_provisioning_tools_enabled_by_flag(monkeypatch):
    """Test that the env flag brings the provisioning tools back."""
    monkeypatch.setattr(settings, "ENABLE_PROVISIONING_TOOLS", True)
    monkeypatch.setattr(settings, "DISABLED_TOOLS", [])
    mcp = FastMCP("test")
    register_site_tools(mcp)
    names = registered_tool_names(mcp)
    assert PROVISIONING_TOOLS <= names
    assert len(names) == 19


def test_disabled_tools_list_removes_individual_tools(monkeypatch):
    """Test that MCP_DISABLED_TOOLS hides specific tools by name."""
    monkeypatch.setattr(settings, "ENABLE_PROVISIONING_TOOLS", False)
    monkeypatch.setattr(
        settings, "DISABLED_TOOLS", ["upload_document", "update_list_item"]
    )
    mcp = FastMCP("test")
    register_site_tools(mcp)
    names = registered_tool_names(mcp)
    assert "upload_document" not in names
    assert "update_list_item" not in names
    assert "search_sharepoint" in names
    assert len(names) == 12


def test_disabled_tools_unknown_name_is_tolerated(monkeypatch):
    """Test that an unknown name in MCP_DISABLED_TOOLS does not break startup."""
    monkeypatch.setattr(settings, "ENABLE_PROVISIONING_TOOLS", False)
    monkeypatch.setattr(settings, "DISABLED_TOOLS", ["no_such_tool"])
    mcp = FastMCP("test")
    register_site_tools(mcp)
    assert len(registered_tool_names(mcp)) == 14
