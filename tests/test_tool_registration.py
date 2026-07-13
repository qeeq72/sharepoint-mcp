"""Tests for conditional tool registration."""

import pytest
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

DOCUMENT_PARSING_TOOLS = {"get_document_content", "get_document_by_path"}


@pytest.fixture
def default_flags(monkeypatch):
    """Pin all registration flags to defaults, ignoring the local .env."""
    monkeypatch.setattr(settings, "ENABLE_PROVISIONING_TOOLS", False)
    monkeypatch.setattr(settings, "ENABLE_DOCUMENT_PARSING_TOOLS", False)
    monkeypatch.setattr(settings, "DISABLED_TOOLS", [])
    return monkeypatch


def registered_tool_names(mcp):
    return {tool.name for tool in mcp._tool_manager.list_tools()}


def test_default_toolset(default_flags):
    """Test the default registry: no provisioning, no document parsing."""
    mcp = FastMCP("test")
    register_site_tools(mcp)
    names = registered_tool_names(mcp)
    assert names.isdisjoint(PROVISIONING_TOOLS)
    assert names.isdisjoint(DOCUMENT_PARSING_TOOLS)
    # Read and write tools are still there
    assert "search_sharepoint" in names
    assert "download_file" in names
    assert "upload_document" in names
    assert len(names) == 12


def test_provisioning_tools_enabled_by_flag(default_flags):
    """Test that the env flag brings the provisioning tools back."""
    default_flags.setattr(settings, "ENABLE_PROVISIONING_TOOLS", True)
    mcp = FastMCP("test")
    register_site_tools(mcp)
    names = registered_tool_names(mcp)
    assert PROVISIONING_TOOLS <= names
    assert len(names) == 17


def test_document_parsing_tools_enabled_by_flag(default_flags):
    """Test that the env flag brings the document parsing tools back."""
    default_flags.setattr(settings, "ENABLE_DOCUMENT_PARSING_TOOLS", True)
    mcp = FastMCP("test")
    register_site_tools(mcp)
    names = registered_tool_names(mcp)
    assert DOCUMENT_PARSING_TOOLS <= names
    assert len(names) == 14


def test_disabled_tools_list_removes_individual_tools(default_flags):
    """Test that MCP_DISABLED_TOOLS hides specific tools by name."""
    default_flags.setattr(
        settings, "DISABLED_TOOLS", ["upload_document", "update_list_item"]
    )
    mcp = FastMCP("test")
    register_site_tools(mcp)
    names = registered_tool_names(mcp)
    assert "upload_document" not in names
    assert "update_list_item" not in names
    assert "search_sharepoint" in names
    assert len(names) == 10


def test_disabled_tools_unknown_name_is_tolerated(default_flags):
    """Test that an unknown name in MCP_DISABLED_TOOLS does not break startup."""
    default_flags.setattr(settings, "DISABLED_TOOLS", ["no_such_tool"])
    mcp = FastMCP("test")
    register_site_tools(mcp)
    assert len(registered_tool_names(mcp)) == 12
