"""
MCP SDK compatibility layer.

The upstream ``mcp`` package renamed its high-level server class between
major releases:

===================  ==========================================
Installed SDK       High-level server class
===================  ==========================================
``mcp`` 1.x         ``mcp.server.fastmcp.FastMCP``
``mcp`` 2.x         ``mcp.server.mcpserver.MCPServer``
===================  ==========================================

Every module in this package imports ``FastMCP`` from *here* instead of
from the SDK, so a fresh ``pip install`` never breaks just because the
SDK moved a symbol. ``MCP_SDK_GENERATION`` reports which one is live and
``describe()`` renders a one-line human/JSON friendly summary used by the
doctor and by the Antigravity installer.
"""

from __future__ import annotations

import logging
from typing import Any, Optional, Tuple

logger = logging.getLogger(__name__)

__all__ = [
    "FastMCP",
    "MCP_SDK_GENERATION",
    "MCP_SDK_LABEL",
    "sdk_version",
    "describe",
]


def _load_server_class() -> Tuple[Any, int, str]:
    """Resolve the high-level MCP server class for the installed SDK."""
    try:  # mcp 1.x (the SDK generation this project was written against)
        from mcp.server.fastmcp import FastMCP as _Server  # type: ignore

        return _Server, 1, "mcp.server.fastmcp.FastMCP"
    except ImportError:
        pass

    try:  # mcp 2.x — FastMCP was renamed to MCPServer
        from mcp.server.mcpserver import MCPServer as _Server  # type: ignore

        return _Server, 2, "mcp.server.mcpserver.MCPServer"
    except ImportError as exc:  # pragma: no cover - only on broken installs
        raise ImportError(
            "Could not import a high-level MCP server class. Install the "
            "Python MCP SDK with: pip install 'mcp>=1.9.0,<3.0.0' "
            f"(original error: {exc})"
        ) from exc


FastMCP, MCP_SDK_GENERATION, MCP_SDK_LABEL = _load_server_class()


def sdk_version() -> Optional[str]:
    """Return the installed ``mcp`` distribution version, if discoverable."""
    try:
        from importlib.metadata import PackageNotFoundError, version

        try:
            return version("mcp")
        except PackageNotFoundError:  # pragma: no cover - editable/odd installs
            return None
    except Exception:  # pragma: no cover - importlib.metadata always present 3.8+
        return None


def describe() -> str:
    """One-line description of the active SDK, e.g. ``mcp 1.29.0 (gen 1)``."""
    ver = sdk_version() or "unknown"
    return f"mcp {ver} (gen {MCP_SDK_GENERATION}: {MCP_SDK_LABEL})"
