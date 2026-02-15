"""Exchange MCP Server.

Exposes OWA email, calendar, directory, and availability tools via MCP.
Uses FastMCP with a lifespan context manager to share a single OWAClient
instance across all tool invocations.
"""

import asyncio
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass, field

from mcp.server.fastmcp import FastMCP

from exchange_mcp.owa_client import OWAClient


@dataclass
class AppContext:
    """Shared application state available to all tools via lifespan context."""
    client: OWAClient
    pending_login: asyncio.Task | None = field(default=None, repr=False)


@asynccontextmanager
async def app_lifespan(server: FastMCP) -> AsyncIterator[AppContext]:
    """Initialize the OWA client on startup, available for the server lifetime."""
    client = OWAClient()
    # Try to pre-load cookies; tolerate missing file so the server can start
    # without an active session (the `login` tool will authenticate first).
    try:
        client._ensure_loaded()
    except Exception:
        pass
    try:
        yield AppContext(client=client)
    finally:
        # No persistent connections to clean up (requests is sync + per-call)
        pass


# Create the MCP server instance
mcp = FastMCP("exchange", lifespan=app_lifespan)

# ------------------------------------------------------------------
# Import tool modules so their @mcp.tool() decorators register tools.
# Each module imports `mcp` from this file and decorates its functions.
# ------------------------------------------------------------------
import exchange_mcp.tools.email      # noqa: E402, F401
import exchange_mcp.tools.calendar   # noqa: E402, F401
import exchange_mcp.tools.people     # noqa: E402, F401
import exchange_mcp.tools.folders    # noqa: E402, F401
import exchange_mcp.tools.availability  # noqa: E402, F401
import exchange_mcp.tools.analytics     # noqa: E402, F401
import exchange_mcp.tools.auth          # noqa: E402, F401


def main():
    """Entry point: run the MCP server over stdio."""
    mcp.run()


if __name__ == "__main__":
    main()
