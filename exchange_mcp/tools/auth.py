"""Authentication tool for the Exchange MCP server.

Provides a `login` tool that handles credential setup, browser-based SSO
login, and 2FA — all from within the MCP session.  Session cookies are
encrypted at rest with the master password.

The login flow is non-blocking for 2FA:
1. First call starts browser login in the background and returns immediately
   with instructions to approve 2FA on the mobile app.
2. Second call checks the background task result and loads cookies on success.
"""

import asyncio
import json

from mcp.server.fastmcp import Context

from exchange_mcp.server import mcp, AppContext
from exchange_mcp.owa_client import OWAClient


def _get_app_ctx(ctx: Context) -> AppContext:
    """Extract the AppContext from the MCP lifespan context."""
    return ctx.request_context.lifespan_context


def _get_client(ctx: Context) -> OWAClient:
    """Extract the OWAClient from the MCP lifespan context."""
    return _get_app_ctx(ctx).client


def _session_is_valid(client: OWAClient) -> bool:
    """Quick check: can we reach the inbox?"""
    try:
        client._ensure_loaded()
        payload = {
            "__type": "GetFolderJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "Exchange2013",
            },
            "Body": {
                "__type": "GetFolderRequest:#Exchange",
                "FolderShape": {
                    "__type": "FolderResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                },
                "FolderIds": [
                    {
                        "__type": "DistinguishedFolderId:#Exchange",
                        "Id": "inbox",
                    }
                ],
            },
        }
        data = client.request("GetFolder", payload)
        items = client.extract_items(data)
        return bool(items)
    except Exception:
        return False


@mcp.tool()
async def login(
    master_password: str,
    username: str = "",
    password: str = "",
    ctx: Context = None,
) -> str:
    """Authenticate to Exchange OWA (handles credential setup and 2FA login).

    Call this tool when the session has expired or before first use.
    It performs browser-based SSO login with 2FA mobile push approval.
    Session cookies are encrypted at rest with the master password.

    **Two-call 2FA flow**: The first call starts the browser login in the
    background and returns immediately asking you to tell the user to approve
    2FA on their phone.  Call login again with the same master_password after
    the user approves — the second call picks up the result.

    Args:
        master_password: Decrypts stored credentials (and cookies), or encrypts
            new ones if username/password are also provided.
        username: Email address. Provide together with password for first-time
            credential setup (replaces `login.py --setup`).
        password: Account password. Required together with username for setup.

    Returns:
        JSON result with success status and any error details.
    """
    app_ctx = _get_app_ctx(ctx)
    client = app_ctx.client

    # ------------------------------------------------------------------
    # If a background login task exists, check its status first
    # ------------------------------------------------------------------
    if app_ctx.pending_login is not None:
        task = app_ctx.pending_login

        if not task.done():
            return json.dumps({
                "status": "awaiting_2fa",
                "message": "Still waiting for 2FA approval. Please approve the login in your authenticator app, then call login again.",
            })

        # Task finished — harvest result and clear
        app_ctx.pending_login = None
        try:
            result = task.result()
        except Exception as e:
            return json.dumps({"success": False, "error": f"Background login failed: {e}"})

        if result.get("success"):
            from exchange_mcp.auth import encrypt_cookie_file

            cookies_str = result.pop("cookies")
            try:
                encrypt_cookie_file(cookies_str, master_password, client.cookie_file)
                client.load_cookies_from_string(cookies_str)
                if _session_is_valid(client):
                    return json.dumps({"success": True, "message": "Logged in and session verified. Cookies encrypted."})
                else:
                    return json.dumps({"success": True, "message": "Cookies saved but session verification failed. Try again."})
            except Exception as e:
                return json.dumps({"success": True, "message": f"Login succeeded but cookie handling failed: {e}"})
        else:
            return json.dumps(result)

    # ------------------------------------------------------------------
    # No pending task — normal login flow
    # ------------------------------------------------------------------

    # 1. Check if already authenticated (cookies already in memory)
    if _session_is_valid(client):
        return json.dumps({"success": True, "message": "Session is already active."})

    # 2. Verify playwright is available
    try:
        import playwright  # noqa: F401
    except ImportError:
        return json.dumps({
            "success": False,
            "error": "playwright is not installed. Run: pip install playwright && playwright install chromium",
        })

    # 3. Resolve credentials
    from exchange_mcp.auth import (
        encrypt_credentials,
        decrypt_credentials,
        CREDS_FILE,
        decrypt_cookie_file,
        encrypt_cookie_file,
        perform_login,
    )

    if username and password:
        # First-time setup: encrypt and save credentials
        encrypt_credentials(username, password, master_password)
    else:
        # Decrypt existing credentials
        if not CREDS_FILE.exists():
            return json.dumps({
                "success": False,
                "error": "No stored credentials found. Provide username and password for first-time setup.",
            })
        username, password = decrypt_credentials(master_password)
        if not username:
            return json.dumps({
                "success": False,
                "error": "Invalid master password — could not decrypt credentials.",
            })

    # Store user email on the client for availability queries
    client.user_email = username

    # 4. Try restoring session from encrypted cookies on disk
    cookies_str = decrypt_cookie_file(master_password, client.cookie_file)
    if cookies_str:
        try:
            client.load_cookies_from_string(cookies_str)
            if _session_is_valid(client):
                return json.dumps({
                    "success": True,
                    "message": "Session restored from encrypted cookies.",
                })
        except Exception:
            pass

    # 5. Start browser login in background (non-blocking for 2FA)
    app_ctx.pending_login = asyncio.create_task(
        perform_login(
            username=username,
            password=password,
            owa_url=client.owa_url,
        )
    )

    return json.dumps({
        "status": "awaiting_2fa",
        "message": "Please approve the login in your 2FA authenticator app, then call login again with the same master_password.",
    })
