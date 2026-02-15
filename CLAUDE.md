# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Exchange MCP is a Model Context Protocol server for any Microsoft Exchange / OWA (Outlook Web Access) deployment. It gives LLM agents full access to email, calendar, directory, availability, and meeting analytics over the standard OWA JSON API. Works with any on-premise or hosted Exchange server that exposes OWA.

## Configuration

The server requires two environment variables:

- `EXCHANGE_OWA_URL` — Base URL of the OWA instance (e.g. `https://owa.example.com`)
- `EXCHANGE_COOKIE_FILE` — (optional) Path to session cookies file. Defaults to `session-cookies.txt` next to the package.

## Structure

- `login.py` — Browser-based login via 2FA, saves session cookies
- `exchange_mcp/` — MCP server package (20 tools)
  - `server.py` — FastMCP server with lifespan context (tolerates missing cookies at startup)
  - `owa_client.py` — Centralized OWA HTTP client
  - `auth.py` — Async Playwright login logic, cookie encryption/decryption (reuses crypto from `login.py`)
  - `tools/` — Tool modules: email, calendar, people, folders, availability, analytics, auth

## Running

```bash
export EXCHANGE_OWA_URL=https://owa.example.com

python3 login.py --setup   # One-time credential setup
python3 login.py            # Login (creates session-cookies.txt)
pip install -e .             # Install MCP server
exchange-mcp-server          # Run MCP server (stdio transport)
```

Dependencies: `mcp`, `requests`, `cryptography`, `playwright` (for login).

## Architecture

**Session-based workflow**: `login.py` (CLI) or the `login` MCP tool authenticates via browser-based 2FA and saves encrypted cookies to `session-cookies.txt`. The `login` tool decrypts cookies into memory; the server never reads plaintext cookies from disk. The server starts without cookies; the `login` tool can authenticate within the MCP session.

**OWA JSON API pattern**:
1. Decrypt cookies from `session-cookies.txt` (encrypted with master password)
2. Extract `X-OWA-CANARY` token for CSRF protection
3. POST JSON to `$EXCHANGE_OWA_URL/owa/service.svc?action=<ACTION>`
4. Request bodies use EWS `__type` annotations (e.g. `"CalendarItem:#Exchange"`)
5. HTTP 401/440 = session expired → call the `login` tool or rerun `login.py`

**RequestServerVersion**: `Exchange2013` for reads, `V2017_08_18` for writes.

**Encryption**: PBKDF2-HMAC-SHA256 (480,000 iterations) + AES-256-Fernet for stored credentials and session cookies.
