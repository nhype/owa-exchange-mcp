# OWA Exchange MCP Server

MCP (Model Context Protocol) server for any Microsoft Exchange / OWA (Outlook Web Access) deployment. Gives LLM agents access to email, calendar, directory search, folders, availability, and meeting analytics via 30 tools.

Works with any on-premise or hosted Exchange server that exposes OWA.

## Quick Start

```bash
# Copy and edit the MCP config with your OWA URL
cp .mcp.json.example .mcp.json

# One-time: set up encrypted credentials
python3 login.py --setup

# Login (opens headless browser, 2FA approval required)
python3 login.py

# Install the MCP server
pip install -e .
```

## Configuration

| Variable | Required | Description |
|---|---|---|
| `EXCHANGE_OWA_URL` | Yes | Base URL of your OWA instance |
| `EXCHANGE_COOKIE_FILE` | No | Path to session cookies file (default: `session-cookies.txt`) |

## Login

### Option A: Via MCP tool (recommended)

The `login` tool handles credential setup and authentication within the MCP session — no separate terminal needed.

First time (setup + login):
```
login(master_password="...", username="user@example.com", password="...")
```

Subsequent logins (decrypts stored credentials):
```
login(master_password="...")
```

### Option B: Via CLI

```bash
python3 login.py --setup   # First time: save encrypted credentials
python3 login.py            # Login with 2FA
```

Both methods:
1. Open a headless browser to your OWA URL
2. Submit credentials
3. Wait for 2FA approval (up to 90 seconds)
4. Save encrypted session cookies to `session-cookies.txt`

Credentials and session cookies are encrypted at rest with AES-256 (PBKDF2 key derivation, 480k iterations).

## Tools (30)

### Email (10)
| Tool | Description |
|---|---|
| `get_emails` | List emails from a folder with filtering |
| `get_email` | Get full email content by ID |
| `send_email` | Send a new email |
| `reply_email` | Reply to an email |
| `forward_email` | Forward an email |
| `delete_email` | Delete an email |
| `move_email` | Move email to another folder |
| `mark_email_read` | Mark email as read/unread |
| `download_attachments` | Download file attachments from an email |
| `get_email_links` | Extract hyperlinks from an email body |

### Calendar (7)
| Tool | Description |
|---|---|
| `get_calendar_events` | Get events in a date range (supports recurring expansion) |
| `create_meeting` | Create a meeting with attendees |
| `update_meeting` | Update an existing meeting |
| `cancel_meeting` | Cancel a meeting and notify attendees |
| `respond_to_meeting` | Accept, decline, or tentatively accept |
| `download_event_attachments` | Download file attachments from a calendar event |
| `get_event_links` | Extract hyperlinks from an event description |

### Directory (1)
| Tool | Description |
|---|---|
| `find_person` | Search people in Active Directory |

### Folders (7)
| Tool | Description |
|---|---|
| `get_folders` | List mail folders with unread counts |
| `create_folder` | Create a new mail folder |
| `rename_folder` | Rename an existing folder |
| `empty_folder` | Empty all items from a folder |
| `delete_folder` | Delete a mail folder |
| `move_folder` | Move a folder to a different parent |
| `check_session` | Check if the OWA session is authenticated |

### Availability (2)
| Tool | Description |
|---|---|
| `find_free_time` | Find free slots in your calendar |
| `find_meeting_time` | Find common free slots for multiple people |

### Analytics (2)
| Tool | Description |
|---|---|
| `get_meeting_stats` | Meeting count statistics for multiple people |
| `get_meeting_contacts` | Connection matrix — who you meet with most |

### Auth (1)
| Tool | Description |
|---|---|
| `login` | Authenticate to OWA (credential setup + 2FA login) |

## Files

```
login.py                  # Browser-based 2FA login (standalone CLI)
exchange_mcp/
  server.py               # FastMCP server entry point
  owa_client.py           # OWA HTTP client
  auth.py                 # Async login logic (shared by MCP tool)
  tools/
    email.py              # Email tools
    calendar.py           # Calendar tools
    people.py             # Directory search
    folders.py            # Folder management & session check
    availability.py       # Free time / meeting time
    analytics.py          # Meeting stats & contacts
    auth.py               # Login tool
pyproject.toml            # Package config
```

## Security

- Credentials and session cookies encrypted with AES-256-Fernet
- Master password never stored
- PBKDF2 with 480,000 iterations for key derivation
- Credential and cookie files have `0600` permissions
- Cookies decrypted into memory only — never written to disk as plaintext (via MCP tool)
- Session cookies never transmitted except to your OWA server
