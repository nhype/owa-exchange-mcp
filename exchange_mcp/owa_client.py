"""Centralized OWA HTTP client for Exchange API.

Extracts and unifies the duplicated request/cookie/session patterns
from all standalone scripts into a single reusable client.
"""

import os
import requests
from pathlib import Path


class SessionExpiredError(Exception):
    """Raised when the OWA session has expired (HTTP 401/440 or HTML redirect)."""
    pass


# Map common folder names (English + Russian) to OWA distinguished folder IDs
DISTINGUISHED_FOLDERS = {
    "inbox": "inbox",
    "входящие": "inbox",
    "sent": "sentitems",
    "отправленные": "sentitems",
    "drafts": "drafts",
    "черновики": "drafts",
    "deleted": "deleteditems",
    "удаленные": "deleteditems",
    "junk": "junkemail",
    "нежелательная почта": "junkemail",
    "outbox": "outbox",
    "исходящие": "outbox",
    "calendar": "calendar",
    "календарь": "calendar",
}


class OWAClient:
    """HTTP client for OWA (Outlook Web Access) JSON API.

    Handles cookie loading, CSRF token extraction, request construction,
    session expiry detection, and automatic cookie reload on first failure.
    """

    def __init__(
        self,
        cookie_file: str | None = None,
        owa_url: str | None = None,
    ):
        self.cookie_file = Path(
            cookie_file
            or os.environ.get("EXCHANGE_COOKIE_FILE", "")
            or str(Path(__file__).parent.parent / "session-cookies.txt")
        )
        self.owa_url = (
            owa_url
            or os.environ.get("EXCHANGE_OWA_URL", "")
        ).rstrip("/")
        if not self.owa_url:
            raise ValueError(
                "OWA URL not configured. Set the EXCHANGE_OWA_URL environment variable."
            )
        self._cookies: dict[str, str] = {}
        self._canary: str = ""
        self._session = requests.Session()
        self._loaded = False
        self.user_email: str = ""

    # ------------------------------------------------------------------
    # Cookie / session helpers
    # ------------------------------------------------------------------

    def _load_cookies(self) -> None:
        """Read session-cookies.txt (name=value per line) and extract X-OWA-CANARY."""
        if not self.cookie_file.exists():
            raise SessionExpiredError(
                f"Cookie file not found: {self.cookie_file}. Call the login tool first."
            )

        raw = self.cookie_file.read_text().strip()

        # Detect encrypted cookie file (Fernet tokens start with gAAAAA)
        if raw.startswith("gAAAAA"):
            raise SessionExpiredError(
                "Cookie file is encrypted. Call the login tool to decrypt and restore the session."
            )

        cookies: dict[str, str] = {}
        for line in raw.split("\n"):
            if "=" in line:
                name, value = line.split("=", 1)
                cookies[name] = value

        if not cookies:
            raise SessionExpiredError("Cookie file is empty. Call the login tool first.")

        self._cookies = cookies
        self._canary = cookies.get("X-OWA-CANARY", "")
        self._session = requests.Session()
        self._session.cookies.update(cookies)
        self._loaded = True

    def _ensure_loaded(self) -> None:
        """Load cookies on first use."""
        if not self._loaded:
            self._load_cookies()

    def reload_cookies(self) -> None:
        """Force-reload cookies from disk (e.g. after re-login).

        Preserves the existing in-memory session if reloading fails
        (e.g. cookie file is encrypted or missing).
        """
        old_cookies = self._cookies
        old_canary = self._canary
        old_session = self._session
        old_loaded = self._loaded
        try:
            self._loaded = False
            self._load_cookies()
        except SessionExpiredError:
            # Restore previous state to avoid corrupting a valid in-memory session
            self._cookies = old_cookies
            self._canary = old_canary
            self._session = old_session
            self._loaded = old_loaded
            raise

    def load_cookies_from_string(self, cookies_str: str) -> None:
        """Load cookies from a decrypted name=value string (one per line).

        Used by the login tool to inject cookies directly into memory
        without writing plaintext to disk.
        """
        cookies: dict[str, str] = {}
        for line in cookies_str.strip().split("\n"):
            if "=" in line:
                name, value = line.split("=", 1)
                cookies[name] = value

        if not cookies:
            raise SessionExpiredError("No cookies in provided data.")

        self._cookies = cookies
        self._canary = cookies.get("X-OWA-CANARY", "")
        self._session = requests.Session()
        self._session.cookies.update(cookies)
        self._loaded = True

    # ------------------------------------------------------------------
    # Core request method
    # ------------------------------------------------------------------

    def request(self, action: str, payload: dict, *, timeout: int = 30) -> dict:
        """POST to /owa/service.svc?action={action}&EP=1&ID=-1&AC=1.

        On session expiry (401, 440, or text/html response), reloads cookies
        once and retries. If that also fails, raises SessionExpiredError.

        Returns the parsed JSON response dict.
        """
        self._ensure_loaded()

        for attempt in range(2):
            try:
                data = self._do_request(action, payload, timeout=timeout)
                return data
            except SessionExpiredError:
                if attempt == 0:
                    # Try reloading cookies (user may have re-logged in)
                    try:
                        self.reload_cookies()
                    except SessionExpiredError:
                        pass  # Keep in-memory state, retry anyway
                else:
                    raise

        # Should not reach here, but just in case
        raise SessionExpiredError("Session expired. Run login.py to login again.")

    def _do_request(self, action: str, payload: dict, *, timeout: int = 30) -> dict:
        """Execute a single OWA API request (no retry)."""
        url = f"{self.owa_url}/owa/service.svc?action={action}&EP=1&ID=-1&AC=1"

        # Prefer our explicitly-tracked canary (updated on every response)
        # over the session cookie jar, which can hold stale domain-less entries.
        canary = self._canary or self._session.cookies.get("X-OWA-CANARY")

        headers = {
            "Content-Type": "application/json; charset=utf-8",
            "Action": action,
            "X-OWA-CANARY": canary,
            "X-Requested-With": "XMLHttpRequest",
        }

        try:
            resp = self._session.post(url, json=payload, headers=headers, timeout=timeout)
        except requests.exceptions.RequestException as exc:
            raise SessionExpiredError(f"Request failed: {exc}") from exc

        # Keep cached canary in sync if OWA rotated it
        new_canary = resp.cookies.get("X-OWA-CANARY")
        if new_canary:
            self._canary = new_canary
            self._session.cookies.set("X-OWA-CANARY", new_canary)

        # Detect session expiry
        if resp.status_code in (401, 440):
            raise SessionExpiredError("Session expired (HTTP {}).".format(resp.status_code))

        if "text/html" in resp.headers.get("Content-Type", ""):
            # Extract a snippet from HTML for diagnostics (login redirects vs API errors)
            body_snippet = resp.text[:300] if resp.text else ""
            raise SessionExpiredError(
                f"Session expired or invalid action (HTML response, HTTP {resp.status_code}). "
                f"Snippet: {body_snippet}"
            )

        # Parse JSON
        try:
            return resp.json()
        except (ValueError, requests.exceptions.JSONDecodeError) as exc:
            raise SessionExpiredError(
                f"Unexpected response (HTTP {resp.status_code}). "
                "Session may have expired."
            ) from exc

    def request_header_payload(
        self, action: str, payload: dict, *, timeout: int = 30
    ) -> dict:
        """POST with payload in x-owa-urlpostdata header (empty body).

        Some OWA actions (CreateFolder, DeleteFolder, RenameFolder, etc.)
        require the JSON payload to be sent as a URL-encoded string in the
        ``x-owa-urlpostdata`` header instead of the POST body.  This method
        handles that pattern with the same retry logic as ``request()``.
        """
        self._ensure_loaded()

        for attempt in range(2):
            try:
                return self._do_request_header_payload(
                    action, payload, timeout=timeout
                )
            except SessionExpiredError:
                if attempt == 0:
                    try:
                        self.reload_cookies()
                    except SessionExpiredError:
                        pass
                else:
                    raise

        raise SessionExpiredError("Session expired. Run login.py to login again.")

    def _do_request_header_payload(
        self, action: str, payload: dict, *, timeout: int = 30
    ) -> dict:
        """Execute a single OWA API request with payload in header."""
        import json as _json
        from urllib.parse import quote

        url = f"{self.owa_url}/owa/service.svc?action={action}&EP=1&ID=-1&AC=1"
        canary = self._canary or self._session.cookies.get("X-OWA-CANARY")

        url_post_data = quote(_json.dumps(payload, separators=(",", ":")))

        headers = {
            "Content-Type": "application/json; charset=UTF-8",
            "Action": action,
            "X-OWA-CANARY": canary,
            "X-OWA-UrlPostData": url_post_data,
            "X-Requested-With": "XMLHttpRequest",
        }

        try:
            resp = self._session.post(url, headers=headers, timeout=timeout)
        except requests.exceptions.RequestException as exc:
            raise SessionExpiredError(f"Request failed: {exc}") from exc

        new_canary = resp.cookies.get("X-OWA-CANARY")
        if new_canary:
            self._canary = new_canary
            self._session.cookies.set("X-OWA-CANARY", new_canary)

        if resp.status_code in (401, 440):
            raise SessionExpiredError(
                "Session expired (HTTP {}).".format(resp.status_code)
            )

        if "text/html" in resp.headers.get("Content-Type", ""):
            body_snippet = resp.text[:300] if resp.text else ""
            raise SessionExpiredError(
                f"Session expired or invalid action (HTML response, HTTP {resp.status_code}). "
                f"Snippet: {body_snippet}"
            )

        try:
            return resp.json()
        except (ValueError, requests.exceptions.JSONDecodeError) as exc:
            raise SessionExpiredError(
                f"Unexpected response (HTTP {resp.status_code}). "
                "Session may have expired."
            ) from exc

    # ------------------------------------------------------------------
    # File download (attachments)
    # ------------------------------------------------------------------

    def download_file(
        self, attachment_id: str, *, timeout: int = 60
    ) -> tuple[bytes, str, str]:
        """Download a file attachment by its AttachmentId.

        Uses the OWA GetFileAttachment endpoint (direct GET).

        Returns:
            (content_bytes, filename, content_type)
        """
        self._ensure_loaded()

        from urllib.parse import quote

        canary = self._canary or self._session.cookies.get("X-OWA-CANARY")
        url = (
            f"{self.owa_url}/owa/service.svc/s/GetFileAttachment"
            f"?id={quote(attachment_id)}&X-OWA-CANARY={quote(canary)}"
        )

        try:
            resp = self._session.get(url, timeout=timeout)
        except requests.exceptions.RequestException as exc:
            raise SessionExpiredError(f"Download failed: {exc}") from exc

        if resp.status_code in (401, 440):
            raise SessionExpiredError(
                f"Session expired (HTTP {resp.status_code})."
            )

        if "text/html" in resp.headers.get("Content-Type", ""):
            raise SessionExpiredError(
                "Session expired (HTML response on attachment download)."
            )

        # Parse filename from Content-Disposition header
        filename = "attachment"
        cd = resp.headers.get("Content-Disposition", "")
        if cd:
            import re as _re
            from urllib.parse import unquote

            # Try filename*= (RFC 5987) first, then filename=
            match = _re.search(r"filename\*=(?:UTF-8''|utf-8'')(.+?)(?:;|$)", cd)
            if match:
                filename = unquote(match.group(1).strip())
            else:
                match = _re.search(r'filename="?([^";]+)"?', cd)
                if match:
                    filename = unquote(match.group(1).strip())

        content_type = resp.headers.get("Content-Type", "application/octet-stream")

        return resp.content, filename, content_type

    # ------------------------------------------------------------------
    # Convenience: extract response items
    # ------------------------------------------------------------------

    @staticmethod
    def extract_items(data: dict) -> list[dict]:
        """Extract Items from standard OWA response envelope.

        Response shape: data["Body"]["ResponseMessages"]["Items"]
        """
        try:
            return data["Body"]["ResponseMessages"]["Items"]
        except (KeyError, TypeError):
            return []

    # ------------------------------------------------------------------
    # Folder helpers
    # ------------------------------------------------------------------

    def get_folder_id(self, folder_name: str) -> str | None:
        """Resolve a folder name to its Exchange folder ID.

        Supports distinguished folder names (inbox, sentitems, drafts, etc.)
        in both English and Russian, plus custom folder names looked up
        via FindFolder on msgfolderroot.
        """
        folder_lower = folder_name.lower()

        # Check distinguished folders first
        distinguished_id = DISTINGUISHED_FOLDERS.get(folder_lower)
        if distinguished_id:
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
                            "Id": distinguished_id,
                        }
                    ],
                },
            }

            data = self.request("GetFolder", payload)
            for msg in self.extract_items(data):
                if "Folders" in msg:
                    for f in msg["Folders"]:
                        fid = f.get("FolderId", {}).get("Id")
                        if fid:
                            return fid

        # Fall back to searching custom folders by name
        payload = {
            "__type": "FindFolderJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "Exchange2013",
            },
            "Body": {
                "__type": "FindFolderRequest:#Exchange",
                "FolderShape": {
                    "__type": "FolderResponseShape:#Exchange",
                    "BaseShape": "Default",
                },
                "ParentFolderIds": [
                    {
                        "__type": "DistinguishedFolderId:#Exchange",
                        "Id": "msgfolderroot",
                    }
                ],
                "Traversal": "Shallow",
                "Paging": {
                    "__type": "IndexedPageView:#Exchange",
                    "BasePoint": "Beginning",
                    "Offset": 0,
                    "MaxEntriesReturned": 200,
                },
            },
        }

        data = self.request("FindFolder", payload)
        for msg in self.extract_items(data):
            if "RootFolder" in msg and "Folders" in msg["RootFolder"]:
                for f in msg["RootFolder"]["Folders"]:
                    if f.get("DisplayName", "").lower() == folder_lower:
                        return f.get("FolderId", {}).get("Id")

        return None

    # ------------------------------------------------------------------
    # ResolveNames (directory search / attendee resolution)
    # ------------------------------------------------------------------

    def resolve_names(
        self, query: str, *, full_contact: bool = True
    ) -> list[dict]:
        """Call ResolveNames to search the directory.

        Returns the list of Resolution dicts from the API, each containing
        Mailbox and optionally Contact data.
        """
        payload = {
            "__type": "ResolveNamesJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "Exchange2013",
            },
            "Body": {
                "__type": "ResolveNamesRequest:#Exchange",
                "UnresolvedEntry": query,
                "ReturnFullContactData": full_contact,
                "SearchScope": "ActiveDirectoryContacts",
                "ContactDataShape": "AllProperties" if full_contact else "Default",
            },
        }

        data = self.request("ResolveNames", payload)
        for msg in self.extract_items(data):
            if "ResolutionSet" in msg and "Resolutions" in msg["ResolutionSet"]:
                return msg["ResolutionSet"]["Resolutions"]

        return []
