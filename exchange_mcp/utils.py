"""Shared utility functions for the Exchange MCP server.

Extracts duplicated helpers from the standalone scripts:
html_to_text, date/time formatting and parsing.
"""

import html
import re
from datetime import datetime


def html_to_text(html_content: str) -> str:
    """Convert HTML to plain text.

    Strips scripts, styles, converts <br>/<p>/<div> to newlines,
    removes remaining tags, and unescapes HTML entities.
    """
    if not html_content:
        return ""
    text = re.sub(
        r"<script[^>]*>.*?</script>", "", html_content, flags=re.DOTALL | re.IGNORECASE
    )
    text = re.sub(
        r"<style[^>]*>.*?</style>", "", text, flags=re.DOTALL | re.IGNORECASE
    )
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<p[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<div[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"\n\s*\n", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def extract_links_from_html(html_content: str) -> list[dict]:
    """Extract hyperlinks from HTML content.

    Finds <a href="...">text</a> patterns, excludes mailto:, cid:,
    javascript:, and fragment-only (#) links. Deduplicates by URL.

    Returns list of {url, text} dicts.
    """
    if not html_content:
        return []

    # Match <a ...href="URL"...>text</a>
    pattern = re.compile(
        r'<a\s[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
        re.DOTALL | re.IGNORECASE,
    )

    seen: set[str] = set()
    links: list[dict] = []

    for url_raw, text_raw in pattern.findall(html_content):
        url = html.unescape(url_raw).strip()

        # Skip non-http links
        if url.startswith(("mailto:", "cid:", "javascript:")) or url == "#":
            continue
        # Skip fragment-only links
        if url.startswith("#"):
            continue

        if url in seen:
            continue
        seen.add(url)

        # Clean link text: strip tags and whitespace
        text = re.sub(r"<[^>]+>", "", text_raw)
        text = html.unescape(text).strip()

        links.append({"url": url, "text": text})

    return links


def format_datetime(dt_str: str) -> str:
    """Format an ISO datetime string as 'YYYY-MM-DD HH:MM'.

    Strips timezone suffixes (Z, +offset) for cleaner display.
    """
    if not dt_str:
        return ""
    if "T" in dt_str:
        date_part, time_part = dt_str.split("T", 1)
        time_part = time_part.split("Z")[0].split("+")[0]
        return f"{date_part} {time_part[:5]}"
    return dt_str


def format_date(dt_str: str) -> str:
    """Extract the date portion from an ISO datetime string."""
    if not dt_str:
        return ""
    if "T" in dt_str:
        return dt_str.split("T")[0]
    return dt_str


def parse_date(date_str: str) -> datetime:
    """Parse a date string in common formats.

    Supports: YYYY-MM-DD, DD.MM.YYYY, DD/MM/YYYY, MM/DD/YYYY.
    """
    formats = ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Could not parse date: {date_str}")


def parse_iso_datetime(dt_str: str) -> datetime:
    """Parse an ISO datetime string to a naive datetime.

    Handles both 'YYYY-MM-DDTHH:MM:SS' and 'YYYY-MM-DD' formats,
    stripping any timezone suffix.
    """
    if "T" in dt_str:
        clean = dt_str.split("Z")[0].split("+")[0]
        return datetime.strptime(clean, "%Y-%m-%dT%H:%M:%S")
    return datetime.strptime(dt_str, "%Y-%m-%d")


def format_attendee(name: str, email: str) -> str:
    """Format an attendee as 'Name <email>' or just the email."""
    if name and email and not email.startswith("/O="):
        return f"{name} <{email}>"
    return name or email or ""
