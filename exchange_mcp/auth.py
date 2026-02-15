"""Shared login logic for Exchange MCP.

Reuses crypto functions from login.py and provides an async Playwright-based
login function for use by the MCP login tool.
"""

import asyncio
import os
import sys
from pathlib import Path

from cryptography.fernet import Fernet

# Ensure login.py (project root) is importable regardless of cwd
_project_root = str(Path(__file__).resolve().parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from login import CREDS_FILE, SALT_FILE, get_key, encrypt_credentials, decrypt_credentials


# ------------------------------------------------------------------
# Cookie encryption (same master password / salt as credentials)
# ------------------------------------------------------------------

def encrypt_cookie_file(cookies_str: str, master_password: str, cookie_file: Path) -> None:
    """Encrypt session cookies and write to disk."""
    if not SALT_FILE.exists():
        raise FileNotFoundError("Salt file not found. Set up credentials first.")
    salt = SALT_FILE.read_bytes()
    key = get_key(master_password, salt)
    f = Fernet(key)
    cookie_file.write_bytes(f.encrypt(cookies_str.encode()))
    os.chmod(cookie_file, 0o600)


def decrypt_cookie_file(master_password: str, cookie_file: Path) -> str | None:
    """Decrypt session cookies from disk. Returns None on failure."""
    if not cookie_file.exists() or not SALT_FILE.exists():
        return None
    salt = SALT_FILE.read_bytes()
    key = get_key(master_password, salt)
    f = Fernet(key)
    try:
        return f.decrypt(cookie_file.read_bytes()).decode()
    except Exception:
        return None


# ------------------------------------------------------------------
# Async browser login
# ------------------------------------------------------------------

async def perform_login(
    username: str,
    password: str,
    owa_url: str,
    progress_callback=None,
) -> dict:
    """Authenticate to OWA via browser-based SSO + 2FA.

    Mirrors login.py's login() but uses playwright.async_api so it can
    run inside the MCP server's event loop.

    Args:
        username: Email address.
        password: Account password.
        owa_url: Base OWA URL (e.g. https://owa.example.com).
        progress_callback: Optional async callable for status updates.

    Returns:
        {"success": True, "cookies": "name=val\\n..."} on success,
        {"success": False, "error": "..."} on failure.
    """
    from playwright.async_api import async_playwright

    owa_host = owa_url.replace("https://", "").replace("http://", "").rstrip("/")

    async def _info(msg: str):
        if progress_callback:
            await progress_callback(msg)

    await _info(f"Logging in as {username}...")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        try:
            # Step 1: Navigate to OWA (redirects to SSO)
            await page.goto(f"{owa_url}/owa/", wait_until="networkidle")

            # Step 2: Fill credentials and submit
            await page.fill('input[name="username"]', username)
            await page.fill('input[name="password"]', password)
            try:
                btn = page.locator(
                    'button:has-text("Continue"), button:has-text("Продолжить")'
                ).first
                await btn.click(timeout=5000)
            except Exception:
                await page.press('input[name="password"]', "Enter")
            await page.wait_for_load_state("networkidle")

            await _info("Credentials submitted, selecting 2FA method...")

            # Step 3: Click 2FA authenticator button (skip nav/menu buttons)
            try:
                # Look for a heading that indicates the 2FA choice screen
                await page.wait_for_selector(
                    'h1:has-text("Choose"), h1:has-text("Выберите")',
                    timeout=10000,
                )
                # Pick the first button inside the choices container (after the heading),
                # skipping nav/menu chrome buttons.
                auth_btn = page.locator("h1 ~ div button").first
                await auth_btn.click(timeout=5000)
                await page.wait_for_load_state("networkidle")
            except Exception as e:
                await _info(f"Could not find 2FA button: {e}")

            # Step 4: Wait for mobile approval
            await _info("Waiting for mobile approval... Check your 2FA app!")
            success = False
            last_url = ""

            for i in range(90):
                await asyncio.sleep(1)

                try:
                    url = page.url

                    if url != last_url:
                        await _info(f"URL changed: {url[:80]}...")
                        last_url = url

                    if owa_host in url and "ofam" not in url and "adfs" not in url:
                        await _info("OWA detected! Waiting for page to load...")
                        await page.wait_for_load_state("load", timeout=15000)
                        success = True
                        break

                    try:
                        if (
                            await page.locator(
                                '[aria-label*="Outlook"], [aria-label*="Почта"]'
                            ).count()
                            > 0
                        ):
                            await _info("OWA elements detected!")
                            success = True
                            break
                    except Exception:
                        pass

                    if i > 0 and i % 15 == 0:
                        await _info(f"Still waiting... ({i}s)")

                except Exception as e:
                    err_str = str(e).lower()
                    if any(
                        kw in err_str
                        for kw in ("navigation", "destroyed", "target closed")
                    ):
                        await _info("Navigation in progress...")
                        try:
                            await page.wait_for_load_state("load", timeout=15000)
                            url = page.url
                            if owa_host in url and "ofam" not in url:
                                success = True
                                break
                        except Exception:
                            pass
                    else:
                        await _info(f"Error: {e}")

            if success:
                cookies = await context.cookies()
                cookies_str = "\n".join(
                    f"{c['name']}={c['value']}" for c in cookies
                )
                await _info("Login successful.")
                return {"success": True, "cookies": cookies_str}
            else:
                return {
                    "success": False,
                    "error": "2FA approval not received within 90 seconds.",
                }

        except Exception as e:
            return {"success": False, "error": str(e)}
        finally:
            await browser.close()
