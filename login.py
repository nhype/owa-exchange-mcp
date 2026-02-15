#!/usr/bin/env python3
"""
Exchange Login Script with Encrypted Credentials
Usage: python3 login.py
Approve login via 2FA mobile app.
"""
import sys
import os
import time
import base64
import getpass
from pathlib import Path
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

CREDS_FILE = Path(__file__).parent / ".credentials.enc"
SALT_FILE = Path(__file__).parent / ".salt"

def get_key(password: str, salt: bytes) -> bytes:
    """Derive encryption key from password"""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=480000,
    )
    return base64.urlsafe_b64encode(kdf.derive(password.encode()))

def encrypt_credentials(username: str, password: str, master_password: str):
    """Encrypt and save credentials"""
    salt = os.urandom(16)
    key = get_key(master_password, salt)
    f = Fernet(key)
    
    data = f"{username}:{password}".encode()
    encrypted = f.encrypt(data)
    
    SALT_FILE.write_bytes(salt)
    CREDS_FILE.write_bytes(encrypted)
    os.chmod(SALT_FILE, 0o600)
    os.chmod(CREDS_FILE, 0o600)
    print("Credentials encrypted and saved.")

def decrypt_credentials(master_password: str) -> tuple:
    """Decrypt and return credentials"""
    if not CREDS_FILE.exists() or not SALT_FILE.exists():
        return None, None
    
    salt = SALT_FILE.read_bytes()
    encrypted = CREDS_FILE.read_bytes()
    key = get_key(master_password, salt)
    f = Fernet(key)
    
    try:
        data = f.decrypt(encrypted).decode()
        username, password = data.split(":", 1)
        return username, password
    except:
        print("ERROR: Invalid master password!")
        return None, None

def setup_credentials():
    """Interactive setup of credentials"""
    print("=== Exchange Mail Setup ===")
    username = input("Enter your email: ").strip()
    password = getpass.getpass("Enter your password: ")
    master = getpass.getpass("Create a master password to encrypt credentials: ")
    master2 = getpass.getpass("Confirm master password: ")
    
    if master != master2:
        print("Passwords don't match!")
        return False
    
    encrypt_credentials(username, password, master)
    return True

def login(username: str, password: str, master_password: str = None):
    """Login to OWA with 2FA (mobile push)"""
    from playwright.sync_api import sync_playwright

    owa_url = os.environ.get("EXCHANGE_OWA_URL", "")
    if not owa_url:
        print("ERROR: EXCHANGE_OWA_URL environment variable is not set.")
        return False
    owa_host = owa_url.replace("https://", "").replace("http://", "").rstrip("/")

    print(f"Logging in as {username}...", flush=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        # Step 1: Navigate to OWA (redirects to SSO)
        page.goto(f"{owa_url}/owa/", wait_until="networkidle")
        
        # Step 2: Fill credentials and submit
        page.fill('input[name="username"]', username)
        page.fill('input[name="password"]', password)
        # Try clicking Continue button (may be in English or Russian)
        try:
            page.locator('button:has-text("Continue"), button:has-text("Продолжить")').first.click(timeout=5000)
        except:
            page.press('input[name="password"]', 'Enter')
        page.wait_for_load_state("networkidle")
        
        print("Credentials submitted, selecting 2FA method...", flush=True)

        # Step 3: Click 2FA authenticator button (first available on the page)
        try:
            auth_btn = page.locator('button').first
            auth_btn.click(timeout=10000)
            page.wait_for_load_state("networkidle")
        except Exception as e:
            print(f"Could not find 2FA button: {e}", flush=True)

        # Step 4: Wait for mobile approval (polls for OWA redirect)
        print("Waiting for mobile approval... Check your 2FA app!", flush=True)
        success = False
        last_url = ""
        for i in range(90):  # Wait up to 90 seconds
            time.sleep(1)
            
            try:
                url = page.url
                
                # Print URL when it changes
                if url != last_url:
                    print(f"  URL changed: {url[:80]}...", flush=True)
                    last_url = url
                
                # Check if we're at OWA (not at SSO pages)
                if owa_host in url and "ofam" not in url and "adfs" not in url:
                    print("  OWA detected! Waiting for page to load...", flush=True)
                    page.wait_for_load_state("load", timeout=15000)
                    success = True
                    break
                
                # Also check if page has OWA elements
                try:
                    if page.locator('[aria-label*="Outlook"], [aria-label*="Почта"]').count() > 0:
                        print("  OWA elements detected!", flush=True)
                        success = True
                        break
                except:
                    pass
                
                if i > 0 and i % 15 == 0:
                    print(f"  Still waiting... ({i}s)", flush=True)
                    
            except Exception as e:
                err_str = str(e).lower()
                if "navigation" in err_str or "destroyed" in err_str or "target closed" in err_str:
                    print(f"  Navigation in progress...", flush=True)
                    try:
                        page.wait_for_load_state("load", timeout=15000)
                        url = page.url
                        if owa_host in url and "ofam" not in url:
                            success = True
                            break
                    except:
                        pass
                else:
                    print(f"  Error: {e}", flush=True)
        
        if success:
            print("\n*** SUCCESS! Logged into OWA! ***", flush=True)

            # Save cookies (encrypted if master password available)
            cookies = context.cookies()
            cookie_file = Path(__file__).parent / "session-cookies.txt"
            cookies_str = "\n".join(
                f"{c['name']}={c['value']}" for c in cookies
            )
            if master_password and SALT_FILE.exists():
                salt = SALT_FILE.read_bytes()
                key = get_key(master_password, salt)
                f = Fernet(key)
                cookie_file.write_bytes(f.encrypt(cookies_str.encode()))
                os.chmod(cookie_file, 0o600)
                print(f"Encrypted cookies saved to {cookie_file}", flush=True)
            else:
                with open(cookie_file, 'w') as f:
                    f.write(cookies_str + "\n")
                print(f"Cookies saved to {cookie_file}", flush=True)
        else:
            print("Login failed - no approval received within 60 seconds", flush=True)
        
        browser.close()
        return success

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--setup":
        setup_credentials()
        return
    
    if not CREDS_FILE.exists():
        print("No credentials found. Run with --setup first.")
        sys.exit(1)
    
    master = getpass.getpass("Master password: ")
    username, password = decrypt_credentials(master)
    if not username:
        sys.exit(1)
    
    login(username, password, master_password=master)

if __name__ == "__main__":
    main()
