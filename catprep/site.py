"""seller-admin.cartup.com browser session shared by the export and upload robots.

Login order:
  1. saved session (auth_state.json) - no login needed while it is valid
  2. username/password from credentials.json, then the OTP:
       - read automatically from the email inbox (config [otp] enabled = true)
       - otherwise typed by you in the browser (needs someone at the PC)
"""

import contextlib
import json
import os
import time

from .common import StepError, info, interactive, warn

LOGIN_URL = "https://seller-admin.cartup.com/auth/signin?callbackUrl=%2F"
APP_HOME = "https://seller-admin.cartup.com/"


# ---------------------------------------------------------------------------
# UI helpers (PrimeReact)
# ---------------------------------------------------------------------------
def click(locator, page, desc, total_timeout_ms=4000):
    """Click with retries: normal, forced, then JavaScript click."""
    end = time.time() + total_timeout_ms / 1000.0
    last_err = None
    while time.time() < end:
        try:
            if locator.count() == 0 or not locator.first.is_visible():
                page.wait_for_timeout(200)
                continue
            for attempt in (lambda: locator.first.click(timeout=600),
                            lambda: locator.first.click(timeout=600, force=True),
                            lambda: locator.first.evaluate("el=>el.click()")):
                try:
                    attempt()
                    return True
                except Exception as e:
                    last_err = e
            page.wait_for_timeout(220)
        except Exception as e:
            last_err = e
            page.wait_for_timeout(220)
    warn(f"Could not click {desc}: {last_err}")
    return False


def wait_overlay_clear(page, loops=60, sleep_ms=250):
    overlay = page.locator("div.p-datatable-loading-overlay").first
    for _ in range(loops):
        try:
            if overlay.count() == 0 or not overlay.is_visible():
                return
        except Exception:
            return
        page.wait_for_timeout(sleep_ms)


# ---------------------------------------------------------------------------
# Credentials + login
# ---------------------------------------------------------------------------
def load_credentials(path):
    data = {}
    if os.path.isfile(path):
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            warn(f"Cannot read {os.path.basename(path)}: {e}")
    return {k: str(v).strip() for k, v in data.items() if not k.startswith("_")}


def _is_logged_in(page):
    url = (page.url or "").lower()
    return "seller-admin.cartup.com" in url and "/auth/" not in url


def _confirm_logged_in(page):
    try:
        page.goto(APP_HOME, wait_until="domcontentloaded", timeout=15000)
    except Exception:
        pass
    page.wait_for_timeout(400)
    return _is_logged_in(page)


def _fill_otp(page, code, cfg_otp):
    """Type the OTP into the OTP page: one box, or one box per digit."""
    page.wait_for_timeout(500)
    inputs = page.locator(cfg_otp.get("input_selector") or "input:visible")
    count = inputs.count()
    if count == 0:
        raise StepError("OTP page has no input box (set [otp] input_selector in config.toml).")
    if count >= len(code) > 1:
        for i, digit in enumerate(code):
            inputs.nth(i).fill(digit)
    else:
        inputs.first.fill(code)
    button = page.locator(cfg_otp.get("submit_selector") or
                          "button:has-text('Verify'), button:has-text('Submit'), "
                          "button:has-text('Confirm'), button:has-text('Sign In'), button[type=submit]")
    if button.count():
        click(button, page, "OTP submit button", 4000)
    page.wait_for_timeout(2500)


def login(page, cfg):
    page.goto(APP_HOME, wait_until="domcontentloaded")
    if _is_logged_in(page):
        info("Already logged in (saved session).")
        return
    creds = load_credentials(cfg.path(cfg["export"]["credentials_file"]))
    username, password = creds.get("username", ""), creds.get("password", "")
    if not (username and password) and interactive():
        import getpass
        username = input("seller-admin username: ").strip()
        password = getpass.getpass("seller-admin password (hidden): ").strip()

    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    clicked_at = time.time()
    if username and password:
        page.fill('input[name="username"]', username)
        page.fill('input[name="password"]', password)
        page.click('button:has-text("Sign In")')
        info("Username/password filled and Sign In clicked.")
    elif not interactive():
        raise StepError("No username/password in credentials.json and nobody at the PC to log in.")

    otp_cfg = cfg.get("otp", {})
    if otp_cfg.get("enabled"):
        from .otp_email import wait_for_otp
        try:
            code = wait_for_otp(creds, otp_cfg, since=clicked_at - 30)
            info("OTP received by email - filling it in.")
            _fill_otp(page, code, otp_cfg)
            if _confirm_logged_in(page):
                info("Login confirmed.")
                return
            warn("Logged-in page not reached after filling the email OTP.")
        except StepError as e:
            warn(str(e))
        if not interactive():
            raise StepError("Automatic login failed (email OTP). See the messages above.")

    if not interactive():
        raise StepError("Login needs an OTP but nobody is at the PC. Turn on [otp] email reading in config.toml.")
    for attempt in range(1, 9):
        input(f"\n>>> Type the OTP in the browser, then press ENTER here (try {attempt}/8) ")
        if _confirm_logged_in(page):
            info("Login confirmed.")
            return
        warn("Not logged in yet. Finish the login in the browser, then press Enter again.")
    raise StepError("Login was not confirmed after 8 tries.")


def display_name(page, cfg):
    """Your name as shown in Export History 'Downloaded By'."""
    name = load_credentials(cfg.path(cfg["export"]["credentials_file"])).get("display_name", "")
    if name:
        return name
    for sel in ("button.p-button.p-component >> nth=-1", "[class*='user'] span",
                "[class*='profile'] span", "header button span"):
        try:
            loc = page.locator(sel).last
            if loc.count() and loc.is_visible(timeout=800):
                text = loc.inner_text(timeout=800).strip()
                if len(text) > 2 and text[0].isalpha() and "\n" not in text:
                    return text
        except Exception:
            continue
    warn("Could not read your display name - Export History will not be filtered.")
    return ""


# ---------------------------------------------------------------------------
@contextlib.contextmanager
def session(cfg):
    """with session(cfg) as page: ... - logged-in page; saves the session at the end."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise StepError("Playwright is not installed. Run:\n"
                        "    pip install playwright\n    python -m playwright install chromium")
    ex = cfg["export"]
    auth_state = cfg.path(ex["auth_state_file"])
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=bool(ex["headless"]), slow_mo=int(ex["slowmo_ms"]))
        kwargs = {"accept_downloads": True, "viewport": {"width": 1400, "height": 900}}
        if os.path.isfile(auth_state):
            kwargs["storage_state"] = auth_state
        context = browser.new_context(**kwargs)
        page = context.new_page()
        try:
            login(page, cfg)
            context.storage_state(path=auth_state)
            yield page
        finally:
            try:
                context.storage_state(path=auth_state)
            except Exception:
                pass
            browser.close()
